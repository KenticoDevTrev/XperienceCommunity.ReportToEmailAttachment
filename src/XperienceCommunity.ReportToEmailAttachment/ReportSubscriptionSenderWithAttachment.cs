using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Hosting;
using System.Web.UI;
using CMS.Base;
using CMS.Base.Web.UI;
using CMS.Core;
using CMS.DataEngine;
using CMS.EmailEngine;
using CMS.FormEngine;
using CMS.Helpers;
using CMS.ImportExport;
using CMS.MacroEngine;
using CMS.Reporting;
using CMS.Reporting.Web.UI;
using CMS.Scheduler;
using CMS.SiteProvider;
using CMS.WebAnalytics;

namespace XperienceCommunity.ReportToEmailAttachment
{
    public class ReportSubscriptionWithAttachmentPage : Page
    {
        //
        // Summary:
        //     Enables render outside form tag
        //
        // Parameters:
        //   control:
        //     Control
        public override void VerifyRenderingInServerForm(Control control)
        {
        }
    }

    //
    // Summary:
    //     Class for subscription email sender
    public class ReportSubscriptionSenderWithAttachment : ITask
    {
        //
        // Summary:
        //     Indicates for how long will be subscription email css file cached (one week)
        public const int EMAIL_CSS_FILE_CACHE_MINUTES = 10080;

        private static string mPath = string.Empty;

        private static ReportSubscriptionWithAttachmentPage mReportPage;

        private static Regex mInlineImageRegex;

        //
        // Summary:
        //     Regex object for search inline images in report.
        private static Regex InlineImageRegex => mInlineImageRegex ?? (mInlineImageRegex = RegexHelper.GetRegex("<img\\s*src=\"\\S*fileguid=(\\S*)\" />"));

        //
        // Summary:
        //     Page object used for render without HTTP context
        private static ReportSubscriptionWithAttachmentPage ReportPage => mReportPage ?? (mReportPage = new ReportSubscriptionWithAttachmentPage());

        //
        // Summary:
        //     Path for default CSS styles document
        public static string Path
        {
            get
            {
                if (string.IsNullOrEmpty(mPath))
                {
                    mPath = HostingEnvironment.MapPath("~/CMSModules/Reporting/CMSPages/ReportSubscription.css");
                }

                return mPath;
            }
            set
            {
                mPath = value;
            }
        }

        //
        // Summary:
        //     Executes the logprocessor action.
        //
        // Parameters:
        //   task:
        //     Task to process
        public string Execute(TaskInfo task)
        {
            try
            {
                var excludedTableNames = DataHelper.GetNotEmpty(task.TaskData, "").ToLower().Split("\n\r\t,|;\t ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                SendSubscriptions(excludedTableNames);
            }
            catch (Exception ex)
            {
                Service.Resolve<IEventLogService>().LogException("Reporting", "ReportSubscriptionSender", ex);
                return ex.Message;
            }

            return null;
        }

        private void SendSubscriptions(string[] excludedTableNames)
        {
            InfoDataSet<ReportSubscriptionInfo> typedResult = ReportSubscriptionInfoProvider.GetSubscriptions().WhereEquals("ReportSubscriptionEnabled", 1).WhereLessThan("ReportSubscriptionNextPostDate", DateTime.Now)
                .TypedResult;
            string result = string.Empty;
            if (!typedResult.Any())
            {
                return;
            }

            try
            {
                using (CachedSection<string> cachedSection = new CachedSection<string>(ref result, 10080.0, true, null, "reportsubscriptioncssfile", Path))
                {
                    if (cachedSection.LoadData)
                    {
                        using (CMS.IO.StreamReader streamReader = CMS.IO.StreamReader.New(Path))
                        {
                            result = streamReader.ReadToEnd();
                        }

                        if (cachedSection.Cached)
                        {
                            cachedSection.CacheDependency = CacheHelper.GetFileCacheDependency(Path);
                        }

                        cachedSection.Data = result;
                    }
                }
            }
            catch (Exception ex)
            {
                Service.Resolve<IEventLogService>().LogException("Report subscription sender", "EXCEPTION", ex);
            }

            foreach (ReportSubscriptionInfo item in typedResult)
            {
                string text = string.Empty;
                if (MacroProcessor.ContainsMacro(item.ReportSubscriptionCondition) && !ValidationHelper.GetBoolean(MacroResolver.Resolve(item.ReportSubscriptionCondition), defaultValue: false))
                {
                    continue;
                }

                ReportInfo reportInfo = ReportInfoProvider.GetReportInfo(item.ReportSubscriptionReportID);
                if (reportInfo == null || !reportInfo.ReportEnableSubscription)
                {
                    continue;
                }

                EmailTemplateInfo emailTemplateInfo = AbstractInfo<EmailTemplateInfo, IEmailTemplateInfoProvider>.Provider.Get("Report_subscription_template", item.ReportSubscriptionSiteID);
                if (emailTemplateInfo == null)
                {
                    Service.Resolve<IEventLogService>().LogError("Administration", "SENDEMAIL", string.Format(ResHelper.GetString("subscription.emailtemplatenotfound"), "Report_subscription_template"));
                    break;
                }

                FormInfo formInfo = new FormInfo(reportInfo.ReportParameters);
                DataRow dataRow = formInfo.GetDataRow(convertMacroColumn: false);
                formInfo.LoadDefaultValues(dataRow, enableMacros: true);
                ReportHelper.ApplySubscriptionParameters(item, dataRow, resolveMacros: true);
                bool flag = false;
                if (item.ReportSubscriptionGraphID != 0)
                {
                    flag = true;
                    ReportGraphInfo reportGraphInfo = ReportGraphInfoProvider.GetReportGraphInfo(item.ReportSubscriptionGraphID);
                    if (reportGraphInfo != null && ValidationHelper.GetBoolean(reportGraphInfo.GraphSettings["SubscriptionEnabled"], defaultValue: true))
                    {
                        AbstractReportControl ctrl = ((!reportGraphInfo.GraphIsHtml) ? (ReportPage.LoadUserControl("~/CMSModules/Reporting/Controls/ReportGraph.ascx") as AbstractReportControl) : (ReportPage.LoadUserControl("~/CMSModules/Reporting/Controls/HtmlBarGraph.ascx") as AbstractReportControl));
                        text = RenderControlToString(ctrl, reportInfo.ReportName + "." + reportGraphInfo.GraphName, dataRow, item);
                    }
                }

                // CUSTOM: Use attachments vs. inline rendering for tables
                var attachments = new List<EmailAttachment>();
                if (item.ReportSubscriptionTableID != 0)
                {
                    flag = true;
                    ReportTableInfo reportTableInfo = ReportTableInfoProvider.GetReportTableInfo(item.ReportSubscriptionTableID);
                    if (reportTableInfo != null && ValidationHelper.GetBoolean(reportTableInfo.TableSettings["SubscriptionEnabled"], defaultValue: true))
                    {
                        if (excludedTableNames.Contains(reportTableInfo.TableName.ToLower())) {
                            AbstractReportControl ctrl2 = ReportPage.LoadUserControl("~/CMSModules/Reporting/Controls/ReportTable.ascx") as AbstractReportControl;
                            text = RenderControlToString(ctrl2, reportInfo.ReportName + "." + reportTableInfo.TableName, dataRow, item);
                        }
                        else
                        {
                            // Generate the query params
                            var queryParams = new QueryDataParameters();
                            foreach (var column in dataRow.Table.Columns.Cast<DataColumn>().Select(x => x.ColumnName))
                            {
                                queryParams.Add(column, dataRow[column]);
                            }
                            var tableData = ReportTableInfoProvider.GetTableData(reportTableInfo, queryParams);
                            if (tableData != null)
                            {
                                var dataExportHelper = new DataExportHelper(tableData);
                                for (int ti = 0; ti < tableData.Tables.Count; ti++)
                                {
                                    var memStream = new MemoryStream();
                                    dataExportHelper.ExportToCSV(tableData, ti, memStream);
                                    attachments.Add(CreateAttachment($"{reportTableInfo.TableName}{(tableData.Tables.Count > 1 ? "-" + (ti + 1) : "")}.csv", "text/plain", isInline: false, Guid.Empty, memStream));
                                }
                                text = "Report Attached";
                            }
                        }
                    }
                }

                if (item.ReportSubscriptionValueID != 0)
                {
                    flag = true;
                    ReportValueInfo reportValueInfo = ReportValueInfoProvider.GetReportValueInfo(item.ReportSubscriptionValueID);
                    if (reportValueInfo != null && ValidationHelper.GetBoolean(reportValueInfo.ValueSettings["SubscriptionEnabled"], defaultValue: true))
                    {
                        AbstractReportControl ctrl3 = ReportPage.LoadUserControl("~/CMSModules/Reporting/Controls/ReportValue.ascx") as AbstractReportControl;
                        text = RenderControlToString(ctrl3, reportInfo.ReportName + "." + reportValueInfo.ValueName, dataRow, item);
                    }
                }

                if (!flag)
                {
                    IDisplayReport displayReport = ReportPage.LoadUserControl("~/CMSModules/Reporting/Controls/DisplayReport.ascx") as IDisplayReport;
                    displayReport.ReportName = reportInfo.ReportName;
                    displayReport.EmailMode = true;
                    displayReport.ReportParameters = dataRow;
                    displayReport.LoadFormParameters = false;
                    displayReport.RenderCssClasses = true;
                    displayReport.SendOnlyNonEmptyDataSource = item.ReportSubscriptionOnlyNonEmpty;
                    displayReport.ReportSubscriptionSiteID = item.ReportSubscriptionSiteID;
                    ((AbstractReportControl)displayReport).SubscriptionInfo = item;
                    string defaultDynamicMacros = item.ReportSubscriptionSettings["reportinterval"] as string;
                    displayReport.SetDefaultDynamicMacros(defaultDynamicMacros);
                    try
                    {
                        text = displayReport.RenderToString(item.ReportSubscriptionSiteID);
                    }
                    catch (Exception ex2)
                    {
                        Service.Resolve<IEventLogService>().LogException("Reporting", "ReportSubscriptionSender", ex2, 0, "Report name: " + reportInfo.ReportName);
                    }
                }

                if (!(text != string.Empty))
                {
                    continue;
                }

                SiteInfo siteInfo = AbstractInfo<SiteInfo, ISiteInfoProvider>.Provider.Get(item.ReportSubscriptionSiteID);
                EmailFormatEnum emailFormat = EmailHelper.GetEmailFormat(siteInfo.SiteName);
                EmailMessage emailMessage = new EmailMessage();
                emailMessage.EmailFormat = (flag ? EmailFormatEnum.Default : EmailFormatEnum.Html);
                emailMessage.From = (string.IsNullOrEmpty(emailTemplateInfo.TemplateFrom) ? EmailHelper.Settings.NotificationsSenderAddress(siteInfo.SiteName) : emailTemplateInfo.TemplateFrom);
                emailMessage.Recipients = item.ReportSubscriptionEmail;
                emailMessage.Subject = item.ReportSubscriptionSubject;
                emailMessage.BccRecipients = emailTemplateInfo.TemplateBcc;
                emailMessage.CcRecipients = emailTemplateInfo.TemplateCc;
                MacroResolver macroResolver = CreateSubscriptionMacroResolver(reportInfo, item, siteInfo, emailMessage.Recipients, result, text);
                string text2 = macroResolver.ResolveMacros(emailTemplateInfo.TemplateText);
                string text3 = macroResolver.ResolveMacros(emailTemplateInfo.TemplatePlainText);

                // CUSTOM: Add table attachments
                foreach (var attachment in attachments)
                {
                    emailMessage.Attachments.Add(attachment);
                }

                if (AbstractStockHelper<RequestStockHelper>.GetItem(reportInfo.ReportName) is Dictionary<string, byte[]> dictionary)
                {
                    foreach (string key in dictionary.Keys)
                    {
                        string text4 = key.Substring(1);
                        byte[] buffer = dictionary[key];
                        if (key.StartsWith("t", StringComparison.OrdinalIgnoreCase))
                        {
                            emailMessage.Attachments.Add(CreateAttachment(text4 + ".csv", "text/plain", isInline: false, Guid.Empty, new MemoryStream(buffer)));
                        }
                        else if (key.StartsWith("g", StringComparison.OrdinalIgnoreCase))
                        {
                            Guid guid = Guid.NewGuid();
                            bool flag2 = emailFormat != EmailFormatEnum.PlainText;
                            string text5;
                            if (!flag2)
                            {
                                text5 = string.Format(ResHelper.GetString("reportsubscription.attachment_img"), text4);
                            }
                            else
                            {
                                Guid guid2 = guid;
                                text5 = "<img src=\"cid:" + guid2.ToString() + "\" alt=\"inlineimage\">";
                            }

                            string newValue = text5;
                            text2 = text2.Replace("##InlineImage##" + text4 + "##InlineImage##", newValue);
                            text3 = text3.Replace("##InlineImage##" + text4 + "##InlineImage##", newValue);
                            MemoryStream ms = new MemoryStream(buffer);
                            if (emailFormat == EmailFormatEnum.Both)
                            {
                                emailMessage.Attachments.Add(CreateAttachment(text4 + ".png", "image/png", isInline: false, guid, ms));
                                emailMessage.Attachments.Add(CreateAttachment(text4 + ".png", "image/png", isInline: true, guid, ms));
                            }
                            else
                            {
                                emailMessage.Attachments.Add(CreateAttachment(text4 + ".png", "image/png", flag2, guid, ms));
                            }
                        }
                    }

                    AbstractStockHelper<RequestStockHelper>.Remove(reportInfo.ReportName);
                }

                emailMessage.Body = URLHelper.MakeLinksAbsolute(text2);
                emailMessage.PlainTextBody = URLHelper.MakeLinksAbsolute(text3);
                EmailHelper.ResolveMetaFileImages(emailMessage, emailTemplateInfo.TemplateID, "cms.emailtemplate", "Template");
                EmailSender.SendEmail(siteInfo.SiteName, emailMessage);
                item.ReportSubscriptionNextPostDate = SchedulingHelper.GetNextTime(SchedulingHelper.DecodeInterval(item.ReportSubscriptionInterval), item.ReportSubscriptionNextPostDate);
                item.ReportSubscriptionLastPostDate = DateTime.Now;
                using (CMSActionContext cMSActionContext = new CMSActionContext())
                {
                    cMSActionContext.LogEvents = false;
                    cMSActionContext.TouchParent = false;
                    ReportSubscriptionInfoProvider.SetReportSubscriptionInfo(item);
                }
            }
        }

        //
        // Summary:
        //     Creates attachment for email
        //
        // Parameters:
        //   name:
        //     Name of the attachment
        //
        //   mediaType:
        //     Type of the attachment (img, text)
        //
        //   isInline:
        //     Indicates whether this attachment is added as inline
        //
        //   guid:
        //     Guid for inline attachment (is ignored for non inline attachments)
        //
        //   ms:
        //     Memory stream with attachment data
        private EmailAttachment CreateAttachment(string name, string mediaType, bool isInline, Guid guid, MemoryStream ms)
        {
            EmailAttachment emailAttachment = new EmailAttachment(ms, name);
            if (isInline)
            {
                emailAttachment.ContentDisposition.Inline = true;
                emailAttachment.ContentDisposition.DispositionType = "inline";
                emailAttachment.ContentId = guid.ToString();
            }

            emailAttachment.ContentType.Name = name;
            emailAttachment.ContentType.MediaType = mediaType;
            emailAttachment.TransferEncoding = EmailHelper.TransferEncoding;
            return emailAttachment;
        }

        //
        // Summary:
        //     Renders control to string
        //
        // Parameters:
        //   ctrl:
        //     Control to render
        //
        //   itemName:
        //     Repor item name
        //
        //   data:
        //     Datarow with report parameters
        //
        //   rsi:
        //     Subscription object
        private string RenderControlToString(AbstractReportControl ctrl, string itemName, DataRow data, ReportSubscriptionInfo rsi)
        {
            try
            {
                ctrl.EmailMode = true;
                ctrl.Parameter = itemName;
                ctrl.RenderCssClasses = true;
                ctrl.ReportParameters = data;
                ctrl.ReportSubscriptionSiteID = rsi.ReportSubscriptionSiteID;
                ctrl.SendOnlyNonEmptyDataSource = rsi.ReportSubscriptionOnlyNonEmpty;
                ctrl.SubscriptionInfo = rsi;
                string value = rsi.ReportSubscriptionSettings["reportinterval"] as string;
                ctrl.SetDefaultDynamicMacros(HitsIntervalEnumFunctions.StringToHitsConversion(value));
                ctrl.AllParameters.Add("CMSContextCurrentSiteID", rsi.ReportSubscriptionSiteID);
                ctrl.ReloadData(forceLoad: true);
                StringBuilder stringBuilder = new StringBuilder();
                Html32TextWriter writer = new Html32TextWriter(new CMS.IO.StringWriter(stringBuilder));
                ctrl.RenderControl(writer);
                return stringBuilder.ToString();
            }
            catch (Exception ex)
            {
                Service.Resolve<IEventLogService>().LogException("Reporting", "ReportSubscriptionSender", ex, 0, "Report item name:" + itemName);
                return string.Empty;
            }
        }

        //
        // Summary:
        //     Creates macro resolver for report subscription
        //
        // Parameters:
        //   ri:
        //     Report info
        //
        //   rsi:
        //     Report subscription info
        //
        //   si:
        //     Site info
        //
        //   rec:
        //     Recipient email
        //
        //   defaultCss:
        //     Report default CSS
        //
        //   content:
        //     Report content
        public static MacroResolver CreateSubscriptionMacroResolver(ReportInfo ri, ReportSubscriptionInfo rsi, SiteInfo si, string rec, string defaultCss = null, string content = null)
        {
            MacroResolver instance = MacroResolver.GetInstance();
            instance.SetNamedSourceData("Report", ri);
            instance.SetNamedSourceData("ReportSubscription", rsi);
            Dictionary<string, object> dictionary = new Dictionary<string, object> {
        {
            "CurrentEmail",
            HttpUtility.UrlEncode(rec)
        } };
            if (si != null)
            {
                string siteDomain = GetSiteDomain(si);
                dictionary.Add("SiteDomain", siteDomain);
                if (rsi != null)
                {
                    dictionary.Add("UnsubscriptionLink", $"{siteDomain}/CMSModules/Reporting/CMSPages/Unsubscribe.aspx?guid={rsi.ReportSubscriptionGUID}&email={HttpUtility.UrlEncode(rec)}");
                }
            }

            dictionary.Add("ItemName", GetItemName(rsi));
            dictionary.Add("defaultsubscriptioncss", defaultCss);
            dictionary.Add("subscriptionbody", content);
            instance.SetNamedSourceData(dictionary, isPrioritized: false);
            return instance;
        }

        private static string GetSiteDomain(SiteInfo siteInfo)
        {
            string text = siteInfo.DomainName.TrimEnd('/');
            if (HttpContext.Current != null && !text.Contains("/"))
            {
                string applicationPath = SystemContext.ApplicationPath;
                if (!string.IsNullOrEmpty(applicationPath))
                {
                    text = text + "/" + applicationPath.Trim('/');
                }
            }

            return RequestContext.CurrentScheme + "://" + text;
        }

        //
        // Summary:
        //     Adds data to collection in HTTP context items
        //
        // Parameters:
        //   key:
        //     Key in http context
        //
        //   item:
        //     Key in item's collections(graph,table,..)
        //
        //   data:
        //     Data to store
        public static void AddToRequest(string key, string item, byte[] data)
        {
            if (AbstractStockHelper<RequestStockHelper>.Contains(key))
            {
                Dictionary<string, byte[]> dictionary = AbstractStockHelper<RequestStockHelper>.GetItem(key) as Dictionary<string, byte[]>;
                dictionary?.Add(item, data);
                AbstractStockHelper<RequestStockHelper>.Add(key, dictionary);
            }
            else
            {
                Dictionary<string, byte[]> dictionary2 = new Dictionary<string, byte[]>();
                dictionary2.Add(item, data);
                AbstractStockHelper<RequestStockHelper>.Add(key, dictionary2);
            }
        }

        //
        // Summary:
        //     Creates type and name string identification of subscription item.
        //
        // Parameters:
        //   rsi:
        //     Report subscription object
        public static string GetItemName(ReportSubscriptionInfo rsi)
        {
            string text = string.Empty;
            string arg = string.Empty;
            if (rsi.ReportSubscriptionGraphID != 0)
            {
                ReportGraphInfo reportGraphInfo = ReportGraphInfoProvider.GetReportGraphInfo(rsi.ReportSubscriptionGraphID);
                if (reportGraphInfo != null)
                {
                    text = reportGraphInfo.GraphDisplayName;
                    arg = ResHelper.GetString("ReportItemType.graph").ToLowerInvariant();
                }
            }

            if (rsi.ReportSubscriptionTableID != 0)
            {
                ReportTableInfo reportTableInfo = ReportTableInfoProvider.GetReportTableInfo(rsi.ReportSubscriptionTableID);
                if (reportTableInfo != null)
                {
                    text = reportTableInfo.TableDisplayName;
                    arg = ResHelper.GetString("ReportItemType.table").ToLowerInvariant();
                }
            }

            if (rsi.ReportSubscriptionValueID != 0)
            {
                ReportValueInfo reportValueInfo = ReportValueInfoProvider.GetReportValueInfo(rsi.ReportSubscriptionValueID);
                if (reportValueInfo != null)
                {
                    text = reportValueInfo.ValueDisplayName;
                    arg = ResHelper.GetString("ReportItemType.value").ToLowerInvariant();
                }
            }

            if (text != string.Empty)
            {
                return string.Format(ResHelper.GetString("reportsubscription.itemnameformat"), arg, text);
            }

            return string.Empty;
        }

        //
        // Summary:
        //     Resolves metafiles in the given report HTML. Returns HTML with resolved metafiles
        //
        //
        // Parameters:
        //   reportName:
        //     Report name
        //
        //   html:
        //     Report HTML
        public static string ResolveMetaFiles(string reportName, string html)
        {
            foreach (Match item in InlineImageRegex.Matches(html))
            {
                Guid guid = ValidationHelper.GetGuid(item.Groups[1], Guid.Empty);
                if (guid != Guid.Empty)
                {
                    MetaFileInfo metaFileInfo = MetaFileInfoProvider.GetMetaFileInfo(guid, null, globalOrLocal: true);
                    if (metaFileInfo != null)
                    {
                        Guid guid2 = guid;
                        AddToRequest(reportName, "g" + guid2.ToString(), metaFileInfo.MetaFileBinary);
                        string text = html;
                        string value = item.Value;
                        guid2 = guid;
                        html = text.Replace(value, "##InlineImage##" + guid2.ToString() + "##InlineImage##");
                    }
                }
            }

            return html;
        }
    }
}