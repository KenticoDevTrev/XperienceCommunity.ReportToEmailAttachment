using CMS;
using CMS.DataEngine;
using CMS.Modules;
using NuGet.Packaging.Core;
using NuGet.Packaging;
using NuGet.Versioning;
using System.Collections.Generic;
using XperienceCommunity.ReportToEmailAttachment;

[assembly: RegisterModule(typeof(NuGetLoadModule))]
namespace XperienceCommunity.ReportToEmailAttachment
{
    public class NuGetLoadModule : Module
    {
        public NuGetLoadModule() : base("XperienceCommunity.ReportToEmailAttachment.NuGetLoadModule")
        {


        }

        protected override void OnInit()
        {
            base.OnInit();

            ModulePackagingEvents.Instance.BuildNuSpecManifest.After += BuildNuSpecManifest_After;
        }

        private void BuildNuSpecManifest_After(object sender, BuildNuSpecManifestEventArgs e)
        {
            if (e.ResourceName.Equals("XperienceCommunity.ReportToEmailAttachment", System.StringComparison.InvariantCultureIgnoreCase))
            {
                // Change the name
                e.Manifest.Metadata.Title = "XperienceCommunity.ReportToEmailAttachment.KX13.Admin";
                e.Manifest.Metadata.SetProjectUrl("https://github.com/KenticoDevTrev/XperienceCommunity.ReportToEmailAttachment");
                e.Manifest.Metadata.SetIconUrl("https://www.hbs.net/HBS/media/Favicon/favicon-96x96.png");
                e.Manifest.Metadata.Tags = "Kentico Xperience 13 Report Email";
                e.Manifest.Metadata.Id = "XperienceCommunity.ReportToEmailAttachment.KX13.Admin";
                e.Manifest.Metadata.ReleaseNotes = "Initial Release";
                e.Manifest.Metadata.Description = "For Kentico Xperience 13 Admin Only.\r\n\r\nProvides a replacement scheduled task assembly to be able to attach report subscription tables as a csv instead of inline.\r\n\r\nEdit the Scheduled Task \"Report subscription sender\", set the Task Provider Assembly to XperienceCommunity.ReportToEmailAttachment, the class XperienceCommunity.ReportToEmailAttachment.ReportSubscriptionSenderWithAttachment\r\n\r\nOptionally, you can include Report Table Code Names in the TaskData that will be excluded from attachments and will still render in the email.\r\n\r\nGraphs and Values will render in email as they are not export-compatible.";
                // Add nuget dependencies

                // Add dependencies
                List<PackageDependency> NetDependencies = new List<PackageDependency>()
                {
                    new PackageDependency("Kentico.Xperience.Libraries", new VersionRange(new NuGetVersion("13.0.13")), new string[] { }, new string[] {"Build","Analyzers"}),
                };
                PackageDependencyGroup PackageGroup = new PackageDependencyGroup(new NuGet.Frameworks.NuGetFramework("net48"), NetDependencies);
                e.Manifest.Metadata.DependencyGroups = new PackageDependencyGroup[] { PackageGroup };
            }
        }
    }
}
