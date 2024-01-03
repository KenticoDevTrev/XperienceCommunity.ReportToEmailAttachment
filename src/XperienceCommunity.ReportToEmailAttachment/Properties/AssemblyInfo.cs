using CMS;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// General Information about an assembly is controlled through the following
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("XperienceCommunity.ReportToEmailAttachment")]
[assembly: AssemblyDescription("For Kentico Xperience 13 Admin Only.\r\n\r\nProvides a replacement scheduled task assembly to be able to attach report subscription tables as a csv instead of inline.\r\n\r\nEdit the Scheduled Task \"Report subscription sender\", set the Task Provider Assembly to XperienceCommunity.ReportToEmailAttachment, the class XperienceCommunity.ReportToEmailAttachment.ReportSubscriptionSenderWithAttachment\r\n\r\nOptionally, you can include Report Table Code Names in the TaskData that will be excluded from attachments and will still render in the email.\r\n\r\nGraphs and Values will render in email as they are not export-compatible.")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("Trevor Fayas (KenticoDevTrev)")]
[assembly: AssemblyProduct("XperienceCommunity.ReportToEmailAttachment")]
[assembly: AssemblyCopyright("Copyright ©  2024")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible
// to COM components.  If you need to access a type in this assembly from
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("2fecb243-e73d-4558-8eb2-2f953aa5085b")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers
// by using the '*' as shown below:
// [assembly: AssemblyVersion("1.0.*")]
[assembly: AssemblyVersion("13.13.0.0")]
[assembly: AssemblyFileVersion("13.13.0.0")]

[assembly: AssemblyDiscoverable]