# XperienceCommunity.ReportToEmailAttachment
Allows reporting tables from Kentico's "Reporting" module subscriptions to be included as attachments instead of HTML tables in the email.

In Kentico Xperience 13, the Reporting Module allows you to create reports (Graphs, Tables, or Values) and subscribe to them.  Often, the report tables can be very large in size.

By default, Kentico presents the tables in these report emails as an inline HTML table.  This is not ideal for large data sets.

This module attaches as a CSV file any tables that were subscribed to in reports, instead of making them inline HTML tables.  It renders Graphs and Values as normal.

## Installation & Requirements
Install the nuget package XperienceCommunity.ReportToEmailAttachment.KX13.Admin nuget package on the Administrator (WebApp) solution.

Must be on a Kentico Xperience 13 solution, hotfix minimum 13.0.13

Once installed, rebuild and load the admin, then:

1. Go to the `Scheduled Tasks` interface in Kentico admin
2. Edit the scheduled task `Report subscription sender`
3. Set the Task Provider Assembly to `XperienceCommunity.ReportToEmailAttachment` and the class to `XperienceCommunity.ReportToEmailAttachment.ReportSubscriptionSenderWithAttachment`
4. [Optional] Add Subscription Table Code Names to the TaskData (space or new line separated) that you wish to still render inline on reports.
5. Save

## Report Subscription Limitation
It seems that this will only work if the Reporting Subscription `Subscription Item` is set to a reporting table, not "(whole report)".

## Excluding Tables
If there are any tables from reports you wish to keep inline to emails, simply modify the `Scheduled tasks` - `Report subscription sender`'s TaskData and include (new line or space separated) the Report table code names (visible when you edit or add a new Report Table in the Reporting module).

# Contributions
If you find a bug, please feel free to submit a pull request!
