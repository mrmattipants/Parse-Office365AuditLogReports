# Parse-Office365AuditLogReports
PowerShell Script to Parse Office 365 Audit Log Reports, in .CSV Format (specifically the "AuditData" Column) and Export a New Report, in Excel .XLSX Format

**Parse-Office365AuditLogReport.vbs**<br>
- The VBS Script Silently Runs the .BAT Script, without any Command Line Windows, etc.

**Parse-Office365AuditLogReport.bat**<br>
- The .BAT Script Runs the .PS1 (PowerShell) Script, with the necessary ExecutionPolicy Parameter, etc.

**Parse-Office365AuditLogReport.ps1**<br>
- The .PS1 Script performs the following Actions.
  - Launches the Open Dialog, allowing the User to Manually Locate and Open the .CSV File, containing the RAW Office 365 Audit Log Data.
  - Imports and Parses the "AuditData" Column, consisting of JSON Keys/Values, from the aforementioned .CSV File.
  - Launches the Save Dialog, allowing the User to Locate the Directory, to which the Report, containing the Parsed Log Data, will be Saved.
  - Exports the Parsed Log Data, in Excel .XLSX Format, to the Chosen Save Location.
 
**The "Inbox" DIrectory:**<br>
- This is the Default "**Open**" Location, to which the "**Open**" Dialog will be Opened.
- To Modify this Setting, simply Update the "**$DefaultOpenFolder**" Variable.

**The "Outbox" Directory:**<br>
- This is the Defauly "**Save**" Location, to which the "**Save**" Dialog will be Opened.
- To Modify this Setting, simply Update the "**$DefaultSaveFolder**" Variable.

**NOTE:** Please refer to the List of Default Location Variables, directly above the "**$DefaultOpenFolder**" and "**$DefaultSaveFolder**" Variables, in the .PS1 (PowerShell) Script.
