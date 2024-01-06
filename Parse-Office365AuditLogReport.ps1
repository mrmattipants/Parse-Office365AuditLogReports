Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *; $error.Clear();

$ExcelModule = Get-Module ImportExcel -ListAvailable -ErrorAction SilentlyContinue
If (!$ExcelModule) {
    Install-Module ImportExcel -Force
    Import-Module ImportExcel -Force
} Else {
    Import-Module ImportExcel -Force
}

$Path = "$($PSScriptRoot)"
$PathInfo=[System.Uri]$Path

if($PathInfo.IsUnc){
    Push-Location "$($PathInfo.OriginalString)"
    $CurrentDirectory = Get-Location
} else {
    Push-Location "$($PSScriptRoot)"
    $CurrentDirectory = Get-Location
}

function Open-CsvFile([string]$initialDirectory) {

    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog

    $OpenFileDialog.InitialDirectory = "$InitialDirectory"
    $OpenFileDialog.Title = "Open O365 Audit Log Report CSV File"
    $OpenFileDialog.Filter = "CSV File|*.csv"        
    $OpenFileDialog.Multiselect=$false
    $Result = $OpenFileDialog.ShowDialog()

    If($Result -eq 'OK') {

        Try {
    
            $InPath = $OpenFileDialog.FileNames
        }

        Catch {

            $InPath = $null
            Break
        }

        $CSVData = Import-Csv -Path $InPath -Header CreationDate,UserIds,Operations,AuditData | Select-Object -Skip 1

        Return $CSVData
    }

}


function Save-ExcelFile([string] $initialDirectory){

    $TodaysDate = Get-Date -Format "MM-dd-yyyy"

    Add-Type -AssemblyName System.Windows.Forms
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.initialDirectory = $initialDirectory
    $SaveFileDialog.Title = "Save Parsed AuditData From O365 Audit Log Report As Excel Spreadsheet"
    $SaveFileDialog.filter = "Excel Workbook|*.xlsx"
    $SaveFileDialog.FileName = "Microsoft_Online_Audit_Log_$($TodaysDate).xlsx"
    $SaveResult = $SaveFileDialog.ShowDialog()

    If($SaveResult -eq 'OK') {

        Try {
    
            $OutPath = $SaveFileDialog.FileNames
        }

        Catch {

            $OutPath = $null
            Break
        }

        Return $OutPath
    }

}

$Inbox = "$($CurrentDirectory)\Inbox"
$Outbox = "$($CurrentDirectory)\Outbox"
$UserProfile = "$($ENV:USERPROFILE)"
$Desktop = "$($ENV:USERPROFILE)\Desktop"
$Documents = "$($ENV:USERPROFILE)\Documents"
$Downloads = "$($ENV:USERPROFILE)\Downloads"
$Pictures = "$($ENV:USERPROFILE)\Pictures"
$Music = "$($ENV:USERPROFILE)\Music"
$Video = "$($ENV:USERPROFILE)\Videos"
$Local = "$($ENV:USERPROFILE)\AppData\Local"
$Roaming = "$($ENV:USERPROFILE)\AppData\Roaming"
$Favorites = "$($ENV:USERPROFILE)\Favorites"
$History = "$($ENV:USERPROFILE)\AppData\Local\Microsoft\Windows\History"
$NetHood = "$($ENV:USERPROFILE)\AppData\Roaming\Microsoft\Windows\Network Shortcuts"
$PrintHood = "$($ENV:USERPROFILE)\AppData\Roaming\Microsoft\Windows\Printer Shortcuts"
$Recent = "$($ENV:USERPROFILE)\AppData\Roaming\Microsoft\Windows\Recent"
$SendTo = "$($ENV:USERPROFILE)\AppData\Roaming\Microsoft\Windows\SendTo"
$Cache = "$($ENV:USERPROFILE)\AppData\Local\Microsoft\Windows\INetCache"
$Cookies = "$($ENV:USERPROFILE)\AppData\Local\Microsoft\Windows\INetCookies"
$StartMenu = "$($ENV:USERPROFILE)\AppData\Roaming\Microsoft\Windows\Start Menu"
$Programs = "$($ENV:USERPROFILE)\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
$Startup = "$($ENV:USERPROFILE)\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
$Templates = "$($ENV:USERPROFILE)\AppData\Roaming\Microsoft\Windows\Templates"

$DefaultOpenFolder = "$($Inbox)"
$DefaultSaveFolder = "$($Outbox)"

$FileData = Open-CsvFile -initialDirectory "$($DefaultOpenFolder)"

$JsonData = @()

Foreach ($Item in $FileData){

    $JsonData += $Item.AuditData | ConvertFrom-Json

}

$SaveFile = Save-ExcelFile -initialDirectory "$($DefaultSaveFolder)"

$JsonData | Export-Excel "$($SaveFile)"

Start-Process "$($ENV:windir)\explorer.exe" -ArgumentList "$($DefaultSaveFolder)"
