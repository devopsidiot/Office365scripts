#===============================================================================================
# AUTHOR:	Brad Wright, Dan Shepard
# DATE:    	05-24-2018
# Version  	v1.2
# COMMENT: 	Used with MFA Office 365
#===============================================================================================

#This script will pull Office 365 reports for forwarding and delegates rules, list of admins, and inactive accounts.
#It will combine these reports into 1 excel spreadsheet and send it in an email to support@bluesprucecapital.com
#To run the report, simply run this script from an admin powershell command line
#The credentials are in LastPass under bscc-reports
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))

{
$arguments = "& '" + $myinvocation.mycommand.definition + "'"
Start-Process powershell -Verb runAs -ArgumentList $arguments
Break
}
Clear-Host
Write-Output "Blue Sprue Capital - Weekly report for Office 365"
Write-Output "Please login to Office 365 with the Global Admin account"

Remove-Item C:\Temp\*.*

#Connect to Office365
$credential = Get-Credential
Install-Module MsOnline
Connect-MsolService -Credential $credential
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking

Clear-Host

############ Delegates & Forwarding Rules ###############

#Find all users within Office 365
$allUsers = @()
$AllUsers = Get-MsolUser -All -EnabledFilter EnabledOnly | Where-Object {($_.UserPrincipalName -notlike "*#EXT#*")}

#Create arrays
$UserInboxRules = @()
$UserDelegates = @()

#Foreach for all users
foreach ($User in $allUsers)
{
    Write-Host "Checking inbox rules and delegates for user: " $User.UserPrincipalName;
	$UserInboxRules += Get-InboxRule -Mailbox $User.UserPrincipalname  |  Where-Object {($_.ForwardTo -ne $null) -or ($_.ForwardAsAttachmentTo -ne $null) -or ($_.RedirectsTo -ne $null)}
	$UserDelegates += Get-MailboxPermission -Identity $User.UserPrincipalName  | Where-Object {($_.IsInherited -ne "True") -and ($_.User -notlike "*SELF*")}
}

#export findings to a file
$UserInboxRules | Select Enabled, Identity, Name, From, SentTo, CopyToFolder, DeleteMessage, ForwardTo, MailboxOwnerID | Export-Csv 'C:\Temp\FwdRulesToExternalDomains.csv' -NoType
$UserDelegates | Select User, Identity, AccessRights | Export-Csv 'C:\Temp\MailboxDelegatePermissions.csv' -NoType

################## List Admin's ########################

#Which role we are looking for
$role = Get-MsolRole -RoleName "Company Administrator"

#export findings to a file
Get-MsolRoleMember -RoleObjectId $role.ObjectId | Select DisplayName, Emailaddress, IsLicensed, LastDirSyncTime, OverallProvisioningStatus, ValidationStatus |Export-CSV 'C:\Temp\list-admins.csv' -NoType

################ Inactive Accounts #####################

#Finding Inactive User accounts
Get-Mailbox -ResultSize Unlimited | Get-MailboxStatistics | where {$_.LastLogonTime -lt ((get-date).AddDays(-90))} | Select displayname, lastlogontime | Export-csv 'C:\Temp\InactiveUsers.csv' -NoType

Clear-Host

#Combine CSV files into 1 excel spreadsheet with multiple tabs
Function Merge-CSVFiles
{
Param(
$CSVPath = "C:\Temp", ## Soruce CSV Folder
$XLOutput="c:\Temp\temp.xlsx" ## Output file name
)

$csvFiles = Get-ChildItem ("$CSVPath\*") -Include *.csv
$Excel = New-Object -ComObject excel.application
$Excel.visible = $false
$Excel.sheetsInNewWorkbook = $csvFiles.Count
$workbooks = $excel.Workbooks.Add()
$CSVSheet = 1

Foreach ($CSV in $Csvfiles)

{
$worksheets = $workbooks.worksheets
$CSVFullPath = $CSV.FullName
$SheetName = ($CSV.name -split "\.")[0]
$worksheet = $worksheets.Item($CSVSheet)
$worksheet.Name = $SheetName
$TxtConnector = ("TEXT;" + $CSVFullPath)
$CellRef = $worksheet.Range("A1")
$Connector = $worksheet.QueryTables.add($TxtConnector,$CellRef)
$worksheet.QueryTables.item($Connector.name).TextFileCommaDelimiter = $True
$worksheet.QueryTables.item($Connector.name).TextFileParseType  = 1
$worksheet.QueryTables.item($Connector.name).Refresh()
$worksheet.QueryTables.item($Connector.name).delete()
$worksheet.UsedRange.EntireColumn.AutoFit()
$CSVSheet++

}

$workbooks.SaveAs($XLOutput,51)
$workbooks.Saved = $true
$workbooks.Close()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbooks) | Out-Null
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

}

#Put date stamp on generated report
Merge-CSVFiles -CSVPath 'C:\Temp' -XLOutput "C:\Temp\WeeklyReport_$(get-date -f MM-dd-yyyy).xlsx"

#Send email to support to generate ticket
Send-MailMessage -To 'helpdesk@somecompany.com' -From 'some.admin@somecompany.com' -Subject "Weekly Office 365 Report" -Body 'Weekly Reports'-Credential $credential -Attachments "C:\Temp\WeeklyReport_$(get-date -f MM-dd-yyyy).xlsx" -UseSsl -SmtpServer 'smtp.office365.com'

Remove-Item C:\Temp\*.*
