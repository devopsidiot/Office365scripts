$credential = Get-Credential
Install-Module MsOnline
Connect-MsolService -Credential $credential
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking

#adding create to things that are audited on the mailbox - using author's email address as example
Set-Mailbox -Identity #emailofperson# -AuditOwner @{Add='Create'}

#pull mailbox audit logs
$contacts = Search-MailboxAuditLog -Identity #emailofperson# -LogonTypes Admin,Owner,Delegate -ResultSize 2000 -ShowDetails -Operations Create

#sort them into something not terrible
$createdBy = $contacts | Select-Object LogonUserDisplayName,ItemSubject,FolderPathName

#This returns JUST the Contact name and who created it
$contactslist = $createdBy | Where-Object {$_.FolderPathName -eq "\Contacts"}

#Export to CSV
$contactslist | Select @{n='Contact Created By';e={$_.LogonUserDisplayName}}, @{n='Contact Created';e={$_.ItemSubject}} | Export-Csv "C:\Scripts\Contacts $(get-date -f yyyy-MM-dd).csv"

# Considering putting this data into an S3 bucket w/ Python and parsing w/ Lambda
