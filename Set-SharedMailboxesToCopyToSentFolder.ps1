#modules
Import-Module ExchangeOnlineManagement

#variables
#SMTP Server for email
$smtpserver = "<SMTP Relay Server>" 
#From Address for email
$fromaddress = "NoReply@example.com" 
#To address for email
$toaddress = "YourEmail@example.com" 
#Email subject
$Subject = "Shared mailboxes set to MessageCopyForSentAsEnabled $($date)" 
#Date used for various functions
$date = Get-Date -Format "MM-dd-yyyy"
#File name for report saved modified mailboxes
$fileName = "SharedMailboxesCopytoSentFolder$($date).csv"
#Path to save report
$path = ".\Reports\"
#Full file name used to save report
$FullFileName = "$($path)$($fileName)"
#Number of days to keep reports in negative days - default '-7'
$ReportTimetokeep = "-7"

#Connection variables
#App ID used for connection
$AppID = "<AppID>"
#Certificate thumbprint used for connection
$CertificateThumbprint = "<CertificateThumbprint>"
#On microsoft domain of M365 tenant
$OrgDomain = "<example>.onmicrosoft.com"

#Connect to Exchange Online PowerShell using certificate auth
Connect-ExchangeOnline -AppId $AppID -CertificateThumbprint $CertificateThumbprint -Organization $OrgDomain

#Retrieve list of shared mailboxes that need to be modified
$SharedMailboxes = Get-mailbox -Filter {RecipientTypeDetails -eq 'SharedMailbox'} | Where-Object {$_.MessageCopyForSentAsEnabled -eq $false -or $_.MessageCopyForSendOnBehalfEnabled -eq $false}

#Create array for the report
$report = New-Object System.Collections.Generic.List[System.Object]

#Set to MessageCopyForSentAsEnabled and MessageCopyForSendOnBehalfEnabled to True
$SharedMailboxes | ForEach-Object {
    Set-Mailbox $_.UserPrincipalName -MessageCopyForSentAsEnabled $true -MessageCopyForSendOnBehalfEnabled $true
    $report.add([pscustomobject]@{Mailbox=$_.UserPrincipalName})
}

#If there is mailboxes in the report, send email
If ($report.count -ge '1') {
    $report | export-csv $FullFileName -NoTypeInformation
    $body = ""
    $attachment = $FullFileName 
    
    $message = new-object System.Net.Mail.MailMessage 
    $message.From = $fromaddress
    $message.To.Add($toaddress)
    $message.IsBodyHtml = $True
    $message.Subject = $Subject
    $attach = new-object Net.Mail.Attachment($attachment)
    $message.Attachments.Add($attach)
    $message.body = $body
    $smtp = new-object Net.Mail.SmtpClient($smtpserver)
    $smtp.Send($message)
}

# Delete all reports older than specified days
$DatetoDelete = $date.AddDays($ReportTimetokeep)
Get-ChildItem $Path | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item