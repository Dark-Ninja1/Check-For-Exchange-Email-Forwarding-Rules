

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Confirm:$false
 
Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -scope CurrentUser

Connect-Exchangeonline

Get-ExoMailbox -ResultSize Unlimited | Select-Object -ExpandProperty UserPrincipalName | Foreach-Object {Get-InboxRule -Mailbox $_ | Select-Object -Property MailboxOwnerID,Name,Enabled,From,Description,RedirectTo,ForwardTo} >> C:\Temp\InboxForwardingRules.txt

Get-Mailbox -ResultSize Unlimited -Filter "ForwardingAddress -like '*' -or ForwardingSmtpAddress -like '*'" | Select-Object Name,ForwardingAddress,ForwardingSmtpAddress >> C:\InboxForwardingRules.txt
