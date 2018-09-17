$exchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential (Get-Credential) -Authentication Basic -AllowRedirection
Import-PSSession $exchSession

cls

Write-Host "
MsExchange - Powershell - DKC - maif@dkcompany.com

  1. Delegate Mailbox Permissions

"

switch(Read-Host "Choose option")
{
    1 { delegate }
}

function delegate
{
    $del_mbox = Read-Host "Enter mailbox"
    $del_user = Read-Host "Enter user"
    Add-MailboxPermission -Identity $del_mbox -AccessRights Reviewer -User $del_user
}

function viewCalendar
{
    
}