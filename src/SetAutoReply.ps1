$objCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchcas01.domain.local/PowerShell/ -Authentication Kerberos -Credential $objCredential
Import-PSSession $Session

# Enter Username
$Username = "rero"

# Fill out message
$Message = @"
Thank you for your email.

Unfortunately, I'm on sick leave, and therefore not able to process and/or reply to Your inquiry.
If the inquiry can be processed by one of my colleagues, please direct your correspondence at my brand manager, Jeanette Thistrup: jeth@dkcompany.com


Your email will not be forwarded.

"@

Set-MailboxAutoReplyConfiguration -Identity "rero@dkcompany.com" -AutoReplyState Enabled -ExternalMessage $Message -ExternalAudience all -InternalMessage $Message

Remove-PSSession $Session