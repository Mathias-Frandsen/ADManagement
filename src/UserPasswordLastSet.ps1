cls
$strUsername = Read-Host "Username"

$objUser = Get-ADUser $strUsername -properties PasswordLastSet,Name,samAccountName
Write-Host "`n`n`n`n`n`n"
Write-Host "$($objUser.Name) ($($objUser.SamAccountName)) last set their password on: $($objUser.PasswordLastSet)`n`n`n`n`n`n"

pause
cls