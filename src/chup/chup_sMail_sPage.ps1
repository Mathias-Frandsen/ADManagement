$imp = Import-CSV C:\IT\PS\opdtprint_uptd.csv -Delimiter ";" -Encoding Default
$imp[0] | FT -AutoSize
switch (Read-Host "Er ovenstående data importeret korrekt? (j/n)")
{
    {("j","ja",1,"y","yes") -contains $_} { Write-Host "Fortsætter..."; $dataOK = $true }
    {("n","nej",0,"no") -contains $_} { Write-Host "Afbryder..."; $dataOK = $false ; break }
}

if ($dataOK)
{
    foreach ($user in $imp)
    {
        Add-Content -Path C:\IT\PS\opdtPrint.log -Value "User: '$($User.Name)'"
        $htmlBody = @"
Hej.
<br>
<br>
Vi er i gang med at rette data i vores brugerdatabase, således printerkode matcher medarbejdernummer.<br><br>
<strong>I den forbindelse, bliver din printkode ændret.</strong> <br>
Fremover vil din printkode være:<strong><p style="color:red">$($user.Number)</p></strong>
Tak for forståelsen.
<br>
<br>
Mvh.<br>
<strong>IT-Afdelingen</strong>
</p>
"@
        
        Send-MailMessage -BodyAsHtml $htmlBody -From "Helpdesk <helpdesk@dkcompany.com>" -To $user.Email -SmtpServer mail.dkcompany.com -Priority High -Subject "Din printkode ændres" -Encoding Default
        Add-Content -Path C:\IT\PS\opdtPrint.log -Value "Sent Mail. ($($User.Email))"

        Set-ADUser -HomePage $user.Nummer -Identity (Get-ADUser -Filter "UserPrincipalName -eq `"$($user.Email)`"" -Properties ObjectGUID).ObjectGUID
        Add-Content -Path C:\IT\PS\opdtPrint.log -Value "Set Code ($($User.Number))"
    }
}