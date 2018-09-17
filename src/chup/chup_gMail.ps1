$imp = Import-CSV C:\IT\PS\opdtprint.csv -Delimiter ";" -Encoding Default
$imp | % { New-Object -TypeName PSCustomObject -Property @{
                                                           Name=$_.navn;
                                                           Number=$_.nummer;
                                                           Email=(Get-ADUser -Filter "Name -like `"$($_.navn)`"" -Properties UserPrincipalName).UserPrincipalName
                                                           }
}| Export-Csv -Path C:\IT\PS\opdtprint_uptd.csv -Delimiter ";" -Encoding Default -NoTypeInformation -Force