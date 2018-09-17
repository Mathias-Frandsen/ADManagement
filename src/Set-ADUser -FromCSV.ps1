#displayname,samaccountname,distinguishedname,Description,Department,Officephone,mobilephone,homePhone,title,office,street,emailaddress,Enabled,thumbnailPhoto,info

#Import-CSV -Path '\\fil01\it$\Teknik\Vejle\streamline\csv\adGennemgang_rettet.csv' -Encoding Default | % { Write-Host "Processing user: $($_.samaccountname)"; Set-ADUser $_} 
$HT = @{}
$data = Import-CSV -Path '\\fil01\it$\Teknik\Vejle\streamline\csv\adGennemgang_test.csv' -Encoding Default -Delimiter ","  | Group-Object -AsHashTable -AsString -Property samAccountName
$data.GetEnumerator() | ForEach{($_.Value).GetEnumerator()} #| Set-ADUser -Replace $_
<#
foreach ($user in $data.GetEnumerator())
{
    Get-ADUser -Filter "samAccountName -eq `"$($user.Name)`"" | Set-ADUser -Replace ($user.Value.GetEnumerator() | Group-Object -AsHashTable -AsString)
}
#>