cls

$User = Read-Host "Username"

$ADobj_user = Get-ADUser -Filter "samAccountName -eq `"$User`"" -Properties Name,samAccountName,enabled,DistinguishedName,objectGUID

switch((Read-Host ("Skal brugeren i `"External Users`"? (y/n)")).ToLower())
{
	"y" { $startOU = "External Users" }
	"n" { $startOU = "Users" }
	default { Write-Host "`"y`" or `"n`" was not provided, assuming no"; $startOU = "Users" }
}

$destOU = "OU=$startOU,OU=Domain,DC=Domain,DC=local"
$runOU = $true
while ($runOU)
{
	cls
	Write-Host $destOU
	$OUs = Get-ADOrganizationalUnit -SearchBase $destOU -SearchScope OneLevel -Filter * | Select-Object Name
	if ($OUs -eq $null)
	{
		$runOU = $false
		break
	}
	$index = 1
	foreach ($OU in $OUs)
	{
		Write-Host ("{0}. {1}" -f $index, $OU.Name)
		$index++
	}
	$selection = Read-Host ("Chose OU")
	
	$destOU = ("OU={0},{1}" -f $OUs[$selection-1].Name,$destOU)
}
Move-ADObject -TargetPath $destOU -Identity $adobj_user.ObjectGUID