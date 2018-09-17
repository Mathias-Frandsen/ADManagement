param(
[string]$Username,
[string]$Name,
[string]$PrintID,
[string]$Phone,
[string]$Mobile,
[switch]$MatchAll
)

#Enter-PSSession -Computername DC01

$AllUsers = Get-ADUser -Filter * -Properties *

if ($MatchAll)
{
    $Result = $AllUsers | Where { ($_.samAccountName -like (if($Username){"*$Username*"}else{"*"}))-and($_.Name -like (if($Name){"*$Name*"}else{"*"}))-and($_.wWWHomePage -like (if($PrintID){"*$PrintID*"}else{"*"}))-and($_.telephoneNumber -like (if($Phone){"*$Phone*"}else{"*"}))-and($_.mobile -like (if($Mobile){"*$Mobile*"}else{"*"})) }
}
else
{
    $Result = $AllUsers | Where{ ($Username -and $_.samAccountName -like "*$Username*")}
}

$Result

#Exit-PSSession