cls
$srcUsr = Read-Host "Source User"
$destUsr = Read-Host "Destination User"

try
{
    try
    {
        $srcGrp_usr = Get-ADUser -Filter "samAccountName -eq `"$srcUsr`"" -Properties MemberOf
    }
    catch {Write-Host "No src!"}
    try
    {
        $dest = Get-ADUser $destUsr
    }
    catch {Write-Host "No dest!"}
    
    foreach ($grp in $srcGrp_usr.MemberOf)
    {
        Add-ADGroupMember -Identity $grp -Members $destUsr
    }
}
catch{Write-Host "No work."}
pause