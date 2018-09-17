param ([Parameter(Mandatory=$true,Position=1)][string]$usr)

if (-not (Get-Module ActiveDirectory))
{
    if ($env:USERNAME -like "*admin*")
    {
        $ADSession = New-PSSession DC01
    }
    else
    {
        $ADSession = New-PSSession DC01 -Credential (Get-Credential)
    }
    Import-Module ActiveDirectory -PSSession $ADSession
}

if ($usr)
{
    $input = $usr
}
else
{
    $input = Read-Host "Enter Username"
}
try
{
    $user = Get-ADUser -Filter "samAccountName -eq `"$input`"" -Properties Name,samAccountName,enabled,DistinguishedName,objectGUID
    cls
    Write-Host "String `"$input`" matched user: $($user.Name)"
    $inpPass = Read-Host "Enter new password for '$($user.Name)'" -AsSecureString
    try
    {
        Set-ADAccountPassword -Identity $user.SamAccountName -NewPassword $inpPass
        Write-Host "`n`nSuccessfully changed password of '$($user.Name)'"
        if (-not $user.Enabled)
        {
            Set-ADUser -Identity $user.SamAccountName -Enabled $True
            switch((Read-Host ("Is user External? (y/n)")).ToLower())
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
            Move-ADObject -TargetPath $destOU -Identity $user.ObjectGUID
        }
    }
    catch { Write-Host "`nCould not change password.`n`nPlease check, and make sure that the password fulfills the requirements." }
}
catch { cls; Write-Host "Error!`nInput '$input' did not match any records.`n`nPlease check, and try again." }
pause