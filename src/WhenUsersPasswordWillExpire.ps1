$Session = New-PSSession DC01
Import-Module ActiveDirectory -PSSession $Session

Function get-pwdset{
    Param([parameter(Mandatory=$true)][string]$user)
    $use = get-aduser $user -properties passwordlastset,passwordneverexpires
    if($use.passwordneverexpires -eq $true) {
        write-host $user "last set their password on " $use.passwordlastset  "this account has a non-expiring password" -foregroundcolor yellow
    }else{
        $til = (([datetime]::FromFileTime((get-aduser $user -properties "msDS-UserPasswordExpiryTimeComputed")."msDS-UserPasswordExpiryTimeComputed"))-(get-date)).days
        if($til -lt "5"){
            write-host $user "last set their password on " $use.passwordlastset "it will expire again in " $til " days" -foregroundcolor red
        }else{
            write-host $user "last set their password on " $use.passwordlastset "it will expire again in " $til " days" -foregroundcolor green
        }
    }
}
$TheUser = Read-Host "Username"

get-pwdset $TheUser
pause
Remove-PSSession $Session