# Global
$Title = "AD Manager V0.1"

function Display-Legal
{

Write-Host @"
AD Manager V0.1
(C) 2017 - Mathias Frandsen, DK Company Vejle A/S
ALL RIGHTS RESERVED

This message MUST be included with all copies of this script,
along with detailed documentation on changes made to the original script.

All changes and additions are considered donations to the original author, and hereby also the property of the author.

Inquiries and general contact:
Mathias Frandsen
DK Company Vejle A/S
E-Mail: maif@dkcompany.com
Phone: (+45) 9642 5086
Mobile: (+45) 2214 3983

"@

#switch (Read-Host)
}

function main
{
$Title_Suffix = "Main Menu"
## START TEXT DEFINE
$main_Menu = @"
$Title -> $Title_Suffix

  1. Manage User

  2. Get user info

  0. Exit


"@

# Initialize
$run= $true
while($run)
{
    cls
    Write-Host $main_Menu
    switch(Read-Host "Enter Number")
    {
        0 { $run = $false;cls; Write-Host "Bye!`n`n";pause;cls;break }
        1 { Manage-Users }
        2 { Get-UserInfo }
    }
}
}

function Get-UserInfo ([switch]$ReturnValue, [string]$strUsername)
{
    cls
    Write-Host "$Title - Get User Info`n`n"
    if (-not $strUsername)
    {
        $strUsername = Read-Host "Enter Username"
    }
    $usr_data = Get-ADUser $strUsername -Properties "msDS-UserPasswordExpiryTimeComputed",Name,SamAccountName,Enabled,Description,PasswordLastSet

    # Write out data:
    cls
    Write-Host "$Title - User Info for: $($usr_data.Name)`n`n"
    Write-Host "Name:                   $($usr_data.Name)"
    Write-Host "Username:               $($usr_data.SamAccountName)"
    Write-Host "Is Enabled:             $($usr_data.Enabled)"
    Write-Host "Password Last Set:      $($usr_data.PasswordLastSet)"
    Write-Host "Password expires in:    $((([datetime]::FromFileTime($usr_data."msDS-UserPasswordExpiryTimeComputed"))-(get-date)).days) days"
    <#Write-Host "OU:                  $($usr_data)"
    Write-Host "Description:         $($usr_data.Description)"
    Write-Host "Desk Phone:          $($usr_data)"
    Write-Host "Mobile Phone:        $($usr_data)"#>
    if ($ReturnValue)
    {
        return $usr_data
    }
    else
    {
        pause
    }

}

function Manage-Users
{
    cls
    $Title_Suffix = "Manage Users"

    Write-Host "$Title -> $Title_Suffix`n"
    $strUser = Read-Host "Enter Username"

    $objUser = Get-UserInfo -ReturnValue

    $strMenu = @"


Available actions:

  1. Reset Password

  2. $(if ($objUser.Enabled){"Disable"}else{"Enable"}) user
"@
}

function Set-UserInfo
{}

function Setup-UserAccount
{}
main