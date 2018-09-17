Write-Host "Initializing custom cmdlets build $ProfileBuild..."

$ProfileBuild = "1.0.2"

$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")


try
{
    if (-not (Get-Module ActiveDirectory))
    {
        if ($isAdmin)   { try { $adSession = New-PSSession DC01 } catch { $adSession = New-PSSession DC01 -Credential (Get-Credential) }}
        else            { $adSession = New-PSSession DC01 -Credential (Get-Credential)}
        Import-Module ActiveDirectory -PSSession $adSession
    }

    function SetLocation-SharedScript
    {
        Push-Location -Path '\\fil01\it$\Teknik\#Shared#\PowerShell Commands\AD\maif'
    }

    function Get-ADExport-NameEmailNumber
    {
        SetLocation-SharedScript
        & '.\AdExport-NameEmailEmpNumber.ps1'
    }

    function Export-ADData
    {
        param(
        [Parameter(Mandatory=$true,Position=1)][ValidateSet("All","BAP","Ikast","TOG","Vejle")][string]$Location)
        $scrPath = "C:\IT\PS\AD_Gennemgang\exportInfo.ps1"
        if ($Location -like "All")
        {
            $i = 0
            $x = 1
            foreach ($loc in ("BAP","Ikast","TOG","Vejle"))
            {
                $i++
                Write-Progress -Activity "Exporting..." -Status "Location: $loc ($x/4)" -PercentComplete ($i/8*100)
                & $scrPath -Location $loc
                $i++
                $x++
            }
        }
        else
        {
            & $scrPath -Location $Location
        }
    }

    function Change-UserPassword
    {
        param(
        [Parameter(Mandatory=$true,Position=1)][string]$User,
        [Parameter(Mandatory=$false)]$NewPassword)
    	try
    	{
    		$ADobj_user = Get-ADUser -Filter "samAccountName -eq `"$User`"" -Properties Name,samAccountName,enabled,DistinguishedName,objectGUID
    		cls
    		Write-Host "Strengen `"$user`" matchede bruger: $($adobj_user.Name)"
            if ($NewPassword)
            {
                $inpPass = ConvertTo-SecureString -String $NewPassword -AsPlainText -Force 
            }
            else
            {
    		    $inpPass = Read-Host "Indtast nyt password for '$($adobj_user.Name)'" -AsSecureString
            }
    		try
    		{
    			Set-ADAccountPassword -Identity $adobj_user.objectGUID -NewPassword $inpPass
    			Write-Host "Success!`nPassword for '$($adobj_user.Name)' er nu skiftet."
    			if (-not $adobj_user.Enabled)
    			{
    				Set-ADUser -Identity $adobj_user.SamAccountName -Enabled $True
                    Move-ADUser $User
    			}
            Get-ADUser $User -Properties Name,SamAccountName,PasswordLastSet,Enabled,DistinguishedName|Select Name,SamAccountName,PasswordLastSet,Enabled,DistinguishedName
    		}
    		catch { Write-Host "`nCould not change password.`n`nPlease check, and make sure that the password fulfills the requirements." }
    	}
    	catch { cls; Write-Host "Error!`nInput '$input' did not match any records.`n`nPlease check, and try again." }
        pause
        Clear-Variable User
        if ($NewPassword){cls}
        Clear-Variable NewPassword
    }
    
    function Get-GroupMembership
    {
        param([Parameter(Mandatory=$true,Position=1)][string]$User)
        Write-Host "Loading`nPlease Wait...`n`n"

        Get-ADUser $user -Properties MemberOf | Select -ExpandProperty MemberOf | % { (Get-ADObject $_).Name } | Sort 
        Write-Host "`nShowing group membership for: $((Get-ADUser $User).Name) ($User)"
        Clear-Variable User
        Write-Host "`n"
    }
    
    function Move-ADUser
    {
        param([Parameter(Mandatory=$true,Position=1)]$User)
        cls
        
        $ADobj_user = Get-ADUser -Filter "samAccountName -eq `"$User`"" -Properties Name,samAccountName,enabled,DistinguishedName,objectGUID
        
        Set-ADUser -Identity $adobj_user.SamAccountName -Enabled $True
    				switch((Read-Host ("Skal brugeren i `"External Users`"? (j/n)")).ToLower())
    				{
    					"j"  { $startOU = "External Users" }
                        "ja" { $startOU = "External Users" }
    					"n"  { $startOU = "Users" }
                        "nej"{ $startOU = "Users" }
    					default { Write-Host "Svar ikke genkendt. Antager `"n`"/`"nej`""; $startOU = "Users" }
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
    					$selection = Read-Host ("Vælg OU")
    					
    					$destOU = ("OU={0},{1}" -f $OUs[$selection-1].Name,$destOU)
    				}
    				Move-ADObject -TargetPath $destOU -Identity $adobj_user.ObjectGUID
                    Get-ADUser $User -Properties Name,DistinguishedName,Enabled,PasswordLastSet,Description | Select Name,DistinguishedName,Enabled,PasswordLastSet,Description
        Clear-Variable User
    }
    
    function Get-PwdLastSet
    {
        param ([Parameter(Mandatory=$true,Position=1)]$Username)
        $pwdLastSet = [datetime](Get-ADUser $Username -Properties PasswordLastSet | Select -ExpandProperty PasswordLastSet)
        $PwdExpire = $pwdLastSet.AddDays(60)
        $daysValid = New-TimeSpan -Start (Get-Date) -End $pwdExpire
        $user = Get-ADUser $Username
        Write-Host "User                 :  $($User.Name) ($($User.SamAccountName))"
        Write-Host "Password Last Set    :  $($pwdLastSet)"
        Write-Host "Password expires on  :  $($PwdExpire)"
        Write-Host "Password validaty    :  $($daysValid.Days) day(s)"
        Clear-Variable Username
        Clear-Variable User
    }

    function Copy-GroupMembership
    {
        param (
        [Parameter(Mandatory=$True,Position=1)][string]$SourceUser,
        [Parameter(Mandatory=$True,Position=2)][string]$TargetUser)
        
        try
        {
            try
            {
                $srcGrp_usr = Get-ADUser -Filter "samAccountName -eq `"$SourceUser`"" -Properties MemberOf
            }
            catch {Write-Host "No src!"}
            try
            {
                $dest = Get-ADUser $TargetUser
            }
            catch {Write-Host "No dest!"}
            
            foreach ($grp in $srcGrp_usr.MemberOf)
            {
                Add-ADGroupMember -Identity $grp -Members $TargetUser
            }
            Get-GroupMembership -User $TargetUser
        }
        catch{Write-Host "No work."}
        pause
    }

    function Get-GroupPolicyStatus
    {
        [CmdletBinding()]
        param([Parameter(Mandatory=$true)][String]$ComputerName,[Parameter(Mandatory=$true)][String]$UserName,[Parameter(Mandatory=$false)][string]$OutputDirectory)
        Write-Host "Initializing and configuring query parameters..."
        if (-not $OutputDirectory) { $OutputDirectory =  "C:"}
        $gpMgmt                  = New-Object -ComObject GPMgmt.GPM
        $gpmConst                = $gpMgmt.GetConstants()
        $gpmRSOP                 = $gpMgmt.GetRSOP($gpmConst.RSOPModeLogging,$null,0)
        $gpmRSOP.LoggingComputer = $ComputerName
        $gpmRSOP.LoggingUser     = $UserName
        Write-Host "Executing query..."
        $gpmRSOP.CreateQueryResults()
        Write-Host "Data Fetched."
        Start-Sleep -Milliseconds 200
        Write-Host "Generating report..."
        $gpmRSOP.GenerateReportToFile($gpmConst.ReportHTML,"$OutputDirectory\GPO_Compliance_Result_$UserName@$ComputerName.html")
        Write-Host "Done. Check path below for result."
    }
}

catch{}
function New-SymbolicLink
{
    Param(
    [Parameter(Position = 0,Mandatory=$true)][string]$Target,
    [Parameter(Position = 1,Mandatory=$true)][string]$Link)

    try{
    [Void]$(New-Item -Path $Link -ItemType SymbolicLink -Value $Target)
    Write-Host ("Successfully created link to: `"{0}`" at location: {1}" -f $Target,$Link)
    }
    catch{
        $errMsg = $_.Exception.Message
        Write-Host ("An error occured during link creation to: `"{0}`" at location: {1}`n`nError details:" -f $Target,$Link)

        $errMsg
    }
}

function Scan-Homefolders
{
    Param(
    [ValidateSet({&{Get-Childitem -Path '\\fil01\it$\Teknik\Vejle\MAIF\Reporter\files' -Name -Directory}})][string]$Version)
}

function Update-Profile
{
	param([switch]$Reload)
	Write-Debug "Setting content of profile..."
	Set-Content $profile -Value (Get-Content "C:\Users\maif\Drive\ADManager\src\initCmdlets.ps1") -Force
	if ($Reload)
	{
	    Write-Debug "Reloading profile..."
	    &$profile
	}
    Write-Host "Profile updated."
}

function ConnectTo-Exchange
{
    $O365Cred = Get-Credential -Message "Indtast Exchange administrator legitimationsoplysninger"
    $O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection
    Import-PSSession $O365Session
}

function Manage-AD
{
    & "C:\Users\maif\Drive\ADManager\addsmgr.ps1"
}
Function Generate-SecurePassword
{
    param([int]$PwLength)
    if (-not $PwLength)
    { & "C:\Users\maif\Drive\ADManager\src\passwordGenerator.ps1" }
    else
    { Powershell "C:\Users\maif\Drive\ADManager\src\passwordGenerator.ps1 -PWLength $PwLength"}
}

function Find-Item
{
    param(
    [Parameter(Mandatory=$true,Position=0,Alias="This")][string]$Item,
    [Parameter(Mandatory=$false,Position=1,Alias="In")][string]$Where
    )

    if (-not $Where)
    {
        $Where = "C:\"
    }

    Write-Host "Searching for `"$What`" in '$Where'"

    $res = Get-ChildItem -Path $Where -Filter "*$What*"

    Write-Host ("...done`nFound {0} result(s):" -f $res.Count)
    $res | % { $_.FullName }
}

$SharedOutputPath = "\\fil01\it´$\Teknik\#Shared#\PowerShell Commands\AD\maif\Output"

New-Alias -Name mklnk -Value New-SymbolicLink -ErrorAction SilentlyContinue
Write-Host "initCmdlets done"
