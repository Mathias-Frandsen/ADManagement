##<init>##
Clear
$adSessionCreated = $false
function StartScript
{
    if (-not (Get-Module ActiveDirectory))
    {
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
        if ($env:USERNAME -like "admin*" -and (([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).ContextType -eq "Domain"))
        {
            $RemoteSession = ConnectTo-ActiveDirectory
        }
        else
        {
            $RemoteSession = ConnectTo-ActiveDirectory -Credential (Get-Credential -Message "Enter Domain Administrator credentials")
        }
    }
    Clear
	# Prepare 
	Write-Host "Vent et øjeblik; ting klargøres..."
    Add-Type -AssemblyName System.Windows.Forms
    Main
    Clear
    if ($adSessionCreated)
    {
	    Cleanup($RemoteSession)
    }
    Write-Debug "Script End."
    Write-Host "`n`n`nFærdig.`n`n`n"
    Pause
    Clear
}
##</init>##

# Global Vars
$LoggingLevel = "default"

##<Structural>##
function Main
{
    param($RemoteSession,$Credentials)
	Clear
	Write-Host "Licence/Credits:

Active Directory Domain Services og Microsoft Exchange Administrations script til PowerShell
Skrevet af: Mathias Frandsen

Scriptet er stadig under udvikling, så tag højde for at fejl kan opstå, samt at nogle funktioner mangler.

Nogle funktioner, er helt eller delvist taget fra andre scripts skrevet af bl.a.:
		
		Mathias Frandsen
		Daniel Christensen
		Martin Rosenkrands

Dette script tilhører DK Company A/S.
Copyright (C) DK Company A/S (Mathias Frandsen), 2016-2018
	"  # License/Legal info
	pause
    $MainMenu = $true
    while ($MainMenu)
    {
		Clear
		Write-Host "
Active Directory Domain Services og Microsoft Exchange Server administrations script til PowerShell

	1. Administrer ADDS
    
    9. Aktiver Debug logging (WIP)

	0. Afslut
" # Main Menu
        switch(Read-Host "Vælg menupunkt")
        {
            0 { $MainMenu = $false; break }
            1 { AD_Main }
            9 { $LoggingLevel = "debug"; cls; Write-Host "`n`n  Debug logging aktiveret"; pause }
            default {[System.Windows.Forms.MessageBox]::Show("Input ikke genkendt","Fejl!","OK","Warning")}
        }
    }
}
function Cleanup
{
    param($RemoteSession)
    Write-Host "Rydder op..."
    Remove-PSSession $RemoteSession
}
function ConnectTo-ActiveDirectory # DONE
{
	param($Credential)
    if (-not $Credential)
    {
        $RemoteSession = New-PSSession DC01
    }
    else
    {
	    $RemoteSession = New-PSSession DC01 -Credential $Credential
    }
    Import-Module ActiveDirectory -PSSession $RemoteSession
    $adSessionCreated = $true
    return $RemoteSession
}
function ConnectTo-Exchange        # DONE
{
    if (-not $ExchCreds)
    {
        $ExchCreds = Get-Credential -Message "Please provide administrator credentials (admin<username>@dkcompany.com, and password)"
    }
    Write-Host "Opretter forbindelse til Microsoft Exchange..."
	$ExchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionURI https://ps.outlook.com/powershell -AllowRedirect -Credential $ExchCreds -Authentication Basic
	Import-PSSession $ExchSession
    Write-Host "Session oprettet."
}
function Write-Log
{
    param(
    [string]$LogLevel,
    [string]$Module,
    [string]$Action,
    [string]$Result,
    [string]$Details
    )

    $doWriteLog = $false

    if ($LoggingLevel -eq "debug")
    {
        $doWriteLog = $true
    }
    elseif ($LoggingLevel -eq "default" -and $LogLevel -ne "debug")
    {
        $doWriteLog = $true
    }
    else 
    {
        $doWriteLog = $false
    }

    if ($doWriteLog)
    {
        $logOut = New-Object -TypeName PSCustomObject -Property @{Time=(Get-Date);Level=$logLevel;Module=$Module;Action=$Action;Result=$Result;Details=$Details}
        $logOut | Export-CSV -Path "X:\Teknik\#Shared#\PowerShell Commands\AD\maif\Output\Log\addsmgr.log" -NoTypeInformation -Encoding Default -Append -Delimiter ";"
    }
}
##</Structural>##

##<ADDS>##
function AD_Main
{
    $AD_Menu = $true
    while($AD_Menu)
    {
        Clear
        Write-Host "
Administration af Active Directory Domain Services

    1. Skift password for brugerkonto
    2. Tjek hvornår en bruger sidst skiftede password
    3. Generér email signatur
    4. Administrér brugerstatus (Under udvikling)
    5. Find næste ledige PC navn
    6. Kopier gruppemedlemskab
    7. Vis gruppemedlemskab for bruger

    0. Tilbage
"
        switch(Read-Host "Vælg menupunkt")
        {
            0 { $AD_Menu   = $false ;break}
            1 { $SkipPause = Change-Password }
			2 { $SkipPause = Get-PasswordLastSet }
            3 { $SkipPause = Generate-Signature }
            4 { $SkipPause = UserStatusManagement }
            5 { $SkipPause = Get-NextFreeComputerName }
            6 { $SkipPause = Copy-ADGroupMembership }
            7 { $SkipPause = Get-ADGroupMembership }
            default {[System.Windows.Forms.MessageBox]::Show("Input ikke genkendt","Fejl!","OK","Warning"); $SkipPause = $true;}
        }
		if ($AD_Menu -and -not($SkipPause))
		{
			pause
		}
        else
        {
            $SkipPause = $false
        }
    }
}
function Change-Password           # TODO: Forms/no forms option
{
    param(
    [Parameter(Position=1)][string]$User)
    Clear
    Write-Host "Skift password for AD konto`n"


    if (-not $User)
    {
        $User = Read-Host "Indtast brugernavn"
    }

    if ($User -eq "0")
    {
        return $true
    }

    try
    {
        Write-Host "Søger efter bruger..."

        $ADobj_user = Get-ADUser -Filter "samaccountname -eq `"$User`"" -Properties Name,samAccountName,enabled,DistinguishedName,objectGUID
        if ($ADobj_user -eq $null)
        {
            [System.Windows.Forms.MessageBox]::Show("Bruger '$($User)' kunne ikke findes. Tjek at brugernavnet er stavet rigtigt, og prøv igen.","Fejl!","OK","Error")
            break
        }
    	Write-Host "Strengen `"$user`" matchede bruger: $($adobj_user.Name)"
        $NewPassword = Read-Host "Indtast nyt password for '$($adobj_user.Name)'" -AsSecureString
    	try
    	{
    		Set-ADAccountPassword -Identity $adobj_user.ObjectGUID -NewPassword $NewPassword
    		[System.Windows.Forms.MessageBox]::Show("Password for $($ADObj_user.Name) er nu opdateret!","Success!","OK","Information")
    		if (-not $adobj_user.Enabled)
    		{
    			Set-ADUser -Identity $adobj_user.SamAccountName -Enabled $True
                Move-ADUser $User
    		}
        cls
        Write-Host "Opdaterede brugeroplysninger:"
        Get-ADUser $User -Properties Name,SamAccountName,PasswordLastSet,Enabled,DistinguishedName|Select Name,SamAccountName,PasswordLastSet,Enabled,DistinguishedName
    	}
    	catch {[System.Windows.Forms.MessageBox]::Show("Fejlbesked:`n$($_.Exception.Message)","Fejl!","OK","Error")}
    }
    catch { [System.Windows.Forms.MessageBox]::Show("Fejlbesked:`n$($_.Exception.Message)","Fejl!","OK","Error") }
    Clear-Variable User
    Clear-Variable NewPassword
    return $false
}
function Get-PasswordLastSet
{
    param ($Username)
	cls
	Write-Host ""
	if (-not $Username)
	{
		$Username = Read-Host "Indtast brugernavn"
	}

    if ($Username -eq "0")
    {
        return $true
    }

    $pwdLastSet = [datetime](Get-ADUser $Username -Properties PasswordLastSet | Select -ExpandProperty PasswordLastSet)
    $PwdExpire = $pwdLastSet.AddDays(60)
    $daysValid = New-TimeSpan -Start (Get-Date) -End $pwdExpire
    $user = Get-ADUser $Username
    Write-Host "Bruger oplysninger            :  $($User.Name) ($($User.SamAccountName))"
    Write-Host "Password sidst sat            :  $($pwdLastSet)"
    Write-Host "Dato for udløb af password    :  $($PwdExpire)"
    Write-Host "Resterende password gyldighed :  $($daysValid.Days) day(s)"
    Clear-Variable Username
    Clear-Variable User
    return $false
}
function Generate-Signature        # DONE
{
    param(
    $User,
    [switch]$CopyToHomeDir)
	cls
	if (-not $User)
	{
		$User = Read-Host "Indtast brugernavn"
	}

    if ($User -eq "0")
    {
        return $true
    }

    $FolderRoot = "\\fil01\it$\Teknik\Vejle\streamline\addsmgr\gensig"

    $objCopyPrompt = ""
    Write-Host "Henter data..."
    $userData = Get-ADUser -Filter "samAccountName -eq `"$User`"" -Properties Name,mail,title,department,company
    
    # Generate "New"
    $Sig_New_HTML = Get-Content "$FolderRoot\_template\new.html" 
    $Sig_New_HTML = $Sig_New_HTML -replace '{userdata.name}',$userData.Name -replace '{userdata.title}',$userData.Title -replace '{userdata.department}',$userData.Department -replace '{userdata.company}',$userData.company -replace '{userdata.mail}',$userData.mail
    
    $Sig_Reply_HTML = Get-Content "$FolderRoot\_template\reply.html"
    $Sig_Reply_HTML = $Sig_Reply_HTML -replace '{userdata.name}',$userData.Name -replace '{userdata.title}',$userData.Title -replace '{userdata.department}',$userData.Department -replace '{userdata.company}',$userData.company -replace '{userdata.mail}',$userData.mail
    
    [void](mkdir $FolderRoot\output\$User)
    $FolderUser = "$FolderRoot\output\$User"
    Write-Host "Genererer html signatur..."
    Copy-Item -Path $FolderRoot\_template\new   -Destination $FolderUser\new_files   -Recurse
    Copy-Item -Path $FolderRoot\_template\reply -Destination $FolderUser\reply_files -Recurse
    Out-File -FilePath $FolderUser\new.htm   -Encoding default -InputObject $Sig_New_HTML
    Out-File -FilePath $FolderUser\reply.htm -Encoding default -InputObject $Sig_Reply_HTML
    Write-Host "Sætter variabler..."
    # Store root work folder in var, so it's available for easy use.
    $FolderRoot = "$FolderRoot\Output\$User"
    
    # Create object for MS Word
    $objWord = New-Object -ComObject Word.Application
    
    $saveFormatText = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatText")
    
    
    #NEW
    ## Open "new.htm"
    $objDocNew = $objWord.Documents.Open("$FolderRoot\new.htm")
    
    ## Save as:
    Write-Host "Konverterer `"New`" fra HTML til RTF..."
    $objDocNew.SaveAs([ref]"$FolderRoot\new.rtf",[ref]$SaveFormat::wdFormatRTF)      #RTF
    Write-Host "Konverterer `"New`" fra HTML til TXT..."
    $objDocNew.SaveAs([ref]"$FolderRoot\new.txt",[ref]$SaveFormatText)               #TXT
    
    #REPLY
    $objDocReply = $objWord.Documents.Open("$FolderRoot\reply.htm")
    Write-Host "Konverterer `"Reply`" fra HTML til RTF..."
    $objDocReply.SaveAs([ref]"$FolderRoot\reply.rtf",[ref]$SaveFormat::wdFormatRTF)  #RTF
    Write-Host "Konverterer `"Reply`" fra HTML til TXT..."
    $objDocReply.SaveAs([ref]"$FolderRoot\reply.txt",[ref]$SaveFormatText)           #TXT
    Write-Host "Ryder op i objekter..."
    $objWord.Quit()
    
    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$objWord)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Remove-Variable objWord
    
    if (Test-Path "\\fil01\users$\$user")
    {
        switch(Read-Host "Skal signaturen kopieres til brugerens homefolder?(j/n)")
        {
            "j" { $CopyToHomeDir = $true }
            "n" { $CopyToHomeDir = $false }
            default { $CopyToHomeDir = $false }
        }
    }

    if ($CopyToHomeDir)
    {
        Copy-Item -Path $FolderRoot -Destination "\\fil01\users$\$user\signatures" -Recurse -Force
        #Copy-Item -Path "\\fil01\it$\Teknik\#Shared#\PowerShell Commands\AD\maif\SetSignature.ps1" -Destination "\\fil01\users$\$usr"
    }
    Write-Host "Færdig."
    return $false
}
function Copy-ADGroupMembership    # Done-arooo
{
    param (
    [string]$SourceUser,
    [string]$TargetUser)
    cls

    Write-Host "Kopier AD Gruppe medlemskab`n`n"

    if (-not $SourceUser)
    {
        $SourceUser = Read-Host "Kopier gruppe(r) fra bruger"
        if ($SourceUser -eq "0") {return $true}
    }
    if (-not $TargetUser)
    {
        $TargetUser = Read-Host "Destinations bruger"
        if ($TargetUser -eq "0") {return $true}
    }

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
        Get-ADGroupMembership -User $TargetUser
    }
    catch{Write-Host "No work."}
    return $false
}
function Get-ADGroupMembership     # DONE
{
    param([Parameter(Position=1)][string]$User,[switch]$scr)
    if (-not $scr) { cls }
    Write-Host "-- Vis AD Gruppemedlemskab --`n"
    if (-not $User) { $User = Read-Host "Indtast brugernavn" }
    if ($user -eq "0") { return $true}
    Write-Host "Indlæser grupper, vent venligst...`n"

    Get-ADUser $user -Properties MemberOf | Select -ExpandProperty MemberOf | Select Name
    Write-Host "`nViser gruppemedlemskab for: $((Get-ADUser $User).Name) ($User)"
    Clear-Variable User
    Write-Host "`n"
    if (-not $scr) { pause } 
    return $false
}
function Setup-User                # DONE but meh.. Add something?
{
    param($RemoteSession,
    [string]$Username,
    [string]$FirstName,
    [string]$LastName,
    [string]$CopyPermissionsFrom,
    [string]$Title,
    [string]$Department,
    [string]$Company,
    [string]$Manager)

    Set-ADUser -Identity (Get-ADUser $Username) -GivenName $FirstName -Surname $LastName -Company $Company -Department $Department -Enabled $True -Manager (Get-ADUser $Manager)
}
function UserStatusManagement      # TODO: Finish "Get";Write "Set"
{
	$StatusMgr = $true
	
	while ($StatusMgr)
	{
		Clear
		Write-Host "
Tjek/Skift status for bruger

	1. Tjek status (WIP)
	2. Skift status (WIP)
	
	0. Tilbage
"
		switch(Read-Host "Vælg menupunkt")
		{
			0 { $StatusMgr = $false; break }
			1 { 
					$objPC = Get-ADComputer -Filter * -Properties Name,Description
					clear
					$objUser = Get-ADUser (Read-Host "Indtast brugernavn") -Properties Name,SamAccountName,DistinguishedName,CanonicalName,Enabled,PasswordLastSet
					
					if ($objUser.CanonicalName -Like "*Bruger Stoppet*")
					{
						$objUser_State = "Deaktiveret"
					}
					elseif ($objUser.CanonicalName -Like "Domain.local/Domain/Users/*/Barsel")
					{
						$objUser_State = "Barsel"
					}
					else
					{
						$objUser_State = "Aktiv"
					}
					
					
					# Try to find users PC, so we can interact with it later.
					$objPC = $objPC | Where { $_.Description -like "*$($objUser.Name)*" }
					if (-not $objPC)
					{
						Write-Host "Advarsel: Der kunne ikke findes en PC for bruger `"$($objUser.Name)`""
						$objPC = New-Object -TypeName PSCustomObject -Property @{Name="N/A"}
					}
	
					# Figure out user location
					if ($objUser.DistinguishedName -like "*Vejle*")
					{
						$objUser_Location = "Vejle"
					}
					elseif ($objUser.DistinguishedName -like "*Ikast*")
					{
						$objUser_Location = "Ikast"
					}
					elseif ($objUser.DistinguishedName -like "*BAP*")
					{
						$objUser_Location = "BAP"
					}
					else
					{
						$objUser_Location_Split = $objUser.DistinguishedName.Split(',')
						for ($i = 0;$i -lt $objUser_Location_Split.Length;$i++)
						{
							if ($i -ge 1)
							{
								$objUser_Location = $objUser_Location + ",$($objUser_Location_Split[$i])"
							}
							else
							{
								$objUser_Location = $objUser_Location_Split[0]
							}
						}
					}
					if ($objUser.Enabled)
					{
						$strUserEnabled = "Ja"
					}
					else
					{
						$strUserEnabled = "Nej"
					}
	
					$objStateOut = New-Object -TypeName PSCustomObject -Property @{
						Navn=$objUser.Name
						Brugernavn=$objUser.SamAccountName
						Status=$objUser_State
						Placering=$objUser_Location
						PC=$objPC.Name
						Aktiveret=$strUserEnabled
						PwdLastSet=$objUser.PasswordLastSet
					}
	
					$objStateOut | Select Navn,Brugernavn,Status,Placering,PC,Aktiveret,@{Expression={$_.PwdLastSet};Label="Sidste password skift"}
					Write-Host ""
	
				}
			2 { WipCmd }
		}
		if ($StatusMgr)
		{
			pause
		}
	}
	return $true
}
function Get-NextFreeComputerName
{
    Clear
    Write-Host "Find næste ledige PC navn
Lokationer:
    
    1. BAP
    2. CPH
    3. Ikast
    4. Vejle
    5. Ext - DK
    6. Ext - UK
    7. Ext - DE
    8. Ext - NO


    0. Afbryd

" # Sæt Ext som et punkt for sig, aktiverer undermenu med DK,UK,De,NO osv. 
    switch(Read-Host "Vælg Lokation")
    {
        0 { return $true }
        1 { $sitecode = "BAP" }
        2 { $sitecode = "CPH" }
        3 { $sitecode = "Ikast" }
        4 { $sitecode = "Vejle" }
        5 { $sitecode = "EXTDK" }
        6 { $sitecode = "EXTUK" }
        7 { $sitecode = "EXTDE" }
        8 { $sitecode = "EXTNO" }
        default {cls; Write-Host "Ikke genkendt.`n`nAfbryder...";pause;break}
    } 
    
    Write-Host "`nLokation: $sitecode`n`nHenter info på eksisterende computere..."
	$SearchComputers = $sitecode
	$PrefixCount = $SearchComputers | Measure-Object -Character | Select-Object -ExpandProperty characters
	$Computers = Get-ADComputer -Filter "Name -like `"$SearchComputers*`"" | where { $_.Name -notlike "*Data*" } | Select-Object Name | Sort-Object Name
    Write-Host "...done.`n"

    write-host " Viser de næste ledige PC navne for lokation '$sitecode'" -ForegroundColor Green
    [int]$PreviousNumber = 0
	foreach ($Computer in $Computers){

		$Computer = $Computer.Name
		[int]$CurrentNumber = $Computer.Substring($PrefixCount)
		if ($CurrentNumber.CompareTo($PreviousNumber+1) -eq 1 )
		{
            $i = 0
			for ($PreviousNumber;$PreviousNumber+1 -lt $CurrentNumber; $PreviousNumber++)
			{
                if ($i -lt 10)
                {
				    $MissingNumber = $PreviousNumber+1
				    Write-host "Available Computer Number: $SearchComputers$MissingNumber"
                    $i++
                }
                
			}
		}
		
		[int]$PreviousNumber = $CurrentNumber
	}
    return $false
}
# TODO: Get-ADInfo for an ADObject. 
##</ADDS>##

##<Exchange>##
function Manage-ExchangePermissions      # TODO: Finish sub-functions (See below functions and TODOs)
{
    if (-not (Get-PSSession -ComputerName ps.outlook.com))
    {
        ConnectTo-Exchange
    }
    $ExchMgr = $true
    while ($ExchMgr)
    {
        Clear
        Write-Host "
Administrer Exchange rettigheder

    1. Mailbox rettigheder
    2. Kalender rettigheder
    
    0. Tilbage
    
    
"
        switch (Read-Host "Vælg menupunkt")
        {
            0 { $ExchMgr = $false ; break }
            1 { Manage-MailboxPermissions }
            2 { Manage-MailboxFolderPermissions }
            default { clear ; Write-Host "`n`n`n    Ikke genkendt!`n`n`n" ;pause}
        }
    }
}
function Manage-MailboxPermissions       # TODO: Validation list of accessrights, assuming "FullAccess";Output Formatting
{	
	$MailboxMgr = $true
	while ($MailboxMgr)
	{
		Clear
		Write-Host "
Administrer Exchange Mailbox rettigheder

	1. Vis rettigheder for en mailbox
	2. Tilføj rettigheder for en mailbox
	3. Fjern rettigheder for en mailbox
	
	0. Tilbage
	
	
"
		Switch (Read-Host "Vælg menupunkt")
		{
				0 { $MailboxMgr = $false; break }
				1 { Get-MailboxPermission (Read-Host "Indtast navn på mailboksen") | Select @{Expression={$_.User};Label="Bruger"},@{Expression={$_.AccessRights};Label="Rettigheds niveau"}}
				2 { Add-MailboxPermission -Identity (Read-Host "Indtast navn på mailboksen") -User (Read-Host "Indtast bruger, som skal have rettighed") -AccessRights FullAccess}
				3 { Remove-MailboxPermission -Identity (Read-Host "Indtast navn på mailboksen") -User (Read-Host "Indtast bruger, som er indehaver af rettigheden, der skal fjernes") -AccessRights FullAccess }
		}
		pause
	}
}
function Manage-MailboxFolderPermissions # TODO: Function.
{
    <#
    $mBoxFolder = $true

    while ($mBoxFolder)
    {
        Clear
        Write-Host "
Administrer Exchange Kalender rettigheder

	1. Vis rettigheder for en brugers kalender
	2. Tilføj rettigheder for en brugers kalender
	3. Fjern rettigheder for en brugers kalender
	
	0. Tilbage
	
	
"

        switch (Read-Host "Vælg menupunkt")
        {
            0 {$mBoxFolder = $false;break}
            1 {
                cls
                $usr = Read-Host "Indtast brugernavn (alt før '@' i e-mail adresse)"
                try {
                    Get-MailboxFolderPermission ""
                }
             }
            2 { 
                
             }
            3 { 
                
             }
        }
    }#>
}
##</Exchange>##

##<AD_FS>##
function Get-PrintJobs # TODO: Fix Params; Cleanup code
{
    cls
    # Search parameters:

    $strComputer = Read-Host "Enter Printserver name (default: uniprint01)" # Print server to scan: "."=Localhost, Default="uniprint01"
    if (-not $strComputer) { $strComputer = "uniprint01" }
    $strUser = Read-Host "Filter on username (default: * [all])"              # Filter on username (SamAccountName, Ex. "maif")
    if (-not $strUser) { $strUser = "*" }
    $strPreFilter = Read-Host "Preselection filter (default: * [all]) Use `"uf`" for uniFLOW only"         # Set to "*" for all, set to "uf" for UniFlow queues only     #  Both filters are applied to the search result
    if (-not $strPreFilter) { $strPreFilter = "*" }
    $strQueueFilter = "*"       # Filter on UF Print Queue (Ex. "OfficePrint","Grafisk" etc.) #
    
    # Processing/Output parameters
    $strOutputMethod = "Host"                # Choose output method. Valid options are: "Host"=Print results on screen;"File"=Send output to file specified in the strOutputFile variable
    $strOutputFile = "C:\Output\Output.csv"  # Set the path of the output file (Default="C:\Output\Output.csv")Note: Only applicable when strOutputMethod is set to "File"
    
    # Add assembly for detection of domain/local user, and use that to evaluate elevation
    Add-Type -AssemblyName System.DirectoryServices.AccountManagement
    
    
    
    ### Proccessing: ###
    $colItems = @()
    if ($strPreFilter -like "uf"){$strPreFilter = "uf_in"} # If UniFlow, filter off output and mobileconvert
    # Run search based on elevation level
    if ($env:USERNAME -like "admin*" -and (([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).ContextType -eq "Domain")){ $colItems = get-wmiobject -class "Win32_PrintJob" -namespace "root\CIMV2" -computername $strComputer | where { $_.Name -like "*$strQueueFilter*" -and ($_.Name -Like "$strPreFilter*") } | where {$_.Owner -Like "$strUser*" -or $_.Notify -Like $strUser} }
    else{$objCreds = Get-Credential -Message "Enter Domain Administrator credentials:"; $colItems = get-wmiobject -class "Win32_PrintJob" -namespace "root\CIMV2" -computername $strComputer -Credential $objCreds | where { $_.Name -like "*$strQueueFilter*" -and ($_.Name -like "$strPreFilter*") } | where {$_.Owner -Like "$strUser*" -or $_.Notify -Like $strUser} } 
    foreach ($objItem in $colItems)
    { 
        # For each result, insert key values into a CustomObject, for ease of handling and reading
        $objOut = New-Object -TypeName PSCustomObject -Property @{
                                              User=$objItem.Owner.Split(' ')[0]
                                              Queue=($objItem.Name.Split('_')[$objItem.Name.Split('_').Length-1]).Split(',')[0]
                                              Submitted=("{0}-{1}-{2}" -f $objItem.TimeSubmitted.Substring(0,4),$objItem.TimeSubmitted.Substring(4,2),$objItem.TimeSubmitted.Substring(6,2))
                                              }
        # If the ActiveDirectory module is loaded, fetch the full name of the Owner(user) of the print job
        if (Get-Module ActiveDirectory)
        {
            # Again, execute based on elevation level
            if ($env:USERNAME -like "admin*" -and (([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).ContextType -eq "Domain"))
            {
                $objOut.User = ("{0} ({1})" -f $objOut.User,(Get-ADUser $objOut.User -Properties Name).Name)
            }
            else
            {
                $objOut.User = ("{0} ({1})" -f $objOut.User,(Get-ADUser $objOut.User -Properties Name -Credential $objCreds).Name)
            }
        }
        # Now that all data for the current print job have been collected, processed and formatted, print it to the screen.
        if ($strOutputMethod -eq "Host")
        {
            $objOut |Select User,Queue,Submitted
        }
        elseif ($strOutputMethod -eq "File")
        {
            $objOut |Select User,Queue,Submitted | Export-CSV -Path $strOutputFile -Encoding Default -Append -NoTypeInformation
        }
    }
}
##</AD_FS>##

function WipCmd
{
    Clear
    Write-Host "`n`n`n    Denne funktion er endnu ikke tilgængelig.`n    Har du behov for funktionen, kontakt udvikleren`n`n`n"
}

##<Execution>##
StartScript
##</Execution>##


### DEV NOTES ###
<#

Overall TODOs:

    * Code cleanup
    * Progress Exchange section
    * Add lookup function for AD section. Possibly expand existing lookups to incorporate new lookup, so search is not limited to SamAccountName
    * Fleshout AD section, create better submenus
    * Log all actions.
        * READ and WRITE from AD/EXCH
        * Debug: Enter/exit menu, capture input (no passwords; only on fail to update?)

#>