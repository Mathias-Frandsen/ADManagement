<#
 Password Generator Plus for PowerShell

 DISCLAIMER:
 Original generator (HTML/js) at: http://passwordsgenerator.net/plus
 Copyright (C) 2017 Webp.net


 Code adaptation for PowerShell by Mathias Frandsen

 Code heavily modified for ease-of-use & Password Compliance with DK Company's Active Directory Domain Services
 TODO: Implement the flexibility, that the original has. (options for length, characters, etc.; make dynamic)
#>

Add-Type -AssemblyName System.Windows.Forms

Function Start-Script
{
	param
	([Parameter(Position=1)]$PWLength)
	if (-not $PWLength) {$PWLength = 8}
	
	$genPW = Generator -l $PWLength
	Write-Host "PW: $genPW"
    #[System.Windows.Forms.MessageBox]::Show("Generated Password:`n$genPW","Done!","OK","Information")
}

function Generator
{
	param
	($l)
	# Param Def.
	$intLength    = $l
	$strLower     = "abcdefghjkmnpqrstuvwxyz"
	$strUpper     = "ABCDEFGHJKLMNPQRSTUVWXYZ"
	$strNumerical = "23456789"
	$strSymbols   = "!;#$%&'()*+,-./?@_"
	
	$strAll = ("{0}{1}{2}{3}" -f $strLower,$strUpper,$strNumerical,$strSymbols)
	
	# Generation
	$strBuffer     = ""
	$intSub = 4
	for ($intPre = 0; $intPre -lt ($intLength-$intSub); $intPre++)
	{
		$intPos = Get-Random -Minimum 0 -Maximum $strAll.Length
		$strBuffer += $strAll.Substring($intPos,1)
	}
	# Make sure that all types of chars are included, and insert them at random.
	$strBuffer = Insert-Char -MaxBuff ($intLength-$intSub) -Buffer $strBuffer -CharSet $strLower
	$intSub--
	$strBuffer = Insert-Char -MaxBuff ($intLength-$intSub) -Buffer $strBuffer -CharSet $strUpper
	$intSub--
	$strBuffer = Insert-Char -MaxBuff ($intLength-$intSub) -Buffer $strBuffer -CharSet $strNumerical
	$intSub--
	$strBuffer = Insert-Char -MaxBuff ($intLength-$intSub) -Buffer $strBuffer -CharSet $strSymbols
	
	return $strBuffer
}
# Insert a random character from supplied CharSet into string Buffer
Function Insert-Char
{
	param(
	$MaxBuff,
	$Buffer,
	$CharSet
	)
	$intPos    = Get-Random -Minimum 0 -Maximum $CharSet.Length
	$intInsPos = Get-Random -Minimum 0 -Maximum $MaxBuff
	$strTmp    = ""
	$strTmp    = ("{0}{1}" -f $strTmp,$Buffer.Substring(0, $intInsPos))
	$strTmp    = ("{0}{1}" -f $strTmp,$CharSet.Substring($intPos, 1))
	$strTmp    = ("{0}{1}" -f $strTmp,$Buffer.Substring($intInsPos))
	
	$Buffer = $strTmp
	
	Clear-Variable strTmp
	
	return $Buffer
}

# Execution
#Start-Script