#
# Script made/modified by Mathias Frandsen
# (C) 2017, DK Company A/S 
#

# Enable remote access:
#$objCreds = Get-Credential
#$Session = New-PSSession -ComputerName DC01 -Credential $objCreds
#Import-PSSession $Session

# Change this to the respective sitecode
$sitecode = "Vejle" 

set executionpolicy bypass –force
cls

Write-Host "`nSitecode: $sitecode`n`nGathering info about existing computers..."
$numbers= @(
(new-object System.DirectoryServices.DirectorySearcher("(&(objectClass=computer)(name=$sitecode*))")).FindAll()  |
%{$_.Properties.name} |
%{$_ -replace "$sitecode",""} |
%{try {[int]$_} catch{$_ -replace "$_",""}} |
sort)  
Write-Host "...done.`n"                     #make sure single items still get treated as a list
Write-Host "Figuring out which numbers are already in use..."
$i= 0
$prev= 0   #Assume you start at "001"... $prev should be one less than your first expected number
$next= $numbers[0]
while (($next- $prev) -le 1     -and     $i -le ($numbers.count-1)) {
    $i++
    $prev= $next
    $next= $numbers[$i]
}
Write-Host "...done.`n"
if ($i -le $numbers.count)
{$nextopen= "$sitecode{0:0##}" -f ($prev+1)}
else 
{$nextopen= "$sitecode{0:0##}" -f ($next+1)} 

write-host "Next Available Workstation: $nextopen`n`n`n" -ForegroundColor Green
Pause

#Remove-PSSession $Session