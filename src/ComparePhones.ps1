$csvPath = '\\fil01\it$\Teknik\Vejle\streamline\csv'

$mobilData = Import-CSV -Path "$csvPath\mobil.csv" -Encoding Default -Delimiter ';'
$fastnetData = Import-Csv -Path "$csvPath\fastnetVejle.csv" -Encoding Default -Delimiter ';'
$usersData = Import-CSV -Path "$csvPath\adGennemgang.csv" -Encoding Default

#$fastnetData
#$usersData | % {$_.officephone}

$out = @()

foreach ($objNumber in $fastnetData)
{
    #$objNumber.number
    foreach ($objUser in $usersData)
    {
        #$objUser.Officephone
        if (($objNumber.name -like "$($objUser.displayname)") -or ($objUser.Officephone -like "*$($objNumber.number)"))
        {
            $objHandler = New-Object -TypeName PSObject
            $objHandler | Add-Member -MemberType NoteProperty -Name ADName -Value $objUser.displayname
            $objHandler | Add-Member -MemberType NoteProperty -Name BMSName -Value $objNumber.name
            $objHandler | Add-Member -MemberType NoteProperty -Name ADNumber -Value $objUser.Officephone
            $objHandler | Add-Member -MemberType NoteProperty -Name BMSNumber -Value $objNumber.number
            $objHandler
            #Write-Host "test1"
        }
    }
}