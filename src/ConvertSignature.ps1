param(
[Parameter(Mandatory=$true)]$usr,
[switch]$CopyToHomeDir)

# Store root work folder in var, so it's available for easy use.
$FolderRoot = "\\fil01\it$\Teknik\Vejle\streamline\test\GenSig\$usr"

# Create object for MS Word
$objWord = New-Object -ComObject Word.Application

$saveFormatText = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatText")


#NEW
## Open "new.htm"
$objDocNew = $objWord.Documents.Open("$FolderRoot\new.htm")

## Save as:
$objDocNew.SaveAs([ref]"$FolderRoot\new.rtf",[ref]$SaveFormat::wdFormatRTF)      #RTF
$objDocNew.SaveAs([ref]"$FolderRoot\new.txt",[ref]$SaveFormatText)               #TXT

#REPLY
$objDocReply = $objWord.Documents.Open("$FolderRoot\reply.htm")

$objDocReply.SaveAs([ref]"$FolderRoot\reply.rtf",[ref]$SaveFormat::wdFormatRTF)  #RTF
$objDocReply.SaveAs([ref]"$FolderRoot\reply.txt",[ref]$SaveFormatText)           #TXT

$objWord.Quit()

$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$objWord)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
Remove-Variable objWord

if ($CopyToHomeDir)
{
    Copy-Item -Path $FolderRoot -Destination "\\fil01\users$\$usr\signatures" -Recurse -Force
    Copy-Item -Path "\\fil01\it$\Teknik\#Shared#\PowerShell Commands\AD\maif\SetSignature.ps1" -Destination "\\fil01\users$\$usr"
}