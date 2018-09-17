Set-ExecutionPolicy bypass -Force

Copy-Item .\signatures\$env:USERNAME $env:APPDATA\Microsoft\signatures -Recurse -Force
#Copy-Item .\signatures\$env:USERNAME H:\signatures -Recurse -Force

if (test-path "HKCU:\\Software\\Microsoft\\Office\\11.0\\Common\\General") {
  get-item -path HKCU:\\Software\\Microsoft\\Office\\11.0\\Common\\General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
  Write-Host "11"
}  

if (test-path "HKCU:\\Software\\Microsoft\\Office\\12.0\\Common\\General") {
  get-item -path HKCU:\\Software\\Microsoft\\Office\\12.0\\Common\\General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
  Write-Host "12"
}
if (test-path "HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\General") {
  get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
  Write-Host "14"
}
if (test-path "HKCU:\\Software\\Microsoft\\Office\\15.0\\Common\\General") {
  get-item -path HKCU:\\Software\\Microsoft\\Office\\15.0\\Common\\General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
  Write-Host "15"
}
if (test-path "HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General") {
  get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
  Write-Host "16"
}

pause