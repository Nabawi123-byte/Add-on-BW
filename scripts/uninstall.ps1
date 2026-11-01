<#  MASTER BW — copot pemasangan (per pengguna).  #>
param([switch]$KeepData, [string]$Target = "$env:LOCALAPPDATA\Programs\MasterBW")
$ErrorActionPreference = 'Continue'
Get-Process MasterBW.Engine -ErrorAction SilentlyContinue | Stop-Process -Force
if (Get-Process WINWORD, EXCEL -ErrorAction SilentlyContinue) { Read-Host "Tutup Word & Excel dulu, lalu tekan Enter" }

foreach ($p in 'MasterBW.WordAddIn', 'MasterBW.ExcelAddIn') {
  Remove-Item "HKCU:\Software\Microsoft\Office\Word\Addins\$p" -Recurse -Force -ErrorAction SilentlyContinue
  Remove-Item "HKCU:\Software\Microsoft\Office\Excel\Addins\$p" -Recurse -Force -ErrorAction SilentlyContinue
}
foreach ($c in '{5d44520c-7476-4ff9-add1-bfedace586de}', '{0e03ea87-46e9-40aa-aeea-16c471df42a7}', '{a0ec4d74-04ef-4d41-9403-139c35f56883}') {
  Remove-Item "HKCU:\Software\Classes\CLSID\$c" -Recurse -Force -ErrorAction SilentlyContinue
}
foreach ($p in 'MasterBW.WordAddIn', 'MasterBW.ExcelAddIn', 'MasterBW.TaskPane') {
  Remove-Item "HKCU:\Software\Classes\$p" -Recurse -Force -ErrorAction SilentlyContinue
}
$keep = Read-Host "Simpan model AI (agar tidak perlu unduh ulang)? (Y/n)"
if ($keep -eq 'n' -or $keep -eq 'N') { Remove-Item $Target -Recurse -Force -ErrorAction SilentlyContinue }
else { Get-ChildItem $Target -Exclude models | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue }
if (-not $KeepData) { Remove-Item "$env:LOCALAPPDATA\MasterBW\cache" -Recurse -Force -ErrorAction SilentlyContinue }
Write-Host "MASTER BW telah dicopot." -ForegroundColor Cyan
