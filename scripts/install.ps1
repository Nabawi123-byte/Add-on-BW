<#
  MASTER BW — Pemasang add-in AI offline untuk Microsoft Word & Excel 2016+
  Dipasang per pengguna (tanpa hak Administrator) ke %LOCALAPPDATA%\Programs\MasterBW.

  Cara pakai: klik ganda INSTALL.cmd, atau jalankan:
    powershell -ExecutionPolicy Bypass -File scripts\install.ps1
#>
param(
  [switch]$SkipModel,           # jangan tawarkan unduh model
  [string]$Target = "$env:LOCALAPPDATA\Programs\MasterBW"
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Split-Path -Parent $here     # folder paket (berisi addin\, engine\, scripts\)

function Info($m) { Write-Host "  $m" -ForegroundColor Gray }
function Ok($m)   { Write-Host "  [OK] $m" -ForegroundColor Green }
function Warn($m) { Write-Host "  [!] $m" -ForegroundColor Yellow }

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "   MASTER BW - Asisten AI Offline Word & Excel" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan

# ---- 1. Prasyarat ----
$net = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue).Release
if (-not $net -or $net -lt 528040) {
  Warn ".NET Framework 4.8 belum terpasang. Unduh: https://dotnet.microsoft.com/download/dotnet-framework/net48"
  Read-Host "Tekan Enter untuk tetap melanjutkan (add-in tidak akan jalan tanpa .NET 4.8)"
} else { Ok ".NET Framework 4.8 tersedia" }

$running = Get-Process WINWORD, EXCEL -ErrorAction SilentlyContinue
if ($running) {
  Warn "Word/Excel sedang terbuka. Tutup semua jendela Word dan Excel dulu."
  Read-Host "Tekan Enter setelah Word & Excel ditutup"
}
Get-Process MasterBW.Engine -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

# ---- 2. Salin berkas ----
Info "Menyalin berkas ke $Target ..."
New-Item -ItemType Directory -Force -Path $Target | Out-Null
foreach ($d in 'addin', 'engine', 'scripts', 'docs') {
  if (Test-Path "$src\$d") {
    if (Test-Path "$Target\$d") { Remove-Item "$Target\$d" -Recurse -Force }
    Copy-Item "$src\$d" "$Target\$d" -Recurse -Force
  }
}
New-Item -ItemType Directory -Force -Path "$Target\models" | Out-Null
if (Test-Path "$src\models") { Get-ChildItem "$src\models\*" -Include *.gguf, *.bin | Copy-Item -Destination "$Target\models" -Force }
# Hapus tanda "diunduh dari internet" agar .NET mau memuat DLL
Get-ChildItem $Target -Recurse -File | Unblock-File -ErrorAction SilentlyContinue
Ok "Berkas tersalin"

# ---- 2b. Runtime Visual C++ (dibutuhkan llama.cpp) ----
$vcDlls = 'msvcp140.dll', 'vcruntime140.dll', 'vcruntime140_1.dll', 'vcomp140.dll'
$vcMissing = $vcDlls | Where-Object { -not (Test-Path "$Target\engine\$_") -and -not (Test-Path "$env:WINDIR\System32\$_") }
if ($vcMissing) {
  Warn "Microsoft Visual C++ Redistributable belum lengkap (tidak ada: $($vcMissing -join ', '))."
  Info "Mengunduh dan memasang VC++ Redistributable x64 (akan muncul permintaan izin Windows)..."
  try {
    $vc = "$env:TEMP\vc_redist.x64.exe"
    Invoke-WebRequest -Uri 'https://aka.ms/vs/17/release/vc_redist.x64.exe' -OutFile $vc -UseBasicParsing
    Start-Process $vc -ArgumentList '/install', '/passive', '/norestart' -Verb RunAs -Wait
    Ok "VC++ Redistributable terpasang"
  } catch {
    Warn "Gagal memasang otomatis. Pasang manual dari https://aka.ms/vs/17/release/vc_redist.x64.exe"
  }
} else { Ok "Runtime Visual C++ tersedia" }

# ---- 3. Daftarkan komponen COM (HKCU, tanpa admin) ----
$dll = "$Target\addin\MasterBW.Office.dll"
if (-not (Test-Path $dll)) { throw "MasterBW.Office.dll tidak ditemukan di $Target\addin" }
$asmName = [System.Reflection.AssemblyName]::GetAssemblyName($dll).FullName
$codeBase = 'file:///' + ($dll -replace '\\', '/')

$classes = @(
  @{ Clsid = '{5d44520c-7476-4ff9-add1-bfedace586de}'; ProgId = 'MasterBW.WordAddIn';  Class = 'MasterBW.Office.Com.WordConnect' },
  @{ Clsid = '{0e03ea87-46e9-40aa-aeea-16c471df42a7}'; ProgId = 'MasterBW.ExcelAddIn'; Class = 'MasterBW.Office.Com.ExcelConnect' },
  @{ Clsid = '{a0ec4d74-04ef-4d41-9403-139c35f56883}'; ProgId = 'MasterBW.TaskPane';   Class = 'MasterBW.Office.Com.TaskPaneHost' }
)
foreach ($c in $classes) {
  $root = "HKCU:\Software\Classes\CLSID\$($c.Clsid)"
  New-Item -Path $root -Force | Out-Null
  Set-ItemProperty -Path $root -Name '(default)' -Value $c.Class
  New-Item -Path "$root\InprocServer32" -Force | Out-Null
  Set-ItemProperty -Path "$root\InprocServer32" -Name '(default)' -Value 'mscoree.dll'
  Set-ItemProperty -Path "$root\InprocServer32" -Name 'ThreadingModel' -Value 'Both'
  Set-ItemProperty -Path "$root\InprocServer32" -Name 'Class' -Value $c.Class
  Set-ItemProperty -Path "$root\InprocServer32" -Name 'Assembly' -Value $asmName
  Set-ItemProperty -Path "$root\InprocServer32" -Name 'RuntimeVersion' -Value 'v4.0.30319'
  Set-ItemProperty -Path "$root\InprocServer32" -Name 'CodeBase' -Value $codeBase
  New-Item -Path "$root\ProgId" -Force | Out-Null
  Set-ItemProperty -Path "$root\ProgId" -Name '(default)' -Value $c.ProgId
  New-Item -Path "$root\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}" -Force | Out-Null
  New-Item -Path "HKCU:\Software\Classes\$($c.ProgId)\CLSID" -Force | Out-Null
  Set-ItemProperty -Path "HKCU:\Software\Classes\$($c.ProgId)" -Name '(default)' -Value $c.Class
  Set-ItemProperty -Path "HKCU:\Software\Classes\$($c.ProgId)\CLSID" -Name '(default)' -Value $c.Clsid
}
Ok "Komponen COM terdaftar"

# ---- 4. Daftarkan add-in ke Word & Excel ----
foreach ($app in @(@{ Name = 'Word'; ProgId = 'MasterBW.WordAddIn' }, @{ Name = 'Excel'; ProgId = 'MasterBW.ExcelAddIn' })) {
  $k = "HKCU:\Software\Microsoft\Office\$($app.Name)\Addins\$($app.ProgId)"
  New-Item -Path $k -Force | Out-Null
  Set-ItemProperty -Path $k -Name 'FriendlyName' -Value 'MASTER BW'
  Set-ItemProperty -Path $k -Name 'Description' -Value 'Asisten AI offline: akademik, perkantoran, keuangan'
  New-ItemProperty -Path $k -Name 'LoadBehavior' -Value 3 -PropertyType DWord -Force | Out-Null
  # Pastikan add-in tidak masuk daftar "Disabled Items" dari percobaan sebelumnya
  foreach ($ver in '16.0', '15.0') {
    $res = "HKCU:\Software\Microsoft\Office\$ver\$($app.Name)\Resiliency\DoNotDisableAddinList"
    New-Item -Path $res -Force | Out-Null
    New-ItemProperty -Path $res -Name $app.ProgId -Value 1 -PropertyType DWord -Force | Out-Null
  }
}
Ok "Add-in terdaftar di Word dan Excel (tab 'MASTER BW')"

# ---- 5. Pengaturan awal ----
$dataDir = "$env:LOCALAPPDATA\MasterBW"
New-Item -ItemType Directory -Force -Path $dataDir | Out-Null
if (-not (Test-Path "$dataDir\settings.json")) {
  $name = Read-Host "Nama Anda (untuk tanda tangan surat, boleh dikosongkan)"
  $inst = Read-Host "Nama instansi/sekolah/kampus (boleh dikosongkan)"
  $city = Read-Host "Kota (boleh dikosongkan)"
  $settings = [ordered]@{ UserName = $name; Institution = $inst; City = $city }
  ($settings | ConvertTo-Json) | Set-Content -Path "$dataDir\settings.json" -Encoding UTF8
}
Ok "Pengaturan tersimpan di $dataDir"

# ---- 6. Model AI ----
$hasModel = Get-ChildItem "$Target\models\*.gguf" -ErrorAction SilentlyContinue
if (-not $hasModel -and -not $SkipModel) {
  Write-Host ""
  Warn "Model AI belum ada. Model hanya diunduh SEKALI, setelah itu semuanya bekerja OFFLINE."
  $ans = Read-Host "Unduh model sekarang? (Y/n)"
  if ($ans -ne 'n' -and $ans -ne 'N') { & "$Target\scripts\download-model.ps1" -ModelsDir "$Target\models" }
  else { Info "Nanti jalankan: $Target\scripts\download-model.ps1" }
} elseif ($hasModel) { Ok "Model AI ditemukan: $($hasModel[0].Name)" }

Write-Host ""
Write-Host "Instalasi selesai! Buka Word atau Excel, lalu cari tab 'MASTER BW'." -ForegroundColor Cyan
Write-Host "Pintasan palet perintah: Ctrl+Alt+Space. Log: $dataDir\logs" -ForegroundColor Cyan
Write-Host ""
