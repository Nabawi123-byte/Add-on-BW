<#
  Unduh model AI untuk MASTER BW (hanya sekali; setelah itu semua offline).
  Rekomendasi untuk laptop RAM 8–10 GB tanpa GPU:
    1) Qwen3-4B Q4_K_M  (~2,5 GB) — paling pintar, mendukung mode nalar mendalam (high reasoning)
    2) Qwen3-1.7B Q4_K_M (~1,1 GB) — lebih cepat, untuk laptop lebih lambat
  Plus model suara Whisper base (~150 MB) untuk perintah suara bahasa Indonesia.
#>
param([string]$ModelsDir = "$env:LOCALAPPDATA\Programs\MasterBW\models")
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
New-Item -ItemType Directory -Force -Path $ModelsDir | Out-Null

$models = @(
  @{ Key = '1'; Name = 'Qwen3-4B (rekomendasi, ~2,5 GB)'; File = 'Qwen3-4B-Q4_K_M.gguf';
     Url = 'https://huggingface.co/Qwen/Qwen3-4B-GGUF/resolve/main/Qwen3-4B-Q4_K_M.gguf';
     Page = 'https://huggingface.co/Qwen/Qwen3-4B-GGUF'; MinMB = 2000 },
  @{ Key = '2'; Name = 'Qwen3-1.7B (cepat, ~1,1 GB)'; File = 'Qwen3-1.7B-Q4_K_M.gguf';
     Url = 'https://huggingface.co/unsloth/Qwen3-1.7B-GGUF/resolve/main/Qwen3-1.7B-Q4_K_M.gguf';
     Page = 'https://huggingface.co/unsloth/Qwen3-1.7B-GGUF'; MinMB = 900 }
)
$whisper = @{ File = 'ggml-base.bin'; Url = 'https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-base.bin'; MinMB = 100 }

function Get-File($url, $dest, $minMB) {
  if ((Test-Path $dest) -and ((Get-Item $dest).Length / 1MB -ge $minMB)) { Write-Host "  Sudah ada: $dest" -ForegroundColor Green; return $true }
  $tmp = "$dest.part"
  Write-Host "  Mengunduh $url"
  Write-Host "  (bisa memakan waktu beberapa menit, tergantung koneksi)" -ForegroundColor Gray
  try {
    Start-BitsTransfer -Source $url -Destination $tmp -DisplayName 'MASTER BW model' -Description $url
  } catch {
    Write-Host "  BITS gagal, mencoba metode lain..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri $url -OutFile $tmp -UseBasicParsing
  }
  if ((Get-Item $tmp).Length / 1MB -lt $minMB) { Remove-Item $tmp -Force; return $false }
  Move-Item $tmp $dest -Force
  Write-Host "  Selesai: $dest" -ForegroundColor Green
  return $true
}

Write-Host ""
Write-Host "Pilih model AI:" -ForegroundColor Cyan
foreach ($m in $models) { Write-Host ("  {0}) {1}" -f $m.Key, $m.Name) }
$choice = Read-Host "Nomor model (Enter = 1)"
if (-not $choice) { $choice = '1' }
$model = $models | Where-Object { $_.Key -eq $choice } | Select-Object -First 1
if (-not $model) { $model = $models[0] }

$ok = Get-File $model.Url (Join-Path $ModelsDir $model.File) $model.MinMB
if (-not $ok) {
  Write-Host "  Unduhan gagal. Buka halaman ini, unduh file $($model.File), lalu taruh di $ModelsDir" -ForegroundColor Yellow
  Start-Process $model.Page
} else {
  # simpan pilihan model di pengaturan
  $settingsFile = "$env:LOCALAPPDATA\MasterBW\settings.json"
  New-Item -ItemType Directory -Force -Path (Split-Path $settingsFile) | Out-Null
  $s = if (Test-Path $settingsFile) { Get-Content $settingsFile -Raw | ConvertFrom-Json } else { [pscustomobject]@{} }
  $s | Add-Member -NotePropertyName ModelFile -NotePropertyValue $model.File -Force
  ($s | ConvertTo-Json) | Set-Content $settingsFile -Encoding UTF8
}

$voice = Read-Host "Unduh juga model suara (Whisper, ~150 MB) untuk perintah suara? (Y/n)"
if ($voice -ne 'n' -and $voice -ne 'N') { [void](Get-File $whisper.Url (Join-Path $ModelsDir $whisper.File) $whisper.MinMB) }

Write-Host ""
Write-Host "Model siap. Mulai sekarang MASTER BW bekerja sepenuhnya offline." -ForegroundColor Cyan
