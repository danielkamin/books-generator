Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$AppName = "program"
$Entry   = "program.py"

Remove-Item -Recurse -Force build, dist, release -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path "release\data","release\config" | Out-Null

python -m venv .venv
. .\.venv\Scripts\Activate.ps1
pip install -U pip wheel setuptools
if (Test-Path requirements.txt) { pip install -r requirements.txt }
pip install pyinstaller

pyinstaller --noconfirm --clean --onefile $Entry

Copy-Item "dist\$AppName.exe" "release\" -Force

# Config: copy example if available, else placeholder
if (Test-Path "config\config.example.ini") {
  Copy-Item "config\config.example.ini" "release\config\config.ini" -Force
} else {
  New-Item -ItemType File -Force -Path "release\config\.keep" | Out-Null
}

# Data: nothing to copy; create placeholder so the empty dir is preserved
New-Item -ItemType File -Force -Path "release\data\.keep" | Out-Null

# Zip
$Version = (git describe --tags --always) 2>$null
if (-not $Version) { $Version = Get-Date -Format "yyyyMMddHHmm" }
$ZipPath = ".\${AppName}_windows_${Version}.zip"
if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }
Compress-Archive -Path "release\*" -DestinationPath $ZipPath
Write-Host "Built artifact: $ZipPath"