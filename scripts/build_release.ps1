param(
    [string]$Version = "0.2.0"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

$python = Join-Path $root ".venv\\Scripts\\python.exe"
if (-not (Test-Path $python)) {
    throw "Python virtual environment was not found at $python"
}

& $python -m pip install pyinstaller

$distRoot = Join-Path $root "dist"
$releaseRoot = Join-Path $distRoot "release"
if (Test-Path $releaseRoot) {
    Remove-Item -LiteralPath $releaseRoot -Recurse -Force
}

& $python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --name "MU Unscramble Bot" `
    --add-data "config.json;." `
    --add-data ".env.example;." `
    --add-data "data;data" `
    --collect-all rapidocr_onnxruntime `
    --collect-all pygetwindow `
    "mu_unscramble_bot\\__main__.py"

$appFolder = Join-Path $distRoot "MU Unscramble Bot"
New-Item -ItemType Directory -Path $releaseRoot | Out-Null
Copy-Item -LiteralPath $appFolder -Destination (Join-Path $releaseRoot "MU Unscramble Bot") -Recurse
Copy-Item -LiteralPath "Start MU Unscramble Bot.vbs" -Destination $releaseRoot

$zipPath = Join-Path $distRoot "MU-Unscramble-Bot-$Version-win64.zip"
if (Test-Path $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}
Compress-Archive -Path (Join-Path $releaseRoot "*") -DestinationPath $zipPath
Write-Output "Built release archive: $zipPath"
