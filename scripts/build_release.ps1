param(
    [string]$Version = "0.3.23",
    [string]$PythonPath = ""
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

$python = $PythonPath
if (-not $python) {
    $venvPython = Join-Path $root ".venv\\Scripts\\python.exe"
    if (Test-Path $venvPython) {
        $python = $venvPython
    }
}
if (-not $python) {
    $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCommand) {
        $python = $pythonCommand.Source
    }
}
if (-not $python) {
    throw "No Python interpreter was found. Pass -PythonPath or create .venv\\Scripts\\python.exe."
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
    --collect-data certifi `
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

$updatePublishRoot = Join-Path $distRoot "update-files\\windows\\latest"
$manifestPath = Join-Path $updatePublishRoot "update-manifest.json"
if (Test-Path $manifestPath) {
    Remove-Item -LiteralPath $manifestPath -Force
}
if (Test-Path $updatePublishRoot) {
    Remove-Item -LiteralPath $updatePublishRoot -Recurse -Force
}

& $python "scripts\\build_update_manifest.py" `
    --release-root $releaseRoot `
    --version $Version `
    --manifest-output $manifestPath `
    --asset-output-dir $updatePublishRoot

Write-Output "Built file update manifest: $manifestPath"
