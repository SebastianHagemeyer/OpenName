# Build a standalone, distributable OpenName.exe with Nuitka.
#
# Output: dist\OpenName.dist\OpenName.exe (and the DLLs/Qt plugins it needs).
#
# Prereqs:
#   pip install nuitka
#   pip install PySide6 openpyxl pypdf reportlab qrcode pymupdf
#
# Usage:
#   .\build.ps1                # build into .\dist
#   .\build.ps1 -Onefile       # build a single OpenName.exe (slower startup)
#   .\build.ps1 -Clean         # delete .\dist before building

[CmdletBinding()]
param(
    [switch]$Onefile,
    [switch]$Clean
)

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
Set-Location $root

if ($Clean -and (Test-Path "$root\dist")) {
    Write-Host "Removing existing dist/ ..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force "$root\dist"
}

$nuitkaArgs = @(
    "-m", "nuitka",
    "--standalone",
    "--assume-yes-for-downloads",
    "--enable-plugin=pyside6",
    "--windows-console-mode=disable",
    "--windows-icon-from-ico=paper.ico",
    "--include-data-files=paper.ico=paper.ico",
    "--product-name=OpenName",
    "--file-description=OpenName - per-student name sheet stamper",
    "--copyright=MIT - https://github.com/SebastianHagemeyer/OpenName",
    "--output-dir=dist",
    "--output-filename=OpenName.exe",
    "--remove-output"
)

if ($Onefile) {
    $nuitkaArgs += "--onefile"
}

$nuitkaArgs += "OpenName.py"

Write-Host "Running: python $($nuitkaArgs -join ' ')" -ForegroundColor Cyan
& python @nuitkaArgs
if ($LASTEXITCODE -ne 0) {
    throw "Nuitka build failed with exit code $LASTEXITCODE"
}

if ($Onefile) {
    Write-Host "`nBuilt: dist\OpenName.exe" -ForegroundColor Green
} else {
    Write-Host "`nBuilt: dist\OpenName.dist\OpenName.exe" -ForegroundColor Green
    Write-Host "Distribute the entire 'OpenName.dist' folder." -ForegroundColor Green
}
