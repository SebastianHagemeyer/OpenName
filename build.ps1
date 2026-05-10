# Build a standalone, distributable OpenName.exe with Nuitka and bundle
# SumatraPDF (used for printing) so the resulting folder runs without any
# extra install on the target machine.
#
# Output: dist\OpenName.dist\OpenName.exe (+ DLLs, Qt plugins, SumatraPDF.exe).
#
# Prereqs:
#   pip install nuitka
#   pip install PySide6 openpyxl pypdf reportlab qrcode pymupdf
#
# Usage:
#   .\build.ps1                    # build, bundle SumatraPDF, slim deps (default)
#   .\build.ps1 -WithPreview       # ALSO compile pymupdf for live preview via
#                                  # Nuitka (adds ~30 min build + ~150 MB,
#                                  # often hangs on the mupdf C compile).
#   .\build.ps1 -WithPreviewFast   # Bundle pre-built pymupdf + fitz + numpy
#                                  # via post-build Copy-Item instead of
#                                  # letting Nuitka compile them (no Zig hang,
#                                  # adds ~60 MB, ~5 sec extra build time).
#                                  # Recommended for Store-bound MSIX builds.
#   .\build.ps1 -Onefile           # single OpenName.exe (slower startup)
#   .\build.ps1 -Clean             # wipe dist\ first
#   .\build.ps1 -NoSumatra         # skip SumatraPDF download (printing won't
#                                  # work unless SumatraPDF is already installed)
#
# Why exclusions
# --------------
# OpenName imports `fitz` (pymupdf) only for the optional live preview pane,
# guarded by a try/except ImportError. Letting Nuitka follow it pulls in
# pymupdf.mupdf which is one C file >100k LOC -- Zig takes 30+ minutes to
# compile it, often hangs. scipy/pandas/matplotlib/numpy/tkinter aren't
# actually used by OpenName at all but get pulled transitively. Excluding
# them cuts the build from ~1.5 hours to ~5 minutes and the dist size
# from ~700 MB to ~130 MB. -WithPreview re-enables pymupdf for users who
# want the live preview pane back.

[CmdletBinding()]
param(
    [switch]$Onefile,
    [switch]$Clean,
    [switch]$NoSumatra,
    [switch]$WithPreview,
    [switch]$WithPreviewFast
)

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
Set-Location $root

# Pinned for reproducibility. Bump when a newer SumatraPDF release is available.
$SumatraVersion = "3.5.2"
$SumatraUrl = "https://www.sumatrapdfreader.org/dl/rel/$SumatraVersion/SumatraPDF-$SumatraVersion-64.exe"

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
    "--product-version=1.0.0",
    "--file-version=1.0.0",
    "--file-description=OpenName - per-student name sheet stamper",
    "--copyright=MIT - https://github.com/SebastianHagemeyer/OpenName",
    "--output-dir=dist",
    "--output-filename=OpenName.exe",
    "--remove-output"
)

# Heavy/unused deps that Nuitka would otherwise pull in transitively. See
# the header for why.
$exclusions = @(
    "scipy", "pandas", "matplotlib", "tkinter", "test", "unittest"
)
if (-not $WithPreview) {
    # -WithPreviewFast also keeps these in --nofollow-import-to so Nuitka
    # doesn't try to compile pymupdf's huge C source. Pre-built binaries
    # are bundled as a data dir below.
    $exclusions += @("fitz", "pymupdf", "numpy")
}
foreach ($pkg in $exclusions) {
    $nuitkaArgs += "--nofollow-import-to=$pkg"
}

# -WithPreviewFast: pymupdf is bundled below as a POST-BUILD copy of the
# pre-built site-packages dir, not via --include-data-dir (which silently
# filters out .pyd/.dll/.py files because it thinks they're code, not
# data, and Nuitka's anti-bloat plugin won't copy them). Nuitka skips
# its own Zig compile of pymupdf.mupdf entirely (saves 30+ min, no hang
# risk). The OpenName runtime imports pymupdf because the standalone
# dist root is on sys.path, so a sibling pymupdf/ folder is importable.

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
    $distDir = "$root\dist"
    $exePath = "$distDir\OpenName.exe"
} else {
    $distDir = "$root\dist\OpenName.dist"
    $exePath = "$distDir\OpenName.exe"
}

if (-not (Test-Path $exePath)) {
    throw "Build claimed success but $exePath does not exist."
}

# -WithPreviewFast post-build: copy pre-built pymupdf + fitz + numpy
# verbatim into the dist tree. We do this here (not via --include-data-dir)
# because Nuitka's anti-bloat plugin filters .pyd/.py/.dll out of data
# dirs. The dist root is on the standalone Python's sys.path, so a
# sibling <pkg>/ folder is importable at runtime.
if ($WithPreviewFast -and -not $Onefile) {
    Write-Host "`nCopying pre-built pymupdf + fitz + numpy into dist ..." -ForegroundColor Cyan
    foreach ($pkg in @("pymupdf", "fitz", "numpy")) {
        $src = & python -c "import $pkg, os; print(os.path.dirname($pkg.__file__))" 2>$null
        if ($LASTEXITCODE -ne 0 -or -not $src -or -not (Test-Path $src)) {
            throw "$pkg not importable; install with 'pip install pymupdf numpy' before using -WithPreviewFast."
        }
        $dst = Join-Path $distDir $pkg
        if (Test-Path $dst) { Remove-Item -Recurse -Force $dst }
        Copy-Item -Recurse -Force $src $dst
        # Strip __pycache__ — wastes space, gets regenerated on first import
        Get-ChildItem $dst -Recurse -Directory -Filter __pycache__ -ErrorAction SilentlyContinue |
            Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        $size = [math]::Round(((Get-ChildItem $dst -Recurse -File | Measure-Object Length -Sum).Sum) / 1MB, 1)
        Write-Host "  $($pkg.PadRight(8)) -> $dst ($size MB)" -ForegroundColor DarkGray
    }
    # Mirror any DLLs found inside the bundled packages to the dist root.
    # _mupdf.pyd depends on mupdfcpp64.dll, but Windows DLL search defaults
    # to the .exe's directory and won't look inside pymupdf/. Without this
    # mirror, `import fitz` silently fails with ImportError and OpenName
    # falls back to the "Live preview disabled" banner.
    $sideDlls = Get-ChildItem -Path (Join-Path $distDir "pymupdf"), (Join-Path $distDir "numpy") `
                              -Filter *.dll -Recurse -ErrorAction SilentlyContinue
    foreach ($dll in $sideDlls) {
        $rootCopy = Join-Path $distDir $dll.Name
        if (-not (Test-Path $rootCopy)) {
            Copy-Item $dll.FullName $rootCopy -Force
            Write-Host "  Mirrored $($dll.Name) -> dist root for DLL search" -ForegroundColor DarkGray
        }
    }
}

if (-not $NoSumatra) {
    $sumatraDest = Join-Path $distDir "SumatraPDF.exe"
    Write-Host "`nDownloading SumatraPDF $SumatraVersion (portable, ~7 MB) ..." -ForegroundColor Cyan
    try {
        Invoke-WebRequest -Uri $SumatraUrl -OutFile $sumatraDest -UseBasicParsing
        Write-Host "  -> $sumatraDest" -ForegroundColor DarkGray
    } catch {
        Write-Warning "Failed to download SumatraPDF: $_"
        Write-Warning "Build is otherwise complete; printing will require a separate SumatraPDF install."
    }

    # GPL-3.0 written-offer notice for the bundled SumatraPDF binary.
    $noticePath = Join-Path $distDir "SumatraPDF.NOTICE.txt"
    $notice = @"
This distribution of OpenName includes SumatraPDF, an unmodified third-party
binary used to send PDFs to the printer.

  SumatraPDF version: $SumatraVersion
  SumatraPDF homepage: https://www.sumatrapdfreader.org/
  SumatraPDF source code: https://github.com/sumatrapdfreader/sumatrapdf
  SumatraPDF license: GPL-3.0 (https://www.gnu.org/licenses/gpl-3.0.txt)

OpenName itself is MIT-licensed; see LICENSE.txt next to OpenName.exe.
"@
    Set-Content -Path $noticePath -Value $notice -Encoding UTF8
}

# Copy the OpenName MIT license into the dist so the whole folder is shippable.
if (Test-Path "$root\LICENSE") {
    Copy-Item "$root\LICENSE" -Destination (Join-Path $distDir "LICENSE.txt") -Force
}

Write-Host ""
Write-Host "Build complete." -ForegroundColor Green
Write-Host "  Executable: $exePath" -ForegroundColor Green
if (-not $Onefile) {
    Write-Host "  Distribute the entire 'OpenName.dist' folder (zip and ship)." -ForegroundColor Green
}
if (-not $NoSumatra) {
    Write-Host "  SumatraPDF bundled. End users do not need a separate install." -ForegroundColor Green
}
