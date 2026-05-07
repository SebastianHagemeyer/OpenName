# OpenName

A small PySide6 (Qt) GUI for teachers: pick a worksheet PDF and a class roster,
and the script stamps each student's name (and optionally class, date, and a
QR code) onto a copy of the PDF, then sends one print job per student to the
configured printer via SumatraPDF.

Per-page behaviour is configurable: stamp page 1 only, stamp every page, or
write just the student's name in a small "top-of-page" position on pages 2+.
A live preview shows page 1 with the current settings as you drag the
position sliders.

## Requirements

- Windows (uses SumatraPDF for printing and `os.startfile` for preview).
- Python 3.10+.
- [SumatraPDF](https://www.sumatrapdfreader.org/) installed at
  `%LOCALAPPDATA%\SumatraPDF\SumatraPDF.exe` (the default per-user install
  location).
- Python packages:
  ```
  pip install PySide6 openpyxl pypdf reportlab qrcode pymupdf
  ```
  `qrcode` is optional (disables the QR checkbox if missing). `pymupdf` is
  optional (disables the live preview if missing).

## Folder layout

```
OpenName/
├── OpenName.py                # GUI + stamping/printing logic
├── OpenName.bat               # Windows double-click launcher
├── build.ps1                  # Nuitka build script -> standalone .exe
├── paper.ico                  # Window/taskbar icon
├── measurementsheet.pdf       # Example worksheet (replace with your own)
├── Classes/
│   ├── ExampleClass.xlsx      # Example roster (column A = student names, row 1 = header)
│   └── <YourClass>.xlsx       # Your own rosters (gitignored)
├── OpenName.settings.json     # Auto-written, gitignored
└── batchprint.ps1             # Original prototype, kept for reference
```

The script discovers `.xlsx` rosters in `Classes/` automatically — every file
in that folder shows up in the class dropdown. Worksheet PDFs can sit
anywhere; you pick one with the **Browse...** button.

## Launching

Either:

- Double-click `OpenName.bat`, or
- Run `python OpenName.py` from this folder.

There are **no command-line arguments** — all configuration happens via the
GUI, and the most-recent state is persisted to `OpenName.settings.json` on
the next print.

## First-run walkthrough

The repo ships with a worked example so you can see the app running before
plugging in real data:

1. Install SumatraPDF and the Python packages above.
2. Launch the app (`OpenName.bat` or `python OpenName.py`).
3. The PDF picker defaults to `measurementsheet.pdf`; the class dropdown
   shows `ExampleClass`. You should see 20 fake names load in the Students
   panel and a live preview render on the right.
4. Drag the X/Y sliders for **Name field**, **Class field**, and **Date
   field** until the stamps align with the boxes on your worksheet.
5. Set the **Printer** to your printer's network share (e.g.
   `\\server\YourPrinter`) and click **Print Selected**, or use **Export
   combined PDF** to dump everything to a single file without printing.

## Adding your own class

1. Create `Classes/<ClassName>.xlsx`.
2. Put the class name (or any header) in cell `A1`. From `A2` downward,
   one student per row.
3. Save and re-open the dropdown — the new class appears.

`Classes/*.xlsx` is gitignored except for files matching `Example*.xlsx`,
so your real rosters never get committed by accident.

## Adding your own worksheet PDF

Drop the PDF anywhere readable and pick it via **Browse...**. PDFs are
gitignored except `measurementsheet.pdf` (the bundled example), so worksheets
with sensitive content won't be committed.

## Stamping options

- **Stamp positions** — drag the X/Y sliders (units are PDF points, origin
  top-left). The right-hand pane redraws live.
- **Stamp full Name+Class+Date on every page** — for multi-page worksheets
  where you want the full header on each sheet.
- **Mark name at top of pages 2+** — small name-only stamp on later pages.
- **Include QR code (every page)** — encodes `<class>/<name>/<page>/<total>`
  so scanned work can be auto-sorted.
- **Print Selected** — sends one print job per checked student. Current
  settings are saved to `OpenName.settings.json` for next launch.
- **Export combined PDF** — concatenates every selected student's stamped
  copy into a single PDF (useful for testing without burning paper).

## Building a standalone .exe

You can compile OpenName into a self-contained Windows executable (with all
the necessary DLLs, Qt plugins, **and SumatraPDF** bundled) using
[Nuitka](https://nuitka.net/):

```powershell
pip install nuitka
.\build.ps1              # dist\OpenName.dist\OpenName.exe + DLLs + SumatraPDF
.\build.ps1 -Onefile     # single dist\OpenName.exe (slower startup)
.\build.ps1 -Clean       # wipe dist\ before rebuilding
.\build.ps1 -NoSumatra   # skip SumatraPDF download (smaller, but printing
                         # then requires a separate SumatraPDF install)
```

The first build downloads the MSVC/MinGW toolchain Nuitka uses and takes a
few minutes; subsequent builds are much faster. To distribute, zip up the
entire `OpenName.dist` folder (or just ship the single `OpenName.exe` if you
used `-Onefile`). End users drop `Classes/` and their worksheet PDFs next to
`OpenName.exe` — no Python, no SumatraPDF install, no DLL hunt.

The bundled SumatraPDF is unmodified and licensed under GPL-3.0; the dist
folder includes a `SumatraPDF.NOTICE.txt` pointing at the source.

## Privacy note

This repo intentionally contains no real student data. If you fork or clone
it, be careful not to commit your `Classes/*.xlsx` rosters or your stamped
PDFs — the `.gitignore` is set up to prevent this, but always check
`git status` before pushing.

## License

MIT — see [`LICENSE`](LICENSE).
