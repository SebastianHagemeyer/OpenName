"""
OpenName - GUI for printing per-student name sheets.

Pick a PDF and a class, optionally configure the stamp positions, and click
Print. For each selected student the script stamps Name (and optionally
Class/Date) onto the PDF and sends one print job per copy via SumatraPDF.

Inspired by batchprint.ps1.
"""

from __future__ import annotations

import io
import json
import os
import re
import subprocess
import sys
import tempfile
import threading
import time
from datetime import datetime
from pathlib import Path

import openpyxl
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.utils import ImageReader

from PySide6.QtCore import Qt, QObject, QTimer, Signal
from PySide6.QtGui import QIcon, QImage, QPixmap
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QFileDialog, QFrame, QGridLayout,
    QGroupBox, QHBoxLayout, QLabel, QLineEdit, QMainWindow, QMessageBox,
    QPushButton, QScrollArea, QSizePolicy, QSlider, QSpinBox, QVBoxLayout,
    QWidget,
)

try:
    import qrcode
    HAS_QR = True
except ImportError:
    HAS_QR = False

try:
    import fitz  # PyMuPDF — used for the live preview pane
    HAS_PREVIEW = True
except ImportError:
    HAS_PREVIEW = False


# --------------------------------------------------------------------------- #
# Paths and defaults                                                          #
# --------------------------------------------------------------------------- #

# When frozen by Nuitka/PyInstaller, __file__ points inside the bundle, so we
# resolve paths next to the executable instead — that way Classes/ and PDFs
# the user drops next to OpenName.exe are picked up.
if "__compiled__" in globals() or getattr(sys, "frozen", False):
    SCRIPT_DIR = Path(sys.executable).resolve().parent
else:
    SCRIPT_DIR = Path(__file__).resolve().parent
BUNDLE_DIR = Path(__file__).resolve().parent  # for assets shipped with the build
CLASSES_DIR = SCRIPT_DIR / "Classes"
ICON_PATH = BUNDLE_DIR / "paper.ico"


def _writable_data_dir() -> Path:
    """Per-user writable directory for OpenName settings + default exports.

    Inside an MSIX install SCRIPT_DIR is read-only (Program Files\\WindowsApps),
    so persisting OpenName.settings.json or defaulting save dialogs there
    silently fails or gets virtualized into the per-package container where
    the user can't easily find it. Use %LOCALAPPDATA%\\OpenName\\ instead.
    """
    base = os.environ.get("LOCALAPPDATA") or str(Path.home())
    d = Path(base) / "OpenName"
    d.mkdir(parents=True, exist_ok=True)
    return d


SETTINGS_PATH = _writable_data_dir() / "OpenName.settings.json"


def _qmark_class_override() -> Path | None:
    """Return the qmark-supplied class roster path if set and valid.

    qmark passes QMARK_CLASS_PATH (an absolute .xlsx) so OpenName uses
    the same roster the dashboard is marking against. When unset, OpenName
    falls back to discovering rosters in its own Classes/ folder.
    """
    p = os.environ.get("QMARK_CLASS_PATH", "").strip()
    if not p:
        return None
    path = Path(p)
    return path if path.is_file() else None


def _class_xlsx_for(name: str) -> Path:
    """Resolve a displayed class name to its .xlsx file."""
    override = _qmark_class_override()
    if override is not None and override.stem == name:
        return override
    return CLASSES_DIR / f"{name}.xlsx"


def _find_sumatra() -> Path:
    """Locate SumatraPDF.exe — prefer a copy bundled next to us, then fall
    back to the user's installed copy. The bundled-first order means the
    distributable .zip works without a separate SumatraPDF install."""
    local_appdata = Path(os.environ.get("LOCALAPPDATA", ""))
    program_files = Path(os.environ.get("ProgramFiles", r"C:\Program Files"))
    candidates = [
        SCRIPT_DIR / "SumatraPDF.exe",
        local_appdata / "SumatraPDF" / "SumatraPDF.exe",
        program_files / "SumatraPDF" / "SumatraPDF.exe",
    ]
    for c in candidates:
        if c.exists():
            return c
    return candidates[1]  # default for the "not found" error message


SUMATRA_PATH = _find_sumatra()

DEFAULTS = {
    "last_pdf": "",
    "last_class": "",
    "class_label": "",
    # Defaults are tuned for namesheetmathmeasurement.pdf (A4, name bar at top).
    # Use the Preview button to adjust for other PDFs.
    "name_x": 130, "name_y": 95,
    "class_x": 475, "class_y": 95,
    "date_x": 540, "date_y": 95,
    "top_x": 40,  "top_y": 40,    # used when mark_top_pages is on
    "qr_x":  40,  "qr_y":  8,     # top-left of QR (used when include_qr is on)
    "qr_size": 32,
    "font_size": 11,
    "date_text": "",
    "include_date": True,
    "printer": r"\\ha-19-print\Follow_Me_BW",
    "monochrome": True,
    "all_pages": False,
    "mark_top_pages": False,
    "include_qr": True,
}


# --------------------------------------------------------------------------- #
# Core helpers                                                                #
# --------------------------------------------------------------------------- #

def load_students(xlsx_path: Path) -> list[str]:
    """Read column A (skipping header row) and return the list of names."""
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        ws = wb.active
        names: list[str] = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:
                continue
            if not row:
                continue
            cell = row[0]
            if cell is None:
                continue
            s = str(cell).strip()
            if s:
                names.append(s)
        return names
    finally:
        wb.close()


def _qr_image(data: str):
    """Build a PIL image of a QR code for `data`."""
    qr = qrcode.QRCode(
        border=1, box_size=10,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
    )
    qr.add_data(str(data))
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white")


def make_overlay_pdf(page_w: float, page_h: float,
                     items: list[tuple[str, float, float]],
                     font_size: float,
                     qr_items: list[tuple[str, float, float, float]] | None = None) -> bytes:
    """Generate a one-page PDF with the given text and QR items.

    items: (text, x_top_left, y_top_left). y is the text baseline from the top.
    qr_items: (data, x_top_left, y_top_left_of_qr, size_pt).
    """
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(page_w, page_h))
    c.setFont("Helvetica", font_size)
    for text, x, y_top in items:
        if not text:
            continue
        c.drawString(x, page_h - y_top, str(text))
    if qr_items and HAS_QR:
        for data, x, y_top, size in qr_items:
            if not data:
                continue
            img = _qr_image(data)
            img_buf = io.BytesIO()
            img.save(img_buf, format="PNG")
            img_buf.seek(0)
            # Reportlab drawImage takes the image's bottom-left in bottom-up coords
            c.drawImage(ImageReader(img_buf), x, page_h - y_top - size,
                        width=size, height=size, mask="auto")
    c.showPage()
    c.save()
    return buf.getvalue()


def stamp_pdf(source: Path, output, name: str, class_label: str,
              date_text: str, positions: dict, font_size: float,
              all_pages: bool = False, mark_top_pages: bool = False,
              top_x: float = 40, top_y: float = 40,
              include_qr: bool = True, qr_size: float = 32,
              qr_x: float = 40, qr_y: float = 8) -> None:
    """Write a stamped copy of `source` to `output` (a path or a writable stream).

    Page 1 always gets the full Name/Class/Date stamp at `positions`.
    If `all_pages`, every page also gets that full stamp.
    If `mark_top_pages`, pages 2+ get the student's name at (top_x, top_y).
    If `include_qr`, every page gets a QR (encoding "<class>/<name>/<page>/<total>")
    at (qr_x, qr_y) — the page suffix lets downstream scan tools sort and detect
    missing pages.
    """
    reader = PdfReader(str(source))
    writer = PdfWriter()
    qr_prefix = f"{class_label}/{name}" if class_label else name
    qr_on = include_qr and HAS_QR and bool(qr_prefix)
    total_pages = len(reader.pages)

    for idx, page in enumerate(reader.pages):
        mb = page.mediabox
        w = float(mb.width)
        h = float(mb.height)

        text_items: list[tuple[str, float, float]] = []
        qr_items: list[tuple[str, float, float, float]] = []

        if qr_on:
            qr_data = f"{qr_prefix}/{idx + 1}/{total_pages}"
            qr_items.append((qr_data, qr_x, qr_y, qr_size))

        if idx == 0 or all_pages:
            text_items = [
                (name,         positions["name_x"],  positions["name_y"]),
                (class_label,  positions["class_x"], positions["class_y"]),
                (date_text,    positions["date_x"],  positions["date_y"]),
            ]
        elif mark_top_pages:
            text_items = [(name, top_x, top_y)]

        has_text = any(t for t, _, _ in text_items)
        if has_text or qr_items:
            overlay_bytes = make_overlay_pdf(w, h, text_items, font_size, qr_items)
            overlay_page = PdfReader(io.BytesIO(overlay_bytes)).pages[0]
            page.merge_page(overlay_page)

        writer.add_page(page)

    if hasattr(output, "write"):
        writer.write(output)
    else:
        with open(output, "wb") as f:
            writer.write(f)


def send_to_printer(pdf_path: Path, printer: str, monochrome: bool) -> None:
    """Spawn SumatraPDF to send a print job (silent, queued)."""
    if not SUMATRA_PATH.exists():
        raise FileNotFoundError(f"SumatraPDF not found at: {SUMATRA_PATH}")
    settings = "monochrome" if monochrome else "color"
    subprocess.run(
        [str(SUMATRA_PATH),
         "-print-to", printer,
         "-print-settings", settings,
         "-silent",
         str(pdf_path)],
        check=False,
    )


def safe_filename(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9 _.-]+", "_", s).strip("_") or "student"


# --------------------------------------------------------------------------- #
# Qt helpers                                                                  #
# --------------------------------------------------------------------------- #

class LinkedSpin(QWidget):
    """A horizontal slider paired with a spinbox; both stay in sync.

    Emits `valueChanged(int)` when either control changes the value.
    The slider's range can be smaller than the spinbox's (the spinbox is
    authoritative for the real value; the slider is just a quick scrubber).
    """

    valueChanged = Signal(int)

    def __init__(self, value: int, slider_max: int, spin_max: int = 9999,
                 parent: QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)
        self.slider = QSlider(Qt.Horizontal)
        self.slider.setRange(0, slider_max)
        self.slider.setMinimumWidth(110)
        self.spin = QSpinBox()
        self.spin.setRange(0, max(spin_max, slider_max))
        self.spin.setFixedWidth(64)
        layout.addWidget(self.slider, 1)
        layout.addWidget(self.spin)
        self._syncing = False
        self.slider.valueChanged.connect(self._from_slider)
        self.spin.valueChanged.connect(self._from_spin)
        self.set_value(value)

    def _from_slider(self, v: int) -> None:
        if self._syncing:
            return
        self._syncing = True
        self.spin.setValue(v)
        self._syncing = False
        self.valueChanged.emit(v)

    def _from_spin(self, v: int) -> None:
        if self._syncing:
            return
        self._syncing = True
        self.slider.setValue(min(v, self.slider.maximum()))
        self._syncing = False
        self.valueChanged.emit(v)

    def set_value(self, v: int) -> None:
        self._syncing = True
        self.slider.setValue(min(v, self.slider.maximum()))
        self.spin.setValue(v)
        self._syncing = False

    def value(self) -> int:
        return self.spin.value()


class PreviewLabel(QLabel):
    """QLabel that re-emits its resize so the preview can be re-rasterised."""

    resized = Signal()

    def resizeEvent(self, event):  # type: ignore[override]
        super().resizeEvent(event)
        self.resized.emit()


class WorkerSignals(QObject):
    """Cross-thread channel for background workers to update the GUI."""

    status = Signal(str)
    info = Signal(str, str)     # title, message
    error = Signal(str, str)    # title, message
    busy = Signal(bool)


# --------------------------------------------------------------------------- #
# GUI                                                                         #
# --------------------------------------------------------------------------- #

class App(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("OpenName")
        if ICON_PATH.exists():
            self.setWindowIcon(QIcon(str(ICON_PATH)))
        self.resize(1280, 860)
        self.setMinimumSize(960, 700)

        self.settings = self._load_settings()
        self.student_checks: list[tuple[QCheckBox, str]] = []
        self._busy = False
        self._preview_pixmap: QPixmap | None = None

        self.signals = WorkerSignals()
        self.signals.status.connect(self._set_status)
        self.signals.info.connect(self._show_info)
        self.signals.error.connect(self._show_error)
        self.signals.busy.connect(self._set_busy)

        self._preview_timer = QTimer(self)
        self._preview_timer.setSingleShot(True)
        self._preview_timer.setInterval(70)
        self._preview_timer.timeout.connect(self._update_preview)

        self._build_ui()
        self._wire_signals()
        self._refresh_classes()
        self._refresh_students()
        QTimer.singleShot(120, self._update_preview)

    # ---- settings -------------------------------------------------------- #

    def _load_settings(self) -> dict:
        s = dict(DEFAULTS)
        if SETTINGS_PATH.exists():
            try:
                with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                    s.update(json.load(f))
            except Exception:
                pass
        if not s.get("date_text"):
            s["date_text"] = datetime.now().strftime("%d/%m/%Y")
        # Decide the initial PDF. Priority:
        #   1. QMARK_SHEET_PATH (qmark passes the session-bound worksheet)
        #   2. The previously-saved last_pdf, IF it still exists AND
        #      (when QMARK_SHEETS_DIR is set) it lives inside Sheets/
        #   3. First PDF in QMARK_SHEETS_DIR
        #   4. First PDF next to OpenName.py (standalone fallback,
        #      typically measurementsheet.pdf)
        sheet_override = os.environ.get("QMARK_SHEET_PATH", "")
        if sheet_override and Path(sheet_override).is_file():
            s["last_pdf"] = sheet_override
            return s

        sheets_dir = os.environ.get("QMARK_SHEETS_DIR", "")
        sheets_path = Path(sheets_dir) if sheets_dir else None
        last = s.get("last_pdf") or ""
        last_ok = bool(last) and Path(last).exists()
        if sheets_path and sheets_path.is_dir():
            try:
                in_sheets = last_ok and sheets_path in Path(last).resolve().parents
            except Exception:
                in_sheets = False
            if not in_sheets:
                pdfs = sorted(sheets_path.glob("*.pdf"))
                s["last_pdf"] = str(pdfs[0]) if pdfs else ""
        elif not last_ok:
            pdfs = sorted(SCRIPT_DIR.glob("*.pdf"))
            s["last_pdf"] = str(pdfs[0]) if pdfs else ""
        return s

    def _save_settings(self) -> None:
        try:
            with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
                json.dump(self.settings, f, indent=2)
        except Exception:
            pass

    # ---- layout ---------------------------------------------------------- #

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        outer_v = QVBoxLayout(central)
        outer_v.setContentsMargins(8, 8, 8, 8)

        cols = QHBoxLayout()
        cols.setSpacing(6)
        outer_v.addLayout(cols, 1)

        # ---------------- LEFT: controls ---------------- #
        left = QWidget()
        left_v = QVBoxLayout(left)
        left_v.setContentsMargins(0, 0, 0, 0)
        left_v.setSpacing(4)
        cols.addWidget(left, 3)

        # PDF picker
        pdf_row = QHBoxLayout()
        pdf_label = QLabel("PDF:")
        pdf_label.setFixedWidth(64)
        self.pdf_edit = QLineEdit(self.settings["last_pdf"])
        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self._browse_pdf)
        pdf_row.addWidget(pdf_label)
        pdf_row.addWidget(self.pdf_edit, 1)
        pdf_row.addWidget(browse_btn)
        left_v.addLayout(pdf_row)

        # Class + class label
        cls_row = QHBoxLayout()
        cls_label = QLabel("Class:")
        cls_label.setFixedWidth(64)
        self.class_combo = QComboBox()
        self.class_combo.setMinimumWidth(160)
        self.class_combo.currentTextChanged.connect(self._on_class_change)
        clbl_label = QLabel("Class label:")
        self.classlabel_edit = QLineEdit(self.settings["class_label"])
        cls_row.addWidget(cls_label)
        cls_row.addWidget(self.class_combo)
        cls_row.addSpacing(12)
        cls_row.addWidget(clbl_label)
        cls_row.addWidget(self.classlabel_edit, 1)
        left_v.addLayout(cls_row)

        # Students
        students_box = QGroupBox("Students  (uncheck to skip)")
        sb_v = QVBoxLayout(students_box)
        toolbar = QHBoxLayout()
        all_btn = QPushButton("All")
        all_btn.setFixedWidth(60)
        all_btn.clicked.connect(lambda: self._set_all_students(True))
        none_btn = QPushButton("None")
        none_btn.setFixedWidth(60)
        none_btn.clicked.connect(lambda: self._set_all_students(False))
        self.count_label = QLabel("")
        self.count_label.setStyleSheet("color: #555;")
        toolbar.addWidget(all_btn)
        toolbar.addWidget(none_btn)
        toolbar.addSpacing(16)
        toolbar.addWidget(self.count_label)
        toolbar.addStretch(1)
        sb_v.addLayout(toolbar)

        self.students_scroll = QScrollArea()
        self.students_scroll.setWidgetResizable(True)
        self.students_scroll.setFrameShape(QFrame.StyledPanel)
        self.students_inner = QWidget()
        self.students_grid = QGridLayout(self.students_inner)
        self.students_grid.setContentsMargins(8, 6, 8, 6)
        self.students_grid.setHorizontalSpacing(8)
        self.students_grid.setVerticalSpacing(2)
        self.students_scroll.setWidget(self.students_inner)
        sb_v.addWidget(self.students_scroll, 1)
        left_v.addWidget(students_box, 1)

        # Stamp positions (with sliders)
        pos_box = QGroupBox("Stamp positions (points; drag sliders to move)")
        pos_grid = QGridLayout(pos_box)
        pos_grid.setHorizontalSpacing(6)
        pos_grid.setVerticalSpacing(4)
        self.name_x, self.name_y   = self._make_pos_row(pos_grid, 0, "Name field:",      int(self.settings["name_x"]),  int(self.settings["name_y"]))
        self.class_x, self.class_y = self._make_pos_row(pos_grid, 1, "Class field:",     int(self.settings["class_x"]), int(self.settings["class_y"]))
        self.date_x, self.date_y   = self._make_pos_row(pos_grid, 2, "Date field:",      int(self.settings["date_x"]),  int(self.settings["date_y"]))
        self.top_x, self.top_y     = self._make_pos_row(pos_grid, 3, "Top of pages 2+:", int(self.settings["top_x"]),   int(self.settings["top_y"]))
        self.qr_x, self.qr_y       = self._make_pos_row(pos_grid, 4, "QR code:",         int(self.settings["qr_x"]),    int(self.settings["qr_y"]))
        left_v.addWidget(pos_box)

        # Multi-page options
        self.allpages_cb = QCheckBox("Stamp full Name+Class+Date on every page")
        self.allpages_cb.setChecked(bool(self.settings["all_pages"]))
        self.marktop_cb = QCheckBox("Mark name at top of pages 2+")
        self.marktop_cb.setChecked(bool(self.settings["mark_top_pages"]))
        self.includeqr_cb = QCheckBox("Include QR code (every page)")
        self.includeqr_cb.setChecked(bool(self.settings["include_qr"]) and HAS_QR)
        if not HAS_QR:
            self.includeqr_cb.setEnabled(False)
        opts_v = QVBoxLayout()
        opts_v.setSpacing(2)
        opts_v.addWidget(self.allpages_cb)
        opts_v.addWidget(self.marktop_cb)
        opts_v.addWidget(self.includeqr_cb)
        left_v.addLayout(opts_v)

        # Font / date
        fd_row = QHBoxLayout()
        fd_row.addWidget(QLabel("Font size:"))
        self.font_spin = QSpinBox()
        self.font_spin.setRange(4, 72)
        self.font_spin.setValue(int(self.settings["font_size"]))
        self.font_spin.setFixedWidth(60)
        fd_row.addWidget(self.font_spin)
        fd_row.addSpacing(16)
        fd_row.addWidget(QLabel("Date text:"))
        self.date_edit = QLineEdit(self.settings["date_text"])
        self.date_edit.setFixedWidth(120)
        fd_row.addWidget(self.date_edit)
        fd_row.addSpacing(8)
        self.include_date_cb = QCheckBox("Include date")
        self.include_date_cb.setChecked(bool(self.settings["include_date"]))
        self.include_date_cb.toggled.connect(self._update_date_state)
        fd_row.addWidget(self.include_date_cb)
        fd_row.addStretch(1)
        left_v.addLayout(fd_row)

        # Printer
        pr_row = QHBoxLayout()
        pr_label = QLabel("Printer:")
        pr_label.setFixedWidth(64)
        self.printer_edit = QLineEdit(self.settings["printer"])
        self.mono_cb = QCheckBox("Monochrome")
        self.mono_cb.setChecked(bool(self.settings["monochrome"]))
        pr_row.addWidget(pr_label)
        pr_row.addWidget(self.printer_edit, 1)
        pr_row.addSpacing(8)
        pr_row.addWidget(self.mono_cb)
        left_v.addLayout(pr_row)

        # ---------------- RIGHT: live preview ---------------- #
        right_box = QGroupBox("Live preview (page 1)")
        right_v = QVBoxLayout(right_box)
        self.preview_label = PreviewLabel()
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setStyleSheet("background-color: #bfbfbf;")
        self.preview_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.preview_label.setMinimumSize(200, 200)
        self.preview_label.resized.connect(self._on_preview_resize)
        if not HAS_PREVIEW:
            self.preview_label.setText(
                "Live preview disabled.\n"
                "Install PyMuPDF to enable:\n  pip install pymupdf"
            )
            self.preview_label.setStyleSheet("background-color: #bfbfbf; color: #444; padding: 12px;")
        right_v.addWidget(self.preview_label, 1)
        cols.addWidget(right_box, 4)

        # ---------------- BOTTOM: status + actions ---------------- #
        bottom = QHBoxLayout()
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #234;")
        bottom.addWidget(self.status_label, 1)
        self.preview_btn = QPushButton("Open in viewer")
        self.preview_btn.clicked.connect(self._on_preview)
        bottom.addWidget(self.preview_btn)
        self.export_btn = QPushButton("Export combined PDF")
        self.export_btn.clicked.connect(self._on_export)
        bottom.addWidget(self.export_btn)
        self.print_btn = QPushButton("Print Selected")
        self.print_btn.clicked.connect(self._on_print)
        bottom.addWidget(self.print_btn)
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        bottom.addWidget(close_btn)
        outer_v.addLayout(bottom)

        self._update_date_state(self.include_date_cb.isChecked())

    def _make_pos_row(self, grid: QGridLayout, row: int, label: str,
                      x_value: int, y_value: int) -> tuple[LinkedSpin, LinkedSpin]:
        grid.addWidget(QLabel(label), row, 0)
        grid.addWidget(QLabel("X"), row, 1)
        x_widget = LinkedSpin(x_value, slider_max=600)
        grid.addWidget(x_widget, row, 2)
        grid.addWidget(QLabel("Y"), row, 3)
        y_widget = LinkedSpin(y_value, slider_max=850)
        grid.addWidget(y_widget, row, 4)
        return x_widget, y_widget

    # ---- signal wiring -------------------------------------------------- #

    def _wire_signals(self) -> None:
        """Schedule a preview redraw whenever any input changes."""
        self.pdf_edit.textChanged.connect(self._schedule_preview)
        self.classlabel_edit.textChanged.connect(self._schedule_preview)
        for w in (self.name_x, self.name_y, self.class_x, self.class_y,
                  self.date_x, self.date_y, self.top_x, self.top_y,
                  self.qr_x, self.qr_y):
            w.valueChanged.connect(self._schedule_preview)
        self.font_spin.valueChanged.connect(self._schedule_preview)
        self.date_edit.textChanged.connect(self._schedule_preview)
        self.include_date_cb.toggled.connect(self._schedule_preview)
        self.allpages_cb.toggled.connect(self._schedule_preview)
        self.marktop_cb.toggled.connect(self._schedule_preview)
        self.includeqr_cb.toggled.connect(self._schedule_preview)

    def _update_date_state(self, checked: bool) -> None:
        """Grey out the date entry when 'Include date' is unchecked."""
        self.date_edit.setEnabled(checked)

    # ---- live preview --------------------------------------------------- #

    def _schedule_preview(self, *_args) -> None:
        if not HAS_PREVIEW:
            return
        self._preview_timer.start()

    def _on_preview_resize(self) -> None:
        self._schedule_preview()

    def _update_preview(self) -> None:
        if not HAS_PREVIEW:
            return
        try:
            state = self._read_form_state()
            pdf_in = Path(state["last_pdf"])
            if not pdf_in.exists():
                return
            sel = self._selected_students()
            name = sel[0] if sel else "Sample Name"

            buf = io.BytesIO()
            stamp_pdf(
                source=pdf_in, output=buf,
                name=name, class_label=state["class_label"],
                date_text=state["date_text"],
                positions=state, font_size=state["font_size"],
                all_pages=state["all_pages"],
                mark_top_pages=state["mark_top_pages"],
                top_x=state["top_x"], top_y=state["top_y"],
                include_qr=state["include_qr"], qr_size=state["qr_size"],
                qr_x=state["qr_x"], qr_y=state["qr_y"],
            )

            doc = fitz.open(stream=buf.getvalue(), filetype="pdf")
            try:
                page = doc[0]
                cw = max(1, self.preview_label.width() - 8)
                ch = max(1, self.preview_label.height() - 8)
                if cw < 50 or ch < 50:
                    return
                scale = min(cw / page.rect.width, ch / page.rect.height)
                pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                # Copy the bytes — the underlying buffer is owned by `pix`.
                qimg = QImage(
                    bytes(pix.samples), pix.width, pix.height, pix.stride,
                    QImage.Format_RGB888,
                )
                self._preview_pixmap = QPixmap.fromImage(qimg)
                self.preview_label.setPixmap(self._preview_pixmap)
            finally:
                doc.close()
        except Exception as e:
            self._set_status(f"Preview error: {e}")

    # ---- class / student data ------------------------------------------- #

    def _refresh_classes(self) -> None:
        # When qmark passes a class roster, that's the only option — qmark
        # owns the roster in this mode. Otherwise glob the local Classes/.
        override = _qmark_class_override()
        if override is not None:
            names = [override.stem]
        elif CLASSES_DIR.exists():
            names = [f.stem for f in sorted(CLASSES_DIR.glob("*.xlsx"))]
        else:
            self.class_combo.clear()
            self._set_status(f"Classes folder not found: {CLASSES_DIR}")
            return
        self.class_combo.blockSignals(True)
        self.class_combo.clear()
        self.class_combo.addItems(names)
        last = self.settings.get("last_class")
        if last in names:
            self.class_combo.setCurrentText(last)
        elif names:
            self.class_combo.setCurrentIndex(0)
        self.class_combo.blockSignals(False)

    def _on_class_change(self, cls: str) -> None:
        # Auto-fill the class label on the PDF if the user hasn't customised it.
        current = self.classlabel_edit.text().strip()
        if not current or current == self.settings.get("last_class", ""):
            self.classlabel_edit.setText(cls)
        self.settings["last_class"] = cls
        self._refresh_students()

    def _refresh_students(self) -> None:
        # Clear existing widgets
        while self.students_grid.count():
            item = self.students_grid.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()
        self.student_checks.clear()

        cls = self.class_combo.currentText()
        if not cls:
            self.count_label.setText("")
            return
        xlsx = _class_xlsx_for(cls)
        if not xlsx.exists():
            self.count_label.setText("")
            self._set_status(f"Class file not found: {xlsx}")
            return
        try:
            names = load_students(xlsx)
        except Exception as e:
            self.count_label.setText("")
            self._set_status(f"Error reading {xlsx.name}: {e}")
            return

        if not self.classlabel_edit.text().strip():
            self.classlabel_edit.setText(cls)

        cols = 3
        for c in range(cols):
            self.students_grid.setColumnStretch(c, 1)
        for idx, name in enumerate(names):
            cb = QCheckBox(name)
            cb.setChecked(True)
            cb.toggled.connect(self._schedule_preview)
            self.students_grid.addWidget(cb, idx // cols, idx % cols)
            self.student_checks.append((cb, name))

        self.count_label.setText(f"{len(names)} students loaded")
        self._set_status("")
        self._schedule_preview()

    def _set_all_students(self, val: bool) -> None:
        for cb, _ in self.student_checks:
            cb.setChecked(val)

    # ---- buttons --------------------------------------------------------- #

    def _browse_pdf(self) -> None:
        # Default to QMARK_SHEETS_DIR when launched from the qmark dashboard,
        # else fall back to the parent of the last-used PDF, else SCRIPT_DIR.
        sheets_dir = os.environ.get("QMARK_SHEETS_DIR", "")
        if sheets_dir and Path(sheets_dir).is_dir():
            initialdir = sheets_dir
        elif self.pdf_edit.text() and Path(self.pdf_edit.text()).exists():
            initialdir = str(Path(self.pdf_edit.text()).parent)
        else:
            initialdir = str(SCRIPT_DIR)
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select PDF to stamp",
            initialdir,
            "PDF files (*.pdf);;All files (*.*)",
        )
        if path:
            self.pdf_edit.setText(path)

    def _read_form_state(self) -> dict:
        include_date = self.include_date_cb.isChecked()
        return {
            "last_pdf":    self.pdf_edit.text(),
            "last_class":  self.class_combo.currentText(),
            "class_label": self.classlabel_edit.text(),
            "name_x":  float(self.name_x.value()),
            "name_y":  float(self.name_y.value()),
            "class_x": float(self.class_x.value()),
            "class_y": float(self.class_y.value()),
            "date_x":  float(self.date_x.value()),
            "date_y":  float(self.date_y.value()),
            "top_x":   float(self.top_x.value()),
            "top_y":   float(self.top_y.value()),
            "qr_x":    float(self.qr_x.value()),
            "qr_y":    float(self.qr_y.value()),
            "font_size": float(self.font_spin.value()),
            "date_text": self.date_edit.text() if include_date else "",
            "include_date": include_date,
            "printer":   self.printer_edit.text(),
            "monochrome": self.mono_cb.isChecked(),
            "all_pages":  self.allpages_cb.isChecked(),
            "mark_top_pages": self.marktop_cb.isChecked(),
            "include_qr": self.includeqr_cb.isChecked() and HAS_QR,
            "qr_size":   float(DEFAULTS["qr_size"]),
        }

    def _selected_students(self) -> list[str]:
        return [n for cb, n in self.student_checks if cb.isChecked()]

    def _on_preview(self) -> None:
        try:
            state = self._read_form_state()
            pdf_in = Path(state["last_pdf"])
            if not pdf_in.exists():
                raise FileNotFoundError(f"PDF not found: {pdf_in}")
            sel = self._selected_students()
            if not sel:
                raise RuntimeError("No student selected.")
            tmp = Path(tempfile.gettempdir()) / f"preview_{safe_filename(sel[0])}_{os.getpid()}.pdf"
            stamp_pdf(
                source=pdf_in,
                output=tmp,
                name=sel[0],
                class_label=state["class_label"],
                date_text=state["date_text"],
                positions=state,
                font_size=state["font_size"],
                all_pages=state["all_pages"],
                mark_top_pages=state["mark_top_pages"],
                top_x=state["top_x"],
                top_y=state["top_y"],
                include_qr=state["include_qr"],
                qr_size=state["qr_size"],
                qr_x=state["qr_x"],
                qr_y=state["qr_y"],
            )
            os.startfile(tmp)  # type: ignore[attr-defined]
            self._set_status(f"Preview opened for: {sel[0]}")
        except Exception as e:
            QMessageBox.critical(self, "Preview error", str(e))

    def _on_export(self) -> None:
        """Stamp every selected student and concatenate into a single PDF.

        For testing: simulates a scanned bundle so the downstream scanner code
        can be exercised end-to-end.
        """
        if self._busy:
            return
        try:
            state = self._read_form_state()
            pdf_in = Path(state["last_pdf"])
            if not pdf_in.exists():
                raise FileNotFoundError(f"PDF not found: {pdf_in}")
            sel = self._selected_students()
            if not sel:
                raise RuntimeError("No students selected.")
        except Exception as e:
            QMessageBox.critical(self, "Export error", str(e))
            return

        cls = state["class_label"] or state["last_class"] or "combined"
        default_name = f"{safe_filename(cls)}_combined.pdf"
        # Default the save dialog to a writable location, not SCRIPT_DIR
        # (read-only under MSIX). Prefer the qmark-supplied Sheets dir
        # when launched from the dashboard so exports land alongside the
        # source worksheet.
        sheets_dir = os.environ.get("QMARK_SHEETS_DIR", "")
        if sheets_dir and Path(sheets_dir).is_dir():
            initial_dir = Path(sheets_dir)
        else:
            initial_dir = _writable_data_dir()
        out_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save combined PDF",
            str(initial_dir / default_name),
            "PDF files (*.pdf)",
        )
        if not out_path:
            return

        self._set_busy(True)
        threading.Thread(
            target=self._export_worker, args=(sel, state, pdf_in, Path(out_path)),
            daemon=True,
        ).start()

    def _export_worker(self, students: list[str], state: dict,
                       pdf_in: Path, out_path: Path) -> None:
        try:
            writer = PdfWriter()
            for i, name in enumerate(students, 1):
                self.signals.status.emit(f"Stamping {i} of {len(students)}: {name}")
                buf = io.BytesIO()
                stamp_pdf(
                    source=pdf_in,
                    output=buf,
                    name=name,
                    class_label=state["class_label"],
                    date_text=state["date_text"],
                    positions=state,
                    font_size=state["font_size"],
                    all_pages=state["all_pages"],
                    mark_top_pages=state["mark_top_pages"],
                    top_x=state["top_x"],
                    top_y=state["top_y"],
                    include_qr=state["include_qr"],
                    qr_size=state["qr_size"],
                    qr_x=state["qr_x"],
                    qr_y=state["qr_y"],
                )
                buf.seek(0)
                reader = PdfReader(buf)
                for page in reader.pages:
                    writer.add_page(page)
            self.signals.status.emit(f"Writing {out_path.name} ...")
            with open(out_path, "wb") as f:
                writer.write(f)
            self.signals.status.emit(f"Exported {len(students)} student(s) to {out_path.name}")
            self.signals.info.emit("Export complete", f"Wrote combined PDF:\n{out_path}")
        except Exception as e:
            self.signals.status.emit(f"Error: {e}")
            self.signals.error.emit("Export error", str(e))
        finally:
            self.signals.busy.emit(False)

    def _on_print(self) -> None:
        if self._busy:
            return
        try:
            state = self._read_form_state()
            pdf_in = Path(state["last_pdf"])
            if not pdf_in.exists():
                raise FileNotFoundError(f"PDF not found: {pdf_in}")
            sel = self._selected_students()
            if not sel:
                raise RuntimeError("No students selected.")
            if not state["printer"].strip():
                raise RuntimeError("No printer configured.")
        except Exception as e:
            QMessageBox.critical(self, "Print error", str(e))
            return

        reply = QMessageBox.question(
            self,
            "Confirm print",
            f"About to send {len(sel)} print job(s) "
            f"(one per student) to:\n{state['printer']}\n\nProceed?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        # Persist current state for next launch.
        self.settings.update({k: v for k, v in state.items() if k in DEFAULTS})
        self._save_settings()

        self._set_busy(True)
        threading.Thread(
            target=self._print_worker, args=(sel, state, pdf_in), daemon=True
        ).start()

    def _print_worker(self, students: list[str], state: dict, pdf_in: Path) -> None:
        tmp_dir = Path(tempfile.mkdtemp(prefix="OpenName_"))
        try:
            for i, name in enumerate(students, 1):
                self.signals.status.emit(f"Printing {i} of {len(students)}: {name}")
                out_pdf = tmp_dir / f"{i:03d}_{safe_filename(name)}.pdf"
                stamp_pdf(
                    source=pdf_in,
                    output=out_pdf,
                    name=name,
                    class_label=state["class_label"],
                    date_text=state["date_text"],
                    positions=state,
                    font_size=state["font_size"],
                    all_pages=state["all_pages"],
                    mark_top_pages=state["mark_top_pages"],
                    top_x=state["top_x"],
                    top_y=state["top_y"],
                    include_qr=state["include_qr"],
                    qr_size=state["qr_size"],
                    qr_x=state["qr_x"],
                    qr_y=state["qr_y"],
                )
                send_to_printer(out_pdf, state["printer"], state["monochrome"])
                time.sleep(3)
            self.signals.status.emit(f"Done. Sent {len(students)} print job(s).")
            self.signals.info.emit(
                "Done", f"Sent {len(students)} print job(s) to {state['printer']}."
            )
        except Exception as e:
            self.signals.status.emit(f"Error: {e}")
            self.signals.error.emit("Print error", str(e))
        finally:
            try:
                for p in tmp_dir.glob("*"):
                    p.unlink(missing_ok=True)
                tmp_dir.rmdir()
            except Exception:
                pass
            self.signals.busy.emit(False)

    # ---- thread-safe UI helpers (slots) --------------------------------- #

    def _set_status(self, msg: str) -> None:
        self.status_label.setText(msg)

    def _show_info(self, title: str, msg: str) -> None:
        QMessageBox.information(self, title, msg)

    def _show_error(self, title: str, msg: str) -> None:
        QMessageBox.critical(self, title, msg)

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        for btn in (self.print_btn, self.preview_btn, self.export_btn):
            btn.setEnabled(not busy)


def main() -> None:
    if not SUMATRA_PATH.exists():
        # Show a warning but still let the GUI start (preview still works).
        print(f"Warning: SumatraPDF not found at {SUMATRA_PATH}. "
              f"Preview will work but printing will fail.", file=sys.stderr)
    # Without an explicit AppUserModelID, Windows groups us under the python.exe
    # icon in the taskbar instead of using the window icon we set below.
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("OpenName")
        except Exception:
            pass
    app = QApplication(sys.argv)
    if ICON_PATH.exists():
        app.setWindowIcon(QIcon(str(ICON_PATH)))
    win = App()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
