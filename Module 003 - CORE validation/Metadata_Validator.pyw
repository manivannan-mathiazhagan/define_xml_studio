#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# ====================================================================================================
# Script Name    : core_validation_gui_app.pyw
#
# Module Name    : CDISC CORE Validation GUI
#
# Author         : Manivannan Mathialagan
#
# Purpose        :
#                  Desktop GUI utility to run CDISC CORE validation on selected XPT datasets only.
#
# Description    :
#                  • Reads CDISC API key from CDISC_API_KEY.json placed next to this script.
#                  • Lets user choose standard, IG version, CT package(s), XPT folder, and optional define.xml.
#                  • For SDTM validation, one SDTM CT package is selected and passed to CORE.
#                  • For ADaM validation, ADaM CT is selected for CORE validation and SDTM CT is shown for
#                    traceability/reference because ADaM define/specs may rely on SDTM CT concepts.
#                  • Uses a temporary isolated validation folder during each run.
#                  • Copies only top-level *.xpt files from the selected folder into the temp folder.
#                  • Copies define.xml only when selected by the user.
#                  • Runs CORE validation against the temp folder, preventing subfolder/old-study contamination.
#                  • Generates XLSX report and copies it back to the originally selected XPT folder.
#                  • Does not open CT browser automatically.
#                  • Opens report only when user clicks "Open Last Report".
#
# API Key File   :
#                  CDISC_API_KEY.json must be available in the same folder as this script.
#
#                  Example:
#                  {
#                      "primary_key": "YOUR_PRIMARY_API_KEY",
#                      "secondary_key": "YOUR_SECONDARY_API_KEY"
#                  }
#
# CORE Location  :
#                  Update CORE_DIR below if your cdisc-rules-engine folder is different.
#
# Important      :
#                  CORE command standard names used here:
#                      SDTM : sdtmig
#                      ADaM : adam
#
#                  The selected XPT folder is not passed directly to CORE validation.
#                  A temp folder is created and only immediate *.xpt files are copied.
#
# Run            :
#                  python core_validation_gui_app.pyw
#
# ====================================================================================================

import os
import sys
import json
import re
import shutil
import tempfile
import subprocess
import importlib
from pathlib import Path
from datetime import datetime

from PyQt5 import QtCore, QtGui, QtWidgets


# ====================================================================================================
# ENVIRONMENT
# ====================================================================================================

os.environ.setdefault("PYTHONUTF8", "1")
os.environ.setdefault("QT_AUTO_SCREEN_SCALE_FACTOR", "0")
os.environ.setdefault("QT_SCALE_FACTOR", "1")
os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "0")

try:
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_DisableHighDpiScaling, True)
except Exception:
    pass


# ====================================================================================================
# CONFIG
# ====================================================================================================

SCRIPT_DIR = Path(__file__).resolve().parent
KEY_FILE = SCRIPT_DIR / "CDISC_API_KEY.json"

# Update this path if needed.
CORE_DIR = r"P:\BSP_LocalDev\Manivannan.Mathialag\zzzz_My_SAS_Files\My GitHub\cdisc-rules-engine"
CORE_PY = os.path.join(CORE_DIR, "core.py")
CORE_PYTHON = sys.executable

APP_TITLE = "CDISC CORE Validation GUI"

THEME = {
    "app_bg": "#f3f7fd",
    "header": "#cfe7ff",
    "header_text": "#103760",
    "subtitle": "#315f8f",
    "border": "#b7cde8",
    "status": "#fff8dc",
    "button": "#4d8ef7",
    "button_hover": "#2f75dd",
    "danger": "#d9534f",
    "panel": "#ffffff",
}


# ====================================================================================================
# HELPERS
# ====================================================================================================

def safe_text(value):
    if value is None:
        return ""
    return str(value).strip()


def load_api_keys():
    if not KEY_FILE.exists():
        return "", ""

    try:
        data = json.loads(KEY_FILE.read_text(encoding="utf-8"))
        primary = data.get("primary_key") or data.get("primary") or data.get("api_key") or ""
        secondary = data.get("secondary_key") or data.get("secondary") or ""
        return safe_text(primary), safe_text(secondary)
    except Exception:
        return "", ""


def standard_to_core_standard(display_standard):
    txt = safe_text(display_standard).upper()
    if txt.startswith("SDTM"):
        return "sdtmig"
    if txt.startswith("ADAM") or txt.startswith("ADA"):
        return "adam"
    return "sdtmig"


def standard_to_ct_prefix(core_standard):
    std = safe_text(core_standard).lower()
    if std == "sdtmig":
        return "sdtmct"
    if std == "adam":
        return "adamct"
    return ""


def default_versions(core_standard):
    std = safe_text(core_standard).lower()
    if std == "sdtmig":
        return ["3-4", "3-3", "3-2", "3-1-3", "3-1-2"]
    if std == "adam":
        return ["1-3", "1-2", "1-1", "1-0"]
    return []


def detect_define_xml(xpt_folder):
    if not xpt_folder or not os.path.exists(xpt_folder):
        return ""

    folder = Path(xpt_folder)

    preferred = [
        folder / "define.xml",
        folder / "define2-0-0.xml",
        folder / "define2-1-0.xml",
    ]

    for candidate in preferred:
        if candidate.exists():
            return str(candidate)

    for candidate in folder.glob("define*.xml"):
        if candidate.is_file():
            return str(candidate)

    return ""


def open_file(path):
    try:
        if path and os.path.exists(path):
            os.startfile(path)
    except Exception:
        pass


def quote_cmd(cmd):
    out = []
    for item in cmd:
        item = str(item)
        if " " in item or "\\" in item or "/" in item:
            out.append(f'"{item}"')
        else:
            out.append(item)
    return " ".join(out)


def parse_ct_date(package_name):
    txt = safe_text(package_name)
    import re
    m = re.search(r"(20\d{2}-\d{2}-\d{2})", txt)
    return m.group(1) if m else ""


def newest_first(packages):
    return sorted(packages, key=lambda x: parse_ct_date(x) or x, reverse=True)



def import_pyreadstat_optional():
    """Import pyreadstat only when CORE-safe XPT rewrite is requested."""
    try:
        return importlib.import_module("pyreadstat")
    except Exception:
        return None


def read_xpt_with_encoding_fallback(path, log_func=None):
    """Read XPT using the same practical fallback approach used by Define XML Generator.

    CORE can fail on XPT metadata containing bytes such as 0xA0 because its builder
    may read with UTF-8 only. This routine tries common SAS/Windows encodings.
    """
    pyreadstat = import_pyreadstat_optional()
    if pyreadstat is None:
        raise RuntimeError("pyreadstat is not installed; cannot create CORE-safe XPT copy.")

    encodings_to_try = [None, "windows-1252", "latin1", "iso-8859-1"]
    errors = []

    for enc in encodings_to_try:
        try:
            kwargs = {}
            if enc is not None:
                kwargs["encoding"] = enc
            df, meta = pyreadstat.read_xport(str(path), **kwargs)
            return df, meta, enc or "default"
        except Exception as e:
            errors.append(f"encoding={enc or 'default'}: {e}")
            msg = str(e).lower()
            retryable = (
                isinstance(e, UnicodeDecodeError)
                or "codec can't decode" in msg
                or "invalid start byte" in msg
                or "utf-8" in msg
                or "utf8" in msg
            )
            if not retryable and enc is None:
                # Not an encoding problem; still try fallback once because CORE commonly fails
                # with metadata encoding even when the first exception text is not clear.
                continue
            continue

    raise RuntimeError(
        "Unable to read XPT using encoding fallback.\n\n"
        f"File: {path}\n\n"
        + "\n".join(errors)
    )


def write_core_safe_xpt(df, meta, output_path, table_name):
    """Write a fresh XPT copy that CORE can read.

    We preserve variable labels where pyreadstat supports column_labels. If the
    installed pyreadstat has an older signature, we fall back to a simpler write.
    """
    pyreadstat = import_pyreadstat_optional()
    if pyreadstat is None:
        raise RuntimeError("pyreadstat is not installed; cannot write CORE-safe XPT copy.")

    table_name = re.sub(r"[^A-Za-z0-9_]", "", safe_text(table_name).upper())[:8] or "DATASET"

    column_labels = None
    try:
        if hasattr(meta, "column_names") and hasattr(meta, "column_labels"):
            column_labels = {
                name: (label or name)
                for name, label in zip(meta.column_names, meta.column_labels)
            }
    except Exception:
        column_labels = None

    file_label = ""
    try:
        file_label = safe_text(getattr(meta, "file_label", ""))
    except Exception:
        file_label = ""

    attempts = []

    # Newer pyreadstat versions.
    kwargs = {"table_name": table_name, "file_format_version": 5}
    if file_label:
        kwargs["file_label"] = file_label[:40]
    if column_labels:
        kwargs["column_labels"] = column_labels
    attempts.append(kwargs)

    # Some versions expect column_labels as list.
    if column_labels:
        labels_list = [column_labels.get(c, c) for c in df.columns]
        kwargs2 = {"table_name": table_name, "column_labels": labels_list, "file_format_version": 5}
        if file_label:
            kwargs2["file_label"] = file_label[:40]
        attempts.append(kwargs2)

    # Minimal fallback for older pyreadstat. If this fallback is used, the written file may not be v5.
    # CORE requires XPORT v5, so prefer the attempts with file_format_version=5 above.
    attempts.append({"table_name": table_name, "file_format_version": 5})
    attempts.append({})

    last_err = None
    for kwargs in attempts:
        try:
            pyreadstat.write_xport(df, str(output_path), **kwargs)
            return
        except TypeError as e:
            last_err = e
            continue
        except Exception as e:
            last_err = e
            continue

    raise RuntimeError(f"Unable to write CORE-safe XPT copy: {last_err}")


def copy_xpt_for_core(src, dst, normalize_xpt=True, log_func=None):
    """Copy XPT into temp folder; optionally rewrite it with encoding fallback.

    If rewrite fails, the original file is copied so the user still gets a clear
    CORE error rather than the GUI stopping without a validation command.
    """
    if not normalize_xpt:
        shutil.copy2(src, dst)
        return "copied"

    try:
        df, meta, encoding_used = read_xpt_with_encoding_fallback(src, log_func=log_func)

        # If default read works, still write a fresh copy. This standardizes text
        # metadata and avoids CORE's stricter decode path.
        write_core_safe_xpt(df, meta, dst, table_name=Path(src).stem)

        if log_func:
            log_func(f"CORE-safe XPT created: {Path(src).name}  [read encoding: {encoding_used}]")
        return f"rewritten:{encoding_used}"

    except Exception as e:
        shutil.copy2(src, dst)
        if log_func:
            log_func(f"WARNING: Could not rewrite {Path(src).name}; copied original. Reason: {e}")
        return "copied_original_after_rewrite_failure"



def create_isolated_validation_folder(source_folder, define_xml_path="", include_define=False, normalize_xpt=True, log_func=None):
    """Create temp folder containing only top-level *.xpt files and optional define.xml.

    This prevents CORE from reading unrelated files or subfolders from the original location.
    """
    source = Path(source_folder)
    if not source.exists():
        raise FileNotFoundError(f"XPT folder does not exist: {source}")

    xpt_files = sorted([p for p in source.glob("*.xpt") if p.is_file()])
    if not xpt_files:
        raise FileNotFoundError(f"No top-level .xpt files found in: {source}")

    temp_dir = Path(tempfile.mkdtemp(prefix="core_validate_xpt_only_"))

    if log_func:
        log_func(f"Temporary validation folder created: {temp_dir}")
        log_func(f"Copying {len(xpt_files)} top-level XPT file(s) only...")
        if normalize_xpt:
            log_func("CORE-safe XPT rewrite is enabled for encoding fallback.")

    for src in xpt_files:
        copy_xpt_for_core(src, temp_dir / src.name, normalize_xpt=normalize_xpt, log_func=log_func)

    temp_define = ""
    if include_define and define_xml_path and os.path.exists(define_xml_path):
        temp_define = str(temp_dir / "define.xml")
        shutil.copy2(define_xml_path, temp_define)
        if log_func:
            log_func(f"Copied define.xml into temp folder: {temp_define}")
    elif log_func:
        log_func("Define.xml not copied to temp folder.")

    return str(temp_dir), temp_define, len(xpt_files)


# ====================================================================================================
# MAIN GUI
# ====================================================================================================

class CoreValidationGui(QtWidgets.QWidget):

    def __init__(self):
        super().__init__()

        self.primary_key, self.secondary_key = load_api_keys()
        self.xpt_folder = ""
        self.define_xml_path = ""
        self.core_report_path = ""
        self.all_ct_packages = []

        self.setWindowTitle(APP_TITLE)
        self.resize(1400, 850)
        self.setMinimumSize(1180, 720)

        self.build_ui()
        self.apply_style()
        self.refresh_api_status()

    # --------------------------------------------------------------------------------------------
    # UI
    # --------------------------------------------------------------------------------------------

    def build_ui(self):
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(10, 8, 10, 8)
        root.setSpacing(8)

        header = QtWidgets.QFrame()
        header.setObjectName("Header")
        h = QtWidgets.QHBoxLayout(header)
        h.setContentsMargins(14, 10, 14, 10)

        title = QtWidgets.QLabel("CDISC CORE Validation GUI")
        title.setObjectName("Title")

        subtitle = QtWidgets.QLabel("Validate selected top-level XPT files only")
        subtitle.setObjectName("Subtitle")

        h.addWidget(title)
        h.addSpacing(18)
        h.addWidget(subtitle)
        h.addStretch()

        root.addWidget(header)

        controls = QtWidgets.QFrame()
        controls.setObjectName("Controls")
        g = QtWidgets.QGridLayout(controls)
        g.setContentsMargins(12, 10, 12, 10)
        g.setHorizontalSpacing(10)
        g.setVerticalSpacing(8)

        self.api_status = QtWidgets.QLabel("API Key: -")
        self.api_status.setObjectName("InfoLabel")

        self.core_path_status = QtWidgets.QLabel(f"CORE: {CORE_PY}")
        self.core_path_status.setObjectName("InfoLabel")

        self.standard_combo = QtWidgets.QComboBox()
        self.standard_combo.addItems(["SDTM (sdtmig)", "ADaM (adam)"])
        self.standard_combo.currentIndexChanged.connect(self.on_standard_changed)

        self.version_combo = QtWidgets.QComboBox()
        self.set_versions_for_standard("sdtmig")

        self.sdtm_ct_combo = QtWidgets.QComboBox()
        self.sdtm_ct_combo.setMinimumWidth(520)

        self.adam_ct_combo = QtWidgets.QComboBox()
        self.adam_ct_combo.setMinimumWidth(520)

        self.xpt_edit = QtWidgets.QLineEdit()
        self.xpt_edit.setPlaceholderText("Select folder containing XPT files")

        self.btn_browse_xpt = QtWidgets.QPushButton("Browse XPT Folder")
        self.btn_browse_xpt.clicked.connect(self.browse_xpt_folder)

        self.define_edit = QtWidgets.QLineEdit()
        self.define_edit.setPlaceholderText("Optional define.xml path")
        self.define_edit.textChanged.connect(self.on_define_text_changed)

        self.btn_browse_define = QtWidgets.QPushButton("Browse Define.xml")
        self.btn_browse_define.clicked.connect(self.browse_define_xml)

        self.include_define_chk = QtWidgets.QCheckBox("Include Define.xml in CORE Validation")
        self.include_define_chk.setChecked(False)
        self.include_define_chk.setToolTip(
            "Unchecked by default. Enable only when you want CORE checks that use define.xml."
        )

        self.xpt_only_chk = QtWidgets.QCheckBox("Use isolated temp folder with top-level XPT files only")
        self.xpt_only_chk.setChecked(True)
        self.xpt_only_chk.setToolTip(
            "Recommended. Copies only immediate *.xpt files and optional define.xml to a temp folder before validation."
        )

        self.normalize_xpt_chk = QtWidgets.QCheckBox("Create CORE-safe XPT v5 copies using encoding fallback")
        self.normalize_xpt_chk.setChecked(True)
        self.normalize_xpt_chk.setToolTip(
            "Recommended when CORE fails with UTF-8 decode errors. Uses pyreadstat fallback encodings and writes fresh temp XPT v5 copies."
        )

        self.btn_load_ct = QtWidgets.QPushButton("Load CT Packages")
        self.btn_load_ct.clicked.connect(self.load_ct_packages)

        self.btn_update_cache = QtWidgets.QPushButton("Update CORE Cache")
        self.btn_update_cache.clicked.connect(self.update_core_cache)

        self.btn_run = QtWidgets.QPushButton("Run CORE Validation")
        self.btn_run.clicked.connect(self.run_core_validation)

        self.btn_open = QtWidgets.QPushButton("Open Last Report")
        self.btn_open.clicked.connect(self.open_last_report)

        self.btn_clear = QtWidgets.QPushButton("Clear Log")
        self.btn_clear.clicked.connect(self.clear_log)

        row = 0
        g.addWidget(self.api_status, row, 0, 1, 4)
        row += 1
        g.addWidget(self.core_path_status, row, 0, 1, 4)
        row += 1

        g.addWidget(QtWidgets.QLabel("Standard"), row, 0)
        g.addWidget(self.standard_combo, row, 1)
        g.addWidget(QtWidgets.QLabel("IG Version"), row, 2)
        g.addWidget(self.version_combo, row, 3)
        row += 1

        g.addWidget(QtWidgets.QLabel("SDTM CT Package"), row, 0)
        g.addWidget(self.sdtm_ct_combo, row, 1, 1, 3)
        row += 1

        g.addWidget(QtWidgets.QLabel("ADaM CT Package"), row, 0)
        g.addWidget(self.adam_ct_combo, row, 1, 1, 3)
        row += 1

        g.addWidget(QtWidgets.QLabel("XPT Folder"), row, 0)
        g.addWidget(self.xpt_edit, row, 1, 1, 2)
        g.addWidget(self.btn_browse_xpt, row, 3)
        row += 1

        g.addWidget(QtWidgets.QLabel("Define.xml"), row, 0)
        g.addWidget(self.define_edit, row, 1, 1, 2)
        g.addWidget(self.btn_browse_define, row, 3)
        row += 1

        g.addWidget(self.include_define_chk, row, 0, 1, 2)
        g.addWidget(self.xpt_only_chk, row, 2, 1, 2)
        row += 1

        g.addWidget(self.normalize_xpt_chk, row, 0, 1, 4)
        row += 1

        g.addWidget(self.btn_load_ct, row, 0)
        g.addWidget(self.btn_update_cache, row, 1)
        g.addWidget(self.btn_run, row, 2)
        g.addWidget(self.btn_open, row, 3)
        row += 1

        g.addWidget(self.btn_clear, row, 0)

        root.addWidget(controls)

        self.log_box = QtWidgets.QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setLineWrapMode(QtWidgets.QTextEdit.NoWrap)
        root.addWidget(self.log_box, stretch=1)

        self.status = QtWidgets.QLabel("Ready.")
        self.status.setObjectName("Status")
        root.addWidget(self.status)

        self.on_standard_changed()

    def apply_style(self):
        self.setStyleSheet(f"""
            QWidget {{
                background-color: {THEME['app_bg']};
                font-family: Segoe UI;
                font-size: 10pt;
                color: #1c2e4a;
            }}

            #Header {{
                background-color: {THEME['header']};
                border-radius: 14px;
                border: 1px solid #9cc8f5;
            }}

            #Title {{
                font-size: 20pt;
                font-weight: bold;
                color: {THEME['header_text']};
            }}

            #Subtitle {{
                color: {THEME['subtitle']};
                font-weight: bold;
            }}

            #Controls {{
                background-color: {THEME['panel']};
                border-radius: 12px;
                border: 1px solid {THEME['border']};
            }}

            QPushButton {{
                background-color: {THEME['button']};
                color: white;
                border-radius: 8px;
                padding: 7px 12px;
                font-weight: bold;
            }}

            QPushButton:hover {{
                background-color: {THEME['button_hover']};
            }}

            QLineEdit, QComboBox {{
                background-color: white;
                border: 1px solid #a8bfdc;
                border-radius: 8px;
                padding: 5px;
            }}

            QTextEdit {{
                background-color: white;
                border: 1px solid #b7cde8;
                border-radius: 8px;
                padding: 8px;
                font-family: Consolas;
                font-size: 9pt;
            }}

            QCheckBox {{
                font-weight: bold;
            }}

            #Status {{
                background-color: {THEME['status']};
                border: 1px solid #dbc46f;
                border-radius: 8px;
                padding: 6px;
                color: #4d3b00;
                font-weight: bold;
            }}

            #InfoLabel {{
                background-color: #fff8dc;
                border: 1px solid #dbc46f;
                border-radius: 8px;
                padding: 6px;
                color: #4d3b00;
                font-weight: bold;
            }}
        """)

    # --------------------------------------------------------------------------------------------
    # Logging
    # --------------------------------------------------------------------------------------------

    def log(self, text=""):
        self.log_box.append(str(text))
        QtWidgets.QApplication.processEvents()

    def set_status(self, text):
        self.status.setText(str(text))
        QtWidgets.QApplication.processEvents()

    def clear_log(self):
        self.log_box.clear()
        self.set_status("Log cleared.")

    # --------------------------------------------------------------------------------------------
    # Status / setup
    # --------------------------------------------------------------------------------------------

    def refresh_api_status(self):
        self.primary_key, self.secondary_key = load_api_keys()

        if self.primary_key:
            self.api_status.setText(f"API Key: Loaded from {KEY_FILE.name} using primary_key")
        elif self.secondary_key:
            self.api_status.setText(f"API Key: Loaded from {KEY_FILE.name} using secondary_key")
        else:
            self.api_status.setText(f"API Key: Not found. Place CDISC_API_KEY.json next to this script.")

        if os.path.exists(CORE_PY):
            self.core_path_status.setText(f"CORE: Found - {CORE_PY}")
        else:
            self.core_path_status.setText(f"CORE: NOT FOUND - {CORE_PY}")

    def get_api_key(self):
        self.primary_key, self.secondary_key = load_api_keys()
        return self.primary_key or self.secondary_key or ""

    def build_core_env(self):
        env = os.environ.copy()
        api_key = self.get_api_key()
        if api_key:
            env["CDISC_LIBRARY_API_KEY"] = api_key
        env["PYTHONUTF8"] = "1"
        return env

    # --------------------------------------------------------------------------------------------
    # Standard / CT selections
    # --------------------------------------------------------------------------------------------

    def set_versions_for_standard(self, core_standard):
        current = self.version_combo.currentText() if hasattr(self, "version_combo") else ""
        self.version_combo.clear()

        versions = default_versions(core_standard)
        self.version_combo.addItems(versions)

        if current in versions:
            self.version_combo.setCurrentText(current)

    def on_standard_changed(self):
        core_standard = standard_to_core_standard(self.standard_combo.currentText())
        self.set_versions_for_standard(core_standard)

        is_adam = core_standard == "adam"

        self.adam_ct_combo.setEnabled(is_adam)
        self.sdtm_ct_combo.setEnabled(True)

        # For SDTM, ADaM CT is not used. Keep visible but disabled for clarity.
        if not is_adam:
            self.adam_ct_combo.setToolTip("Not used for SDTM validation.")
        else:
            self.adam_ct_combo.setToolTip("Used for ADaM validation.")
            self.sdtm_ct_combo.setToolTip("Shown for reference/fallback traceability. CORE validate uses one -ct package.")

        self.filter_ct_packages()

    def load_ct_packages(self):
        self.refresh_api_status()

        if not os.path.exists(CORE_PY):
            QtWidgets.QMessageBox.critical(self, "CORE Missing", f"core.py not found:\n\n{CORE_PY}")
            return

        self.log("=" * 100)
        self.log("Loading CT packages from CORE cache")
        self.log("=" * 100)

        cmd = [CORE_PYTHON, CORE_PY, "list-ct"]
        self.log(quote_cmd(cmd))
        self.log("-" * 100)

        try:
            result = subprocess.run(
                cmd,
                cwd=CORE_DIR,
                capture_output=True,
                text=True,
                env=self.build_core_env(),
                shell=False,
            )

            if result.stdout:
                self.log(result.stdout)
            if result.stderr:
                self.log(result.stderr)

            packages = [x.strip() for x in result.stdout.splitlines() if x.strip()]
            packages = [p for p in packages if p.lower().startswith(("sdtmct-", "adamct-"))]

            self.all_ct_packages = packages
            self.filter_ct_packages()

            self.log(f"Loaded {len(packages)} CT package(s) from CORE cache.")
            self.set_status("CT packages loaded from CORE cache.")

        except Exception as e:
            self.log(str(e))
            self.set_status("Failed to load CT packages.")

    def filter_ct_packages(self):
        sdtm_packages = newest_first([p for p in self.all_ct_packages if p.lower().startswith("sdtmct-")])
        adam_packages = newest_first([p for p in self.all_ct_packages if p.lower().startswith("adamct-")])

        current_sdtm = self.sdtm_ct_combo.currentText()
        current_adam = self.adam_ct_combo.currentText()

        self.sdtm_ct_combo.clear()
        self.adam_ct_combo.clear()

        self.sdtm_ct_combo.addItems(sdtm_packages)
        self.adam_ct_combo.addItems(adam_packages)

        if current_sdtm in sdtm_packages:
            self.sdtm_ct_combo.setCurrentText(current_sdtm)

        if current_adam in adam_packages:
            self.adam_ct_combo.setCurrentText(current_adam)

    # --------------------------------------------------------------------------------------------
    # Browse
    # --------------------------------------------------------------------------------------------

    def browse_xpt_folder(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Select XPT Folder")
        if folder:
            self.xpt_folder = folder
            self.xpt_edit.setText(folder)

            detected = detect_define_xml(folder)
            if detected and not self.define_edit.text().strip():
                self.define_xml_path = detected
                self.define_edit.setText(detected)

            top_xpts = list(Path(folder).glob("*.xpt"))
            self.set_status(f"XPT folder selected. Top-level XPT files detected: {len(top_xpts)}")

    def browse_define_xml(self):
        start_dir = self.xpt_edit.text().strip() or str(SCRIPT_DIR)
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Select Define.xml",
            start_dir,
            "XML Files (*.xml);;All Files (*.*)"
        )
        if path:
            self.define_xml_path = path
            self.define_edit.setText(path)
            self.include_define_chk.setChecked(True)
            self.set_status("Define.xml selected.")

    def on_define_text_changed(self):
        self.define_xml_path = self.define_edit.text().strip()

    # --------------------------------------------------------------------------------------------
    # Cache update
    # --------------------------------------------------------------------------------------------

    def update_core_cache(self):
        self.refresh_api_status()

        api_key = self.get_api_key()
        if not api_key:
            QtWidgets.QMessageBox.warning(
                self,
                "Missing API Key",
                f"CDISC API key not found.\n\nPlease create:\n{KEY_FILE}"
            )
            return

        if not os.path.exists(CORE_PY):
            QtWidgets.QMessageBox.critical(self, "CORE Missing", f"core.py not found:\n\n{CORE_PY}")
            return

        confirm = QtWidgets.QMessageBox.question(
            self,
            "Update CORE Cache",
            "This will update CORE cache using CDISC Library API.\n\nContinue?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No,
        )
        if confirm != QtWidgets.QMessageBox.Yes:
            return

        self.log_box.clear()
        self.log("=" * 100)
        self.log("Updating CDISC CORE cache")
        self.log("=" * 100)

        cmd = [CORE_PYTHON, CORE_PY, "update-cache"]
        self.log(quote_cmd(cmd))
        self.log("-" * 100)

        try:
            result = subprocess.run(
                cmd,
                cwd=CORE_DIR,
                capture_output=True,
                text=True,
                env=self.build_core_env(),
                shell=False,
            )

            if result.stdout:
                self.log(result.stdout)
            if result.stderr:
                self.log(result.stderr)

            if result.returncode == 0:
                self.log("CORE cache updated successfully.")
                self.set_status("CORE cache updated. Loading CT package list...")
                self.load_ct_packages()
            else:
                self.log("CORE cache update failed.")
                self.set_status("CORE cache update failed.")

        except Exception as e:
            self.log(str(e))
            self.set_status("CORE cache update failed.")

    # --------------------------------------------------------------------------------------------
    # Validation
    # --------------------------------------------------------------------------------------------

    def selected_ct_for_core(self, core_standard):
        if core_standard == "adam":
            return self.adam_ct_combo.currentText().strip()
        return self.sdtm_ct_combo.currentText().strip()

    def run_core_validation(self):
        self.refresh_api_status()

        self.xpt_folder = self.xpt_edit.text().strip()
        self.define_xml_path = self.define_edit.text().strip()

        if not self.xpt_folder:
            QtWidgets.QMessageBox.warning(self, "Missing XPT Folder", "Please select XPT folder.")
            return

        if not os.path.exists(self.xpt_folder):
            QtWidgets.QMessageBox.warning(self, "Invalid XPT Folder", f"Folder does not exist:\n\n{self.xpt_folder}")
            return

        top_level_xpts = sorted([p for p in Path(self.xpt_folder).glob("*.xpt") if p.is_file()])
        if not top_level_xpts:
            QtWidgets.QMessageBox.warning(
                self,
                "No XPT Files",
                f"No top-level .xpt files found in:\n\n{self.xpt_folder}"
            )
            return

        if not os.path.exists(CORE_PY):
            QtWidgets.QMessageBox.critical(self, "CORE Missing", f"core.py not found:\n\n{CORE_PY}")
            return

        core_standard = standard_to_core_standard(self.standard_combo.currentText())
        display_standard = "ADaM" if core_standard == "adam" else "SDTM"
        version = self.version_combo.currentText().strip()
        core_ct_package = self.selected_ct_for_core(core_standard)

        if not core_ct_package:
            QtWidgets.QMessageBox.warning(
                self,
                "Missing CT Package",
                "Please load/select CT package before running validation."
            )
            return

        include_define = (
            self.include_define_chk.isChecked()
            and self.define_xml_path
            and os.path.exists(self.define_xml_path)
        )

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        final_output_file = os.path.join(
            self.xpt_folder,
            f"CORE_Validation_Report_{display_standard}_{timestamp}.xlsx"
        )

        temp_dir = ""
        temp_define = ""
        temp_output_file = ""

        self.log_box.clear()
        self.log("=" * 100)
        self.log("Running CDISC CORE Validation")
        self.log("=" * 100)
        self.log(f"Standard          : {display_standard} ({core_standard})")
        self.log(f"IG Version        : {version}")
        self.log(f"CORE CT Package   : {core_ct_package}")
        self.log(f"SDTM CT Selected  : {self.sdtm_ct_combo.currentText().strip() or '-'}")
        self.log(f"ADaM CT Selected  : {self.adam_ct_combo.currentText().strip() or '-'}")
        self.log(f"Source XPT Folder : {self.xpt_folder}")
        self.log(f"Top-level XPTs    : {len(top_level_xpts)}")
        self.log(f"Include Define    : {'Yes' if include_define else 'No'}")
        self.log(f"CORE-safe XPT     : {'Yes' if self.normalize_xpt_chk.isChecked() else 'No'}")
        self.log(f"Define.xml        : {self.define_xml_path if self.define_xml_path else '-'}")
        self.log(f"Final Report      : {final_output_file}")
        self.log("-" * 100)

        try:
            if self.xpt_only_chk.isChecked():
                temp_dir, temp_define, copied_count = create_isolated_validation_folder(
                    self.xpt_folder,
                    define_xml_path=self.define_xml_path,
                    include_define=include_define,
                    normalize_xpt=self.normalize_xpt_chk.isChecked(),
                    log_func=self.log,
                )
                validation_folder = temp_dir
                validation_define = temp_define
                self.log(f"Validation will use isolated temp folder with {copied_count} XPT file(s).")
            else:
                validation_folder = self.xpt_folder
                validation_define = self.define_xml_path if include_define else ""

            temp_output_file = os.path.join(
                validation_folder,
                f"CORE_Validation_Report_{display_standard}_{timestamp}.xlsx"
            )

            cmd = [
                CORE_PYTHON,
                CORE_PY,
                "validate",
                "-s", core_standard,
                "-v", version,
                "-d", validation_folder,
                "-ft", "xpt",
                "-of", "XLSX",
                "-o", temp_output_file,
                "-ct", core_ct_package,
            ]

            if include_define and validation_define:
                cmd.extend(["-dxp", validation_define])

            self.log("=" * 100)
            self.log("CORE command")
            self.log("=" * 100)
            self.log(quote_cmd(cmd))
            self.log("-" * 100)

            creation_flags = 0
            if os.name == "nt":
                creation_flags = subprocess.CREATE_NO_WINDOW

            self.set_status("CORE validation running...")

            result = subprocess.run(
                cmd,
                cwd=CORE_DIR,
                capture_output=True,
                text=True,
                env=self.build_core_env(),
                shell=False,
                creationflags=creation_flags,
            )

            if result.stdout:
                self.log(result.stdout)

            if result.stderr:
                self.log(result.stderr)

            if result.returncode == 0 and os.path.exists(temp_output_file):
                shutil.copy2(temp_output_file, final_output_file)
                self.core_report_path = final_output_file

                self.log("=" * 100)
                self.log("CORE validation completed successfully.")
                self.log(f"Report copied to: {final_output_file}")
                self.log("=" * 100)
                self.set_status("CORE validation completed.")

                QtWidgets.QMessageBox.information(
                    self,
                    "CORE Validation Complete",
                    f"CORE validation completed successfully:\n\n{final_output_file}"
                )
            else:
                self.log("=" * 100)
                self.log("CORE validation failed.")
                self.log(f"Return code: {result.returncode}")
                self.log("=" * 100)
                self.set_status("CORE validation failed.")

                QtWidgets.QMessageBox.critical(
                    self,
                    "CORE Validation Failed",
                    "CORE validation failed. Please check the log."
                )

        except Exception as e:
            self.log("=" * 100)
            self.log("CORE validation failed with exception.")
            self.log(str(e))
            self.log("=" * 100)
            self.set_status("CORE validation failed.")
            QtWidgets.QMessageBox.critical(self, "CORE Error", str(e))

        finally:
            if temp_dir and os.path.exists(temp_dir):
                try:
                    shutil.rmtree(temp_dir, ignore_errors=True)
                    self.log(f"Temporary validation folder removed: {temp_dir}")
                except Exception as e:
                    self.log(f"Could not remove temporary validation folder: {e}")

    def open_last_report(self):
        if self.core_report_path and os.path.exists(self.core_report_path):
            open_file(self.core_report_path)
        else:
            QtWidgets.QMessageBox.information(
                self,
                "No CORE Report",
                "No CORE report has been generated in this session yet."
            )


# ====================================================================================================
# MAIN
# ====================================================================================================

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")

    win = CoreValidationGui()

    try:
        win.showMaximized()
    except Exception:
        win.show()

    sys.exit(app.exec_())
