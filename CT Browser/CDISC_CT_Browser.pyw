#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# ====================================================================================================
# Script Name    : cdisc_ct_browser.py
#
# Author         : Manivannan Mathialagan
# Created Date   : 12-May-2026
#
# Description    :
#                  Fast and modern CDISC Controlled Terminology (CT) Browser built using Python
#                  and PyQt5 for SDTM and ADaM metadata browsing, mapping, and terminology review.
#
#                  This application connects directly to the CDISC Library API and allows users to:
#
#                  • Browse SDTM and ADaM Controlled Terminology packages
#                  • Search CT values interactively
#                  • View:
#                        - Extensible status
#                        - Codelist Name
#                        - Codelist Code
#                        - CDISC Submission Value
#                        - CDISC Code
#                        - CDISC Synonyms
#                        - CDISC Definitions
#                  • Load different CT package versions
#                  • Review detailed Definition and Synonym information
#                  • Use the application for:
#                        - SDTM mapping workflows
#                        - ADaM mapping workflows
#                        - Metadata review
#                        - Define.xml preparation
#                        - Future CT validation engines
#
# Key Features   :
#                  • Fast PyQt5 table rendering
#                  • Vibrant modern GUI
#                  • Wrapped Definition display
#                  • Cached API responses for better performance
#                  • Automatic cleanup of cache on application close
#                  • Automatic fallback from Primary API Key to Secondary API Key
#                  • Opens maximized by default
#                  • Compact content-focused layout
#
# Cache Location :
#                  C:\TEMP\CDISC_CT_BROWSER_CACHE
#
# API Key File   :
#                  A JSON file named:
#
#                      CDISC_API_KEY.json
#
#                  must be available in the same folder as this script.
#
#                  Example:
#
#                  {
#                      "primary_key": "YOUR_PRIMARY_API_KEY",
#                      "secondary_key": "YOUR_SECONDARY_API_KEY"
#                  }
#
# Usage          :
#
#                  Run directly:
#
#                      python cdisc_ct_browser.py
#
#                  or double-click the .py/.pyw file if Python is associated.
#
#                  Recommended:
#                      • Python 3.10+
#                      • PyQt5
#                      • requests
#
# Auto Installation:
#                  Required Python packages are automatically installed if missing.
#
# Notes          :
#                  • Cache files are automatically deleted when application closes.
#                  • API keys are never displayed in the GUI.
#                  • Primary API Key is used by default.
#                  • Secondary API Key is automatically used if Primary fails.
#
# Future Planned Enhancements:
#                  • SDTM variable browser
#                  • ADaM variable browser
#                  • CT version comparison
#                  • CT migration utilities
#                  • Define.xml generation support
#                  • Pinnacle/CDISC rule validation
#                  • Biomedical Concepts integration
#                  • Sponsor-specific CT extensions
#
# ====================================================================================================

import importlib
import json
import os
import subprocess
import sys
import threading
import shutil
from pathlib import Path

# =========================================================
# 100% UI scaling setup BEFORE PyQt import
# =========================================================
os.environ.setdefault("QT_AUTO_SCREEN_SCALE_FACTOR", "0")
os.environ.setdefault("QT_SCALE_FACTOR", "1")
os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "0")

# =========================================================
# AUTO INSTALL
# =========================================================
for pkg, imp in [("requests", "requests"), ("PyQt5", "PyQt5")]:
    try:
        importlib.import_module(imp)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])

import requests
from PyQt5 import QtCore, QtGui, QtWidgets

# Force normal 100% rendering instead of oversized high DPI scaling
try:
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_DisableHighDpiScaling, True)
except Exception:
    pass

# =========================================================
# CONFIG
# =========================================================
BASE_URL = "https://library.cdisc.org/api"

SCRIPT_DIR = Path(__file__).resolve().parent

# =========================================================
# TEMP CACHE DIRECTORY
# Cache is automatically deleted when app closes
# =========================================================
CACHE_DIR = Path(r"C:\TEMP\CDISC_CT_BROWSER_CACHE")
CACHE_DIR.mkdir(parents=True, exist_ok=True)

KEY_FILE = SCRIPT_DIR / "CDISC_API_KEY.json"

TABLE_COLUMNS = [
    "extensible",
    "codelist_name",
    "codelist_code",
    "submission_value",
    "code",
    "synonyms",
    "definition"
]

TABLE_HEADERS = [
    "Ext",
    "Codelist Name",
    "Codelist Code",
    "Submission Value",
    "Code",
    "CDISC Synonyms",
    "CDISC Definition"
]

# =========================================================
# HELPERS
# =========================================================
def load_api_keys():
    if not KEY_FILE.exists():
        return "", ""

    try:
        data = json.loads(KEY_FILE.read_text(encoding="utf-8"))

        primary = (
            data.get("primary_key")
            or data.get("primary")
            or data.get("api_key")
            or ""
        )

        secondary = (
            data.get("secondary_key")
            or data.get("secondary")
            or ""
        )

        return str(primary).strip(), str(secondary).strip()

    except Exception:
        return "", ""


def cache_file(name):
    safe = "".join(ch if ch.isalnum() or ch in ("_", "-", ".") else "_" for ch in name)
    return CACHE_DIR / f"{safe}.json"


def normalize_extensible(value):
    if isinstance(value, bool):
        return "Yes" if value else "No"

    v = str(value or "").strip().lower()

    if v in ("yes", "y", "true"):
        return "Yes"

    if v in ("no", "n", "false"):
        return "No"

    return str(value or "")


def first_value(*vals):
    for v in vals:
        if v is None:
            continue

        if isinstance(v, str) and v.strip() == "":
            continue

        return v

    return ""


def flatten_synonyms(v):
    if isinstance(v, list):
        return "; ".join(str(x) for x in v)

    return str(v or "")


def parse_package(data, package_title, standard):
    rows = []

    codelists = data.get("codelists", [])

    if not codelists:
        codelists = data.get("_embedded", {}).get("codelists", [])

    if isinstance(codelists, dict):
        codelists = [codelists]

    for cl in codelists:
        cl_name = str(first_value(
            cl.get("name"),
            cl.get("label"),
            cl.get("preferredTerm")
        ))

        cl_code = str(first_value(
            cl.get("conceptId"),
            cl.get("code"),
            cl.get("codelistCode")
        ))

        extensible = normalize_extensible(
            cl.get("extensible")
        )

        terms = (
            cl.get("terms", [])
            or cl.get("enumeratedItems", [])
            or cl.get("items", [])
        )

        if isinstance(terms, dict):
            terms = [terms]

        for term in terms:
            definition = str(first_value(
                term.get("definition"),
                term.get("description"),
                term.get("preferredTerm"),
                term.get("label"),
                term.get("name")
            ))

            synonyms = flatten_synonyms(
                first_value(
                    term.get("synonyms"),
                    term.get("synonym")
                )
            )

            row = {
                "package": package_title,
                "standard": standard,
                "extensible": extensible,
                "codelist_name": cl_name,
                "codelist_code": cl_code,
                "submission_value": str(first_value(
                    term.get("submissionValue"),
                    term.get("value")
                )),
                "code": str(first_value(
                    term.get("conceptId"),
                    term.get("code"),
                    term.get("termCode")
                )),
                "definition": definition,
                "synonyms": synonyms
            }

            row["_search"] = " ".join(
                str(row.get(k, "")).lower()
                for k in row.keys()
            )

            rows.append(row)

    return rows


# =========================================================
# TABLE MODEL
# =========================================================
class CtTableModel(QtCore.QAbstractTableModel):
    def __init__(self):
        super().__init__()
        self.rows = []

    def set_rows(self, rows):
        self.beginResetModel()
        self.rows = rows
        self.endResetModel()

    def rowCount(self, parent=None):
        return len(self.rows)

    def columnCount(self, parent=None):
        return len(TABLE_COLUMNS)

    def data(self, index, role):
        if not index.isValid():
            return None

        row = self.rows[index.row()]
        key = TABLE_COLUMNS[index.column()]
        value = str(row.get(key, ""))

        if role == QtCore.Qt.DisplayRole:
            return value

        if role == QtCore.Qt.ToolTipRole:
            return value

        if role == QtCore.Qt.TextAlignmentRole:
            if key == "extensible":
                return QtCore.Qt.AlignCenter
            return QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter

        if role == QtCore.Qt.BackgroundRole:
            ext = str(row.get("extensible", "")).lower()

            if ext == "yes":
                return QtGui.QColor("#e7f8ec")

            if ext == "no":
                return QtGui.QColor("#fff0f0")

            return QtGui.QColor("#ffffff")

        if role == QtCore.Qt.ForegroundRole:
            ext = str(row.get("extensible", "")).lower()

            if key == "extensible" and ext == "yes":
                return QtGui.QColor("#116329")

            if key == "extensible" and ext == "no":
                return QtGui.QColor("#9a1b1b")

        if role == QtCore.Qt.FontRole:
            if key == "extensible":
                f = QtGui.QFont()
                f.setBold(True)
                return f

        return None

    def headerData(self, section, orientation, role):
        if role != QtCore.Qt.DisplayRole:
            return None

        if orientation == QtCore.Qt.Horizontal:
            return TABLE_HEADERS[section]

        return str(section + 1)


# =========================================================
# MAIN APP
# =========================================================
class CtBrowser(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.primary_key, self.secondary_key = load_api_keys()

        self.rows = []
        self.filtered_rows = []
        self.packages = []

        self.model = CtTableModel()

        self.setWindowTitle("CDISC CT Browser")
        self.resize(1600, 900)
        self.setMinimumSize(1250, 760)

        self.build_ui()
        self.apply_style()

        self.load_package_list()

    # -----------------------------------------------------
    # UI
    # -----------------------------------------------------
    def build_ui(self):
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)

        # HEADER - compact
        header = QtWidgets.QFrame()
        header.setObjectName("Header")
        h_layout = QtWidgets.QHBoxLayout(header)
        h_layout.setContentsMargins(12, 8, 12, 8)

        title = QtWidgets.QLabel("CDISC CT Browser")
        title.setObjectName("Title")

        subtitle = QtWidgets.QLabel("Fast Controlled Terminology Browser for Mapping Workflows")
        subtitle.setObjectName("Subtitle")

        self.api_status = QtWidgets.QLabel("Connected using Primary API Key")
        self.api_status.setObjectName("ApiHeaderStatus")

        h_layout.addWidget(title)
        h_layout.addSpacing(18)
        h_layout.addWidget(subtitle)
        h_layout.addStretch()
        h_layout.addWidget(self.api_status)

        root.addWidget(header)

        # CONTROLS - compact single area
        controls = QtWidgets.QFrame()
        controls.setObjectName("Controls")
        g = QtWidgets.QGridLayout(controls)
        g.setContentsMargins(10, 8, 10, 8)
        g.setHorizontalSpacing(8)
        g.setVerticalSpacing(6)

        self.standard_combo = QtWidgets.QComboBox()
        self.standard_combo.addItems(["SDTM", "ADaM"])
        self.standard_combo.currentIndexChanged.connect(self.filter_package_dropdown)
        self.standard_combo.setFixedWidth(90)

        self.package_combo = QtWidgets.QComboBox()
        self.package_combo.setMinimumWidth(620)

        self.load_btn = QtWidgets.QPushButton("Load CT")
        self.load_btn.clicked.connect(self.load_selected_package)
        self.load_btn.setFixedWidth(95)

        self.search = QtWidgets.QLineEdit()
        self.search.setPlaceholderText("Search codelist, submission value, code, synonyms, definition...")
        self.search.textChanged.connect(self.apply_filter)

        g.addWidget(QtWidgets.QLabel("Standard:"), 0, 0)
        g.addWidget(self.standard_combo, 0, 1)

        g.addWidget(QtWidgets.QLabel("Package:"), 0, 2)
        g.addWidget(self.package_combo, 0, 3)
        g.addWidget(self.load_btn, 0, 4)

        g.addWidget(QtWidgets.QLabel("Search:"), 1, 0)
        g.addWidget(self.search, 1, 1, 1, 4)

        root.addWidget(controls)

        self.summary = QtWidgets.QLabel("Standard: - | Package: -")
        self.summary.setObjectName("Summary")
        root.addWidget(self.summary)

        # TABLE
        self.table = QtWidgets.QTableView()
        self.table.setModel(self.model)

        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)

        # Word wrap ON because Definition is now in table
        self.table.setWordWrap(True)
        self.table.setTextElideMode(QtCore.Qt.ElideRight)
        self.table.setShowGrid(True)
        self.table.setGridStyle(QtCore.Qt.SolidLine)

        # Avoid unnecessary horizontal drag/scroll bar.
        # Columns are stretched to available view width.
        self.table.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.table.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)

        self.table.verticalHeader().setDefaultSectionSize(58)

        header = self.table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(QtWidgets.QHeaderView.Stretch)

        # Give practical starting widths before stretch takes over.
        widths = [65, 230, 115, 190, 115, 300, 500]
        for i, w in enumerate(widths):
            self.table.setColumnWidth(i, w)

        self.table.selectionModel().selectionChanged.connect(self.update_details)

        root.addWidget(self.table, stretch=7)

        # DETAILS PANEL
        details_frame = QtWidgets.QFrame()
        details_frame.setObjectName("DefinitionFrame")
        d_layout = QtWidgets.QVBoxLayout(details_frame)
        d_layout.setContentsMargins(8, 6, 8, 6)
        d_layout.setSpacing(4)

        d_title = QtWidgets.QLabel("Selected Row Details: CDISC Definition and Synonyms")
        d_title.setObjectName("DefinitionTitle")

        self.details_text = QtWidgets.QTextEdit()
        self.details_text.setReadOnly(True)
        self.details_text.setMinimumHeight(170)
        self.details_text.setMaximumHeight(230)
        self.details_text.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.details_text.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)

        d_layout.addWidget(d_title)
        d_layout.addWidget(self.details_text)

        root.addWidget(details_frame, stretch=2)

        self.status = QtWidgets.QLabel("Ready.")
        self.status.setObjectName("Status")
        root.addWidget(self.status)

    # -----------------------------------------------------
    # STYLE
    # -----------------------------------------------------
    def apply_style(self):
        self.setStyleSheet("""
            QWidget {
                background-color: #f3f7fd;
                font-family: 'Segoe UI';
                font-size: 10pt;
                color: #1c2e4a;
            }

            #Header {
                background-color: #cfe7ff;
                border-radius: 12px;
                border: 1px solid #9cc8f5;
            }

            #Title {
                font-size: 18pt;
                font-weight: bold;
                color: #103760;
            }

            #Subtitle {
                color: #315f8f;
                font-size: 10pt;
            }

            #ApiHeaderStatus {
                background-color: #ffffff;
                border: 1px solid #9cc8f5;
                border-radius: 8px;
                padding: 6px 10px;
                color: #116329;
                font-weight: bold;
            }

            #Controls {
                background-color: #ffffff;
                border-radius: 10px;
                border: 1px solid #b7cde8;
            }

            #Summary {
                background-color: #fff8dc;
                border-radius: 8px;
                border: 1px solid #dbc46f;
                padding: 6px;
                font-weight: bold;
                color: #4d3b00;
            }

            #DefinitionFrame {
                background-color: white;
                border-radius: 10px;
                border: 1px solid #b7cde8;
            }

            #DefinitionTitle {
                font-size: 11pt;
                font-weight: bold;
                color: #103760;
            }

            #Status {
                color: #385a7c;
                font-style: italic;
            }

            QPushButton {
                background-color: #4d8ef7;
                color: white;
                border-radius: 8px;
                padding: 6px 10px;
                font-weight: bold;
            }

            QPushButton:hover {
                background-color: #2f75dd;
            }

            QLineEdit, QComboBox {
                background-color: white;
                border: 1px solid #a8bfdc;
                border-radius: 8px;
                padding: 5px;
            }

            QTableView {
                background-color: white;
                border: 1px solid #8fb2d6;
                gridline-color: #9fb7d4;
                selection-background-color: #cfe7ff;
                selection-color: #102a43;
            }

            QHeaderView::section {
                background-color: #d8ebff;
                color: #103760;
                padding: 6px;
                border: 1px solid #8fb2d6;
                font-weight: bold;
            }

            QTextEdit {
                background-color: #fbfdff;
                border: 1px solid #d7e4f3;
                border-radius: 8px;
                padding: 6px;
            }
        """)

    # -----------------------------------------------------
    # API
    # -----------------------------------------------------
    def api_get(self, href):
        url = BASE_URL + href

        if self.primary_key:
            try:
                r = requests.get(
                    url,
                    headers={
                        "api-key": self.primary_key,
                        "Accept": "application/json"
                    },
                    timeout=90
                )

                if r.status_code == 200:
                    self.api_status.setText("Connected using Primary API Key")
                    return r.json()

            except Exception:
                pass

        if self.secondary_key:
            r = requests.get(
                url,
                headers={
                    "api-key": self.secondary_key,
                    "Accept": "application/json"
                },
                timeout=90
            )

            if r.status_code == 200:
                self.api_status.setText("Connected using Backup API Access")
                return r.json()

        raise RuntimeError("Unable to access CDISC Library API")

    # -----------------------------------------------------
    # PACKAGE LIST
    # -----------------------------------------------------
    def load_package_list(self):
        self.status.setText("Loading package list...")

        def worker():
            cache = cache_file("package_list")

            if cache.exists():
                return json.loads(cache.read_text(encoding="utf-8"))

            data = self.api_get("/mdr/ct/packages")
            links = data.get("_links", {})
            packages = links.get("packages", [])

            out = []

            for p in packages:
                href = p.get("href", "")
                title = p.get("title", "")

                if "sdtmct-" in href.lower() or "adamct-" in href.lower():
                    out.append({
                        "href": href,
                        "title": title
                    })

            cache.write_text(json.dumps(out), encoding="utf-8")
            return out

        threading.Thread(
            target=lambda: self.finish_package_load(worker()),
            daemon=True
        ).start()

    def finish_package_load(self, packages):
        self.packages = packages
        QtCore.QTimer.singleShot(0, self.filter_package_dropdown)

    def filter_package_dropdown(self):
        standard = self.standard_combo.currentText()
        self.package_combo.clear()

        filtered = []

        for p in self.packages:
            href = p["href"].lower()

            if standard == "SDTM" and "sdtmct-" in href:
                filtered.append(p)
            elif standard == "ADaM" and "adamct-" in href:
                filtered.append(p)

        # latest at top by reversing original CDISC order if needed
        filtered = list(reversed(filtered))

        for p in filtered:
            self.package_combo.addItem(p["title"], p["href"])

        self.status.setText(f"Loaded {len(filtered)} {standard} CT package options.")

    # -----------------------------------------------------
    # LOAD CT
    # -----------------------------------------------------
    def load_selected_package(self):
        href = self.package_combo.currentData()
        title = self.package_combo.currentText()
        standard = self.standard_combo.currentText()

        if not href:
            QtWidgets.QMessageBox.warning(self, "No Package", "Please select a package.")
            return

        self.summary.setText(f"Standard: {standard} | Package: {title}")
        self.status.setText("Loading CT...")

        def worker():
            cache = cache_file(f"{standard}_{title}")

            if cache.exists():
                return json.loads(cache.read_text(encoding="utf-8"))

            data = self.api_get(href)
            rows = parse_package(data, title, standard)

            cache.write_text(json.dumps(rows), encoding="utf-8")
            return rows

        threading.Thread(
            target=lambda: self.finish_load(worker()),
            daemon=True
        ).start()

    def finish_load(self, rows):
        self.rows = rows
        self.filtered_rows = rows

        QtCore.QTimer.singleShot(0, self.refresh_table)

    def refresh_table(self):
        self.model.set_rows(self.filtered_rows)
        self.status.setText(f"Loaded {len(self.filtered_rows)} rows.")

        # Keep row height consistent for wrapped definitions.
        self.table.verticalHeader().setDefaultSectionSize(58)

    # -----------------------------------------------------
    # FILTER
    # -----------------------------------------------------
    def apply_filter(self):
        q = self.search.text().strip().lower()

        if not q:
            self.filtered_rows = self.rows
        else:
            self.filtered_rows = [
                row for row in self.rows
                if q in row["_search"]
            ]

        self.refresh_table()

    # -----------------------------------------------------
    # DETAILS
    # -----------------------------------------------------
    def update_details(self):
        indexes = self.table.selectionModel().selectedRows()

        if not indexes:
            return

        row = self.filtered_rows[indexes[0].row()]

        definition = row.get("definition", "")
        synonyms = row.get("synonyms", "")

        txt = f"""
        <b>CDISC Definition:</b><br>
        {definition if definition else "-"}<br><br>

        <b>CDISC Synonym(s):</b><br>
        {synonyms if synonyms else "-"}
        """

        self.details_text.setHtml(txt)




    # -----------------------------------------------------
    # CLEANUP CACHE ON CLOSE
    # -----------------------------------------------------
    def closeEvent(self, event):
        try:
            shutil.rmtree(CACHE_DIR, ignore_errors=True)
        except Exception:
            pass

        event.accept()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)

    # Ensures first display is not zoomed/overscaled.
    try:
        app.setAttribute(QtCore.Qt.AA_DisableHighDpiScaling, True)
    except Exception:
        pass

    win = CtBrowser()

    # Open maximized by default so the layout is correct immediately.
    # User can restore/resize normally after opening.
    win.showMaximized()

    sys.exit(app.exec_())
