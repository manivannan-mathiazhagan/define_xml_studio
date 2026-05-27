#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# ====================================================================================================
# Script Name    : cdisc_ct_domain_browser.py
#
# Author         : Manivannan Mathialagan
# Updated Date   : 20-May-2026
#
# Description    :
#                  CDISC Library Browser with two tabs:
#                    1) Controlled Terminology Browser for SDTM/ADaM CT packages
#                    2) SDTM/SEND/ADaM Domain Metadata Browser for domain labels/classes/structures
#
#                  The Domain Browser is intended for checks like:
#                    APCM -> Associated Persons Concomitant Medications
#                    APMH -> Associated Persons Medical History
#
#                  It connects directly to CDISC Library API using CDISC_API_KEY.json.
#
# API Key File   :
#                  Keep a JSON file named CDISC_API_KEY.json in the same folder as this script:
#                  {
#                      "primary_key": "YOUR_PRIMARY_API_KEY",
#                      "secondary_key": "YOUR_SECONDARY_API_KEY"
#                  }
#
# Usage          :
#                  python cdisc_ct_domain_browser.py
#                  or rename to .pyw and double-click where Python association is configured.
#
# Dependencies   : requests, PyQt5. Auto-installed if missing.
# ====================================================================================================

import importlib
import json
import os
import re
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

try:
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_DisableHighDpiScaling, True)
except Exception:
    pass

# =========================================================
# CONFIG
# =========================================================
BASE_URL = "https://library.cdisc.org/api"
SCRIPT_DIR = Path(__file__).resolve().parent
CACHE_DIR = Path(r"C:\TEMP\CDISC_CT_DOMAIN_BROWSER_CACHE")
CACHE_DIR.mkdir(parents=True, exist_ok=True)
KEY_FILE = SCRIPT_DIR / "CDISC_API_KEY.json"

CT_COLUMNS = [
    "extensible",
    "codelist_name",
    "codelist_code",
    "submission_value",
    "code",
    "synonyms",
    "definition",
]

CT_HEADERS = [
    "Ext",
    "Codelist Name",
    "Codelist Code",
    "Submission Value",
    "Code",
    "CDISC Synonyms",
    "CDISC Definition",
]

DOMAIN_COLUMNS = [
    "domain",
    "label",
    "class",
    "structure",
    "purpose",
    "model",
    "version",
    "source_href",
    "definition",
]

DOMAIN_HEADERS = [
    "Domain",
    "Label / Description",
    "Class",
    "Structure",
    "Purpose",
    "Model",
    "Version",
    "Source API Path",
    "Definition / Notes",
]

# Manual fallback endpoints. These are used when the API package-list endpoint shape differs.
# Add/remove versions as needed for your environment.
STANDARD_VERSION_ENDPOINTS = {
    "SDTMIG": {
        "3.4": ["/mdr/sdtmig/3-4", "/mdr/sdtmig/3-4/domains", "/mdr/sdtmig/3-4/datasets"],
        "3.3": ["/mdr/sdtmig/3-3", "/mdr/sdtmig/3-3/domains", "/mdr/sdtmig/3-3/datasets"],
        "3.2": ["/mdr/sdtmig/3-2", "/mdr/sdtmig/3-2/domains", "/mdr/sdtmig/3-2/datasets"],
    },
    "SDTMIG-AP": {
        "1.0": ["/mdr/sdtmig-ap/1-0", "/mdr/sdtmig-ap/1-0/domains", "/mdr/sdtmig-ap/1-0/datasets"],
    },
    "ADAMIG": {
        "1.3": ["/mdr/adamig/1-3", "/mdr/adamig/1-3/datasets", "/mdr/adamig/1-3/domains"],
        "1.2": ["/mdr/adamig/1-2", "/mdr/adamig/1-2/datasets", "/mdr/adamig/1-2/domains"],
    },
    "SENDIG": {
        "3.1": ["/mdr/sendig/3-1", "/mdr/sendig/3-1/domains", "/mdr/sendig/3-1/datasets"],
    },
}

AP_DOMAIN_BASE_LABELS = {
    "AE": "Adverse Events",
    "CM": "Concomitant Medications",
    "DS": "Disposition",
    "DV": "Protocol Deviations",
    "EC": "Exposure as Collected",
    "EG": "ECG Test Results",
    "EX": "Exposure",
    "FA": "Findings About",
    "LB": "Laboratory Test Results",
    "MH": "Medical History",
    "PR": "Procedures",
    "QS": "Questionnaires",
    "SC": "Subject Characteristics",
    "VS": "Vital Signs",
}

# =========================================================
# HELPERS
# =========================================================
def safe_text(value):
    if value is None:
        return ""
    return str(value).strip()


def safe_upper(value):
    return safe_text(value).upper()


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


def cache_file(name):
    safe = "".join(ch if ch.isalnum() or ch in ("_", "-", ".") else "_" for ch in name)
    return CACHE_DIR / f"{safe}.json"


def first_value(*vals):
    for v in vals:
        if v is None:
            continue
        if isinstance(v, str) and v.strip() == "":
            continue
        if isinstance(v, list) and len(v) == 0:
            continue
        if isinstance(v, dict) and len(v) == 0:
            continue
        return v
    return ""


def flatten_list(value):
    if isinstance(value, list):
        out = []
        for item in value:
            if isinstance(item, dict):
                txt = first_value(item.get("name"), item.get("label"), item.get("value"), item.get("text"), item.get("title"))
                if txt:
                    out.append(str(txt))
            else:
                out.append(str(item))
        return "; ".join(out)
    return safe_text(value)


def normalize_extensible(value):
    if isinstance(value, bool):
        return "Yes" if value else "No"
    v = safe_text(value).lower()
    if v in {"yes", "y", "true"}:
        return "Yes"
    if v in {"no", "n", "false"}:
        return "No"
    return safe_text(value)


def iter_dicts(obj):
    if isinstance(obj, dict):
        yield obj
        for v in obj.values():
            yield from iter_dicts(v)
    elif isinstance(obj, list):
        for v in obj:
            yield from iter_dicts(v)


def href_to_url(href):
    href = safe_text(href)
    if not href:
        return ""
    if href.startswith("http://") or href.startswith("https://"):
        return href
    if not href.startswith("/"):
        href = "/" + href
    return BASE_URL + href


def api_path_from_url_or_href(value):
    value = safe_text(value)
    if value.startswith(BASE_URL):
        return value.replace(BASE_URL, "")
    return value


def extract_links(data):
    links = []
    if not isinstance(data, dict):
        return links

    def add_link(obj, key_hint=""):
        if not isinstance(obj, dict):
            return
        href = obj.get("href") or obj.get("url") or obj.get("path")
        title = obj.get("title") or obj.get("name") or obj.get("label") or key_hint or href
        if href:
            links.append({"href": safe_text(href), "title": safe_text(title)})

    embedded = data.get("_embedded")
    if isinstance(embedded, dict):
        for k, v in embedded.items():
            if isinstance(v, list):
                for item in v:
                    add_link(item, k)
            elif isinstance(v, dict):
                add_link(v, k)

    raw_links = data.get("_links")
    if isinstance(raw_links, dict):
        for key, val in raw_links.items():
            if isinstance(val, list):
                for item in val:
                    add_link(item, key)
            elif isinstance(val, dict):
                add_link(val, key)

    # de-duplicate
    seen = set()
    out = []
    for link in links:
        key = (link["href"], link["title"])
        if key not in seen:
            seen.add(key)
            out.append(link)
    return out


# =========================================================
# PARSERS
# =========================================================
def parse_ct_package(data, package_title, standard):
    rows = []
    codelists = data.get("codelists", []) if isinstance(data, dict) else []
    if not codelists and isinstance(data, dict):
        codelists = data.get("_embedded", {}).get("codelists", [])
    if isinstance(codelists, dict):
        codelists = [codelists]

    for cl in codelists:
        cl_name = safe_text(first_value(cl.get("name"), cl.get("label"), cl.get("preferredTerm"), cl.get("submissionValue")))
        cl_code = safe_text(first_value(cl.get("conceptId"), cl.get("code"), cl.get("codelistCode")))
        extensible = normalize_extensible(cl.get("extensible"))
        terms = cl.get("terms", []) or cl.get("enumeratedItems", []) or cl.get("items", [])
        if isinstance(terms, dict):
            terms = [terms]

        for term in terms:
            definition = safe_text(first_value(term.get("definition"), term.get("description"), term.get("preferredTerm"), term.get("label"), term.get("name")))
            synonyms = flatten_list(first_value(term.get("synonyms"), term.get("synonym"), term.get("cdiscSynonyms")))
            row = {
                "package": package_title,
                "standard": standard,
                "extensible": extensible,
                "codelist_name": cl_name,
                "codelist_code": cl_code,
                "submission_value": safe_text(first_value(term.get("submissionValue"), term.get("value"), term.get("codedValue"))),
                "code": safe_text(first_value(term.get("conceptId"), term.get("code"), term.get("termCode"))),
                "definition": definition,
                "synonyms": synonyms,
            }
            row["_search"] = " ".join(str(row.get(k, "")).lower() for k in row)
            rows.append(row)
    return rows


def derive_ap_label(domain):
    ds = safe_upper(domain)
    if ds.startswith("AP") and len(ds) == 4:
        base = ds[2:]
        base_label = AP_DOMAIN_BASE_LABELS.get(base)
        if base_label:
            return f"Associated Persons {base_label}"
    return ""


def looks_like_domain_record(d):
    if not isinstance(d, dict):
        return False
    keys = {safe_upper(k) for k in d.keys()}
    possible_domain = first_value(
        d.get("domain"), d.get("name"), d.get("shortName"), d.get("dataset"),
        d.get("datasetName"), d.get("domainCode"), d.get("submissionValue"),
    )
    dom = safe_upper(possible_domain)
    if not re.fullmatch(r"[A-Z][A-Z0-9]{1,7}", dom):
        return False
    if any(k in keys for k in ["CLASS", "STRUCTURE", "PURPOSE", "DATASET", "DOMAIN", "DESCRIPTION", "LABEL"]):
        return True
    # AP records can be sparse in some Library payloads.
    if dom.startswith("AP") and len(dom) == 4:
        return True
    return False


def parse_domain_metadata(data, model_name="", version="", source_href=""):
    rows = []
    seen = set()

    # The CDISC Library response shape varies by endpoint and version. Search recursively.
    for d in iter_dicts(data):
        if not looks_like_domain_record(d):
            continue

        domain = safe_upper(first_value(
            d.get("domain"), d.get("name"), d.get("shortName"), d.get("dataset"),
            d.get("datasetName"), d.get("domainCode"), d.get("submissionValue"),
        ))
        if not domain:
            continue

        label = safe_text(first_value(
            d.get("label"), d.get("description"), d.get("title"), d.get("longName"),
            d.get("datasetLabel"), d.get("preferredTerm"), d.get("definition"),
        ))

        # If CDISC payload does not directly expose AP label, derive from base domain as useful fallback.
        derived_ap = derive_ap_label(domain)
        if derived_ap and (not label or label.upper() == domain):
            label = derived_ap

        dclass = safe_text(first_value(d.get("class"), d.get("datasetClass"), d.get("domainClass"), d.get("category")))
        structure = safe_text(first_value(d.get("structure"), d.get("recordStructure"), d.get("datasetStructure")))
        purpose = safe_text(first_value(d.get("purpose"), d.get("datasetPurpose")))
        definition = safe_text(first_value(d.get("definition"), d.get("notes"), d.get("comment"), d.get("description")))

        key = (domain, label, dclass, structure, purpose, model_name, version)
        if key in seen:
            continue
        seen.add(key)

        row = {
            "domain": domain,
            "label": label,
            "class": dclass,
            "structure": structure,
            "purpose": purpose,
            "model": model_name,
            "version": version,
            "source_href": api_path_from_url_or_href(source_href),
            "definition": definition,
        }
        row["_search"] = " ".join(str(row.get(k, "")).lower() for k in row)
        rows.append(row)

    rows.sort(key=lambda r: (r.get("domain", ""), r.get("label", "")))
    return rows


# =========================================================
# TABLE MODELS
# =========================================================
class GenericTableModel(QtCore.QAbstractTableModel):
    def __init__(self, columns, headers):
        super().__init__()
        self.columns = columns
        self.headers = headers
        self.rows = []

    def set_rows(self, rows):
        self.beginResetModel()
        self.rows = rows or []
        self.endResetModel()

    def rowCount(self, parent=None):
        return len(self.rows)

    def columnCount(self, parent=None):
        return len(self.columns)

    def data(self, index, role):
        if not index.isValid():
            return None
        row = self.rows[index.row()]
        key = self.columns[index.column()]
        value = safe_text(row.get(key, ""))

        if role == QtCore.Qt.DisplayRole:
            return value
        if role == QtCore.Qt.ToolTipRole:
            return value
        if role == QtCore.Qt.TextAlignmentRole:
            if key in {"extensible", "domain"}:
                return QtCore.Qt.AlignCenter
            return QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter
        if role == QtCore.Qt.BackgroundRole:
            if key == "extensible":
                ext = value.lower()
                if ext == "yes":
                    return QtGui.QColor("#e7f8ec")
                if ext == "no":
                    return QtGui.QColor("#fff0f0")
            if key == "domain":
                return QtGui.QColor("#eef7ff")
            return QtGui.QColor("#ffffff")
        if role == QtCore.Qt.ForegroundRole:
            if key == "extensible" and value.lower() == "yes":
                return QtGui.QColor("#116329")
            if key == "extensible" and value.lower() == "no":
                return QtGui.QColor("#9a1b1b")
            if key == "domain":
                return QtGui.QColor("#103760")
        if role == QtCore.Qt.FontRole:
            if key in {"extensible", "domain"}:
                f = QtGui.QFont()
                f.setBold(True)
                return f
        return None

    def headerData(self, section, orientation, role):
        if role != QtCore.Qt.DisplayRole:
            return None
        if orientation == QtCore.Qt.Horizontal:
            return self.headers[section]
        return str(section + 1)



class WorkerSignals(QtCore.QObject):
    finished = QtCore.pyqtSignal(object)
    failed = QtCore.pyqtSignal(str)


# =========================================================
# MAIN APP
# =========================================================
class CdiscBrowser(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.primary_key, self.secondary_key = load_api_keys()

        self.ct_rows = []
        self.ct_filtered_rows = []
        self.ct_packages = []
        self.ct_model = GenericTableModel(CT_COLUMNS, CT_HEADERS)

        self.domain_rows = []
        self.domain_filtered_rows = []
        self.domain_model = GenericTableModel(DOMAIN_COLUMNS, DOMAIN_HEADERS)

        self.setWindowTitle("CDISC Library Browser - CT and Domain Metadata")
        self.resize(1650, 930)
        self.setMinimumSize(1250, 760)

        self.build_ui()
        self.apply_style()
        self.load_ct_package_list()
        self.populate_domain_versions()

    # -----------------------------------------------------
    # API
    # -----------------------------------------------------
    def api_get(self, href_or_url):
        url = href_to_url(href_or_url)
        errors = []

        if self.primary_key:
            try:
                r = requests.get(url, headers={"api-key": self.primary_key, "Accept": "application/json"}, timeout=90)
                if r.status_code == 200:
                    self.api_status.setText("Connected using Primary API Key")
                    return r.json()
                errors.append(f"Primary {r.status_code}: {r.text[:250]}")
            except Exception as e:
                errors.append(f"Primary error: {e}")

        if self.secondary_key:
            try:
                r = requests.get(url, headers={"api-key": self.secondary_key, "Accept": "application/json"}, timeout=90)
                if r.status_code == 200:
                    self.api_status.setText("Connected using Backup API Access")
                    return r.json()
                errors.append(f"Backup {r.status_code}: {r.text[:250]}")
            except Exception as e:
                errors.append(f"Backup error: {e}")

        raise RuntimeError("Unable to access CDISC Library API. " + " | ".join(errors))

    # -----------------------------------------------------
    # UI
    # -----------------------------------------------------
    def build_ui(self):
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)

        header = QtWidgets.QFrame()
        header.setObjectName("Header")
        h_layout = QtWidgets.QHBoxLayout(header)
        h_layout.setContentsMargins(12, 8, 12, 8)

        title = QtWidgets.QLabel("CDISC Library Browser")
        title.setObjectName("Title")
        subtitle = QtWidgets.QLabel("Controlled Terminology + Domain Metadata lookup")
        subtitle.setObjectName("Subtitle")
        self.api_status = QtWidgets.QLabel("API key not checked yet")
        self.api_status.setObjectName("ApiHeaderStatus")

        h_layout.addWidget(title)
        h_layout.addSpacing(18)
        h_layout.addWidget(subtitle)
        h_layout.addStretch()
        h_layout.addWidget(self.api_status)
        root.addWidget(header)

        self.tabs = QtWidgets.QTabWidget()
        self.tabs.addTab(self.build_ct_tab(), "Controlled Terminology")
        self.tabs.addTab(self.build_domain_tab(), "Domain Metadata")
        root.addWidget(self.tabs, stretch=1)

        self.status = QtWidgets.QLabel("Ready.")
        self.status.setObjectName("Status")
        root.addWidget(self.status)

    def build_ct_tab(self):
        tab = QtWidgets.QWidget()
        root = QtWidgets.QVBoxLayout(tab)
        root.setContentsMargins(0, 6, 0, 0)
        root.setSpacing(6)

        controls = QtWidgets.QFrame()
        controls.setObjectName("Controls")
        g = QtWidgets.QGridLayout(controls)
        g.setContentsMargins(10, 8, 10, 8)
        g.setHorizontalSpacing(8)
        g.setVerticalSpacing(6)

        self.ct_standard_combo = QtWidgets.QComboBox()
        self.ct_standard_combo.addItems(["SDTM", "ADaM"])
        self.ct_standard_combo.currentIndexChanged.connect(self.filter_ct_package_dropdown)
        self.ct_standard_combo.setFixedWidth(90)

        self.ct_package_combo = QtWidgets.QComboBox()
        self.ct_package_combo.setMinimumWidth(620)

        self.ct_load_btn = QtWidgets.QPushButton("Load CT")
        self.ct_load_btn.clicked.connect(self.load_selected_ct_package)
        self.ct_load_btn.setFixedWidth(95)

        self.ct_search = QtWidgets.QLineEdit()
        self.ct_search.setPlaceholderText("Search codelist, submission value, code, synonyms, definition...")
        self.ct_search.textChanged.connect(self.apply_ct_filter)

        g.addWidget(QtWidgets.QLabel("Standard:"), 0, 0)
        g.addWidget(self.ct_standard_combo, 0, 1)
        g.addWidget(QtWidgets.QLabel("Package:"), 0, 2)
        g.addWidget(self.ct_package_combo, 0, 3)
        g.addWidget(self.ct_load_btn, 0, 4)
        g.addWidget(QtWidgets.QLabel("Search:"), 1, 0)
        g.addWidget(self.ct_search, 1, 1, 1, 4)
        root.addWidget(controls)

        self.ct_summary = QtWidgets.QLabel("Standard: - | Package: -")
        self.ct_summary.setObjectName("Summary")
        root.addWidget(self.ct_summary)

        self.ct_table = QtWidgets.QTableView()
        self.ct_table.setModel(self.ct_model)
        self.setup_table(self.ct_table)
        root.addWidget(self.ct_table, stretch=7)

        details_frame = QtWidgets.QFrame()
        details_frame.setObjectName("DefinitionFrame")
        d_layout = QtWidgets.QVBoxLayout(details_frame)
        d_layout.setContentsMargins(8, 6, 8, 6)
        d_title = QtWidgets.QLabel("Selected CT Row Details")
        d_title.setObjectName("DefinitionTitle")
        self.ct_details_text = QtWidgets.QTextEdit()
        self.ct_details_text.setReadOnly(True)
        self.ct_details_text.setMinimumHeight(145)
        self.ct_details_text.setMaximumHeight(210)
        d_layout.addWidget(d_title)
        d_layout.addWidget(self.ct_details_text)
        root.addWidget(details_frame, stretch=2)
        self.ct_table.selectionModel().selectionChanged.connect(self.update_ct_details)

        return tab

    def build_domain_tab(self):
        tab = QtWidgets.QWidget()
        root = QtWidgets.QVBoxLayout(tab)
        root.setContentsMargins(0, 6, 0, 0)
        root.setSpacing(6)

        controls = QtWidgets.QFrame()
        controls.setObjectName("Controls")
        g = QtWidgets.QGridLayout(controls)
        g.setContentsMargins(10, 8, 10, 8)
        g.setHorizontalSpacing(8)
        g.setVerticalSpacing(6)

        self.domain_standard_combo = QtWidgets.QComboBox()
        self.domain_standard_combo.addItems(["SDTMIG", "SDTMIG-AP", "ADAMIG", "SENDIG"])
        self.domain_standard_combo.currentIndexChanged.connect(self.populate_domain_versions)
        self.domain_standard_combo.setFixedWidth(120)

        self.domain_version_combo = QtWidgets.QComboBox()
        self.domain_version_combo.setFixedWidth(130)

        self.domain_load_btn = QtWidgets.QPushButton("Load Domains")
        self.domain_load_btn.clicked.connect(self.load_domain_metadata)
        self.domain_load_btn.setFixedWidth(125)

        self.domain_search = QtWidgets.QLineEdit()
        self.domain_search.setPlaceholderText("Search domain or label, e.g. APCM, APMH, Associated Persons, Medical History...")
        self.domain_search.textChanged.connect(self.apply_domain_filter)

        self.ap_hint_btn = QtWidgets.QPushButton("APCM/APMH Hint")
        self.ap_hint_btn.clicked.connect(self.show_ap_hint)
        self.ap_hint_btn.setFixedWidth(130)

        g.addWidget(QtWidgets.QLabel("Standard:"), 0, 0)
        g.addWidget(self.domain_standard_combo, 0, 1)
        g.addWidget(QtWidgets.QLabel("Version:"), 0, 2)
        g.addWidget(self.domain_version_combo, 0, 3)
        g.addWidget(self.domain_load_btn, 0, 4)
        g.addWidget(self.ap_hint_btn, 0, 5)
        g.addWidget(QtWidgets.QLabel("Search:"), 1, 0)
        g.addWidget(self.domain_search, 1, 1, 1, 5)
        root.addWidget(controls)

        self.domain_summary = QtWidgets.QLabel("Standard: - | Version: - | Domains loaded: -")
        self.domain_summary.setObjectName("Summary")
        root.addWidget(self.domain_summary)

        self.domain_table = QtWidgets.QTableView()
        self.domain_table.setModel(self.domain_model)
        self.setup_table(self.domain_table)
        self.domain_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Interactive)
        self.domain_table.setColumnWidth(0, 90)
        self.domain_table.setColumnWidth(1, 320)
        self.domain_table.setColumnWidth(2, 150)
        self.domain_table.setColumnWidth(3, 420)
        self.domain_table.setColumnWidth(4, 120)
        self.domain_table.setColumnWidth(5, 100)
        self.domain_table.setColumnWidth(6, 90)
        self.domain_table.setColumnWidth(7, 250)
        self.domain_table.setColumnWidth(8, 450)
        root.addWidget(self.domain_table, stretch=7)

        details_frame = QtWidgets.QFrame()
        details_frame.setObjectName("DefinitionFrame")
        d_layout = QtWidgets.QVBoxLayout(details_frame)
        d_layout.setContentsMargins(8, 6, 8, 6)
        d_title = QtWidgets.QLabel("Selected Domain Details")
        d_title.setObjectName("DefinitionTitle")
        self.domain_details_text = QtWidgets.QTextEdit()
        self.domain_details_text.setReadOnly(True)
        self.domain_details_text.setMinimumHeight(145)
        self.domain_details_text.setMaximumHeight(210)
        d_layout.addWidget(d_title)
        d_layout.addWidget(self.domain_details_text)
        root.addWidget(details_frame, stretch=2)
        self.domain_table.selectionModel().selectionChanged.connect(self.update_domain_details)

        return tab

    def setup_table(self, table):
        table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        table.verticalHeader().setVisible(False)
        table.setWordWrap(True)
        table.setTextElideMode(QtCore.Qt.ElideRight)
        table.setShowGrid(True)
        table.setGridStyle(QtCore.Qt.SolidLine)
        table.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        table.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        table.verticalHeader().setDefaultSectionSize(58)
        header = table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(QtWidgets.QHeaderView.Stretch)

    def run_in_thread(self, worker, on_success, busy_text=None):
        """Run API/cache work outside the UI thread and update UI safely via Qt signals."""
        if busy_text:
            self.status.setText(busy_text)
        signals = WorkerSignals()
        signals.finished.connect(on_success)
        signals.failed.connect(lambda msg: self.status.setText(msg))

        # Keep signal objects alive until the worker is finished.
        if not hasattr(self, "_worker_signals"):
            self._worker_signals = []
        self._worker_signals.append(signals)

        def runner():
            try:
                result = worker()
                signals.finished.emit(result)
            except Exception as e:
                signals.failed.emit(str(e))
            finally:
                try:
                    self._worker_signals.remove(signals)
                except Exception:
                    pass

        threading.Thread(target=runner, daemon=True).start()


    # -----------------------------------------------------
    # CT TAB FUNCTIONS
    # -----------------------------------------------------
    def load_ct_package_list(self):
        self.status.setText("Loading CT package list...")

        def worker():
            cache = cache_file("ct_package_list")
            if cache.exists():
                return json.loads(cache.read_text(encoding="utf-8"))
            data = self.api_get("/mdr/ct/packages")
            links = data.get("_links", {}) if isinstance(data, dict) else {}
            packages = links.get("packages", []) if isinstance(links, dict) else []
            out = []
            for p in packages:
                href = p.get("href", "")
                title = p.get("title", "")
                if "sdtmct-" in href.lower() or "adamct-" in href.lower():
                    out.append({"href": href, "title": title})
            cache.write_text(json.dumps(out), encoding="utf-8")
            return out

        def done(packages):
            self.ct_packages = packages
            self.filter_ct_package_dropdown()
            self.status.setText("CT package list loaded.")

        self.run_in_thread(worker, done, "Loading CT package list...")

    def filter_ct_package_dropdown(self):
        standard = self.ct_standard_combo.currentText()
        self.ct_package_combo.clear()
        filtered = []
        for p in self.ct_packages:
            href = p["href"].lower()
            if standard == "SDTM" and "sdtmct-" in href:
                filtered.append(p)
            elif standard == "ADaM" and "adamct-" in href:
                filtered.append(p)
        filtered = list(reversed(filtered))
        for p in filtered:
            self.ct_package_combo.addItem(p["title"], p["href"])
        self.status.setText(f"Loaded {len(filtered)} {standard} CT package options.")

    def load_selected_ct_package(self):
        href = self.ct_package_combo.currentData()
        title = self.ct_package_combo.currentText()
        standard = self.ct_standard_combo.currentText()
        if not href:
            QtWidgets.QMessageBox.warning(self, "No Package", "Please select a CT package.")
            return
        self.ct_summary.setText(f"Standard: {standard} | Package: {title}")
        self.status.setText("Loading CT...")

        def worker():
            cache = cache_file(f"CT_{standard}_{title}")
            if cache.exists():
                return json.loads(cache.read_text(encoding="utf-8"))
            data = self.api_get(href)
            rows = parse_ct_package(data, title, standard)
            cache.write_text(json.dumps(rows), encoding="utf-8")
            return rows

        def done(rows):
            self.ct_rows = rows
            self.ct_filtered_rows = self.ct_rows
            self.refresh_ct_table()

        self.run_in_thread(worker, done, "Loading CT...")

    def apply_ct_filter(self):
        q = self.ct_search.text().strip().lower()
        if not q:
            self.ct_filtered_rows = self.ct_rows
        else:
            self.ct_filtered_rows = [row for row in self.ct_rows if q in row.get("_search", "")]
        self.refresh_ct_table()

    def refresh_ct_table(self):
        self.ct_model.set_rows(self.ct_filtered_rows)
        self.status.setText(f"Loaded {len(self.ct_filtered_rows)} CT rows.")
        self.ct_table.verticalHeader().setDefaultSectionSize(58)

    def update_ct_details(self):
        indexes = self.ct_table.selectionModel().selectedRows()
        if not indexes:
            return
        row = self.ct_filtered_rows[indexes[0].row()]
        txt = f"""
        <b>Codelist:</b> {row.get('codelist_name','')} ({row.get('codelist_code','')})<br>
        <b>Submission Value:</b> {row.get('submission_value','')}<br>
        <b>Term Code:</b> {row.get('code','')}<br><br>
        <b>CDISC Definition:</b><br>{row.get('definition','-') or '-'}<br><br>
        <b>CDISC Synonym(s):</b><br>{row.get('synonyms','-') or '-'}
        """
        self.ct_details_text.setHtml(txt)

    # -----------------------------------------------------
    # DOMAIN TAB FUNCTIONS
    # -----------------------------------------------------
    def populate_domain_versions(self):
        if not hasattr(self, "domain_version_combo"):
            return
        std = self.domain_standard_combo.currentText()
        self.domain_version_combo.clear()
        versions = list(STANDARD_VERSION_ENDPOINTS.get(std, {}).keys())
        for v in versions:
            self.domain_version_combo.addItem(v)

    def domain_candidate_endpoints(self, standard, version):
        endpoints = list(STANDARD_VERSION_ENDPOINTS.get(standard, {}).get(version, []))
        # Extra guesses; harmless if API returns 404 because loader catches per endpoint.
        slug = standard.lower().replace("_", "-")
        vslug = version.replace(".", "-")
        endpoints.extend([
            f"/mdr/{slug}/{vslug}",
            f"/mdr/{slug}/{vslug}/domains",
            f"/mdr/{slug}/{vslug}/datasets",
            f"/mdr/{slug}/{vslug}/classes",
        ])
        # de-duplicate
        out = []
        for e in endpoints:
            if e not in out:
                out.append(e)
        return out

    def load_domain_metadata(self):
        standard = self.domain_standard_combo.currentText()
        version = self.domain_version_combo.currentText()
        if not version:
            QtWidgets.QMessageBox.warning(self, "No Version", "Please select a domain metadata version.")
            return
        self.status.setText(f"Loading {standard} {version} domain metadata...")
        self.domain_summary.setText(f"Standard: {standard} | Version: {version} | Loading...")

        def worker():
            cache = cache_file(f"DOMAIN_{standard}_{version}")
            if cache.exists():
                return json.loads(cache.read_text(encoding="utf-8"))

            all_rows = []
            errors = []
            tried = []
            for endpoint in self.domain_candidate_endpoints(standard, version):
                if endpoint in tried:
                    continue
                tried.append(endpoint)
                try:
                    data = self.api_get(endpoint)
                    rows = parse_domain_metadata(data, standard, version, endpoint)
                    all_rows.extend(rows)

                    # If root endpoint has useful links, fetch likely domain/dataset/class links too.
                    for link in extract_links(data):
                        href = link.get("href", "")
                        hlow = href.lower()
                        if any(token in hlow for token in ["domain", "dataset", "class"]):
                            try:
                                sub_data = self.api_get(href)
                                all_rows.extend(parse_domain_metadata(sub_data, standard, version, href))
                            except Exception:
                                pass
                except Exception as e:
                    errors.append(f"{endpoint}: {e}")

            # De-duplicate final rows.
            seen = set()
            final_rows = []
            for row in all_rows:
                key = (row.get("domain"), row.get("label"), row.get("class"), row.get("structure"), row.get("purpose"))
                if key not in seen:
                    seen.add(key)
                    row["_search"] = " ".join(str(row.get(k, "")).lower() for k in row)
                    final_rows.append(row)
            final_rows.sort(key=lambda r: (r.get("domain", ""), r.get("label", "")))

            if not final_rows:
                # Useful fallback rows for AP check even if endpoint shape/API access is not exposing domains.
                # These are marked clearly as derived fallback by Source API Path.
                for dom in ["APCM", "APMH"]:
                    label = derive_ap_label(dom)
                    final_rows.append({
                        "domain": dom,
                        "label": label,
                        "class": "",
                        "structure": "",
                        "purpose": "",
                        "model": standard,
                        "version": version,
                        "source_href": "Derived fallback - verify against CDISC Library/CORE",
                        "definition": "Derived from AP + base SDTM domain label because API domain payload was not found.",
                        "_search": f"{dom.lower()} {label.lower()} associated persons",
                    })

            cache.write_text(json.dumps(final_rows), encoding="utf-8")
            return final_rows

        def done(rows):
            self.domain_rows = rows
            self.domain_filtered_rows = self.domain_rows
            self.refresh_domain_table()
            self.domain_summary.setText(
                f"Standard: {standard} | Version: {version} | Domains loaded: {len(self.domain_rows)}"
            )

        self.run_in_thread(worker, done, "Loading domain metadata...")

    def apply_domain_filter(self):
        q = self.domain_search.text().strip().lower()
        if not q:
            self.domain_filtered_rows = self.domain_rows
        else:
            self.domain_filtered_rows = [row for row in self.domain_rows if q in row.get("_search", "")]
        self.refresh_domain_table()

    def refresh_domain_table(self):
        self.domain_model.set_rows(self.domain_filtered_rows)
        self.status.setText(f"Loaded {len(self.domain_filtered_rows)} domain rows.")
        self.domain_table.verticalHeader().setDefaultSectionSize(58)

    def update_domain_details(self):
        indexes = self.domain_table.selectionModel().selectedRows()
        if not indexes:
            return
        row = self.domain_filtered_rows[indexes[0].row()]
        txt = f"""
        <b>Domain:</b> {row.get('domain','')}<br>
        <b>Label:</b> {row.get('label','')}<br>
        <b>Class:</b> {row.get('class','') or '-'}<br>
        <b>Structure:</b> {row.get('structure','') or '-'}<br>
        <b>Purpose:</b> {row.get('purpose','') or '-'}<br>
        <b>Model/Version:</b> {row.get('model','')} {row.get('version','')}<br>
        <b>Source:</b> {row.get('source_href','')}<br><br>
        <b>Definition / Notes:</b><br>{row.get('definition','') or '-'}
        """
        self.domain_details_text.setHtml(txt)

    def show_ap_hint(self):
        self.domain_search.setText("AP")
        QtWidgets.QMessageBox.information(
            self,
            "Associated Persons Domain Hint",
            "For AP domains, search AP or the specific domain.\n\n"
            "Expected examples:\n"
            "APCM = Associated Persons Concomitant Medications\n"
            "APMH = Associated Persons Medical History\n\n"
            "If CDISC Library metadata is not returned by the selected endpoint, the tool shows a derived fallback row clearly marked as fallback."
        )

    # -----------------------------------------------------
    # STYLE / CLOSE
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
            QTabWidget::pane {
                border: 1px solid #b7cde8;
                background: #f3f7fd;
                border-radius: 8px;
            }
            QTabBar::tab {
                background: #d8ebff;
                color: #103760;
                padding: 8px 14px;
                border: 1px solid #9cc8f5;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                font-weight: bold;
            }
            QTabBar::tab:selected {
                background: #ffffff;
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

    def closeEvent(self, event):
        try:
            shutil.rmtree(CACHE_DIR, ignore_errors=True)
        except Exception:
            pass
        event.accept()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    win = CdiscBrowser()
    win.showMaximized()
    sys.exit(app.exec_())
