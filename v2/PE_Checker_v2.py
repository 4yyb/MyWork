import os
import sys
import json
import csv
import re
import socket
import subprocess
import time
import uuid
import hmac
import hashlib
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string

from PySide6.QtCore import Qt, QTimer, QRect, QSize
from PySide6.QtGui import QColor, QFont, QFontMetrics
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QMessageBox, QHeaderView, QDialog, QFormLayout, QSpinBox,
    QComboBox, QCheckBox, QGroupBox, QScrollArea, QInputDialog, QFileDialog,
    QTabWidget, QSplitter, QListWidget, QListWidgetItem, QStackedWidget, QToolButton,
    QFrame, QSizePolicy, QStyledItemDelegate, QStyle, QMenu
)

# ---------------------------------------------------------------------------
# LOKASI CONFIG & SESSION
# config.json dan session.json SELALU dibaca/ditulis di folder yang sama dengan
# program ini (clone1.py atau clone1.exe). Tidak ada lokasi bersama di folder
# user, jadi setiap clone (folder berbeda) punya config & session sendiri.
# ---------------------------------------------------------------------------
def get_app_dir():
    if getattr(sys, "frozen", False):
        # PyInstaller (onefile/onedir) -> lokasi file .exe asli
        return Path(sys.executable).resolve().parent
    if "__compiled__" in globals():
        # Nuitka -> sys.argv[0] menunjuk ke lokasi .exe asli
        return Path(sys.argv[0]).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = get_app_dir()
CONFIG_FILE = APP_DIR / "config.json"
SESSION_FILE = APP_DIR / "session.json"
# File operator bersama. Jika config.json tidak menyebut "operators_file", dipakai
# operators.json di folder program ini. Path dapat diubah lewat menu Tambah Operator
# supaya banyak clone memakai SATU file yang sama.
DEFAULT_OPERATORS_FILE = APP_DIR / "operators.json"
MAX_OPERATORS = 1000

# Warna baris step: kuning = menunggu scan, hijau = OK, merah = NG.
COLOR_WAITING = QColor("#FFFF00")
COLOR_OK = QColor("#66DD66")
COLOR_NG = QColor("#FF6B6B")

# Pesan bawaan validasi Excel (berbeda tergantung hasil "ditemukan" = OK atau NG).
EXCEL_MSG_FOUND_NG = "Barcode sudah terdaftar di database Excel."
EXCEL_MSG_NOTFOUND_NG = "Barcode tidak ditemukan di database Excel."

EMPTY_CONFIG = {
    "operators_file": "",
    "steps": [],
    "saving": {
        "enabled": False,
        "targets": [{
            "destination": "",
            "filename": "",
            "filename_part1": "none",
            "filename_part2": "none",
            "filename_text": "hasil_pengujian",
            "date_prefix": False,
            "filename_source": "",
            "columns": [],
        }],
        "line": "",
        "result_text": "OK",
        "column_names": {},
    },
}



def load_json(path, default):
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError, RecursionError):
        return json.loads(json.dumps(default))  # salinan baru, bukan objek bersama


def save_json(path, data):
    """Tulis JSON ke path. Return True jika berhasil; tampilkan pesan jika gagal."""
    text = json.dumps(data, indent=2, ensure_ascii=False)
    # Nama sementara unik per proses/pemanggilan supaya dua proses tidak saling menimpa.
    tmp = path.with_name(f"{path.name}.{os.getpid()}.{uuid.uuid4().hex}.tmp")
    try:
        # Tulis ke file sementara lalu ganti, supaya file tidak rusak jika terputus.
        tmp.write_text(text, encoding="utf-8")
        os.replace(tmp, path)
        return True
    except OSError:
        try:
            tmp.unlink()
        except OSError:
            pass
        try:
            path.write_text(text, encoding="utf-8")  # cara langsung sebagai cadangan
            return True
        except OSError as exc:
            if QApplication.instance() is not None:
                QMessageBox.critical(
                    None, "Gagal menulis file",
                    f"Tidak bisa menulis:\n{path}\n\n{exc}\n\n"
                    "Pastikan folder program tidak read-only "
                    "(hindari Program Files / Desktop yang dilindungi)."
                )
            return False


def int_or(value, default):
    """int(value), atau default jika config rusak (bukan angka)."""
    try:
        return int(value)
    except (TypeError, ValueError, OverflowError):
        return default


ILLEGAL_NAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
RESERVED_NAMES = re.compile(r"^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$", re.IGNORECASE)


def safe_file_stem(stem, folder, extension, default="hasil_pengujian"):
    """Nama file aman (tanpa folder) untuk folder tujuan. Input normal tidak berubah.
    Karakter ilegal -> '_'; nama perangkat Windows (CON, NUL, COM1, ...) diawali '_';
    hasil harus tetap berada langsung di dalam folder, jika tidak dipakai nama default."""
    stem = ILLEGAL_NAME_CHARS.sub("_", str(stem))
    if RESERVED_NAMES.match(stem):
        stem = "_" + stem
    try:
        base = Path(folder).resolve()
        if (base / (stem + extension)).resolve().parent != base:
            stem = default
    except (OSError, ValueError, RuntimeError):
        stem = default
    return stem + extension


# Hash password (PBKDF2-SHA256, salt tetap di kode); password yang diterima tidak berubah.
_PW_SALT = b"PE-Checker-v2/pw-salt"
_PW_HASHES = {
    "operator": bytes.fromhex("8b3f3e40f4dac7846e80480cb1e8a6dab230ae68eda6f83d4bea0ce587723630"),
    "edit": bytes.fromhex("84ad3f2ba1c8f692a9d5feee62a5d925de744762c9291ad5a55a735d5e4c761f"),
}


def password_ok(kind, typed):
    digest = hashlib.pbkdf2_hmac("sha256", str(typed).encode("utf-8"), _PW_SALT, 200000)
    return hmac.compare_digest(digest, _PW_HASHES[kind])


def resolve_operators_path(value):
    """Ubah teks path dari config menjadi Path. Kosong = operators.json di folder program;
    path relatif dihitung dari folder program."""
    value = (value or "").strip()
    if not value:
        return DEFAULT_OPERATORS_FILE
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = APP_DIR / path
    return path


def read_operators_file(path):
    """Baca file operator. Return (dict barcode -> nama, pesan_error atau None).
    File yang belum ada dianggap kosong (bukan error)."""
    try:
        if not path.exists():
            return {}, None
    except (OSError, ValueError) as exc:
        return {}, str(exc)
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError, RecursionError) as exc:
        return {}, str(exc)
    ops = data.get("operators") if isinstance(data, dict) else None
    if not isinstance(ops, dict):
        return {}, 'Format file tidak dikenali (isian "operators" tidak ditemukan).'
    return {str(k).strip(): str(v).strip() for k, v in ops.items() if str(k).strip()}, None


def write_operators_file(path, operators):
    """Tulis file operator. Return True jika berhasil."""
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        QMessageBox.critical(None, "Gagal menulis file", f"Tidak bisa membuat folder:\n{path.parent}\n\n{exc}")
        return False
    return save_json(path, {"operators": operators})


def check_ip_connection(host, port=0, timeout=1.5):
    """Cek koneksi ke host/IP. port > 0: sambung TCP ke port itu; port 0: ping.
    Return (True, "") jika tersambung, atau (False, alasan)."""
    host = (host or "").strip()
    if not host:
        return False, "Alamat IP belum diisi."
    if not re.fullmatch(r"[A-Za-z0-9._:\-]+", host) or host.startswith("-"):
        return False, "Alamat IP/host tidak valid."
    if port and port > 0:
        try:
            with socket.create_connection((host, int(port)), timeout=timeout):
                return True, ""
        except OSError as exc:
            return False, str(exc)
    if os.name == "nt":
        ping_exe = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "ping.exe")
        cmd = [ping_exe if os.path.isfile(ping_exe) else "ping", "-n", "1", "-w", str(int(timeout * 1000)), host]
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    else:
        cmd = ["ping", "-c", "1", "-W", str(max(1, int(round(timeout)))), host]
        flags = 0
    try:
        result = subprocess.run(
            cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            timeout=timeout + 2, creationflags=flags
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        return False, str(exc)
    if result.returncode == 0:
        return True, ""
    return False, "Tidak ada balasan (ping gagal)."


# ---------------------------------------------------------------------------
# DIALOG "EDIT PROGRAM" (v2): daftar step di kiri, editor step terpilih di kanan,
# navigasi samping (Step / Penyimpanan CSV / File .dat / Umum), validasi inline.
# Struktur hasil get_config() SAMA dengan versi lama (config.json tidak berubah).
# ---------------------------------------------------------------------------
MAX_GROUPS = 12
MAX_STEPS = 1000
BAD_FILENAME_CHARS = '<>:"/\\|?*'
FS_TTL = 4.0  # detik: hasil cek file/folder dipakai ulang supaya UI tidak menunggu share jaringan

APP_QSS = """
* { font-size: 13px; }
QDialog { background: #F4F6F8; }
QLabel#hint { color: #5B6573; }
QLabel#err { color: #B42318; font-size: 12px; }
QLabel#h1 { font-size: 18px; font-weight: bold; color: #111827; }
QLabel#status { color: #6B7280; padding: 4px 8px; }
QLabel#status[toast="true"] { background: #E7F5EC; color: #14532D; border-radius: 4px; }
QLabel#status[toast="true"][bad="true"] { background: #FDECEA; color: #B42318; }
QLabel#banner { background: #FEF3C7; color: #92400E; border: 1px solid #F59E0B; border-radius: 4px; padding: 6px 10px; }
QLabel#preview { background: #F3F4F6; color: #374151; border-radius: 4px; padding: 8px 10px; font-size: 13px; }
QSpinBox { min-height: 24px; }
QLineEdit, QComboBox { border: 1px solid #8C97A5; border-radius: 4px; padding: 3px 6px;
    background: white; min-height: 24px; selection-background-color: #2563EB; }
QLineEdit:focus, QComboBox:focus { border: 2px solid #2563EB; }
QLineEdit:disabled, QComboBox:disabled { background: #E9EDF1; color: #6B7684; }
QLineEdit[invalid="true"], QComboBox[invalid="true"] { border: 2px solid #D92D20; background: #FFF5F4; }
QGroupBox[invalid="true"] { border: 2px solid #D92D20; }
QPushButton, QToolButton#more { border: 1px solid #8C97A5; border-radius: 4px; padding: 4px 12px; background: white; min-height: 24px; }
QPushButton:hover, QToolButton#more:hover { background: #EAF0F8; }
QPushButton:focus, QToolButton#more:focus { border: 2px solid #2563EB; }
QPushButton:disabled { color: #8A94A1; background: #ECEFF2; border-color: #C5CCD3; }
QToolButton#more::menu-indicator { image: none; }
QPushButton#primary { background: #1D4ED8; color: white; border-color: #1E40AF; font-weight: bold; padding: 4px 18px; }
QPushButton#primary:hover { background: #1E40AF; }
QPushButton#primary:disabled { background: #E5E7EB; border-color: #D1D5DB; color: #6B7280; }
QPushButton#link { border: none; background: transparent; color: #1D4ED8; text-decoration: underline; padding: 2px 6px; min-height: 0; }
QPushButton#link:hover { color: #1E40AF; background: transparent; }
QMenu { background: white; border: 1px solid #C5CCD3; padding: 4px; }
QMenu::item { padding: 6px 26px; }
QMenu::item:selected { background: #DCE8FF; color: #111827; }
QListWidget#nav { background: #1F2937; color: #E5E7EB; border: none; padding: 8px; outline: 0; }
QListWidget#nav::item { padding: 10px 12px; border-radius: 6px; margin-bottom: 2px; }
QListWidget#nav::item:selected { background: #2563EB; color: white; }
QListWidget#steps { background: white; border: 1px solid #C5CCD3; border-radius: 6px; outline: 0; }
QFrame#card { background: white; border: 1px solid #C5CCD3; border-radius: 8px; }
QFrame#section { background: #F8FAFC; border: 1px solid #DDE3E9; border-radius: 6px; }
QToolButton#fold { border: none; font-weight: bold; padding: 4px 2px; }
QToolButton#fold:focus { border: 2px solid #2563EB; border-radius: 4px; }
QGroupBox { border: 1px solid #C5CCD3; border-radius: 8px; margin-top: 12px; padding-top: 8px; background: white; }
QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; font-weight: bold; }
QScrollArea { border: none; background: transparent; }
QTabWidget::pane { border: 1px solid #C5CCD3; border-radius: 8px; background: white; top: -1px; }
QTabBar::tab { background: #EEF1F4; border: 1px solid #C5CCD3; border-bottom: none; border-top-left-radius: 6px;
    border-top-right-radius: 6px; padding: 6px 16px; margin-right: 2px; color: #374151; }
QTabBar::tab:selected { background: white; color: #111827; font-weight: bold; }
QCheckBox:focus { color: #1D4ED8; }
"""

MAIN_QSS = """
QPushButton { border: 1px solid #8C97A5; border-radius: 4px; padding: 6px 14px; background: white; min-height: 24px; }
QPushButton:hover { background: #EAF0F8; }
QPushButton:disabled { color: #8A94A1; background: #ECEFF2; border-color: #C5CCD3; }
QLineEdit { border: 1px solid #8C97A5; border-radius: 4px; padding: 3px 6px; }
QLineEdit:focus { border: 2px solid #2563EB; }
QHeaderView::section { background: #E5EAF0; border: 1px solid #C5CCD3; padding: 6px; font-weight: bold; }
"""


class Button(QPushButton):
    """Tombol dialog: bukan tombol default, supaya Enter di kolom isian tidak menekannya."""

    def __init__(self, *args):
        super().__init__(*args)
        self.setAutoDefault(False)


class Form(QFormLayout):
    """Form dengan jarak label-field 12px dan lebar label seragam."""

    def __init__(self, parent=None, label_width=180):
        super().__init__(parent)
        self._label_width = label_width
        self.setHorizontalSpacing(12)
        self.setVerticalSpacing(8)
        self.setLabelAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)

    def addRow(self, *args):
        if len(args) == 2 and isinstance(args[0], str):
            label = QLabel(args[0])
            label.setMinimumWidth(self._label_width)
            label.setContentsMargins(0, 5, 0, 0)  # sejajar vertikal dengan tengah field
            return super().addRow(label, args[1])
        return super().addRow(*args)


class ElidedLabel(QLabel):
    """Label satu baris yang dipotong dengan "…" di kanan; teks penuh ada di tooltip."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._full = ""
        self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        self.setMinimumWidth(10)

    def set_full_text(self, text):
        self._full = text
        self.setToolTip(text)
        self._elide()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._elide()

    def _elide(self):
        self.setText(self.fontMetrics().elidedText(self._full, Qt.ElideRight, max(10, self.width() - 4)))


def _clamp_int(value, default=1, lo=1, hi=9999):
    try:
        number = int(value)
    except (TypeError, ValueError):
        number = default
    return max(lo, min(hi, number))


_SHEET_CACHE = {}


def workbook_sheetnames(path):
    """(daftar nama sheet, pesan_error) sebuah .xlsx; di-cache per path + waktu ubah."""
    try:
        text = str(path)
        key = (text, os.stat(text).st_mtime_ns)
    except (OSError, ValueError):
        return [], "File tidak dapat dibaca."
    if key not in _SHEET_CACHE:
        try:
            wb = load_workbook(text, read_only=True, data_only=True)
            try:
                _SHEET_CACHE[key] = (list(wb.sheetnames), "")
            finally:
                wb.close()
        except Exception as exc:
            _SHEET_CACHE[key] = ([], str(exc) or "File Excel tidak dapat dibaca.")
    return _SHEET_CACHE[key]


def default_group(j):
    return {
        "name": f"Data {j + 1}", "start": 1, "length": 1, "mode": "length",
        "excel_enabled": False, "xlsx_path": "", "xlsx_column": "G", "xlsx_match_chars": 8,
        "xlsx_sheet_mode": "specific", "xlsx_sheet_name": "Sheet1",
        "excel_error": EXCEL_MSG_FOUND_NG, "excel_popup": False,
        "char_enabled": False, "char_start": 1, "char_length": 3, "char_expected": "ABC",
        "char_error": "Format karakter barcode tidak sesuai.", "char_popup": False,
        "unique": False, "unique_error": "Barcode sudah pernah digunakan.", "unique_popup": False,
        "judge_ok": "OK", "judge_ng": "NG",
    }


def normalize_group(g, j, kind):
    """Nilai kontrol sebuah Data (sama dengan yang dibuat widget versi lama dari config)."""
    if not isinstance(g, dict):
        g = {}
    found = g.get("excel_found_result", "ng")
    found = found if found in ("ng", "ok") else "ng"
    err = g.get("excel_error", EXCEL_MSG_NOTFOUND_NG if found == "ok" else EXCEL_MSG_FOUND_NG)
    cur = str(err).strip()
    if found == "ok" and cur in ("", EXCEL_MSG_FOUND_NG):
        err = EXCEL_MSG_NOTFOUND_NG
    elif found == "ng" and cur in ("", EXCEL_MSG_NOTFOUND_NG):
        err = EXCEL_MSG_FOUND_NG
    modes = ["length", "all"] + (["operator_name"] if kind == "operator" else [])
    mode = g.get("mode", "length")
    mode = mode if mode in modes else "length"
    path = str(g.get("xlsx_path", ""))
    sheet = g.get("xlsx_sheet_name", "Sheet1")
    if not sheet:
        # Kombo nama sheet versi lama memilih sheet pertama file jika nama kosong.
        names = []
        try:
            if path.strip() and Path(path.strip()).is_file():
                names = workbook_sheetnames(path.strip())[0]
        except (OSError, ValueError):
            pass
        sheet = names[0] if names else ""
    sheet_mode = g.get("xlsx_sheet_mode", "specific")
    legacy_expected = g.get("expected", g.get("xlsx_required_prefix", ""))
    legacy_char = bool(g.get("validate", False) and g.get("source", "text") == "text" and legacy_expected)
    legacy_prefix = bool(g.get("xlsx_required_prefix", ""))
    return {
        "name": str(g.get("name", f"Data {j + 1}")), "mode": mode,
        "start": _clamp_int(g.get("start", 1)), "length": _clamp_int(g.get("length", 1)),
        "excel_enabled": bool(g.get("excel_enabled", g.get("validate", False) and g.get("source") == "excel")),
        "xlsx_path": path.strip(),
        "xlsx_column": str(g.get("xlsx_column", "G")).strip().upper(),
        "xlsx_match_chars": _clamp_int(g.get("xlsx_match_chars", 8), 8),
        "xlsx_sheet_mode": sheet_mode if sheet_mode in ("specific", "all") else "specific",
        "xlsx_sheet_name": str(sheet).strip(),
        "excel_error": str(err), "excel_popup": bool(g.get("excel_popup", False)),
        "excel_found_result": found,
        "char_enabled": bool(g.get("char_enabled", legacy_char or legacy_prefix)),
        "char_start": _clamp_int(g.get("char_start", 1)),
        "char_length": _clamp_int(g.get("char_length", len(legacy_expected) if legacy_expected else 3), 3),
        "char_expected": str(g.get("char_expected", legacy_expected)),
        "char_error": str(g.get("char_error", "Format karakter barcode tidak sesuai.")),
        "char_popup": bool(g.get("char_popup", False)),
        "unique": bool(g.get("unique", False)),
        "unique_error": str(g.get("unique_error", "Barcode sudah pernah digunakan.")),
        "unique_popup": bool(g.get("unique_popup", False)),
        "judge_ok": str(g.get("judge_ok", "OK")), "judge_ng": str(g.get("judge_ng", "NG")),
    }


def normalize_step(raw, index):
    if not isinstance(raw, dict):
        raw = {}
    kind = raw.get("input_type", "barcode")
    kind = kind if kind in ("operator", "barcode") else "operator"
    groups = raw.get("groups", [])
    groups = groups if isinstance(groups, list) else []
    return {
        "name_before": str(raw.get("name_before", "")), "name_after": str(raw.get("name_after", "")),
        "message_before": str(raw.get("message_before", "")), "input_type": kind,
        "groups": [normalize_group(g, j, kind) for j, g in enumerate(groups[:MAX_GROUPS])],
    }


def default_step(index):
    return normalize_step({"name_before": f"Step {index + 1}", "name_after": f"Step {index + 1} selesai",
                           "input_type": "barcode", "groups": []}, index)


def export_group(g):
    """Bentuk akhir sebuah Data di config.json (sama dengan versi lama)."""
    return {
        "name": g["name"].strip() or "Data", "mode": g["mode"],
        "start": g["start"], "length": g["length"],
        "excel_enabled": g["excel_enabled"], "xlsx_path": g["xlsx_path"].strip(),
        "xlsx_column": g["xlsx_column"].strip().upper(), "xlsx_match_chars": g["xlsx_match_chars"],
        "xlsx_sheet_mode": g["xlsx_sheet_mode"], "xlsx_sheet_name": g["xlsx_sheet_name"].strip(),
        "excel_error": g["excel_error"], "excel_popup": g["excel_popup"],
        "excel_found_result": g["excel_found_result"] or "ng",
        "char_enabled": g["char_enabled"], "char_start": g["char_start"], "char_length": g["char_length"],
        "char_expected": g["char_expected"], "char_error": g["char_error"], "char_popup": g["char_popup"],
        "unique": g["unique"], "unique_error": g["unique_error"], "unique_popup": g["unique_popup"],
        "judge_ok": g["judge_ok"].strip() or "OK", "judge_ng": g["judge_ng"].strip() or "NG",
    }


def truncation_info(raw_steps):
    """Daftar catatan data yang akan terpotong (batas tetap seperti versi lama)."""
    notes = []
    if not isinstance(raw_steps, list):
        return notes
    if len(raw_steps) > MAX_STEPS:
        notes.append(f"{len(raw_steps) - MAX_STEPS} step terakhir (batas {MAX_STEPS} step)")
    for i, step in enumerate(raw_steps[:MAX_STEPS], start=1):
        groups = step.get("groups") if isinstance(step, dict) else None
        if isinstance(groups, list) and len(groups) > MAX_GROUPS:
            notes.append(f"{len(groups) - MAX_GROUPS} Data pada Step {i} (batas {MAX_GROUPS} Data per step)")
    return notes


def repolish(widget):
    widget.style().unpolish(widget)
    widget.style().polish(widget)
    widget.update()


class CsvRemap:
    """Memetakan ulang pilihan CSV yang terikat nomor urut step (kode_produk_n,
    data_step_S_group_G, nama kolom) mengikuti perpindahan/penghapusan step lewat id step."""

    def __init__(self, before, after):
        self.b, self.a = before, after
        self.a_prod = {sid: n for n, sid in after["prod"].items()}
        self.a_pos = {sid: s for s, sid in after["pos"].items()}
        self.b_pos_of = {sid: s for s, sid in before["pos"].items()}

    def col(self, key):
        """Key kolom baru, atau None jika step-nya sudah tidak ada/bukan barcode lagi."""
        if not key.startswith("kode_produk_"):
            return key
        try:
            n = int(key.rsplit("_", 1)[1])
        except ValueError:
            return None
        sid = self.b["prod"].get(n)
        m = self.a_prod.get(sid) if sid is not None else None
        return None if m is None else f"kode_produk_{m}"

    def data(self, key):
        found = re.fullmatch(r"data_step_(\d+)_group_(\d+)", key or "")
        if not found:
            return key or ""
        sid = self.b["pos"].get(int(found.group(1)))
        pos = self.a_pos.get(sid) if sid is not None else None
        return "" if pos is None else f"data_step_{pos}_group_{found.group(2)}"

    def _removed_pos(self, key):
        found = re.fullmatch(r"kode_produk_(\d+)", key) or re.fullmatch(r"data_step_(\d+)_group_\d+", key)
        if not found:
            return None
        n = int(found.group(1))
        sid = self.b["prod"].get(n) if key.startswith("kode_produk_") else self.b["pos"].get(n)
        return self.b_pos_of.get(sid)

    def apply(self, state):
        removed = set()
        targets = []
        for t in state["targets"]:
            cols = []
            for key in t["columns"]:
                new = self.col(key)
                if new is None:
                    pos = self._removed_pos(key)
                    if pos is not None:
                        removed.add(pos)
                else:
                    cols.append(new)
            entry = {"columns": cols}
            for n in (1, 2):
                old = t[f"data{n}"]
                new = self.data(old)
                if old and not new:
                    pos = self._removed_pos(old)
                    if pos is not None:
                        removed.add(pos)
                entry[f"data{n}"] = new
            targets.append(entry)
        names = {}
        for key, text in state["names"].items():
            new = self.col(key)
            if new is None:
                continue
            if text == self.b["labels"].get(key):  # nama bawaan lama memuat nomor step: pakai bawaan baru
                continue
            names[new] = text
        return {"targets": targets, "names": names}, removed


class StepDelegate(QStyledItemDelegate):
    """Baris daftar step: nomor rata kanan, judul tebal, subbaris abu-abu, chip validasi E K U,
    garis merah di kiri + "!" jika bermasalah."""
    ROLE = Qt.UserRole + 1

    def sizeHint(self, option, index):
        return QSize(0, 48)

    def paint(self, painter, option, index):
        d = index.data(self.ROLE)
        if not isinstance(d, dict):
            return super().paint(painter, option, index)
        painter.save()
        r = option.rect
        selected = bool(option.state & QStyle.StateFlag.State_Selected)
        painter.fillRect(r, QColor("#DCE8FF") if selected else QColor("#FFFFFF"))
        painter.setPen(QColor("#EEF1F4"))
        painter.drawLine(r.left(), r.bottom(), r.right(), r.bottom())
        base = QFont(option.font)
        base.setPixelSize(13)
        bold = QFont(base)
        bold.setBold(True)
        if d["bad"]:
            painter.fillRect(r.left(), r.top(), 3, r.height(), QColor("#D92D20"))
            painter.setFont(bold)
            painter.setPen(QColor("#D92D20"))
            painter.drawText(QRect(r.left() + 6, r.top(), 10, r.height()), Qt.AlignVCenter | Qt.AlignLeft, "!")
        painter.setFont(base)
        painter.setPen(QColor("#374151"))
        painter.drawText(QRect(r.left() + 16, r.top(), 30, r.height()), Qt.AlignVCenter | Qt.AlignRight, str(d["no"]))
        text_x = r.left() + 54
        right = r.right() - 8
        small = QFont(base)
        small.setPixelSize(11)
        small.setBold(True)
        for letter in reversed(d["chips"]):
            chip = QRect(right - 17, r.top() + (r.height() - 17) // 2, 17, 17)
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor("#E0E7FF"))
            painter.drawRoundedRect(chip, 4, 4)
            painter.setFont(small)
            painter.setPen(QColor("#3730A3"))
            painter.drawText(chip, Qt.AlignCenter, letter)
            right -= 21
        width = max(20, right - text_x - 6)
        painter.setFont(bold)
        painter.setPen(QColor("#111827"))
        title = QFontMetrics(bold).elidedText(d["title"], Qt.ElideRight, width)
        painter.drawText(QRect(text_x, r.top() + 6, width, 18), Qt.AlignVCenter | Qt.AlignLeft, title)
        sub_font = QFont(base)
        sub_font.setPixelSize(11)
        painter.setFont(sub_font)
        painter.setPen(QColor("#6B7280"))
        sub = QFontMetrics(sub_font).elidedText(d["sub"], Qt.ElideRight, width)
        painter.drawText(QRect(text_x, r.top() + 26, width, 16), Qt.AlignVCenter | Qt.AlignLeft, sub)
        painter.restore()


class Fold(QFrame):
    """Bagian yang bisa dilipat: judul, ringkasan satu baris (dipotong …), lencana, dan saklar opsional.
    Isi dibuat saat pertama dibuka (hemat waktu untuk banyak step/data)."""

    def __init__(self, title, build_body=None, switch_text=None, card=False, parent=None):
        super().__init__(parent)
        self.setObjectName("card" if card else "section")
        self._build_body = build_body
        self._body = None
        self._open = False
        outer = QVBoxLayout(self)
        outer.setContentsMargins(8, 4, 8, 6)
        outer.setSpacing(4)
        head = QHBoxLayout()
        head.setSpacing(8)
        self.arrow = QToolButton()
        self.arrow.setObjectName("fold")
        self.arrow.setText(title)
        self.arrow.setArrowType(Qt.RightArrow)
        self.arrow.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.arrow.setCursor(Qt.PointingHandCursor)
        self.arrow.clicked.connect(lambda: self.set_open(not self._open))
        self.summary = ElidedLabel()
        self.summary.setObjectName("hint")
        self.badge = QLabel()
        self.badge.setFixedSize(18, 18)
        self.badge.setAlignment(Qt.AlignCenter)
        self.switch = QCheckBox(switch_text) if switch_text else None
        head.addWidget(self.arrow)
        head.addWidget(self.summary, 1)
        head.addWidget(self.badge)
        if self.switch is not None:
            head.addWidget(self.switch)
        outer.addLayout(head)
        self._holder = QVBoxLayout()
        outer.addLayout(self._holder)

    def set_builder(self, build_body):
        self._build_body = build_body

    def set_title(self, text):
        self.arrow.setText(text)

    def set_summary(self, text):
        self.summary.set_full_text(text)

    def set_badge(self, text, kind="ok"):
        """Lencana "!" merah 18px (teks penjelasan di tooltip) di kolom kanan lebar tetap."""
        if text and kind == "err":
            self.badge.setText("!")
            self.badge.setToolTip(text)
            self.badge.setStyleSheet("color:white;background:#D92D20;border-radius:9px;font-weight:bold;font-size:12px;")
        else:
            self.badge.setText("")
            self.badge.setToolTip("")
            self.badge.setStyleSheet("")

    def is_open(self):
        return self._open

    def body(self):
        return self._body

    def set_open(self, value):
        value = bool(value)
        if value and self._body is None and self._build_body is not None:
            self._body = self._build_body()
            self._holder.addWidget(self._body)
        self._open = value
        if self._body is not None:
            self._body.setVisible(value)
        self.arrow.setArrowType(Qt.DownArrow if value else Qt.RightArrow)


class GroupCard(Fold):
    """Satu Data pada sebuah step, sebagai kartu yang bisa dilipat."""

    def __init__(self, dlg, si, gi, step, g):
        super().__init__("", None, card=True)
        self.dlg, self.si, self.gi, self.step, self.g = dlg, si, gi, step, g
        self.w = {}
        self.sections = {}
        self._form = None
        self.set_builder(self._build)
        self.refresh_head()

    # ---------------------------------------------------------------- tampilan ringkas
    def active_labels(self):
        g = self.g
        return [label for label, on in (("Excel", g["excel_enabled"]), ("Karakter", g["char_enabled"]),
                                         ("Unik", g["unique"])) if on]

    def refresh_head(self):
        g = self.g
        name = g["name"].strip()
        self.set_title(f"Data {self.gi + 1}" + (f" — {name}" if name else ""))
        if g["mode"] == "all":
            mode_text = "Seluruh barcode"
        elif g["mode"] == "operator_name":
            mode_text = "Nama operator"
        else:
            mode_text = f"Karakter ke-{g['start']}, {g['length']} buah"
        active = self.active_labels()
        self.set_summary(mode_text + ("  ·  Validasi: " + ", ".join(active) if active else ""))
        for key, section in self.sections.items():
            section.set_summary(self._section_summary(key))
        self.refresh_badges()

    def _section_summary(self, key):
        g = self.g
        if key == "excel":
            if not g["excel_enabled"]:
                return "Nonaktif"
            scope = "Semua sheet" if g["xlsx_sheet_mode"] == "all" else (g["xlsx_sheet_name"] or "?")
            return (f"Kol {g['xlsx_column'] or '?'} · {scope} · {g['xlsx_match_chars']} awal · "
                    f"{'OK' if g['excel_found_result'] == 'ok' else 'NG'} jika ada")
        if key == "char":
            if not g["char_enabled"]:
                return "Nonaktif"
            return f"Ke-{g['char_start']} s/d {g['char_start'] + g['char_length'] - 1} = {g['char_expected']}"
        if not g["unique"]:
            return "Nonaktif"
        return "Tidak boleh sama dalam satu siklus"

    def own_issues(self):
        return [k for k in self.dlg.issues if k[0] == "step" and k[1] == self.si and k[2] == self.gi]

    def refresh_badges(self):
        own = self.own_issues()
        self.set_badge("Ada isian yang perlu diperbaiki" if own else "", "err")
        for key, section in self.sections.items():
            bad = key == "excel" and any(k[3].startswith("xlsx") for k in own)
            section.set_badge("Ada isian yang perlu diperbaiki" if bad else "", "err")

    # ---------------------------------------------------------------- isi kartu
    def _field(self, label_key, widget):
        return self.dlg.field(("step", self.si, self.gi, label_key), widget)

    def _bind(self, widget):
        if isinstance(widget, QLineEdit):
            widget.textChanged.connect(self._changed)
        elif isinstance(widget, QSpinBox):
            widget.valueChanged.connect(self._changed)
        elif isinstance(widget, QComboBox):
            widget.currentTextChanged.connect(self._changed)
        elif isinstance(widget, QCheckBox):
            widget.toggled.connect(self._changed)
        return widget

    def _update_range_rows(self):
        visible = self.g["mode"] == "length"
        for key in ("start", "length"):
            self._form.setRowVisible(self.w[key], visible)

    def _build(self):
        g, w = self.g, self.w
        body = QWidget()
        form = Form(body)
        self._form = form
        w["name"] = QLineEdit(g["name"])
        w["name"].setPlaceholderText("Nama kolom, misalnya Kode")
        w["mode"] = QComboBox()
        w["mode"].addItem("Jumlah karakter", "length")
        w["mode"].addItem("Seluruh barcode", "all")
        if self.step["input_type"] == "operator":
            w["mode"].addItem("Nama operator", "operator_name")
        w["mode"].setCurrentIndex(max(0, w["mode"].findData(g["mode"])))
        w["start"] = QSpinBox(); w["start"].setRange(1, 9999); w["start"].setValue(g["start"])
        w["length"] = QSpinBox(); w["length"].setRange(1, 9999); w["length"].setValue(g["length"])
        w["judge_ok"] = QLineEdit(g["judge_ok"]); w["judge_ok"].setPlaceholderText("Contoh: PASS / ACCEPTED")
        w["judge_ng"] = QLineEdit(g["judge_ng"]); w["judge_ng"].setPlaceholderText("Contoh: FAIL / REJECT")
        form.addRow("Nama kolom", w["name"])
        form.addRow("Mode pengambilan", w["mode"])
        form.addRow("Mulai karakter ke-", w["start"])
        form.addRow("Jumlah karakter", w["length"])
        self._update_range_rows()
        bad_excel = any(k[3].startswith("xlsx") for k in self.own_issues())
        for key, title, builder in (("excel", "1. Database Excel", self._build_excel),
                                    ("char", "2. Karakter / format barcode", self._build_char),
                                    ("unique", "3. Data unik", self._build_unique)):
            section = Fold(title, builder, switch_text="Aktif")
            flag = {"excel": "excel_enabled", "char": "char_enabled", "unique": "unique"}[key]
            section.switch.setChecked(g[flag])
            section.switch.toggled.connect(lambda on, k=key, f=flag: self._section_toggled(k, f, on))
            self.sections[key] = section
            form.addRow(section)
            if key == "excel" and bad_excel:  # tertutup kecuali yang bermasalah
                section.set_open(True)
        form.addRow("Judge jika valid", w["judge_ok"])
        form.addRow("Judge jika tidak valid", w["judge_ng"])
        for key in ("name", "mode", "start", "length", "judge_ok", "judge_ng"):
            self._bind(w[key])
        self.refresh_head()
        return body

    def _section_toggled(self, key, flag, on):
        self.g[flag] = bool(on)
        section = self.sections[key]
        if section.body() is not None:
            section.body().setEnabled(bool(on))
        if on:
            section.set_open(True)
        self.refresh_head()
        self.dlg.group_changed(self.si)

    def _build_excel(self):
        g, w = self.g, self.w
        body = QWidget()
        f = Form(body)
        w["excel_found"] = QComboBox()
        w["excel_found"].addItem("NG (ditolak)", "ng")
        w["excel_found"].addItem("OK (diterima)", "ok")
        w["excel_found"].setCurrentIndex(max(0, w["excel_found"].findData(g["excel_found_result"])))
        path_row = QWidget()
        pl = QHBoxLayout(path_row)
        pl.setContentsMargins(0, 0, 0, 0)
        w["xlsx_path"] = QLineEdit(g["xlsx_path"])
        w["xlsx_path"].setPlaceholderText("Pilih file database .xlsx")
        browse = Button("Pilih file…")
        pl.addWidget(w["xlsx_path"], 1)
        pl.addWidget(browse)
        w["xlsx_column"] = QLineEdit(g["xlsx_column"]); w["xlsx_column"].setMaximumWidth(100)
        w["xlsx_match_chars"] = QSpinBox(); w["xlsx_match_chars"].setRange(1, 9999)
        w["xlsx_match_chars"].setValue(g["xlsx_match_chars"])
        w["xlsx_sheet_mode"] = QComboBox()
        w["xlsx_sheet_mode"].addItem("Sheet tertentu", "specific")
        w["xlsx_sheet_mode"].addItem("Semua sheet", "all")
        w["xlsx_sheet_mode"].setCurrentIndex(max(0, w["xlsx_sheet_mode"].findData(g["xlsx_sheet_mode"])))
        w["xlsx_sheet_name"] = QComboBox()
        w["xlsx_sheet_name"].setEditable(True)
        w["xlsx_sheet_name"].setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        w["xlsx_sheet_name"].setMinimumWidth(180)
        self._load_sheets(g["xlsx_sheet_name"], quiet=True)
        w["excel_error"] = QLineEdit(g["excel_error"])
        w["excel_popup"] = QCheckBox("Tampilkan pop-up untuk error Excel")
        w["excel_popup"].setChecked(g["excel_popup"])
        f.addRow("Hasil jika barcode ADA di Excel", w["excel_found"])
        f.addRow("File database Excel", self._field_path(path_row))
        f.addRow("Kolom Excel", self._field("xlsx_column", w["xlsx_column"]))
        f.addRow("Cocokkan karakter awal", w["xlsx_match_chars"])
        f.addRow("Cakupan sheet", w["xlsx_sheet_mode"])
        f.addRow("Nama sheet", self._field("xlsx_sheet_name", w["xlsx_sheet_name"]))
        f.addRow("Pesan jika NG", w["excel_error"])
        f.addRow(w["excel_popup"])
        browse.clicked.connect(self._browse_xlsx)
        w["xlsx_path"].editingFinished.connect(lambda: self._load_sheets(w["xlsx_sheet_name"].currentText()))
        w["excel_found"].currentIndexChanged.connect(self._found_changed)
        for key in ("excel_found", "xlsx_path", "xlsx_column", "xlsx_match_chars", "xlsx_sheet_mode",
                    "xlsx_sheet_name", "excel_error", "excel_popup"):
            self._bind(w[key])
        w["xlsx_sheet_mode"].currentIndexChanged.connect(self._sheet_mode_enable)
        self._sheet_mode_enable()
        body.setEnabled(g["excel_enabled"])
        return body

    def _field_path(self, row):
        return self.dlg.field(("step", self.si, self.gi, "xlsx_path"), row, mark=self.w["xlsx_path"])

    def _sheet_mode_enable(self, *_):
        self.w["xlsx_sheet_name"].setEnabled(self.w["xlsx_sheet_mode"].currentData() == "specific")

    def _found_changed(self, *_):
        w = self.w
        ok_mode = w["excel_found"].currentData() == "ok"
        current = w["excel_error"].text().strip()
        if ok_mode and current in ("", EXCEL_MSG_FOUND_NG):
            w["excel_error"].setText(EXCEL_MSG_NOTFOUND_NG)
        elif not ok_mode and current in ("", EXCEL_MSG_NOTFOUND_NG):
            w["excel_error"].setText(EXCEL_MSG_FOUND_NG)

    def _load_sheets(self, keep, quiet=False):
        combo = self.w["xlsx_sheet_name"]
        path = self.w["xlsx_path"].text().strip() if "xlsx_path" in self.w else self.g["xlsx_path"]
        names = []
        try:
            if path and Path(path).is_file():
                names = workbook_sheetnames(path)[0]
        except (OSError, ValueError):
            pass
        combo.blockSignals(True)
        combo.clear()
        combo.addItems(names)
        if keep:
            idx = combo.findText(keep)
            if idx >= 0:
                combo.setCurrentIndex(idx)
            else:
                combo.setEditText(keep)
        combo.blockSignals(False)
        if not quiet:
            self._changed()

    def _browse_xlsx(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Pilih database Excel", self.w["xlsx_path"].text().strip() or str(Path.home()),
            "Excel Workbook (*.xlsx)")
        if path:
            self.w["xlsx_path"].setText(path)
            self._load_sheets(self.w["xlsx_sheet_name"].currentText())

    def _build_char(self):
        g, w = self.g, self.w
        body = QWidget()
        f = Form(body)
        w["char_start"] = QSpinBox(); w["char_start"].setRange(1, 9999); w["char_start"].setValue(g["char_start"])
        w["char_length"] = QSpinBox(); w["char_length"].setRange(1, 9999); w["char_length"].setValue(g["char_length"])
        w["char_expected"] = QLineEdit(g["char_expected"]); w["char_expected"].setPlaceholderText("Contoh: ABC")
        w["char_error"] = QLineEdit(g["char_error"])
        w["char_popup"] = QCheckBox("Tampilkan pop-up untuk error karakter"); w["char_popup"].setChecked(g["char_popup"])
        f.addRow("Mulai dari karakter ke-", w["char_start"])
        f.addRow("Jumlah karakter", w["char_length"])
        f.addRow("Karakter yang diwajibkan", w["char_expected"])
        f.addRow("Pesan error", w["char_error"])
        f.addRow(w["char_popup"])
        for key in ("char_start", "char_length", "char_expected", "char_error", "char_popup"):
            self._bind(w[key])
        body.setEnabled(g["char_enabled"])
        return body

    def _build_unique(self):
        g, w = self.g, self.w
        body = QWidget()
        f = Form(body)
        w["unique_error"] = QLineEdit(g["unique_error"])
        w["unique_popup"] = QCheckBox("Tampilkan pop-up untuk error duplikat"); w["unique_popup"].setChecked(g["unique_popup"])
        f.addRow("Pesan error", w["unique_error"])
        f.addRow(w["unique_popup"])
        for key in ("unique_error", "unique_popup"):
            self._bind(w[key])
        body.setEnabled(g["unique"])
        return body

    # ---------------------------------------------------------------- sinkron ke model
    def _changed(self, *_):
        w, g = self.w, self.g
        if "name" in w:
            g["name"] = w["name"].text(); g["mode"] = w["mode"].currentData()
            g["start"] = w["start"].value(); g["length"] = w["length"].value()
            g["judge_ok"] = w["judge_ok"].text(); g["judge_ng"] = w["judge_ng"].text()
            self._update_range_rows()
        if "xlsx_path" in w:
            g["excel_found_result"] = w["excel_found"].currentData() or "ng"
            g["xlsx_path"] = w["xlsx_path"].text().strip()
            g["xlsx_column"] = w["xlsx_column"].text().strip().upper()
            g["xlsx_match_chars"] = w["xlsx_match_chars"].value()
            g["xlsx_sheet_mode"] = w["xlsx_sheet_mode"].currentData()
            g["xlsx_sheet_name"] = w["xlsx_sheet_name"].currentText().strip()
            g["excel_error"] = w["excel_error"].text(); g["excel_popup"] = w["excel_popup"].isChecked()
            self.sections["excel"].set_title(
                "1. Database Excel — ditemukan = " + ("OK, tidak ditemukan = NG" if g["excel_found_result"] == "ok" else "NG"))
        if "char_start" in w:
            g["char_start"] = w["char_start"].value(); g["char_length"] = w["char_length"].value()
            g["char_expected"] = w["char_expected"].text(); g["char_error"] = w["char_error"].text()
            g["char_popup"] = w["char_popup"].isChecked()
        if "unique_error" in w:
            g["unique_error"] = w["unique_error"].text(); g["unique_popup"] = w["unique_popup"].isChecked()
        self.refresh_head()
        self.dlg.group_changed(self.si)


class StepEditor(QWidget):
    """Editor untuk SATU step terpilih (hanya ini yang dibuat, bukan semua step)."""

    def __init__(self, dlg, index):
        super().__init__()
        self.dlg, self.index = dlg, index
        self.step = dlg.steps[index]
        self.cards = []
        step = self.step
        root = QVBoxLayout(self)
        root.setContentsMargins(4, 4, 12, 4)
        form = Form(label_width=180)
        self.before = QLineEdit(step["name_before"])
        self.after = QLineEdit(step["name_after"])
        self.message = QLineEdit(step["message_before"])
        self.message.setPlaceholderText("Kosong = pesan scan default")
        self.kind = QComboBox()
        self.kind.addItem("Operator ID", "operator")
        self.kind.addItem("Barcode produk", "barcode")
        self.kind.setCurrentIndex(max(0, self.kind.findData(step["input_type"])))
        form.addRow("Nama sebelum scan", self.before)
        form.addRow("Nama sesudah scan", self.after)
        form.addRow("Jenis input", self.kind)
        root.addLayout(form)
        self.advanced = Fold("Lanjutan")
        self.advanced.set_builder(self._advanced_body)
        self._update_advanced_summary()
        root.addWidget(self.advanced)
        if step["message_before"]:
            self.advanced.set_open(True)
        bar = QHBoxLayout()
        self.count = QSpinBox()
        self.count.setRange(0, MAX_GROUPS)
        self.count.setValue(len(step["groups"]))
        expand = Button("Buka semua validasi")
        collapse = Button("Ciutkan semua")
        bar.addWidget(QLabel("Jumlah kelompok data"))
        bar.addWidget(self.count)
        bar.addStretch()
        bar.addWidget(expand)
        bar.addWidget(collapse)
        root.addLayout(bar)
        self.cards_layout = QVBoxLayout()
        self.cards_layout.setSpacing(8)
        root.addLayout(self.cards_layout)
        root.addStretch()
        self._rebuild_cards(open_first=True)
        for widget in (self.before, self.after, self.message):
            widget.textChanged.connect(self._header_changed)
        self.kind.currentIndexChanged.connect(self._kind_changed)
        self.count.valueChanged.connect(self._count_changed)
        expand.clicked.connect(self.expand_all)
        collapse.clicked.connect(self.collapse_all)

    def _advanced_body(self):
        body = QWidget()
        form = Form(body)
        form.addRow("Pesan sebelum scan", self.message)
        return body

    def _update_advanced_summary(self):
        text = self.message.text().strip()
        self.advanced.set_summary(f"Pesan: {text}" if text else "Pesan scan default")

    def expand_all(self):
        for card in self.cards:
            card.set_open(True)
            for section in card.sections.values():
                section.set_open(True)

    def collapse_all(self):
        for card in self.cards:
            for section in card.sections.values():
                section.set_open(False)
            card.set_open(False)

    def _rebuild_cards(self, open_first=False):
        self.dlg.purge_fields()
        for card in self.cards:
            card.setParent(None)
            card.deleteLater()
        self.cards = []
        for gi, g in enumerate(self.step["groups"]):
            card = GroupCard(self.dlg, self.index, gi, self.step, g)
            self.cards.append(card)
            self.cards_layout.addWidget(card)
        if open_first and self.cards:
            self.cards[0].set_open(True)
        self.dlg.apply_issues()

    def _header_changed(self, *_):
        s = self.step
        s["name_before"], s["name_after"], s["message_before"] = self.before.text(), self.after.text(), self.message.text()
        self._update_advanced_summary()
        self.dlg.group_changed(self.index)

    def _kind_changed(self, *_):
        kind = self.kind.currentData()
        before = self.dlg.csv_maps()
        self.dlg.push_undo()
        self.step["input_type"] = kind
        if kind != "operator":
            for g in self.step["groups"]:
                if g["mode"] == "operator_name":
                    g["mode"] = "all"
        self.dlg.after_kind_change(before)

    def _count_changed(self, value):
        groups = self.step["groups"]
        self.dlg.push_undo(("count", self.step["_id"]))  # perubahan beruntun spinbox = satu langkah undo
        while len(groups) > value:
            groups.pop()
        while len(groups) < value:
            groups.append(normalize_group(default_group(len(groups)), len(groups), self.step["input_type"]))
        self._rebuild_cards()
        for card in self.cards[-1:]:
            card.set_open(True)
        self.dlg.group_changed(self.index, structural=True)


class EditProgramDialog(QDialog):
    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Program")
        self.resize(1180, 760)
        self.setMinimumSize(900, 600)
        self.setSizeGripEnabled(True)
        self.setStyleSheet(APP_QSS)
        self.config = json.loads(json.dumps(config))
        raw_steps = self.config.get("steps", [])
        raw_steps = raw_steps if isinstance(raw_steps, list) else []
        self._truncation = truncation_info(raw_steps)
        self._next_id = 0
        self.steps = [self._with_id(normalize_step(s, i)) for i, s in enumerate(raw_steps[:MAX_STEPS])]
        if not self.steps:
            self.steps = [self._with_id(default_step(0))]
        self.issues = {}
        self._fields = {}
        self._fs_cache = {}
        self.fs_checks = 0
        self._undo = None
        self._undo_tag = None
        self._undo_time = 0.0
        self._dirty = False
        self._copied = None
        self._editor = None
        self._current = -1
        self.column_names = {}
        self.target_widgets = []
        self._validate_timer = QTimer(self)
        self._validate_timer.setSingleShot(True)
        self._validate_timer.timeout.connect(self.validate_now)
        self._toast_timer = QTimer(self)
        self._toast_timer.setSingleShot(True)
        self._toast_timer.timeout.connect(self._show_status)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 8)
        root.setSpacing(6)
        main = QHBoxLayout()
        main.setSpacing(0)
        self.nav = QListWidget()
        self.nav.setObjectName("nav")
        self.nav.setFixedWidth(190)
        for text in ("Step", "Penyimpanan CSV", "File .dat", "Umum"):
            self.nav.addItem(text)
        self.stack = QStackedWidget()
        main.addWidget(self.nav)
        main.addWidget(self.stack, 1)
        root.addLayout(main, 1)

        self._build_steps_page()
        self._build_csv_page()
        self._build_dat_page()
        self._build_general_page()

        footer = QHBoxLayout()
        footer.setContentsMargins(12, 0, 12, 0)
        self.toast_label = QLabel()
        self.toast_label.setObjectName("status")
        self.issue_link = Button("Ke masalah pertama ›")
        self.issue_link.setObjectName("link")
        self.issue_link.setCursor(Qt.PointingHandCursor)
        self.issue_link.clicked.connect(self._goto_first_issue)
        self.issue_link.hide()
        self.undo_button = Button("Urungkan perubahan")
        self.undo_button.setToolTip("Membatalkan aksi struktur terakhir (tambah, hapus, geser, tempel, jumlah data). "
                                    "Tidak tersedia setelah kamu mengedit isian.")
        self.undo_button.clicked.connect(self.undo)
        self.undo_button.setEnabled(False)
        self.cancel_button = Button("Batal")
        self.save_button = Button("Simpan Konfigurasi")
        self.save_button.setObjectName("primary")
        self.cancel_button.clicked.connect(self.reject)
        self.save_button.clicked.connect(self.validate_and_accept)
        footer.addWidget(self.toast_label, 1)
        footer.addWidget(self.issue_link)
        footer.addWidget(self.undo_button)
        footer.addWidget(self.cancel_button)
        footer.addWidget(self.save_button)
        root.addLayout(footer)

        self.nav.currentRowChanged.connect(self._page_changed)
        self.nav.setCurrentRow(0)
        self._refresh_list(select=0)
        self.validate_now()
        self._show_status()
        if self._truncation:
            self.toast(self._truncation_text(), bad=True)

    # ------------------------------------------------------------------ util umum
    def _with_id(self, step):
        self._next_id += 1
        step["_id"] = self._next_id
        return step

    def _truncation_text(self):
        shown = "; ".join(self._truncation[:4]) + (f"; dan {len(self._truncation) - 4} lainnya" if len(self._truncation) > 4 else "")
        return f"Peringatan: {shown} akan terpotong jika konfigurasi disimpan."

    def _show_status(self):
        self.toast_label.setProperty("toast", False)
        self.toast_label.setProperty("bad", False)
        repolish(self.toast_label)
        self.toast_label.setText("Ada perubahan belum disimpan" if self._dirty else "Belum ada perubahan")

    def toast(self, text, bad=False):
        self.toast_label.setProperty("toast", True)
        self.toast_label.setProperty("bad", bad)
        repolish(self.toast_label)
        self.toast_label.setText(text)
        self._toast_timer.start(7000)

    def _set_dirty(self):
        if not self._dirty:
            self._dirty = True
            if not self._toast_timer.isActive():
                self._show_status()

    def mark_edited(self, *_):
        """Edit isian: riwayat urungkan dikosongkan supaya tidak memulihkan snapshot lama."""
        self._undo = None
        self._undo_tag = None
        self.undo_button.setEnabled(False)
        self._set_dirty()

    def field(self, key, widget, mark=None):
        """Bungkus widget dengan label error di bawahnya; kotak merah jika key bermasalah."""
        box = QWidget()
        lay = QVBoxLayout(box)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(2)
        label = QLabel()
        label.setObjectName("err")
        label.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        label.hide()
        lay.addWidget(widget)
        lay.addWidget(label)
        self._fields[key] = (mark or widget, label)
        return box

    def purge_fields(self, step_index=None):
        for key in [k for k in self._fields if k[0] == "step"]:
            del self._fields[key]

    def apply_issues(self):
        for key, (widget, label) in list(self._fields.items()):
            message = self.issues.get(key, "")
            try:
                label.setText(message)
                label.setToolTip(message)
                label.setVisible(bool(message))
                if label.parentWidget() is not None:
                    label.parentWidget().updateGeometry()
                if bool(widget.property("invalid")) != bool(message):
                    widget.setProperty("invalid", bool(message))
                    repolish(widget)
            except RuntimeError:
                del self._fields[key]
        if self._editor is not None:
            for card in self._editor.cards:
                try:
                    card.refresh_badges()
                except RuntimeError:
                    pass

    def schedule_validate(self, *_):
        self._validate_timer.start(350)

    def has_issue(self, si, gi):
        return any(k[0] == "step" and k[1] == si and k[2] == gi for k in self.issues)

    def group_changed(self, step_index, structural=False):
        if structural:
            self._set_dirty()
        else:
            self.mark_edited()
        self._update_item(step_index)
        if step_index == self._current:
            self._update_title()
        self.schedule_validate()

    def push_undo(self, tag=None):
        now = time.monotonic()
        if tag is not None and tag == self._undo_tag and self._undo is not None and now - self._undo_time < 1.5:
            self._undo_time = now  # lanjutan perubahan yang sama: tetap satu langkah undo
            return
        self._undo = {"steps": json.loads(json.dumps(self.steps)), "csv": self._csv_live_state(),
                      "current": self._current}
        self._undo_tag, self._undo_time = tag, now
        self.undo_button.setEnabled(True)
        self._set_dirty()

    def undo(self):
        snap = self._undo
        if snap is None:
            return
        self._undo = None
        self._undo_tag = None
        self.undo_button.setEnabled(False)
        self.steps = snap["steps"]
        self._current = max(0, min(snap["current"], len(self.steps) - 1))
        self._after_structure(self._current, state=snap["csv"])
        self.toast("Aksi terakhir diurungkan.")

    def _page_changed(self, row):
        if row < 0:
            return
        if row == 1:
            self._sync_csv_options()
        self.stack.setCurrentIndex(row)

    @staticmethod
    def _scroll(widget):
        area = QScrollArea()
        area.setWidgetResizable(True)
        area.setWidget(widget)
        return area

    # ------------------------------------------------------------------ halaman Step
    def _build_steps_page(self):
        page = QWidget()
        outer = QVBoxLayout(page)
        outer.setContentsMargins(10, 10, 10, 4)
        self.banner = QLabel()
        self.banner.setObjectName("banner")
        self.banner.setWordWrap(True)
        self.banner.setVisible(bool(self._truncation))
        if self._truncation:
            self.banner.setText(self._truncation_text())
        outer.addWidget(self.banner)
        split = QSplitter(Qt.Horizontal)
        left = QWidget()
        ll = QVBoxLayout(left)
        ll.setContentsMargins(0, 0, 6, 0)
        self.step_search = QLineEdit()
        self.step_search.setPlaceholderText("Cari step (nomor atau nama)…")
        self.step_search.setClearButtonEnabled(True)
        self.step_filter = QComboBox()
        for text, data in (("Semua jenis", "all"), ("Hanya Operator ID", "operator"),
                           ("Hanya Barcode produk", "barcode"), ("Yang bermasalah", "bad")):
            self.step_filter.addItem(text, data)
        self.step_list = QListWidget()
        self.step_list.setObjectName("steps")
        self.step_list.setItemDelegate(StepDelegate(self.step_list))
        self.step_list.setUniformItemSizes(True)
        ll.addWidget(self.step_search)
        ll.addWidget(self.step_filter)
        ll.addWidget(self.step_list, 1)
        row = QHBoxLayout()
        row.setSpacing(6)
        buttons = {}
        for name, text, tip, handler in (
                ("add", "+ Tambah", "Tambah step di akhir", self.add_step),
                ("dup", "Duplikat", "Gandakan step terpilih", self.duplicate_step),
                ("up", "▲", "Geser ke atas", lambda: self.move_step(-1)),
                ("down", "▼", "Geser ke bawah", lambda: self.move_step(1))):
            button = Button(text)
            button.setToolTip(tip)
            button.clicked.connect(lambda _=False, h=handler: h())
            row.addWidget(button, 1 if name in ("add", "dup") else 0)
            buttons[name] = button
        self.step_buttons = buttons
        more = QToolButton()
        more.setObjectName("more")
        more.setText("⋯")
        more.setToolTip("Salin, tempel, hapus, alat cepat")
        more.setPopupMode(QToolButton.InstantPopup)
        menu = QMenu(more)
        self.step_actions = {
            "copy": menu.addAction("Salin konfigurasi step ini", self.copy_step),
            "paste": menu.addAction("Tempel konfigurasi ke step ini", self.paste_step),
        }
        menu.addSeparator()
        self.step_actions["del"] = menu.addAction("Hapus step ini", self.delete_step)
        menu.addSeparator()
        self.step_actions["tools"] = menu.addAction("Alat cepat: perintah copy step…", self.show_command_tool)
        self.step_actions["paste"].setEnabled(False)
        more.setMenu(menu)
        row.addWidget(more)
        ll.addLayout(row)
        self.command_fold = Fold("Alat cepat: perintah copy step")
        cbody = QWidget()
        cl = QVBoxLayout(cbody)
        help_label = QLabel("Contoh: copy step 2 for step 3 to 73  |  copy step 3 to step 10\n"
                            "Rentang tujuan inklusif. Bisa dijalankan berkali-kali.")
        help_label.setObjectName("hint")
        help_label.setWordWrap(True)
        self.step_command_input = QLineEdit()
        self.step_command_input.setPlaceholderText("copy step 2 for step 3 to 73")
        self.step_command_button = Button("Jalankan")
        crow = QHBoxLayout()
        crow.addWidget(self.step_command_input, 1)
        crow.addWidget(self.step_command_button)
        cl.addWidget(help_label)
        cl.addLayout(crow)
        self.command_fold.set_builder(lambda: cbody)
        self.command_fold.setVisible(False)
        self.step_command_button.clicked.connect(self.execute_step_command)
        self.step_command_input.returnPressed.connect(self.execute_step_command)
        ll.addWidget(self.command_fold)
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.setSpacing(4)
        self.editor_title = QLabel()
        self.editor_title.setObjectName("h1")
        self.editor_title.setContentsMargins(8, 2, 8, 2)
        self.editor_area = QScrollArea()
        self.editor_area.setWidgetResizable(True)
        rl.addWidget(self.editor_title)
        rl.addWidget(self.editor_area, 1)
        split.addWidget(left)
        split.addWidget(right)
        split.setStretchFactor(0, 0)
        split.setStretchFactor(1, 1)
        split.setSizes([340, 800])
        left.setMinimumWidth(280)
        outer.addWidget(split, 1)
        self.stack.addWidget(page)
        self.step_search.textChanged.connect(self._apply_filter)
        self.step_filter.currentIndexChanged.connect(self._apply_filter)
        self.step_list.currentRowChanged.connect(self._on_row)

    def show_command_tool(self):
        self.command_fold.setVisible(True)
        self.command_fold.set_open(True)
        self.step_command_input.setFocus()

    def _update_title(self):
        step = self.steps[self._current]
        self.editor_title.setText(f"Step {self._current + 1} — {step['name_before'].strip() or '(tanpa nama)'}")

    def _update_step_buttons(self):
        self.step_buttons["up"].setEnabled(self._current > 0)
        self.step_buttons["down"].setEnabled(self._current < len(self.steps) - 1)
        self.step_actions["del"].setEnabled(len(self.steps) > 1)
        self.step_actions["paste"].setEnabled(self._copied is not None)

    def _step_info(self, i, bad):
        step = self.steps[i]
        groups = step["groups"]
        chips = [(letter, label) for letter, label, key in (("E", "Excel", "excel_enabled"),
                 ("K", "Karakter", "char_enabled"), ("U", "Unik", "unique")) if any(g[key] for g in groups)]
        name = step["name_before"].strip() or "(tanpa nama)"
        tag = "Operator" if step["input_type"] == "operator" else "Barcode"
        tip = f"Step {i + 1}: {name}"
        if chips:
            tip += "\nValidasi aktif: " + ", ".join(label for _, label in chips)
        if bad:
            tip += "\nAda isian yang perlu diperbaiki"
        return {"no": i + 1, "title": name, "sub": f"{tag} · {len(groups)} data", "bad": bad,
                "chips": [letter for letter, _ in chips]}, tip

    def _decorate_item(self, item, i, bad):
        info, tip = self._step_info(i, bad)
        item.setData(StepDelegate.ROLE, info)
        item.setText(f"{i + 1}. {info['title']}")
        item.setToolTip(tip)

    def _bad_steps(self):
        return {k[1] for k in self.issues if k[0] == "step"}

    def _update_item(self, i):
        item = self.step_list.item(i)  # baris daftar = nomor step (disembunyikan, tidak dihapus)
        if item is not None:
            self._decorate_item(item, i, i in self._bad_steps())
            self.step_list.update(self.step_list.indexFromItem(item))

    def _refresh_list(self, select=None):
        bad = self._bad_steps()
        self.step_list.blockSignals(True)
        self.step_list.clear()
        for i in range(len(self.steps)):
            item = QListWidgetItem()
            item.setData(Qt.UserRole, i)
            self._decorate_item(item, i, i in bad)
            self.step_list.addItem(item)
        self.step_list.blockSignals(False)
        self._apply_filter()
        if select is not None:
            self._select(select, force=True)

    def _apply_filter(self, *_):
        query = self.step_search.text().strip().lower()
        mode = self.step_filter.currentData()
        bad = self._bad_steps()
        for row in range(self.step_list.count()):
            item = self.step_list.item(row)
            i = item.data(Qt.UserRole)
            step = self.steps[i]
            ok = True
            if query:
                ok = query == str(i + 1) or query in step["name_before"].lower() or query in step["name_after"].lower()
            if ok and mode in ("operator", "barcode"):
                ok = step["input_type"] == mode
            if ok and mode == "bad":
                ok = i in bad
            item.setHidden(not ok)

    def _on_row(self, row):
        item = self.step_list.item(row)
        if item is not None:
            self._select(item.data(Qt.UserRole))

    def _select(self, index, force=False):
        index = max(0, min(index, len(self.steps) - 1))
        self.step_list.blockSignals(True)
        self.step_list.setCurrentRow(index)
        self.step_list.blockSignals(False)
        if index == self._current and not force and self._editor is not None:
            return
        self._current = index
        self.refresh_step_editor()

    def refresh_step_editor(self):
        self.purge_fields()
        self._editor = StepEditor(self, self._current)
        old = self.editor_area.takeWidget()
        if old is not None:
            old.deleteLater()
        self.editor_area.setWidget(self._editor)
        self._update_title()
        self._update_step_buttons()
        self.apply_issues()
        self._update_item(self._current)
        self.schedule_validate()

    # ------------------------------------------------------------------ aksi step
    def csv_maps(self):
        """Peta nomor urut -> id step untuk pilihan CSV (sebelum/sesudah aksi struktur)."""
        prod, pos, n = {}, {}, 0
        for i, step in enumerate(self.steps):
            pos[i + 1] = step["_id"]
            if step["input_type"] == "barcode":
                n += 1
                prod[n] = step["_id"]
        return {"prod": prod, "pos": pos, "labels": dict(self.column_options())}

    def _after_structure(self, select, before=None, state=None):
        """Bangun ulang daftar/editor, petakan ulang pilihan CSV; return nomor step lama yang pilihan CSV-nya hilang."""
        self._refresh_list(select=select)
        remap = CsvRemap(before, self.csv_maps()) if before is not None else None
        removed = self._sync_csv_options(state=state, remap=remap)
        self.validate_now()
        return removed

    def after_kind_change(self, before):
        self._set_dirty()
        removed = self._sync_csv_options(remap=CsvRemap(before, self.csv_maps()))
        self.refresh_step_editor()
        self.validate_now()
        if removed:
            self.toast("Pilihan CSV untuk " + ", ".join(f"Step {n}" for n in removed) + " dihapus (step bukan barcode lagi).", bad=True)

    def add_step(self):
        if len(self.steps) >= MAX_STEPS:
            self.toast(f"Maksimal {MAX_STEPS} step.", bad=True)
            return
        before = self.csv_maps()
        self.push_undo()
        self.steps.append(self._with_id(default_step(len(self.steps))))
        self._after_structure(len(self.steps) - 1, before)
        self.toast(f"Step {len(self.steps)} ditambahkan.")

    def duplicate_step(self):
        if len(self.steps) >= MAX_STEPS:
            self.toast(f"Maksimal {MAX_STEPS} step.", bad=True)
            return
        before = self.csv_maps()
        self.push_undo()
        copy = json.loads(json.dumps(self.steps[self._current]))
        self.steps.insert(self._current + 1, self._with_id(copy))
        self._after_structure(self._current + 1, before)
        self.toast(f"Step {self._current} diduplikat menjadi Step {self._current + 1}.")

    def delete_step(self):
        if len(self.steps) <= 1:
            self.toast("Minimal harus ada 1 step.", bad=True)
            return
        before = self.csv_maps()
        self.push_undo()
        number = self._current + 1
        del self.steps[self._current]
        removed = self._after_structure(min(self._current, len(self.steps) - 1), before)
        text = f"Step {number} dihapus"
        if removed:
            text += ". " + "; ".join(f"Pilihan CSV untuk Step {n} dihapus" for n in removed)
        self.toast(text + ". Klik Urungkan perubahan untuk membatalkan (selama belum mengedit isian).", bad=bool(removed))

    def move_step(self, delta):
        target = self._current + delta
        if target < 0 or target >= len(self.steps):
            return
        before = self.csv_maps()
        self.push_undo()
        self.steps[self._current], self.steps[target] = self.steps[target], self.steps[self._current]
        self._after_structure(target, before)
        self.toast(f"Step dipindah ke posisi {target + 1}; pilihan CSV ikut berpindah.")

    def copy_step(self):
        self._copied = json.loads(json.dumps(self.steps[self._current]))
        self.step_actions["paste"].setEnabled(True)
        self.toast(f"Konfigurasi Step {self._current + 1} disalin. Pilih step tujuan lalu Tempel (menu ⋯).")

    def paste_step(self):
        if not self._copied:
            self.toast("Belum ada konfigurasi yang disalin. Pilih Salin pada step sumber dulu.", bad=True)
            return
        before = self.csv_maps()
        self.push_undo()
        new = json.loads(json.dumps(self._copied))
        new["_id"] = self.steps[self._current]["_id"]  # slot yang sama: pilihan CSV tetap menunjuk slot ini
        self.steps[self._current] = new
        self._after_structure(self._current, before)
        self.toast(f"Konfigurasi ditempel ke Step {self._current + 1}.")

    def execute_step_command(self):
        """Perintah copy step (perilaku sama dengan versi lama), tanpa pop-up."""
        command = self.step_command_input.text().strip()
        pattern = re.compile(
            r"^copy\s+(?:step\s*)?(\d+)\s+"
            r"(?:for\s+step\s+|to\s+(?:step\s*)?)"
            r"(\d+)(?:\s+to\s+(?:step\s*)?(\d+)|\s*-\s*(?:step\s*)?(\d+))?$",
            re.IGNORECASE,
        )
        match = pattern.match(command)
        if not match:
            self.toast("Format perintah tidak dikenali. Contoh: copy step 2 for step 3 to 73", bad=True)
            self.step_command_input.setFocus()
            return
        source = int(match.group(1))
        first_target = int(match.group(2))
        last_target = int(match.group(3) or match.group(4) or first_target)
        total = len(self.steps)
        if source < 1 or source > total or first_target < 1 or last_target < 1 or last_target > total:
            self.toast(f"Nomor Step harus berada antara 1 sampai {total}.", bad=True)
            return
        if last_target < first_target:
            self.toast("Nomor tujuan akhir harus lebih besar atau sama dengan tujuan awal.", bad=True)
            return
        before = self.csv_maps()
        self.push_undo()
        source_config = json.loads(json.dumps(self.steps[source - 1]))
        for number in range(first_target, last_target + 1):
            new = json.loads(json.dumps(source_config))
            new["_id"] = self.steps[number - 1]["_id"]
            self.steps[number - 1] = new
        self.step_command_input.clear()
        self._after_structure(self._current, before)
        count = last_target - first_target + 1
        self.toast(f"Step {source} disalin ke Step {first_target}"
                   + (f" sampai {last_target}" if last_target != first_target else "") + f" ({count} step).")

    # ------------------------------------------------------------------ halaman CSV
    def column_options(self):
        barcode_steps = [i for i, step in enumerate(self.steps) if step["input_type"] == "barcode"]
        options = [("no", "No"), ("hasil_tes", "Hasil Tes"), ("waktu_mulai", "Waktu tes dimulai"),
                   ("kode_pekerja", "Kode Pekerja"), ("nama_pekerja", "Nama Pekerja")]
        options += [(f"kode_produk_{n}", f"Kode produk {n} (Step {idx + 1})")
                    for n, idx in enumerate(barcode_steps, start=1)]
        options += [("lini", "Lini"), ("waktu_berakhir", "Waktu tes berakhir")]
        return options

    def _data_choices(self):
        choices = []
        for step_no, step in enumerate(self.steps, start=1):
            if step["input_type"] != "barcode":
                continue
            for group_no, group in enumerate(step["groups"], start=1):
                name = group["name"].strip() or f"Data {group_no}"
                choices.append((f"{name} — Data {group_no}, Step {step_no}", f"data_step_{step_no}_group_{group_no}"))
        return choices

    def _build_csv_page(self):
        saving = self.config.get("saving")
        saving = saving if isinstance(saving, dict) else {}
        legacy_filename = saving.get("filename", "hasil_pengujian.csv")
        saved_targets = saving.get("targets")
        if not saved_targets:
            saved_targets = [{
                "destination": saving.get("destination", ""), "filename": legacy_filename,
                "filename_part1": "none", "filename_part2": "none",
                "filename_text": Path(str(legacy_filename)).stem or "hasil_pengujian",
                "date_prefix": saving.get("date_prefix", False),
            }]
        self.saved_targets = [t for t in saved_targets if isinstance(t, dict)] or [{}]
        self._legacy_columns = saving.get("columns", [])
        self._configured_names = saving.get("column_names", {})
        if not isinstance(self._configured_names, dict):
            self._configured_names = {}

        content = QWidget()
        lay = QVBoxLayout(content)
        lay.setContentsMargins(16, 12, 16, 12)
        title = QLabel("Penyimpanan CSV")
        title.setObjectName("h1")
        lay.addWidget(title)
        self.save_enabled = QCheckBox("Simpan hasil setiap siklus selesai")
        self.save_enabled.setChecked(bool(saving.get("enabled", False)))
        lay.addWidget(self.save_enabled)
        self.csv_body = QWidget()
        bl = QVBoxLayout(self.csv_body)
        bl.setContentsMargins(0, 0, 0, 0)
        top = QHBoxLayout()
        self.target_count = QSpinBox()
        self.target_count.setRange(1, 1000)
        self.target_count.setValue(max(1, min(1000, len(self.saved_targets))))
        top.addWidget(QLabel("Jumlah file tujuan"))
        top.addWidget(self.target_count)
        top.addStretch()
        bl.addLayout(top)
        self.targets_tabs = QTabWidget()
        self.targets_tabs.setMinimumHeight(400)
        bl.addWidget(self.targets_tabs)
        fixed = QGroupBox("Isi tetap")
        fl = Form(fixed, label_width=170)
        self.line_input = QLineEdit(str(saving.get("line", "")))
        self.line_input.setPlaceholderText("contoh: Lini 1")
        self.line_input.setMaximumWidth(320)
        self.result_input = QLineEdit(str(saving.get("result_text", "OK")))
        self.result_input.setPlaceholderText("Contoh: LULUS")
        self.result_input.setMaximumWidth(320)
        fl.addRow("Lini", self.line_input)
        fl.addRow("Isi kolom Hasil Tes", self.field(("csv", "*", "result_text"), self.result_input))
        bl.addWidget(fixed)
        self.columns_fold = Fold("Nama kolom CSV (berlaku untuk semua file)")
        cbody = QWidget()
        self.columns_layout = Form(cbody, label_width=170)
        self.columns_fold.set_builder(lambda: cbody)
        bl.addWidget(self.columns_fold)
        bl.addStretch()
        lay.addWidget(self.csv_body, 1)
        self.stack.addWidget(self._scroll(content))
        self.rebuild_targets(initial=True)
        self._csv_enabled_changed(edit=False)
        self.save_enabled.toggled.connect(self._csv_enabled_changed)
        self.target_count.valueChanged.connect(self.rebuild_targets)
        self.line_input.textChanged.connect(self._csv_edited)
        self.result_input.textChanged.connect(self._csv_edited)

    def _csv_edited(self, *_):
        self.mark_edited()
        self.schedule_validate()

    def _csv_enabled_changed(self, *_, edit=True):
        self.csv_body.setEnabled(self.save_enabled.isChecked())
        if edit:
            self.mark_edited()
        self.schedule_validate()

    def _live_targets(self):
        return [{
            "destination": t["destination"].text().strip(),
            "filename": t["filename"].text().strip(),
            "filename_part1": t["filename_part1"].currentData(),
            "filename_part2": t["filename_part2"].currentData(),
            "filename_data1": t["filename_data1"].currentData(),
            "filename_data2": t["filename_data2"].currentData(),
            "filename_text": t["filename_text"].text().strip(),
            "date_prefix": t["date_prefix"].isChecked(),
            "filename_source": t["filename_source"].currentData(),
            "columns": [k for k, box in t["column_checks"].items() if box.isChecked()],
        } for t in self.target_widgets]

    def _csv_live_state(self):
        """Pilihan CSV saat ini (kolom, data nama file, nama kolom) untuk dipetakan ulang/dipulihkan."""
        targets = []
        for t in self.target_widgets:
            if t.get("_populated"):
                targets.append({"columns": [k for k, b in t["column_checks"].items() if b.isChecked()],
                                "data1": t["filename_data1"].currentData() or "",
                                "data2": t["filename_data2"].currentData() or ""})
            else:
                old = t["_old"]
                columns = old.get("columns", self._legacy_columns)
                targets.append({"columns": [k for k in columns if isinstance(k, str)] if isinstance(columns, list) else [],
                                "data1": old.get("filename_data1", "") or "", "data2": old.get("filename_data2", "") or ""})
        return {"targets": targets, "names": {k: f.text() for k, f in self.column_names.items()}}

    def rebuild_targets(self, *_args, initial=False):
        old_targets = self._live_targets()
        if not old_targets:
            old_targets = [dict(item) for item in self.saved_targets]
            for item in old_targets:
                item.setdefault("columns", self._legacy_columns)
        self.targets_tabs.clear()
        self.target_widgets = []
        for key in [k for k in self._fields if k[0] == "csv" and k[1] != "*"]:
            del self._fields[key]
        for index in range(self.target_count.value()):
            old = old_targets[index] if index < len(old_targets) else {
                "destination": old_targets[0].get("destination", "") if old_targets else "",
                "filename": f"hasil_pengujian_{index + 1}.csv" if index else "hasil_pengujian.csv",
                "filename_part1": "none", "filename_part2": "none",
                "filename_text": f"hasil_pengujian_{index + 1}" if index else "hasil_pengujian",
                "date_prefix": False, "filename_source": "", "columns": [],
            }
            self._make_target(index, old)
        self._sync_csv_options()
        if not initial:
            self.mark_edited()
            self.schedule_validate()

    def _make_target(self, index, old):
        tab = QWidget()
        tl = QHBoxLayout(tab)
        left = QWidget()
        form = Form(left, label_width=140)
        dest_row = QWidget()
        dl = QHBoxLayout(dest_row)
        dl.setContentsMargins(0, 0, 0, 0)
        destination = QLineEdit(str(old.get("destination", "")))
        destination.setPlaceholderText("Pilih folder yang sudah ada")
        browse = Button("Pilih folder…")
        browse.clicked.connect(lambda _=False, f=destination: self.browse_target_destination(f))
        dl.addWidget(destination, 1)
        dl.addWidget(browse)
        form.addRow("Folder tujuan", self.field(("csv", index, "destination"), dest_row, mark=destination))
        part1, part2 = QComboBox(), QComboBox()
        for combo in (part1, part2):
            combo.addItem("Tidak ada", "none")
            combo.addItem("Tanggal (ddmmyyyy)", "date")
            combo.addItem("Data hasil scan", "data")
        part1.setCurrentIndex(max(0, part1.findData(old.get("filename_part1", "none"))))
        part2.setCurrentIndex(max(0, part2.findData(old.get("filename_part2", "none"))))
        text = QLineEdit(str(old.get("filename_text", Path(str(old.get("filename", "hasil_pengujian.csv"))).stem
                                      or "hasil_pengujian")))
        text.setPlaceholderText("Contoh: hasil_pengujian")
        data1, data2 = QComboBox(), QComboBox()
        for combo in (data1, data2):
            combo.addItem("Pilih data ke x step ke x", "")
        form.addRow("Bagian nama file 1", part1)
        form.addRow("Data untuk bagian 1", data1)
        form.addRow("Bagian nama file 2", part2)
        form.addRow("Data untuk bagian 2", data2)
        form.addRow("Teks bebas (bagian 3)", self.field(("csv", index, "filename_text"), text))
        preview = QLabel()
        preview.setObjectName("preview")
        preview.setTextFormat(Qt.PlainText)
        preview.setWordWrap(True)
        form.addRow("Pratinjau nama file", preview)
        # Kontrol lama yang disimpan demi kompatibilitas config lama.
        source = QComboBox()
        source.addItem("Tidak ada tambahan", "")
        date_prefix = QCheckBox()
        columns_group = QGroupBox("Kolom yang disimpan ke file ini")
        cl = QVBoxLayout(columns_group)
        search = QLineEdit()
        search.setPlaceholderText("Cari kolom…")
        search.setClearButtonEnabled(True)
        cbar = QHBoxLayout()
        all_btn, none_btn = Button("Pilih semua"), Button("Kosongkan")
        cbar.addWidget(all_btn)
        cbar.addWidget(none_btn)
        cbar.addStretch()
        cl.addWidget(search)
        cl.addLayout(cbar)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        cwidget = QWidget()
        cgrid = QVBoxLayout(cwidget)
        scroll.setWidget(cwidget)
        cl.addWidget(scroll, 1)
        columns_error = QLabel()
        columns_error.setObjectName("err")
        columns_error.hide()
        cl.addWidget(columns_error)
        self._fields[("csv", index, "columns")] = (columns_group, columns_error)
        tl.addWidget(left, 3)
        tl.addWidget(columns_group, 2)
        target = {
            "destination": destination, "filename": text, "filename_part1": part1, "filename_part2": part2,
            "filename_data1": data1, "filename_data2": data2, "filename_text": text, "filename_source": source,
            "date_prefix": date_prefix, "browse": browse, "column_checks": {},
            "_old": old, "_cgrid": cgrid, "_preview": preview, "_form": form, "_search": search,
        }
        self.target_widgets.append(target)
        for combo in (part1, part2, data1, data2):
            combo.currentIndexChanged.connect(lambda *_a, t=target: self._target_edited(t))
        text.textChanged.connect(lambda *_a, t=target: self._target_edited(t))
        destination.textChanged.connect(self._csv_edited)
        search.textChanged.connect(lambda *_a, t=target: self._filter_columns(t))
        all_btn.clicked.connect(lambda _=False, t=target: self._set_all_columns(t, True))
        none_btn.clicked.connect(lambda _=False, t=target: self._set_all_columns(t, False))
        self.targets_tabs.addTab(tab, f"File {index + 1}")

    def _filter_columns(self, target):
        query = target["_search"].text().strip().lower()
        for box in target["column_checks"].values():
            box.setVisible(not query or query in box.text().lower())

    def _set_all_columns(self, target, on):
        for box in target["column_checks"].values():
            if not box.isHidden():
                box.setChecked(on)

    def _target_edited(self, target):
        self.mark_edited()
        self._target_changed(target)

    def _target_changed(self, target):
        for n in (1, 2):
            target["_form"].setRowVisible(target[f"filename_data{n}"], target[f"filename_part{n}"].currentData() == "data")
        self._update_preview(target)
        self.schedule_validate()

    def _update_preview(self, target):
        parts = []
        for n in (1, 2):
            kind = target[f"filename_part{n}"].currentData()
            if kind == "date":
                parts.append(datetime.now().strftime("%d%m%Y"))
            elif kind == "data" and target[f"filename_data{n}"].currentData():
                label = target[f"filename_data{n}"].currentText().split(" — ")[0]
                parts.append(f"<{label}>")
            else:
                parts.append("")
        base = f"{parts[0]}_{parts[1]}" if parts[0] and parts[1] else (parts[0] or parts[1])
        base += target["filename_text"].text().strip()
        target["_preview"].setText("Contoh: " + (base or "hasil_pengujian") + ".csv")

    def _sync_csv_options(self, state=None, remap=None):
        """Isi ulang pilihan data nama file, kolom CSV, dan nama kolom dari step terbaru.
        remap: petakan ulang pilihan yang terikat nomor step. Return nomor step lama yang pilihannya hilang."""
        if state is None:
            state = self._csv_live_state()
        removed = set()
        if remap is not None:
            state, removed = remap.apply(state)
            self._configured_names = {k: v for k, v in self._configured_names.items() if not k.startswith("kode_produk_")}
        choices = self._data_choices()
        dropped = False
        for target, tstate in zip(self.target_widgets, state["targets"]):
            first = not target.get("_populated")
            for n in (1, 2):
                combo = target[f"filename_data{n}"]
                saved = tstate[f"data{n}"]
                combo.blockSignals(True)
                combo.clear()
                combo.addItem("Pilih data ke x step ke x", "")
                for label, key in choices:
                    combo.addItem(label, key)
                index = combo.findData(saved)
                if saved and index < 0 and not first:
                    dropped = True
                combo.setCurrentIndex(max(0, index))
                combo.blockSignals(False)
            grid = target["_cgrid"]
            while grid.count():
                item = grid.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            target["column_checks"] = {}
            for key, label in self.column_options():
                box = QCheckBox(label)
                box.setChecked(key in tstate["columns"])
                box.toggled.connect(self._csv_edited)
                grid.addWidget(box)
                target["column_checks"][key] = box
            grid.addStretch()
            target["_populated"] = True
            self._filter_columns(target)
            self._target_changed(target)
        live_names = state["names"]
        while self.columns_layout.rowCount():
            self.columns_layout.removeRow(0)
        self.column_names = {}
        for key, default_label in self.column_options():
            field = QLineEdit(live_names.get(key, self._configured_names.get(key, default_label)))
            field.setPlaceholderText("Nama kolom CSV")
            field.textChanged.connect(self.mark_edited)
            self.columns_layout.addRow(default_label, field)
            self.column_names[key] = field
        if dropped and not removed:
            self.toast("Pilihan data nama file yang sudah tidak ada dikosongkan.", bad=True)
        return sorted(removed)

    def browse_target_destination(self, field):
        folder = QFileDialog.getExistingDirectory(self, "Pilih folder tujuan", field.text().strip() or str(Path.home()))
        if folder:
            field.setText(folder)

    # ------------------------------------------------------------------ halaman .dat & umum
    def _build_dat_page(self):
        dat = self.config.get("dat_saving", {}) if isinstance(self.config.get("dat_saving"), dict) else {}
        content = QWidget()
        lay = QVBoxLayout(content)
        lay.setContentsMargins(16, 12, 16, 12)
        title = QLabel("File .dat")
        title.setObjectName("h1")
        help_label = QLabel("Satu kolom ke bawah. Baris pertama berisi nama file; setiap sequence yang selesai menambahkan "
                            "barcode lengkap dari step ber-mode “Seluruh barcode” di baris berikutnya.")
        help_label.setObjectName("hint")
        help_label.setWordWrap(True)
        self.dat_enabled = QCheckBox("Simpan file .dat setiap sequence selesai")
        self.dat_enabled.setChecked(bool(dat.get("enabled", False)))
        self.dat_destination = QLineEdit(str(dat.get("destination", "")))
        self.dat_destination.setPlaceholderText("Pilih folder tujuan")
        self.dat_browse = Button("Pilih folder…")
        self.dat_browse.clicked.connect(lambda: self.browse_target_destination(self.dat_destination))
        row = QWidget()
        rl = QHBoxLayout(row)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.addWidget(self.dat_destination, 1)
        rl.addWidget(self.dat_browse)
        self.dat_filename = QLineEdit(str(dat.get("filename", "")))
        self.dat_filename.setPlaceholderText("Nama file tanpa .dat, contoh: LOT_A")
        self.dat_body = QWidget()
        form = Form(self.dat_body, label_width=140)
        form.addRow("Folder tujuan", self.field(("dat", "destination"), row, mark=self.dat_destination))
        form.addRow("Nama file", self.field(("dat", "filename"), self.dat_filename))
        lay.addWidget(title)
        lay.addWidget(help_label)
        lay.addWidget(self.dat_enabled)
        lay.addWidget(self.dat_body)
        lay.addStretch()
        self.update_dat_controls(edit=False)
        self.dat_enabled.toggled.connect(self.update_dat_controls)
        self.dat_destination.textChanged.connect(self._csv_edited)
        self.dat_filename.textChanged.connect(self._csv_edited)
        self.stack.addWidget(self._scroll(content))

    def update_dat_controls(self, *_, edit=True):
        self.dat_body.setEnabled(self.dat_enabled.isChecked())
        if edit:
            self.mark_edited()
        self.schedule_validate()

    def _build_general_page(self):
        content = QWidget()
        lay = QVBoxLayout(content)
        lay.setContentsMargins(16, 12, 16, 12)
        title = QLabel("Umum")
        title.setObjectName("h1")
        lay.addWidget(title)
        form = Form(label_width=140)
        self.title_input = QLineEdit(str(self.config.get("display_title", "BARCODE DATA INPUT")))
        self.title_input.setPlaceholderText("Judul pada tampilan utama")
        form.addRow("Judul tampilan utama", self.title_input)
        lay.addLayout(form)
        end_cfg = self.config.get("end_popup", {}) if isinstance(self.config.get("end_popup"), dict) else {}
        end_box = QGroupBox("Pop up akhir sequence")
        el = Form(end_box, label_width=140)
        self.end_popup_enabled = QCheckBox("Tampilkan pop up setelah step terakhir valid")
        self.end_popup_enabled.setChecked(bool(end_cfg.get("enabled", False)))
        self.end_popup_message = QLineEdit(str(end_cfg.get("message", "Data telah tersimpan.")))
        self.end_popup_message.setPlaceholderText("Contoh: Data telah tersimpan.")
        self.end_popup_enabled.toggled.connect(self.end_popup_message.setEnabled)
        self.end_popup_message.setEnabled(self.end_popup_enabled.isChecked())
        el.addRow(self.end_popup_enabled)
        el.addRow("Isi pesan", self.end_popup_message)
        lay.addWidget(end_box)
        ip_cfg = self.config.get("ip_link", {}) if isinstance(self.config.get("ip_link"), dict) else {}
        ip_box = QGroupBox("Link IP")
        il = Form(ip_box, label_width=140)
        self.ip_enabled = QCheckBox("Program hanya bisa bekerja jika tersambung ke IP ini")
        self.ip_enabled.setChecked(bool(ip_cfg.get("enabled", False)))
        self.ip_host = QLineEdit(str(ip_cfg.get("host", "")))
        self.ip_host.setPlaceholderText("Contoh: 192.168.1.10")
        self.ip_port = QSpinBox()
        self.ip_port.setRange(0, 65535)
        try:
            self.ip_port.setValue(int(ip_cfg.get("port", 0)))
        except (TypeError, ValueError):
            self.ip_port.setValue(0)
        self.ip_port.setToolTip("0 = cek dengan ping. Isi port (mis. 445 untuk folder share) untuk cek koneksi TCP.")
        self.ip_test_button = Button("Tes Koneksi")
        self.ip_test_button.clicked.connect(self.test_ip_connection)
        self.ip_result = QLabel()
        trow = QHBoxLayout()
        trow.addWidget(self.ip_test_button)
        trow.addWidget(self.ip_result, 1)
        il.addRow(self.ip_enabled)
        il.addRow("IP / Host", self.ip_host)
        il.addRow("Port (0 = hanya ping)", self.ip_port)
        il.addRow(trow)
        lay.addWidget(ip_box)
        lay.addStretch()
        self.ip_enabled.toggled.connect(self._ip_controls)
        self._ip_controls(self.ip_enabled.isChecked())
        for widget in (self.title_input, self.end_popup_message, self.ip_host):
            widget.textChanged.connect(self.mark_edited)
        for widget in (self.end_popup_enabled, self.ip_enabled):
            widget.toggled.connect(self.mark_edited)
        self.ip_port.valueChanged.connect(self.mark_edited)
        self.stack.addWidget(self._scroll(content))

    def _ip_controls(self, on):
        for widget in (self.ip_host, self.ip_port, self.ip_test_button):
            widget.setEnabled(bool(on))

    def test_ip_connection(self):
        ok, reason = check_ip_connection(self.ip_host.text(), self.ip_port.value())
        host = self.ip_host.text().strip() or "-"
        if ok:
            self.ip_result.setText(f"Tersambung ke {host}.")
            self.ip_result.setStyleSheet("color:#14532D;")
        else:
            self.ip_result.setText(f"Tidak tersambung ke {host}. {reason}")
            self.ip_result.setStyleSheet("color:#B42318;")

    # ------------------------------------------------------------------ validasi
    def _fs(self, kind, text):
        """Cek file/folder/sheet dengan cache singkat (FS_TTL): path share yang lambat tidak membekukan UI
        pada setiap perubahan; path yang berubah dicek ulang, dan Simpan selalu memaksa cek baru."""
        key = (kind, text)
        now = time.monotonic()
        hit = self._fs_cache.get(key)
        if hit is not None and now - hit[0] < FS_TTL:
            return hit[1]
        self.fs_checks += 1
        try:
            path = Path(text).expanduser()
            if kind == "file":
                value = path.is_file()
            elif kind == "dir":
                value = path.is_dir()
            else:
                value = workbook_sheetnames(path)
        except (OSError, ValueError):
            value = ([], "File tidak dapat dibaca.") if kind == "sheets" else False
        if len(self._fs_cache) > 500:
            self._fs_cache.clear()
        self._fs_cache[key] = (now, value)
        return value

    def collect_issues(self):
        issues = {}
        for si, step in enumerate(self.steps):
            for gi, g in enumerate(step["groups"]):
                if not g["excel_enabled"]:
                    continue
                path_text = g["xlsx_path"].strip()
                file_ok = self._fs("file", path_text)
                if not file_ok:
                    issues[("step", si, gi, "xlsx_path")] = "Pilih file .xlsx yang tersedia."
                column = g["xlsx_column"].strip().upper()
                default_msg = "Isi huruf kolom, mis. G atau AA."
                if not column or not column.isalpha():
                    issues[("step", si, gi, "xlsx_column")] = default_msg
                else:
                    try:
                        column_index_from_string(column)
                    except Exception as exc:
                        issues[("step", si, gi, "xlsx_column")] = str(exc) or default_msg
                if g["xlsx_sheet_mode"] == "specific":
                    name = g["xlsx_sheet_name"].strip()
                    if not name:
                        issues[("step", si, gi, "xlsx_sheet_name")] = "Isi nama sheet."
                    elif file_ok:
                        names, error = self._fs("sheets", path_text)
                        if error:
                            issues[("step", si, gi, "xlsx_sheet_name")] = error
                        elif name not in names:
                            issues[("step", si, gi, "xlsx_sheet_name")] = f"Sheet '{name}' tidak ditemukan."
        if self.save_enabled.isChecked():
            for index, target in enumerate(self._live_targets()):
                if not self._fs("dir", target["destination"]):
                    issues[("csv", index, "destination")] = "Folder tujuan tidak ditemukan. Pilih folder yang sudah ada."
                if any(c in target["filename_text"] for c in BAD_FILENAME_CHARS):
                    issues[("csv", index, "filename_text")] = 'Tidak boleh memuat karakter \\ / : * ? " < > |.'
                if not target["columns"]:
                    issues[("csv", index, "columns")] = "Pilih minimal satu kolom."
            if not self.result_input.text().strip():
                issues[("csv", "*", "result_text")] = "Isi nilai yang disimpan pada kolom Hasil Tes."
        if self.dat_enabled.isChecked():
            folder_text = self.dat_destination.text().strip()
            name = self.dat_filename.text().strip()
            if not folder_text or not self._fs("dir", folder_text):
                issues[("dat", "destination")] = "Pilih folder tujuan .dat yang sudah ada."
            if not name or name.lower() == ".dat":
                issues[("dat", "filename")] = "Isi nama file .dat."
            elif any(c in name for c in BAD_FILENAME_CHARS):
                issues[("dat", "filename")] = 'Tidak boleh memuat karakter \\ / : * ? " < > |.'
        return issues

    def validate_now(self, force=False):
        self._validate_timer.stop()
        if force:
            self._fs_cache.clear()
        self.issues = self.collect_issues()
        self.apply_issues()
        counts = {"step": 0, "csv": 0, "dat": 0}
        for key in self.issues:
            counts[key[0]] += 1
        for row, section in ((0, "step"), (1, "csv"), (2, "dat")):
            item = self.nav.item(row)
            base = ("Step", "Penyimpanan CSV", "File .dat")[row]
            item.setText(base + (f"   ({counts[section]} masalah)" if counts[section] else ""))
            item.setForeground(QColor("#FCA5A5") if counts[section] else QColor("#E5E7EB"))
        bad = self._bad_steps()
        for i in range(self.step_list.count()):
            item = self.step_list.item(i)
            info = item.data(StepDelegate.ROLE)
            if info is not None and info["bad"] != (i in bad):
                self._decorate_item(item, i, i in bad)
        self.step_list.viewport().update()
        self._apply_filter()
        total = len(self.issues)
        self.save_button.setEnabled(total == 0)
        self.save_button.setToolTip(f"{total} isian belum valid" if total else "Simpan dan tutup")
        self.issue_link.setVisible(total > 0)
        return self.issues

    def _goto_first_issue(self):
        if not self.issues:
            return
        key = next(iter(self.issues))
        widget = None
        if key[0] == "step":
            self.nav.setCurrentRow(0)
            self._select(key[1], force=self._current != key[1])
            if self._editor is not None and key[2] < len(self._editor.cards):
                card = self._editor.cards[key[2]]
                card.set_open(True)
                section = card.sections.get("excel")
                if section is not None:
                    section.set_open(True)
                self.apply_issues()
            widget = self._fields.get(key, (None, None))[0]
        elif key[0] == "csv":
            self.nav.setCurrentRow(1)
            if key[1] != "*":
                self.targets_tabs.setCurrentIndex(key[1])
            widget = self._fields.get(key, (None, None))[0]
        else:
            self.nav.setCurrentRow(2)
            widget = self._fields.get(key, (None, None))[0]
        if widget is not None:
            def show():
                try:
                    widget.setFocus()
                    if key[0] == "step":
                        self.editor_area.ensureWidgetVisible(widget)
                except RuntimeError:
                    pass
            QTimer.singleShot(0, show)

    def validate_and_accept(self):
        self.validate_now(force=True)
        if self.issues:
            self._goto_first_issue()
            self.toast("Perbaiki isian yang bertanda merah sebelum menyimpan.", bad=True)
            return
        self.accept()

    # ------------------------------------------------------------------ hasil
    def get_column_options(self):
        return self.column_options()

    def get_config(self):
        self.config["steps"] = [{
            "name_before": step["name_before"], "name_after": step["name_after"],
            "message_before": step["message_before"], "input_type": step["input_type"],
            "groups": [export_group(g) for g in step["groups"]],
        } for step in self.steps]
        self.config["display_title"] = self.title_input.text().strip() or "BARCODE DATA INPUT"
        self.config["end_popup"] = {
            "enabled": self.end_popup_enabled.isChecked(),
            "message": self.end_popup_message.text().strip() or "Data telah tersimpan.",
        }
        self.config["ip_link"] = {
            "enabled": self.ip_enabled.isChecked(),
            "host": self.ip_host.text().strip(),
            "port": self.ip_port.value(),
        }
        self.config["saving"] = {
            "enabled": self.save_enabled.isChecked(),
            "targets": self._live_targets(),
            "line": self.line_input.text().strip(),
            "result_text": self.result_input.text().strip(),
            "column_names": {key: field.text().strip() or key for key, field in self.column_names.items()},
        }
        self.config["dat_saving"] = {
            "enabled": self.dat_enabled.isChecked(),
            "destination": self.dat_destination.text().strip(),
            "filename": self.dat_filename.text().strip(),
        }
        return self.config


class OperatorDialog(QDialog):
    def __init__(self, operators, file_path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Tambah Operator")
        self.resize(700, 540)
        self.operators = {}
        self.operators_file = ""

        root = QVBoxLayout(self)
        root.addWidget(QLabel(
            "Kelola daftar barcode operator dan nama pekerja.\n"
            "Daftar disimpan di file JSON di bawah ini. Clone lain cukup diarahkan ke file yang sama."
        ))

        # Pilihan file JSON operator (refer / update)
        file_row = QHBoxLayout()
        file_row.addWidget(QLabel("File operator:"))
        self.file_input = QLineEdit(str(file_path))
        self.file_input.setPlaceholderText("Contoh: D:\\Shared\\operators.json")
        browse_btn = QPushButton("Pilih File...")
        file_row.addWidget(self.file_input, 1)
        file_row.addWidget(browse_btn)
        root.addLayout(file_row)
        self._last_path_text = self.file_input.text()
        self.loaded_mtime = self.file_mtime(self.current_path())

        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["Barcode Operator", "Nama Pekerja"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        for barcode, name in operators.items():
            self.add_row(barcode, name)
        root.addWidget(self.table)

        self.count_label = QLabel()
        root.addWidget(self.count_label)
        self.table.model().rowsInserted.connect(self.update_count)
        self.table.model().rowsRemoved.connect(self.update_count)
        self.update_count()

        buttons = QHBoxLayout()
        add_btn = QPushButton("Tambah Baris")
        remove_btn = QPushButton("Hapus Baris Terpilih")
        save_btn = QPushButton("Simpan")
        cancel_btn = QPushButton("Batal")
        add_btn.clicked.connect(self.add_row_checked)
        remove_btn.clicked.connect(self.remove_rows)
        browse_btn.clicked.connect(self.browse_file)
        self.file_input.editingFinished.connect(self.on_path_edited)
        save_btn.clicked.connect(self.validate_and_accept)
        cancel_btn.clicked.connect(self.reject)
        for btn in (add_btn, remove_btn): buttons.addWidget(btn)
        buttons.addStretch()
        buttons.addWidget(cancel_btn)
        buttons.addWidget(save_btn)
        root.addLayout(buttons)

    @staticmethod
    def file_mtime(path):
        try:
            return path.stat().st_mtime_ns
        except OSError:
            return None

    def current_path(self):
        return resolve_operators_path(self.file_input.text())

    def update_count(self, *_args):
        self.count_label.setText(f"Jumlah baris: {self.table.rowCount()} / {MAX_OPERATORS}")

    def add_row(self, barcode="", name=""):
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setItem(row, 0, QTableWidgetItem(str(barcode)))
        self.table.setItem(row, 1, QTableWidgetItem(str(name)))

    def add_row_checked(self):
        if self.table.rowCount() >= MAX_OPERATORS:
            QMessageBox.warning(self, "Batas operator", f"Maksimal {MAX_OPERATORS} operator.")
            return
        self.add_row()
        self.table.setCurrentCell(self.table.rowCount() - 1, 0)

    def remove_rows(self):
        rows = sorted({i.row() for i in self.table.selectionModel().selectedRows()}, reverse=True)
        if not rows and self.table.currentRow() >= 0:
            rows = [self.table.currentRow()]
        for row in rows:
            self.table.removeRow(row)

    def table_has_data(self):
        for row in range(self.table.rowCount()):
            for col in (0, 1):
                item = self.table.item(row, col)
                if item and item.text().strip():
                    return True
        return False

    def browse_file(self):
        chosen, _ = QFileDialog.getSaveFileName(
            self, "Pilih file operator (.json)", str(self.current_path()),
            "File JSON (*.json)", options=QFileDialog.DontConfirmOverwrite
        )
        if not chosen:
            return
        if not chosen.lower().endswith(".json"):
            chosen += ".json"
        self.file_input.setText(chosen)
        self._last_path_text = chosen
        self.load_from_file(Path(chosen))

    def on_path_edited(self):
        text = self.file_input.text()
        if text == self._last_path_text:
            return
        self._last_path_text = text
        self.load_from_file(self.current_path())

    def load_from_file(self, path):
        """Dipanggil setelah file diganti: jika file sudah ada, isinya jadi acuan."""
        operators, error = read_operators_file(path)
        if error:
            QMessageBox.warning(
                self, "File tidak bisa dibaca",
                f"File operator tidak bisa dibaca:\n{path}\n\n{error}\n\n"
                "Tabel tidak diubah. Menyimpan akan menimpa file tersebut."
            )
            self.loaded_mtime = self.file_mtime(path)
            return
        if not path.exists():
            self.loaded_mtime = None
            QMessageBox.information(
                self, "File baru",
                f"File belum ada. File baru akan dibuat saat Simpan:\n{path}"
            )
            return
        if self.table_has_data():
            reply = QMessageBox.question(
                self, "Muat isi file",
                f"File ini berisi {len(operators)} operator.\n\n"
                "Ya = ganti tabel dengan isi file (acuan bersama).\n"
                "Tidak = pertahankan tabel saat ini (isi file akan ditimpa saat Simpan).",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
            )
            if reply != QMessageBox.Yes:
                self.loaded_mtime = self.file_mtime(path)
                return
        self.table.setRowCount(0)
        for barcode, name in operators.items():
            self.add_row(barcode, name)
        self.loaded_mtime = self.file_mtime(path)

    def validate_and_accept(self):
        if not self.file_input.text().strip():
            QMessageBox.warning(self, "File belum diisi", "Isi atau pilih file operator terlebih dahulu.")
            return
        result = {}
        for row in range(self.table.rowCount()):
            a = self.table.item(row, 0)
            b = self.table.item(row, 1)
            barcode = a.text().strip() if a else ""
            name = b.text().strip() if b else ""
            if not barcode and not name:
                continue
            if not barcode or not name:
                QMessageBox.warning(self, "Data belum lengkap", f"Lengkapi barcode dan nama pada baris {row + 1}.")
                return
            if barcode in result:
                QMessageBox.warning(self, "Barcode duplikat", f"Barcode {barcode} terdaftar lebih dari sekali.")
                return
            result[barcode] = name
        if len(result) > MAX_OPERATORS:
            QMessageBox.warning(self, "Batas operator", f"Maksimal {MAX_OPERATORS} operator (saat ini {len(result)}).")
            return
        path = self.current_path()
        if self.file_mtime(path) != self.loaded_mtime:
            reply = QMessageBox.question(
                self, "File berubah",
                f"File operator diubah (mungkin dari clone lain) sejak dibuka:\n{path}\n\n"
                "Menyimpan akan menimpa perubahan tersebut. Lanjutkan?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return
        self.operators = result
        self.operators_file = self.file_input.text().strip()
        self.accept()


class BarcodeApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle(f"Barcode Data Input  -  {APP_DIR}")
        self.resize(1200, 760)

        self.config = load_json(CONFIG_FILE, EMPTY_CONFIG)
        if not isinstance(self.config, dict):
            self.config = json.loads(json.dumps(EMPTY_CONFIG))
        if not isinstance(self.config.get("steps"), list):
            self.config["steps"] = []
        # Step yang bukan objek (config diedit manual) diabaikan supaya tidak crash.
        self.config["steps"] = [s for s in self.config["steps"] if isinstance(s, dict)]
        if not isinstance(self.config.get("saving", {}), dict):
            self.config["saving"] = json.loads(json.dumps(EMPTY_CONFIG["saving"]))
        if not CONFIG_FILE.exists():
            save_json(CONFIG_FILE, self.config)

        old_session = load_json(SESSION_FILE, {"total": 0})
        try:
            self.total = max(0, int(old_session.get("total", 0)))
        except (TypeError, ValueError, AttributeError):
            self.total = 0
        if not SESSION_FILE.exists():
            save_json(SESSION_FILE, {"total": self.total})

        self.current_step = 0
        self.current_operator = ""
        self.running = False
        self.last_barcode = ""
        self.cycle_started_at = None
        self.cycle_operator_code = ""
        self.cycle_operator_name = ""
        self.cycle_product_codes = {}
        self._cycle_seen_values = set()
        # Cache isi database Excel supaya file tidak dibaca ulang pada setiap scan.
        self._xlsx_validation_cache = {}
        # Indeks awalan per (daftar nilai, jumlah karakter) supaya scan tidak memindai seluruh kolom.
        self._xlsx_prefix_index = {}
        # Cache file operator; dibaca ulang otomatis jika file berubah (mis. diedit clone lain).
        self._operators_cache = {}
        self._operators_path = None
        self._operators_mtime = None
        self.row_status = {}  # {indeks step: "ok" / "ng"} untuk warna hijau / merah
        self._ip_ok_until = 0.0  # cache hasil cek IP yang berhasil (detik, time.monotonic)

        self.setup_ui()
        self.refresh_table()
        self.show_step_message()
        self.update_controls()

    def setup_ui(self):
        central = QWidget()
        root = QVBoxLayout(central)

        title = QLabel(self.config.get("display_title", "BARCODE DATA INPUT"))
        self.title_label = title
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("""
            background: #54D7DC;
            color: black;
            font-size: 24px;
            font-weight: bold;
            padding: 16px;
            border: 2px solid black;
        """)
        root.addWidget(title)

        self.table = QTableWidget()
        self.table.setEditTriggers(
            QTableWidget.NoEditTriggers
        )
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch
        )
        root.addWidget(self.table, 5)

        self.message = QLabel()
        self.message.setAlignment(Qt.AlignCenter)
        self.message.setMinimumHeight(80)
        self.message.setStyleSheet("""
            background: #F1F5F9;
            border: 2px solid #64748B;
            font-size: 20px;
            font-weight: bold;
            padding: 10px;
        """)
        root.addWidget(self.message)

        self.scan_input = QLineEdit()
        self.scan_input.setPlaceholderText("Scan barcode di sini...")
        self.scan_input.setMinimumHeight(42)
        self.scan_input.returnPressed.connect(self.process_scan)
        root.addWidget(self.scan_input)

        controls = QHBoxLayout()

        self.start_button = QPushButton("Start")
        self.stop_button = QPushButton("Stop")
        self.reset_button = QPushButton("Reset Total")
        self.edit_button = QPushButton("Edit Program")
        self.operator_button = QPushButton("Tambah Operator")
        self.total_label = QLabel()

        self.start_button.clicked.connect(self.start_program)
        self.stop_button.clicked.connect(self.stop_program)
        self.reset_button.clicked.connect(self.reset_total)
        self.edit_button.clicked.connect(self.edit_program)
        self.operator_button.clicked.connect(self.manage_operators)

        controls.addWidget(self.start_button)
        controls.addWidget(self.stop_button)
        controls.addWidget(self.reset_button)
        controls.addWidget(self.edit_button)
        controls.addWidget(self.operator_button)
        controls.addStretch()
        controls.addWidget(self.total_label)

        root.addLayout(controls)
        self.setCentralWidget(central)

    def get_operators(self):
        """Daftar operator dari file JSON bersama. Config lama (key "operators") diabaikan."""
        path = resolve_operators_path(self.config.get("operators_file", ""))
        try:
            mtime = path.stat().st_mtime_ns
        except (OSError, ValueError):
            mtime = None
        if path != self._operators_path or mtime != self._operators_mtime:
            operators, error = read_operators_file(path)
            if error is None:
                self._operators_cache = operators
                self._operators_path = path
                self._operators_mtime = mtime
            elif path != self._operators_path:
                self._operators_cache = {}
            # Jika error tetapi path sama: pakai cache lama, coba lagi pada scan berikutnya.
        return self._operators_cache

    def refresh_table(self):
        steps = self.config["steps"]
        self.row_status = {}

        # Jumlah kolom data mengikuti step dengan kelompok terbanyak
        max_groups = max(
            len(step.get("groups", []))
            for step in steps
        ) if steps else 0

        headers = ["No", "Process Step"]
        headers += [f"Data {i + 1}" for i in range(max_groups)]
        headers.append("Judge")

        self.table.clear()
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setRowCount(len(steps))

        for i, step in enumerate(steps):
            self.put(i, 0, str(i + 1))
            self.put(i, 1, step.get("name_before", ""))

            # Siapkan semua kolom agar tidak menyisakan data lama
            for col in range(2, len(headers) - 1):
                self.put(i, col, "")

            self.put(i, len(headers) - 1, "")

        self.total_label.setText(f"Total: {self.total}")
        self.highlight_step()


    def put(self, row, col, value):
        item = QTableWidgetItem(str(value))
        item.setTextAlignment(Qt.AlignCenter)
        self.table.setItem(row, col, item)

    def highlight_step(self):
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item is None:
                    item = QTableWidgetItem("")
                    item.setTextAlignment(Qt.AlignCenter)
                    self.table.setItem(row, col, item)

                status = self.row_status.get(row)
                if status == "ng":
                    item.setBackground(COLOR_NG)
                elif status == "ok":
                    item.setBackground(COLOR_OK)
                elif row == self.current_step:
                    item.setBackground(COLOR_WAITING)
                else:
                    item.setBackground(Qt.white)

    def show_step_message(self):
        if not self.config["steps"]:
            self.message.setText("Belum ada step. Buka Edit Program.")
            return

        step = self.config["steps"][self.current_step]
        if step.get("input_type") == "operator":
            default_msg = "Silakan scan ID operator barcode"
        else:
            default_msg = "Silakan scan barcode produk"
        msg = step.get("message_before") or default_msg

        if not self.running:
            msg = "STOP — Tekan Start untuk melanjutkan. " + msg

        self.message.setText(msg)
        self.highlight_step()

    def update_controls(self):
        self.start_button.setEnabled(not self.running)
        self.stop_button.setEnabled(self.running)
        self.edit_button.setEnabled(not self.running)
        self.scan_input.setEnabled(self.running)

    def ensure_ip_link(self):
        """True jika Link IP nonaktif atau IP tersambung. Jika gagal, tampilkan error pop up."""
        ip_cfg = self.config.get("ip_link")
        if not isinstance(ip_cfg, dict) or not ip_cfg.get("enabled"):
            return True
        now = time.monotonic()
        if now < self._ip_ok_until:
            return True
        host = str(ip_cfg.get("host", "")).strip()
        try:
            port = int(ip_cfg.get("port", 0))
        except (TypeError, ValueError):
            port = 0
        ok, reason = check_ip_connection(host, port)
        if ok:
            self._ip_ok_until = now + 5.0  # hindari cek berulang pada setiap scan
            return True
        self._ip_ok_until = 0.0
        target = f"{host}:{port}" if port else host
        QMessageBox.critical(
            self, "Tidak tersambung ke IP",
            f"Program tidak tersambung ke {target or '-'}.\n\n{reason}\n\n"
            "Program tidak dapat bekerja sampai koneksi tersedia. "
            "Periksa kabel/jaringan, lalu coba lagi."
        )
        return False

    def start_program(self):
        if not self.ensure_ip_link():
            return
        # Start always begins a fresh cycle from the first configured step.
        self.current_step = 0
        self.clear_cycle_data()
        self.current_operator = ""
        self.cycle_started_at = None
        self.cycle_operator_code = ""
        self.cycle_operator_name = ""
        self.cycle_product_codes = {}
        self._cycle_seen_values = set()
        self.running = True
        self.update_controls()
        self.show_step_message()
        self.scan_input.setFocus()

    def stop_program(self):
        # Stop discards any incomplete cycle; the next Start begins again at step 1.
        self.running = False
        self.current_step = 0
        self.clear_cycle_data()
        self.current_operator = ""
        self.cycle_started_at = None
        self.cycle_operator_code = ""
        self.cycle_operator_name = ""
        self.cycle_product_codes = {}
        self._cycle_seen_values = set()
        save_json(SESSION_FILE, {"total": self.total})
        self.update_controls()
        self.show_step_message()

    def reset_total(self):
        answer = QMessageBox.question(
            self, "Reset Total",
            f"Yakin ingin mereset total dari {self.total} menjadi 0?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        self.total = 0
        save_json(SESSION_FILE, {"total": self.total})
        self.total_label.setText("Total: 0")

    def clear_cycle_data(self):
        self.row_status = {}
        for row, step in enumerate(self.config["steps"]):
            # Kembalikan Name ke kondisi sebelum input
            self.put(row, 1, step.get("name_before", ""))

            # Bersihkan seluruh kolom data
            for col in range(2, self.table.columnCount() - 1):
                self.put(row, col, "")

            # Bersihkan kolom Judge
            self.put(row, self.table.columnCount() - 1, "")


    def get_xlsx_column_values(self, file_path, column_letter, sheet_mode="specific", sheet_name="Sheet1"):
        """Baca kolom dari sheet tertentu atau seluruh sheet; hasil di-cache."""
        path = Path(file_path).expanduser()
        if not path.is_file():
            raise FileNotFoundError(f"File database Excel tidak ditemukan:\n{path}")
        column_letter = column_letter.strip().upper()
        if not column_letter or not column_letter.isalpha():
            raise ValueError("Kolom Excel harus berupa huruf, misalnya G atau AA.")
        column_index = column_index_from_string(column_letter)
        sheet_mode = sheet_mode or "specific"
        sheet_name = (sheet_name or "Sheet1").strip()
        cache_key = (str(path.resolve()), path.stat().st_mtime, column_letter, sheet_mode, sheet_name)
        if cache_key in self._xlsx_validation_cache:
            return self._xlsx_validation_cache[cache_key]
        workbook = None
        try:
            workbook = load_workbook(str(path), read_only=True, data_only=True)
            if sheet_mode == "all":
                worksheets = workbook.worksheets
            else:
                if sheet_name not in workbook.sheetnames:
                    raise ValueError(f"Sheet '{sheet_name}' tidak ditemukan di file Excel.")
                worksheets = [workbook[sheet_name]]
            values = []
            for worksheet in worksheets:
                for row_number, row in enumerate(worksheet.iter_rows(
                    min_row=1, min_col=column_index, max_col=column_index, values_only=True
                ), start=1):
                    value = row[0]
                    if value is None:
                        continue
                    text_value = str(value).strip().upper()
                    if text_value:
                        values.append((text_value, worksheet.title, row_number))
        finally:
            if workbook is not None:
                workbook.close()
        resolved_path = str(path.resolve())
        self._xlsx_validation_cache = {
            key: value for key, value in self._xlsx_validation_cache.items()
            if not (key[0] == resolved_path and key[2] == column_letter and key[3] == sheet_mode and key[4] == sheet_name)
        }
        self._xlsx_validation_cache[cache_key] = values
        # Buang indeks awalan milik daftar nilai yang sudah tidak ada di cache.
        live_ids = {id(v) for v in self._xlsx_validation_cache.values()}
        self._xlsx_prefix_index = {k: e for k, e in self._xlsx_prefix_index.items() if k[0] in live_ids}
        return values

    def validate_xlsx_group(self, group, value):
        """Cek barcode ke database Excel. Default (excel_found_result = "ng"): ditemukan = NG,
        tidak ditemukan = OK. Dengan "ok": ditemukan = OK, tidak ditemukan = NG."""
        found_ok = group.get("excel_found_result", "ng") == "ok"
        file_path = group.get("xlsx_path", "").strip()
        column_letter = group.get("xlsx_column", "G").strip().upper()
        match_chars = max(1, int_or(group.get("xlsx_match_chars", 8), 8))
        sheet_mode = group.get("xlsx_sheet_mode", "specific")
        sheet_name = group.get("xlsx_sheet_name", "Sheet1")
        normalized = value.strip().upper()
        if not file_path:
            return False, "Pilih file database Excel untuk data ini."
        if len(normalized) < match_chars:
            return False, f"Data hanya {len(normalized)} karakter; dibutuhkan {match_chars} karakter untuk pemeriksaan Excel."
        values = self.get_xlsx_column_values(file_path, column_letter, sheet_mode, sheet_name)
        prefix = normalized[:match_chars]
        entry = self._xlsx_prefix_index.get((id(values), match_chars))
        if entry is None or entry[0] is not values:
            index = {}
            for item, name, row_number in values:
                if len(item) >= match_chars:
                    index.setdefault(item[:match_chars], (name, row_number))  # kemunculan pertama
            if len(self._xlsx_prefix_index) > 64:
                self._xlsx_prefix_index.clear()
            entry = (values, index)
            self._xlsx_prefix_index[(id(values), match_chars)] = entry
        match = entry[1].get(prefix)
        if match:
            sheet, row_number = match
            if found_ok:
                return True, f"Barcode ditemukan di database Excel kolom {column_letter}, sheet '{sheet}' (baris {row_number}); valid."
            return False, f"Barcode sudah terdaftar di database Excel kolom {column_letter}, sheet '{sheet}' (baris {row_number}). Barcode ditolak."
        scope_text = f"sheet '{sheet_name}'" if sheet_mode == "specific" else "semua sheet"
        if found_ok:
            return False, f"Barcode tidak ditemukan di database Excel kolom {column_letter} ({scope_text}). Barcode ditolak."
        return True, f"Barcode tidak ditemukan di database Excel kolom {column_letter} ({scope_text}); valid."

    def process_scan(self):
        if not self.running:
            return

        code = self.scan_input.text().strip()
        self.scan_input.clear()

        if not code or not self.config.get("steps"):
            return

        if not self.ensure_ip_link():
            self.scan_input.setFocus()
            return

        steps = self.config["steps"]
        # Pastikan indeks step masih valid jika konfigurasi berubah
        self.current_step = min(self.current_step, len(steps) - 1)
        step = steps[self.current_step]
        remark_col = self.table.columnCount() - 1

        # STEP OPERATOR
        if step.get("input_type") == "operator":
            operator = self.get_operators().get(code)

            if operator is None:
                self.put(self.current_step, remark_col, "NG")
                error_message = "ID operator tidak terdaftar. Silakan scan ulang."
                operators_path = resolve_operators_path(self.config.get("operators_file", ""))
                if not operators_path.exists():
                    error_message = (
                        f"File operator tidak ditemukan: {operators_path}\n"
                        "Buka 'Tambah Operator' untuk memilih atau membuat file operator."
                    )
                self.message.setText(error_message)
                self.row_status[self.current_step] = "ng"
                self.highlight_step()
                self.scan_input.setFocus()
                return

            self.current_operator = operator
            self.cycle_operator_code = code
            self.cycle_operator_name = operator
            if self.cycle_started_at is None:
                self.cycle_started_at = datetime.now()

            # Keep the process-step description unchanged. Data columns are configured
            # per group: use "Seluruh barcode" for the scanned ID and "Nama operator"
            # for the operator's registered name (e.g. Data 1 and Data 4).
            for col in range(2, remark_col):
                self.put(self.current_step, col, "")
            for index, group in enumerate(step.get("groups", [])):
                mode = group.get("mode", "length")
                if mode == "operator_name":
                    value = operator
                elif mode == "all":
                    value = code
                else:
                    start = max(0, int_or(group.get("start", 1), 1) - 1)
                    length = max(0, int_or(group.get("length", 1), 1))
                    value = code[start:start + length]
                column = index + 2
                if column < remark_col:
                    self.put(self.current_step, column, value)
            self.put(self.current_step, remark_col, "OK")
            self.row_status[self.current_step] = "ok"
            self.message.setText(step.get("name_after", ""))
            self.advance_step()
            self.scan_input.setFocus()
            return

        # STEP BARCODE / INPUT DATA
        groups = step.get("groups", [])
        valid = True

        # Bersihkan data sebelumnya pada baris step aktif
        for col in range(2, remark_col):
            self.put(self.current_step, col, "")

        validation_message = ""
        popup_error = False
        judge_ok_text = "OK"
        judge_ng_text = "NG"
        for index, group in enumerate(groups):
            if group.get("mode") == "all":
                value = code
            else:
                start = max(0, int_or(group.get("start", 1), 1) - 1)
                length = max(0, int_or(group.get("length", 1), 1))
                value = code[start:start + length]

            column = index + 2
            if column < remark_col:
                self.put(self.current_step, column, value)

            # Setiap jenis validasi diperiksa terpisah dan memiliki pesan error sendiri.
            if group.get("excel_enabled", group.get("validate", False) and group.get("source") == "excel"):
                try:
                    group_valid, _detail = self.validate_xlsx_group(group, value)
                except Exception as error:
                    group_valid, _detail = False, f"Gagal memvalidasi Excel: {error}"
                if not group_valid:
                    valid = False
                    judge_ng_text = group.get("judge_ng", "NG") or "NG"
                    validation_message = group.get("excel_error", "").strip() or _detail
                    popup_error = bool(group.get("excel_popup", False))
                    break

            if group.get("char_enabled", False):
                char_start = max(1, int_or(group.get("char_start", 1), 1))
                char_length = max(1, int_or(group.get("char_length", 1), 1))
                expected_chars = str(group.get("char_expected", ""))
                actual_chars = code[char_start - 1:char_start - 1 + char_length]
                if len(actual_chars) != char_length or actual_chars != expected_chars:
                    valid = False
                    judge_ng_text = group.get("judge_ng", "NG") or "NG"
                    validation_message = group.get("char_error", "").strip() or (
                        f"Karakter ke-{char_start} sampai ke-{char_start + char_length - 1} harus {expected_chars}."
                    )
                    popup_error = bool(group.get("char_popup", False))
                    break

            if group.get("unique", False):
                normalized_value = code.strip().upper()
                if normalized_value and normalized_value in self._cycle_seen_values:
                    valid = False
                    judge_ng_text = group.get("judge_ng", "NG") or "NG"
                    validation_message = group.get("unique_error", "").strip() or (
                        f"Data {group.get('name', index + 1)} duplikat; barcode sudah pernah dipindai pada siklus ini."
                    )
                    popup_error = bool(group.get("unique_popup", False))
                    break

            judge_ok_text = group.get("judge_ok", "OK") or "OK"

        if not valid:
            self.put(self.current_step, remark_col, judge_ng_text)
            display_error = validation_message or "Input NG. Periksa data dan silakan scan ulang."
            self.message.setText(display_error)
            self.row_status[self.current_step] = "ng"
            self.highlight_step()
            if popup_error:
                QMessageBox.warning(self, "Data tidak valid", display_error)
            self.scan_input.setFocus()
            return

        # Catat nilai unik hanya setelah seluruh validasi pada Step ini berhasil.
        for index, group in enumerate(groups):
            if group.get("unique", False):
                unique_value = code.strip().upper()
                if unique_value:
                    self._cycle_seen_values.add(unique_value)

        self.put(self.current_step, remark_col, judge_ok_text)
        self.row_status[self.current_step] = "ok"
        self.message.setText(step.get("name_after", ""))
        self.put(self.current_step, 1, step.get("name_after", ""))
        if self.cycle_started_at is None:
            self.cycle_started_at = datetime.now()
        self.cycle_product_codes[self.current_step] = code

        # Setiap input hanya diproses satu kali.
        self.advance_step()
        self.scan_input.setFocus()

    def advance_step(self):
        steps = self.config.get("steps", [])
        total_steps = len(steps)

        if total_steps == 0:
            self.current_step = 0
            return

        # Jika masih ada step berikutnya, pindah tanpa menambah Total.
        if self.current_step < total_steps - 1:
            self.current_step += 1
            self.show_step_message()
            self.highlight_step()
            return

        # Hanya input valid pada step terakhir yang menyelesaikan satu siklus.
        completed_at = datetime.now()
        self.total += 1
        save_json(SESSION_FILE, {"total": self.total})
        self.save_completed_cycle(completed_at)
        self.save_dat_cycle()

        # Tampilkan siklus selesai (semua step hijau) selagi pop up tampil, baru reset.
        self.total_label.setText(f"Total: {self.total}")
        self.highlight_step()
        end_popup = self.config.get("end_popup")
        if isinstance(end_popup, dict) and end_popup.get("enabled"):
            QMessageBox.information(
                self, "Sequence selesai",
                str(end_popup.get("message") or "").strip() or "Data telah tersimpan."
            )

        self.current_step = 0
        self.current_operator = ""
        self.cycle_started_at = None
        self.cycle_operator_code = ""
        self.cycle_operator_name = ""
        self.cycle_product_codes = {}
        self._cycle_seen_values = set()
        self.clear_cycle_data()
        self.refresh_table()
        self.show_step_message()


    def get_filename_data_value(self, source):
        """Return the selected extracted data value for a filename component."""
        try:
            _, _, step_part, _, group_part = source.split("_")
            step_index = int(step_part) - 1
            group_index = int(group_part) - 1
            step = self.config.get("steps", [])[step_index]
            group = step.get("groups", [])[group_index]
            scanned_code = self.cycle_product_codes.get(step_index, "")
            mode = group.get("mode", "all")
            if mode == "operator_name":
                return self.cycle_operator_name
            if mode == "all":
                return scanned_code
            start = max(0, int_or(group.get("start", 1), 1) - 1)
            length = max(0, int_or(group.get("length", 1), 1))
            return scanned_code[start:start + length]
        except (ValueError, IndexError, KeyError, TypeError, AttributeError):
            return ""

    def save_dat_cycle(self):
        """Tambahkan barcode lengkap sequence ini ke file .dat (satu kolom, ke bawah).
        Baris pertama = nama file, hanya ditulis saat file baru/kosong."""
        dat = self.config.get("dat_saving")
        if not isinstance(dat, dict) or not dat.get("enabled"):
            return

        # Barcode lengkap dari step barcode yang punya Data ber-mode "Seluruh barcode".
        lines = []
        for index, step in enumerate(self.config.get("steps", [])):
            if step.get("input_type") == "operator":
                continue
            if not any(group.get("mode") == "all" for group in step.get("groups", [])):
                continue
            lines.append(str(self.cycle_product_codes.get(index, "")))
        if not lines:
            return

        name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(dat.get("filename", "")).strip())
        if name.lower().endswith(".dat"):
            name = name[:-4]
        name = name.strip(" .") or "hasil_pengujian"

        folder = Path(str(dat.get("destination", ""))).expanduser()
        if not folder.is_dir():
            QMessageBox.critical(
                self, "Gagal menyimpan .dat",
                f"Folder tujuan .dat tidak ditemukan:\n{folder}\nData siklus ini tidak tersimpan ke file .dat."
            )
            return

        output_path = folder / safe_file_stem(name, folder, ".dat")
        try:
            needs_header = not output_path.exists() or output_path.stat().st_size == 0
            text = ""
            if not needs_header:
                with output_path.open("rb") as existing:
                    existing.seek(-1, os.SEEK_END)
                    if existing.read(1) not in (b"\n", b"\r"):
                        text += "\r\n"  # file lama tanpa baris baru di akhir
            if needs_header:
                text += name + "\r\n"
            text += "".join(line + "\r\n" for line in lines)
            with output_path.open("a", newline="", encoding="utf-8") as file:
                file.write(text)
        except OSError as exc:
            QMessageBox.critical(
                self, "Gagal menyimpan .dat",
                f"Data siklus selesai, tetapi file .dat tidak dapat ditulis:\n{output_path}\n{exc}"
            )

    def save_completed_cycle(self, completed_at):
        saving = self.config.get("saving", {})
        if not saving.get("enabled", False):
            return

        targets = saving.get("targets")
        if not targets:
            # Compatibility with older single-file configurations.
            targets = [{
                "destination": saving.get("destination", ""),
                "filename": saving.get("filename", "hasil_pengujian.csv"),
                "date_prefix": saving.get("date_prefix", False),
            }]

        start_time = self.cycle_started_at or completed_at
        values = {
            "no": self.total,
            "hasil_tes": saving.get("result_text", "OK"),
            "waktu_mulai": start_time.strftime("%Y-%m-%d %H:%M:%S"),
            "kode_pekerja": self.cycle_operator_code,
            "nama_pekerja": self.cycle_operator_name,
            "lini": saving.get("line", ""),
            "waktu_berakhir": completed_at.strftime("%Y-%m-%d %H:%M:%S"),
        }
        for number, step_index in enumerate(
            [i for i, step in enumerate(self.config.get("steps", []))
             if step.get("input_type") == "barcode"], start=1
        ):
            values[f"kode_produk_{number}"] = self.cycle_product_codes.get(step_index, "")

        labels = {
            "no": "No", "hasil_tes": "Hasil Tes",
            "waktu_mulai": "Waktu tes dimulai",
            "kode_pekerja": "Kode Pekerja", "nama_pekerja": "Nama Pekerja",
            "lini": "Lini", "waktu_berakhir": "Waktu tes berakhir",
        }
        today_prefix = completed_at.strftime("%d%m%Y")
        for target in targets:
            columns = target.get("columns", saving.get("columns", []))
            if not columns:
                columns = list(values.keys())
            headers = []
            row = []
            for key in columns:
                if key.startswith("kode_produk_"):
                    number = key.rsplit("_", 1)[-1]
                    default_header = f"Kode produk {number}"
                    header = saving.get("column_names", {}).get(key, default_header)
                elif key in labels:
                    header = saving.get("column_names", {}).get(key, labels[key])
                else:
                    continue
                headers.append(header)
                row.append(values.get(key, ""))
            if not headers:
                continue
            folder = Path(target.get("destination", "")).expanduser()
            # Build filename as: part1_part2 + free text.
            # An underscore is used only between part 1 and part 2 when both exist.
            # Free text is appended exactly as entered, so users may add their own underscore
            # and may also leave the field empty.
            filename_components = []
            for part_number, part_key in enumerate(("filename_part1", "filename_part2"), start=1):
                part_value = ""
                part_type = target.get(part_key, "none")
                if part_type == "date":
                    part_value = today_prefix
                elif part_type == "data":
                    source = target.get(f"filename_data{part_number}", "")
                    source_value = self.get_filename_data_value(source)
                    if source_value:
                        part_value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(source_value)).strip(" ._")
                filename_components.append(part_value)

            part1, part2 = filename_components
            if part1 and part2:
                filename = f"{part1}_{part2}"
            else:
                filename = part1 or part2

            free_text = (target.get("filename_text") or "").strip()
            filename += ILLEGAL_NAME_CHARS.sub("_", free_text)
            # Keep a usable filename if all three configurable components are empty.
            filename = filename or "hasil_pengujian"
            filename = safe_file_stem(filename, folder, ".csv")

            if not folder.is_dir():
                QMessageBox.critical(
                    self, "Gagal menyimpan data",
                    f"Folder tujuan tidak ditemukan:\n{folder}\nData siklus ini tidak tersimpan ke tujuan tersebut."
                )
                continue

            output_path = folder / filename
            try:
                needs_header = not output_path.exists() or output_path.stat().st_size == 0
                if not needs_header:
                    with output_path.open("r", newline="", encoding="utf-8-sig") as existing_file:
                        existing_rows = csv.reader(existing_file)
                        existing_header = next(existing_rows, [])
                    # Remove the legacy title-only row from older versions, then write a clean CSV header.
                    if existing_header == ["hasil_pengujian"]:
                        needs_header = True
                        output_path.write_text("", encoding="utf-8-sig")
                    elif existing_header != headers:
                        clean_existing = [h.strip() for h in existing_header]
                        clean_new = [h.strip() for h in headers]
                        if sorted(clean_existing) == sorted(clean_new):
                            # Kolom sama, hanya urutan berbeda: susun data mengikuti
                            # urutan header yang sudah ada di file.
                            by_name = dict(zip(clean_new, row))
                            row = [by_name[h] for h in clean_existing]
                        else:
                            missing = [h for h in clean_existing if h not in clean_new]
                            extra = [h for h in clean_new if h not in clean_existing]
                            QMessageBox.critical(
                                self, "Kolom CSV berbeda",
                                f"Kolom konfigurasi berbeda dari header file:\n{output_path}\n\n"
                                f"Ada di file, tidak ada di konfigurasi: {', '.join(missing) or '-'}\n"
                                f"Ada di konfigurasi, tidak ada di file: {', '.join(extra) or '-'}\n\n"
                                "Samakan nama kolom di Edit Program, atau gunakan nama file baru."
                            )
                            continue
                with output_path.open("a", newline="", encoding="utf-8-sig") as file:
                    writer = csv.writer(file)
                    if needs_header:
                        writer.writerow(headers)
                    writer.writerow(row)
            except OSError as exc:
                QMessageBox.critical(
                    self, "Gagal menyimpan data",
                    f"Data siklus selesai, tetapi file tidak dapat ditulis:\n{output_path}\n{exc}"
                )

    def manage_operators(self):
        password, accepted = QInputDialog.getText(
            self, "Autentikasi Tambah Operator",
            "Masukkan password untuk membuka Tambah Operator:", QLineEdit.Password
        )
        if not accepted:
            return
        if not password_ok("operator", password):
            QMessageBox.warning(self, "Password salah", "Password tidak sesuai. Menu Tambah Operator tidak dibuka.")
            return
        path = resolve_operators_path(self.config.get("operators_file", ""))
        operators, error = read_operators_file(path)
        if error:
            QMessageBox.warning(
                self, "File operator tidak bisa dibaca",
                f"{path}\n\n{error}\n\nTabel dibuka kosong. Pilih file lain, atau Simpan untuk menimpa file ini."
            )
        dialog = OperatorDialog(operators, str(path), self)
        if dialog.exec() == QDialog.Accepted:
            new_path = resolve_operators_path(dialog.operators_file)
            if not write_operators_file(new_path, dialog.operators):
                return
            self.config["operators_file"] = "" if new_path == DEFAULT_OPERATORS_FILE else dialog.operators_file
            save_json(CONFIG_FILE, self.config)
            self._operators_path = None  # paksa baca ulang pada scan berikutnya
            QMessageBox.information(
                self, "Berhasil",
                f"{len(dialog.operators)} operator disimpan ke:\n{new_path}"
            )

    def edit_program(self):
        password, accepted = QInputDialog.getText(
            self,
            "Autentikasi Edit Program",
            "Masukkan password untuk membuka Edit Program:",
            QLineEdit.Password,
        )
        if not accepted:
            return
        if not password_ok("edit", password):
            QMessageBox.warning(self, "Password salah", "Password tidak sesuai. Edit Program tidak dibuka.")
            return
        dialog = EditProgramDialog(self.config, self)
        if dialog.exec() == QDialog.Accepted:
            self.config = dialog.get_config()
            self._ip_ok_until = 0.0
            save_json(CONFIG_FILE, self.config)
            self.title_label.setText(self.config.get("display_title", "BARCODE DATA INPUT"))
            self.current_step = 0
            self.current_operator = ""
            self.cycle_started_at = None
            self.cycle_operator_code = ""
            self.cycle_operator_name = ""
            self.cycle_product_codes = {}
            self._cycle_seen_values = set()
            self.refresh_table()
            self.show_step_message()

    def closeEvent(self, event):
        if self.running:
            QMessageBox.warning(
                self,
                "Program masih berjalan",
                "Tekan Stop sebelum menutup program."
            )
            event.ignore()
            return

        save_json(SESSION_FILE, {"total": self.total})
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(MAIN_QSS)
    window = BarcodeApp()
    window.show()
    sys.exit(app.exec())
