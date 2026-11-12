import os
import sys
import json
import csv
import re
import socket
import subprocess
import time
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QMessageBox, QHeaderView, QDialog, QFormLayout, QSpinBox,
    QComboBox, QCheckBox, QGroupBox, QScrollArea, QInputDialog, QFileDialog,
    QTabWidget
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
    except (OSError, json.JSONDecodeError):
        return json.loads(json.dumps(default))  # salinan baru, bukan objek bersama


def save_json(path, data):
    """Tulis JSON ke path. Return True jika berhasil; tampilkan pesan jika gagal."""
    text = json.dumps(data, indent=2, ensure_ascii=False)
    tmp = path.with_name(path.name + ".tmp")
    try:
        # Tulis ke file sementara lalu ganti, supaya file tidak rusak jika terputus.
        tmp.write_text(text, encoding="utf-8")
        os.replace(tmp, path)
        return True
    except OSError:
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
    if not path.exists():
        return {}, None
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
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
        cmd = ["ping", "-n", "1", "-w", str(int(timeout * 1000)), host]
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


class EditProgramDialog(QDialog):
    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Program")
        self.resize(850, 680)
        self.setMinimumSize(650, 480)
        self.config = json.loads(json.dumps(config))
        self.group_widgets = []
        self.column_checks = {}
        self.column_names = {}

        root = QVBoxLayout(self)
        self.tabs = QTabWidget()
        root.addWidget(self.tabs)

        # TAB 2: Barcode / workflow steps
        barcode_tab = QWidget()
        barcode_layout = QVBoxLayout(barcode_tab)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QWidget()
        self.form = QVBoxLayout(content)
        self.step_count = QSpinBox()
        self.step_count.setRange(1, 1000)
        self.step_count.setValue(len(self.config.get("steps", [])) or 1)
        self.form.addWidget(QLabel("Jumlah step"))
        self.form.addWidget(self.step_count)

        # Command prompt untuk menyalin konfigurasi step ke satu atau banyak step.
        command_box = QGroupBox("Command Prompt — Copy Step")
        command_layout = QVBoxLayout(command_box)
        command_help = QLabel(
            "Contoh: copy step 2 for step 3 to 73  |  "
            "atau: copy step 3 to step 10\n"
            "Perintah dapat dijalankan berkali-kali. Rentang tujuan bersifat inklusif."
        )
        command_help.setWordWrap(True)
        command_layout.addWidget(command_help)
        command_row = QHBoxLayout()
        self.step_command_input = QLineEdit()
        self.step_command_input.setPlaceholderText("Ketik perintah copy, contoh: copy step 2 for step 3 to 73")
        self.step_command_button = QPushButton("Jalankan")
        self.step_command_button.clicked.connect(self.execute_step_command)
        self.step_command_input.returnPressed.connect(self.execute_step_command)
        command_row.addWidget(self.step_command_input, 1)
        command_row.addWidget(self.step_command_button)
        command_layout.addLayout(command_row)
        self.form.addWidget(command_box)

        # Pop up yang muncul setelah step terakhir valid (siklus selesai).
        end_cfg = self.config.get("end_popup", {}) if isinstance(self.config.get("end_popup"), dict) else {}
        end_box = QGroupBox("Pop Up Akhir Sequence")
        end_layout = QFormLayout(end_box)
        self.end_popup_enabled = QCheckBox("Tampilkan pop up setelah step terakhir valid")
        self.end_popup_enabled.setChecked(bool(end_cfg.get("enabled", False)))
        self.end_popup_message = QLineEdit(str(end_cfg.get("message", "Data telah tersimpan.")))
        self.end_popup_message.setPlaceholderText("Contoh: Data telah tersimpan.")
        self.end_popup_enabled.toggled.connect(self.end_popup_message.setEnabled)
        self.end_popup_message.setEnabled(self.end_popup_enabled.isChecked())
        end_layout.addRow(self.end_popup_enabled)
        end_layout.addRow("Isi pesan", self.end_popup_message)
        self.form.addWidget(end_box)

        self.step_area = QVBoxLayout()
        self.form.addLayout(self.step_area)
        scroll.setWidget(content)
        barcode_layout.addWidget(scroll)
        self.tabs.addTab(barcode_tab, "1. Barcode")

        # TAB 3: Storage
        saving_tab = QWidget()
        saving_layout = QVBoxLayout(saving_tab)
        saving_form_widget = QWidget()
        saving_form = QFormLayout(saving_form_widget)
        saving = self.config.get("saving", {})
        self.title_input = QLineEdit(self.config.get("display_title", "BARCODE DATA INPUT"))
        self.title_input.setPlaceholderText("Judul pada tampilan utama")
        saving_form.addRow("Judul tampilan utama", self.title_input)

        # Link IP: jika diaktifkan, program hanya bisa berjalan saat IP/host ini tersambung.
        ip_cfg = self.config.get("ip_link", {}) if isinstance(self.config.get("ip_link"), dict) else {}
        ip_box = QGroupBox("Link IP")
        ip_layout = QFormLayout(ip_box)
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
        self.ip_test_button = QPushButton("Tes Koneksi")
        self.ip_test_button.clicked.connect(self.test_ip_connection)
        ip_layout.addRow(self.ip_enabled)
        ip_layout.addRow("IP / Host", self.ip_host)
        ip_layout.addRow("Port (0 = ping)", self.ip_port)
        ip_layout.addRow(self.ip_test_button)
        saving_form.addRow(ip_box)
        self.save_enabled = QCheckBox("Simpan hasil setiap siklus selesai")
        self.save_enabled.setChecked(saving.get("enabled", False))
        saving_form.addRow(self.save_enabled)

        # Multiple CSV destinations. Legacy single-destination config is migrated automatically.
        legacy_destination = saving.get("destination", "")
        legacy_filename = saving.get("filename", "hasil_pengujian.csv")
        saved_targets = saving.get("targets")
        if not saved_targets:
            saved_targets = [{
                "destination": legacy_destination,
                "filename": legacy_filename,
                "filename_part1": "none",
                "filename_part2": "none",
                "filename_text": Path(legacy_filename).stem or "hasil_pengujian",
                "date_prefix": saving.get("date_prefix", False),
            }]
        self.saved_targets = saved_targets
        self.target_count = QSpinBox()
        self.target_count.setRange(1, 1000)
        self.target_count.setValue(max(1, min(1000, len(saved_targets))))
        saving_form.addRow("Jumlah file tujuan", self.target_count)
        self.targets_tabs = QTabWidget()
        self.targets_tabs.setMinimumHeight(360)
        saving_form.addRow("Pengaturan setiap file", self.targets_tabs)
        self.target_widgets = []
        self.target_count.valueChanged.connect(self.rebuild_targets)
        self.line_input = QLineEdit(saving.get("line", ""))
        self.line_input.setPlaceholderText("contoh: Lini 1")
        saving_form.addRow("Lini", self.line_input)
        self.result_input = QLineEdit(saving.get("result_text", "OK"))
        self.result_input.setPlaceholderText("Contoh: LULUS")
        saving_form.addRow("Isi kolom Hasil Tes", self.result_input)

        self.columns_box = QGroupBox("Nama kolom CSV (berlaku untuk semua file)")
        self.columns_layout = QVBoxLayout(self.columns_box)
        saving_form.addRow(self.columns_box)

        # Simpan .dat: satu kolom vertikal. Baris pertama = nama file, lalu barcode lengkap
        # dari step ber-mode "Seluruh barcode" ditambahkan ke bawah pada setiap sequence.
        dat_cfg = self.config.get("dat_saving", {}) if isinstance(self.config.get("dat_saving"), dict) else {}
        self.dat_box = QGroupBox("Simpan .dat (satu kolom, ke bawah)")
        dat_layout = QFormLayout(self.dat_box)
        dat_help = QLabel(
            "Baris pertama berisi nama file. Setiap sequence yang selesai menambahkan barcode lengkap "
            "dari step ber-mode \"Seluruh barcode\" ke baris berikutnya, tanpa menulis nama lagi."
        )
        dat_help.setWordWrap(True)
        self.dat_enabled = QCheckBox("Simpan file .dat setiap sequence selesai")
        self.dat_enabled.setChecked(bool(dat_cfg.get("enabled", False)))
        self.dat_destination = QLineEdit(str(dat_cfg.get("destination", "")))
        self.dat_destination.setPlaceholderText("Pilih folder tujuan")
        self.dat_browse = QPushButton("Pilih folder...")
        self.dat_browse.clicked.connect(lambda: self.browse_target_destination(self.dat_destination))
        dat_dest_row = QWidget()
        dat_dest_layout = QHBoxLayout(dat_dest_row)
        dat_dest_layout.setContentsMargins(0, 0, 0, 0)
        dat_dest_layout.addWidget(self.dat_destination, 1)
        dat_dest_layout.addWidget(self.dat_browse)
        self.dat_filename = QLineEdit(str(dat_cfg.get("filename", "")))
        self.dat_filename.setPlaceholderText("Nama file tanpa .dat, contoh: LOT_A")
        dat_layout.addRow(dat_help)
        dat_layout.addRow(self.dat_enabled)
        dat_layout.addRow("Folder tujuan", dat_dest_row)
        dat_layout.addRow("Nama file", self.dat_filename)
        self.dat_enabled.toggled.connect(self.update_dat_controls)
        self.update_dat_controls()
        saving_form.addRow(self.dat_box)

        # Build targets only after all widgets used by update_saving_controls exist.
        self.rebuild_targets()
        saving_scroll = QScrollArea()
        saving_scroll.setWidgetResizable(True)
        saving_scroll.setWidget(saving_form_widget)
        saving_layout.addWidget(saving_scroll)
        self.tabs.addTab(saving_tab, "2. Penyimpanan")

        buttons = QHBoxLayout()
        save_button = QPushButton("Simpan Konfigurasi")
        cancel_button = QPushButton("Batal")
        save_button.clicked.connect(self.validate_and_accept)
        cancel_button.clicked.connect(self.reject)
        buttons.addWidget(save_button)
        buttons.addWidget(cancel_button)
        root.addLayout(buttons)

        self.step_count.valueChanged.connect(self.rebuild_steps)
        self.save_enabled.toggled.connect(self.update_saving_controls)
        self.rebuild_steps()
        self.update_saving_controls()

    def execute_step_command(self):
        """Jalankan perintah copy step; bisa dipakai berulang kali."""
        command = self.step_command_input.text().strip()
        # Format yang didukung:
        # copy step 2 for step 3 to 73
        # copy step 2 to step 3-73
        # copy step 3 to step 10
        # copy 3 to 10-15
        pattern = re.compile(
            r"^copy\s+(?:step\s*)?(\d+)\s+"
            r"(?:for\s+step\s+|to\s+(?:step\s*)?)"
            r"(\d+)(?:\s+to\s+(?:step\s*)?(\d+)|\s*-\s*(?:step\s*)?(\d+))?$",
            re.IGNORECASE,
        )
        match = pattern.match(command)
        if not match:
            QMessageBox.warning(
                self, "Format perintah tidak dikenali",
                "Gunakan salah satu format berikut:\n"
                "• copy step 2 for step 3 to 73\n"
                "• copy step 2 to step 3-73\n"
                "• copy step 3 to step 10"
            )
            self.step_command_input.setFocus()
            return

        source = int(match.group(1))
        first_target = int(match.group(2))
        last_target = int(match.group(3) or match.group(4) or first_target)
        total = self.step_count.value()
        if source < 1 or source > total or first_target < 1 or last_target < 1 or last_target > total:
            QMessageBox.warning(
                self, "Nomor Step di luar batas",
                f"Nomor Step harus berada antara 1 sampai {total}."
            )
            return
        if last_target < first_target:
            QMessageBox.warning(
                self, "Rentang Step tidak valid",
                "Nomor tujuan akhir harus lebih besar atau sama dengan tujuan awal."
            )
            return

        # Ambil seluruh perubahan yang belum disimpan dari tampilan saat ini.
        steps = self._collect_live_steps()
        if len(steps) != total:
            steps = json.loads(json.dumps(self.config.get("steps", [])))
            while len(steps) < total:
                steps.append({
                    "name_before": f"Step {len(steps) + 1}",
                    "name_after": f"Step {len(steps) + 1} selesai",
                    "input_type": "barcode", "message_before": "", "groups": [],
                })
        source_config = json.loads(json.dumps(steps[source - 1]))
        for target_number in range(first_target, last_target + 1):
            # Setiap tujuan menerima salinan mandiri dari konfigurasi sumber.
            steps[target_number - 1] = json.loads(json.dumps(source_config))

        self.config["steps"] = steps
        # Paksa rebuild membaca hasil perintah dari config, bukan widget lama.
        self.group_widgets = []
        self.rebuild_steps()
        self.step_command_input.clear()
        self.step_command_input.setFocus()
        copied_count = last_target - first_target + 1
        QMessageBox.information(
            self, "Perintah copy berhasil",
            f"Konfigurasi Step {source} berhasil disalin ke Step "
            f"{first_target}" + (f" sampai {last_target}" if last_target != first_target else "")
            + f".\nTotal tujuan: {copied_count} Step.\n\n"
            "Kamu dapat menjalankan perintah copy lainnya kapan saja sebelum menyimpan konfigurasi."
        )

    def browse_step_xlsx(self, path_input, sheet_combo=None):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Pilih database Excel",
            path_input.text().strip() or str(Path.home()),
            "Excel Workbook (*.xlsx)"
        )
        if path:
            path_input.setText(path)
            if sheet_combo is not None:
                try:
                    wb = load_workbook(path, read_only=True, data_only=True)
                    names = list(wb.sheetnames)
                    wb.close()
                    current = sheet_combo.currentText()
                    sheet_combo.clear()
                    sheet_combo.addItems(names)
                    if current:
                        idx = sheet_combo.findText(current)
                        if idx >= 0: sheet_combo.setCurrentIndex(idx)
                        else: sheet_combo.setEditText(current)
                except Exception as exc:
                    QMessageBox.warning(self, "Excel tidak dapat dibaca", str(exc))

    def rebuild_targets(self):
        # Each CSV destination gets its own tab so the controls do not stack
        # vertically and remain readable in the Edit Program window.
        old_targets = []
        for target in getattr(self, "target_widgets", []):
            old_targets.append({
                "destination": target["destination"].text().strip(),
                "filename": target["filename"].text().strip(),
                "filename_part1": target["filename_part1"].currentData(),
                "filename_part2": target["filename_part2"].currentData(),
                "filename_data1": target["filename_data1"].currentData(),
                "filename_data2": target["filename_data2"].currentData(),
                "filename_text": target["filename_text"].text().strip(),
                "date_prefix": target["date_prefix"].isChecked(),
                "filename_source": target["filename_source"].currentData(),
                "columns": [key for key, box in target.get("column_checks", {}).items() if box.isChecked()],
            })
        if not old_targets:
            old_targets = [dict(item) for item in getattr(self, "saved_targets", [])]
            legacy_columns = self.config.get("saving", {}).get("columns", [])
            for item in old_targets:
                item.setdefault("columns", legacy_columns)

        self.targets_tabs.clear()
        self.target_widgets = []
        for index in range(self.target_count.value()):
            old = old_targets[index] if index < len(old_targets) else {
                "destination": old_targets[0].get("destination", "") if old_targets else "",
                "filename": f"hasil_pengujian_{index + 1}.csv" if index else "hasil_pengujian.csv",
                "filename_part1": "none",
                "filename_part2": "none",
                "filename_text": f"hasil_pengujian_{index + 1}" if index else "hasil_pengujian",
                "date_prefix": False,
                "filename_source": "",
                "columns": [],
            }
            tab = QWidget()
            tab_layout = QVBoxLayout(tab)
            form = QFormLayout()
            tab_layout.addLayout(form)

            destination_row = QWidget()
            destination_layout = QHBoxLayout(destination_row)
            destination_layout.setContentsMargins(0, 0, 0, 0)
            destination = QLineEdit(old.get("destination", ""))
            destination.setPlaceholderText("Pilih folder yang sudah ada")
            browse = QPushButton("Pilih folder...")
            browse.clicked.connect(lambda _, field=destination: self.browse_target_destination(field))
            destination_layout.addWidget(destination, 1)
            destination_layout.addWidget(browse)
            form.addRow("Folder tujuan", destination_row)

            filename_part1 = QComboBox()
            filename_part2 = QComboBox()
            for combo in (filename_part1, filename_part2):
                combo.addItem("None", "none")
                combo.addItem("Tanggal (ddmmyyyy)", "date")
                combo.addItem("Data", "data")
            filename_part1.setCurrentIndex(max(0, filename_part1.findData(old.get("filename_part1", "none"))))
            filename_part2.setCurrentIndex(max(0, filename_part2.findData(old.get("filename_part2", "none"))))
            filename_text = QLineEdit(old.get("filename_text", Path(old.get("filename", "hasil_pengujian.csv")).stem or "hasil_pengujian"))
            filename_text.setPlaceholderText("Contoh: hasil_pengujian")
            form.addRow("Bagian nama file 1", filename_part1)
            data_source1 = QComboBox()
            data_source2 = QComboBox()
            for selector in (data_source1, data_source2):
                selector.addItem("Pilih data ke x step ke x", "")
            source_steps = getattr(self, "group_widgets", None)
            if source_steps:
                step_defs = [
                    {"input_type": st["kind"].currentData(), "groups": st["groups"]}
                    for st in source_steps
                ]
            else:
                step_defs = self.config.get("steps", [])
            for step_no, step_def in enumerate(step_defs, start=1):
                if step_def.get("input_type") != "barcode":
                    continue
                for group_no, group in enumerate(step_def.get("groups", []), start=1):
                    group_name = f"Data {group_no}"
                    if isinstance(group, dict):
                        name_widget = group.get("name")
                        if hasattr(name_widget, "text"):
                            group_name = name_widget.text().strip() or group_name
                        elif isinstance(name_widget, str):
                            group_name = name_widget or group_name
                    label = f"{group_name} — Data {group_no}, Step {step_no}"
                    key = f"data_step_{step_no}_group_{group_no}"
                    data_source1.addItem(label, key)
                    data_source2.addItem(label, key)
            for selector, saved_key in ((data_source1, "filename_data1"), (data_source2, "filename_data2")):
                saved_index = selector.findData(old.get(saved_key, ""))
                selector.setCurrentIndex(max(0, saved_index))
            form.addRow("Pilih data untuk bagian 1", data_source1)
            form.addRow("Bagian nama file 2", filename_part2)
            form.addRow("Pilih data untuk bagian 2", data_source2)
            form.addRow("Bagian nama file 3 (teks bebas)", filename_text)

            def update_data_selector_visibility(_index=None, part=filename_part1, selector=data_source1):
                selector.setVisible(part.currentData() == "data")
            def update_data_selector_visibility2(_index=None, part=filename_part2, selector=data_source2):
                selector.setVisible(part.currentData() == "data")
            filename_part1.currentIndexChanged.connect(update_data_selector_visibility)
            filename_part2.currentIndexChanged.connect(update_data_selector_visibility2)
            update_data_selector_visibility()
            update_data_selector_visibility2()

            # Legacy controls are retained only for compatibility with older config files.
            filename_source = QComboBox()
            filename_source.addItem("Tidak ada tambahan", "")
            date_prefix = QCheckBox("Tanggal lama")
            date_prefix.setChecked(False)
            date_prefix.setVisible(False)
            filename_source.setVisible(False)

            columns_group = QGroupBox("Kolom yang disimpan ke file ini")
            columns_layout = QVBoxLayout(columns_group)
            columns_scroll = QScrollArea()
            columns_scroll.setWidgetResizable(True)
            columns_content = QWidget()
            columns_content_layout = QVBoxLayout(columns_content)
            columns_content_layout.setContentsMargins(4, 4, 4, 4)
            columns_scroll.setWidget(columns_content)
            columns_layout.addWidget(columns_scroll)
            tab_layout.addWidget(columns_group, 1)

            column_checks = {}
            old_columns = old.get("columns", self.config.get("saving", {}).get("columns", []))
            for key, label in self.get_column_options():
                check = QCheckBox(label)
                check.setChecked(key in old_columns)
                columns_content_layout.addWidget(check)
                column_checks[key] = check
            columns_content_layout.addStretch()

            self.targets_tabs.addTab(tab, f"File {index + 1}")
            self.target_widgets.append({
                "destination": destination,
                "filename": filename_text,
                "filename_part1": filename_part1,
                "filename_part2": filename_part2,
                "filename_data1": data_source1,
                "filename_data2": data_source2,
                "filename_text": filename_text,
                "filename_source": filename_source,
                "date_prefix": date_prefix,
                "browse": browse,
                "column_checks": column_checks,
            })

        self.update_saving_controls()

    def browse_target_destination(self, destination_input):
        folder = QFileDialog.getExistingDirectory(
            self, "Pilih folder tujuan", destination_input.text().strip() or str(Path.home())
        )
        if folder:
            destination_input.setText(folder)

    def update_dat_controls(self, _checked=None):
        enabled = self.dat_enabled.isChecked()
        for widget in (self.dat_destination, self.dat_browse, self.dat_filename):
            widget.setEnabled(enabled)

    def update_saving_controls(self):
        enabled = self.save_enabled.isChecked()
        self.target_count.setEnabled(enabled)
        for target in getattr(self, "target_widgets", []):
            for key in ("destination", "filename", "filename_part1", "filename_part2", "filename_data1", "filename_data2", "filename_text", "filename_source", "date_prefix", "browse"):
                target[key].setEnabled(enabled)
            for box in target.get("column_checks", {}).values():
                box.setEnabled(enabled)
        self.line_input.setEnabled(enabled)
        self.result_input.setEnabled(enabled)
        # Nama judul tetap bisa diubah walaupun penyimpanan CSV dimatikan.
        self.title_input.setEnabled(True)
        self.columns_box.setEnabled(enabled)

    def validate_and_accept(self):
        # Validasi konfigurasi Excel dilakukan di masing-masing Data.
        for step_index, step in enumerate(self.group_widgets, start=1):
            for data_index, group in enumerate(step["groups"], start=1):
                if not group["excel_enabled"].isChecked():
                    continue
                excel_path = Path(group["xlsx_path"].text().strip()).expanduser()
                column_letter = group["xlsx_column"].text().strip().upper()
                if not excel_path.is_file():
                    QMessageBox.warning(self, "Database Excel tidak valid",
                        f"Step {step_index}, Data {data_index}: pilih file .xlsx yang tersedia.")
                    self.tabs.setCurrentIndex(0)
                    return
                try:
                    if not column_letter or not column_letter.isalpha():
                        raise ValueError
                    column_index_from_string(column_letter)
                    if group["xlsx_sheet_mode"].currentData() == "specific":
                        sheet_name = group["xlsx_sheet_name"].currentText().strip()
                        if not sheet_name:
                            raise ValueError
                        _wb = load_workbook(str(excel_path), read_only=True, data_only=True)
                        try:
                            if sheet_name not in _wb.sheetnames:
                                raise ValueError(f"Sheet '{sheet_name}' tidak ditemukan.")
                        finally:
                            _wb.close()
                except Exception as validation_error:
                    QMessageBox.warning(self, "Kolom Excel tidak valid",
                        f"Step {step_index}, Data {data_index}: {str(validation_error) or 'kolom/sheet Excel harus diisi (kolom berupa huruf, misalnya G atau AA).'}")
                    self.tabs.setCurrentIndex(0)
                    return

        if self.save_enabled.isChecked():
            for index, target in enumerate(self.target_widgets, start=1):
                destination = Path(target["destination"].text().strip()).expanduser()
                filename_text = target["filename_text"].text().strip()
                if not destination.is_dir():
                    QMessageBox.warning(self, "Folder tidak valid", f"Folder tujuan file {index} tidak ditemukan. Pilih folder yang sudah ada; program tidak akan membuat folder baru.")
                    self.tabs.setCurrentIndex(1)
                    return
                if any(char in filename_text for char in '<>:"/\\|?*'):
                    QMessageBox.warning(self, "Teks nama file tidak valid", f'Teks nama file tujuan {index} tidak boleh memuat karakter \\ / : * ? " < > |.')
                    self.tabs.setCurrentIndex(1)
                    return
                if not any(box.isChecked() for box in target.get("column_checks", {}).values()):
                    QMessageBox.warning(self, "Kolom belum dipilih", f"Pilih minimal satu kolom untuk file tujuan {index}.")
                    self.tabs.setCurrentIndex(1)
                    return
            if not self.result_input.text().strip():
                QMessageBox.warning(self, "Hasil Tes kosong", "Isi nilai yang akan disimpan pada kolom Hasil Tes.")
                self.tabs.setCurrentIndex(1)
                return
        if self.dat_enabled.isChecked():
            dat_folder = Path(self.dat_destination.text().strip()).expanduser()
            dat_name = self.dat_filename.text().strip()
            if not self.dat_destination.text().strip() or not dat_folder.is_dir():
                QMessageBox.warning(self, "Folder .dat tidak valid", "Pilih folder tujuan .dat yang sudah ada.")
                self.tabs.setCurrentIndex(1)
                return
            if not dat_name or dat_name.lower() == ".dat":
                QMessageBox.warning(self, "Nama file .dat kosong", "Isi nama file untuk file .dat.")
                self.tabs.setCurrentIndex(1)
                return
            if any(char in dat_name for char in '<>:"/\\|?*'):
                QMessageBox.warning(self, "Nama file .dat tidak valid", 'Nama file tidak boleh memuat karakter \\ / : * ? " < > |.')
                self.tabs.setCurrentIndex(1)
                return
        self.accept()

    def get_column_options(self):
        steps = getattr(self, "group_widgets", None)
        if steps:
            barcode_steps = [i for i, step in enumerate(steps)
                             if step["kind"].currentData() == "barcode"]
        else:
            barcode_steps = [i for i, step in enumerate(self.config.get("steps", []))
                             if step.get("input_type") == "barcode"]
        options = [
            ("no", "No"), ("hasil_tes", "Hasil Tes"),
            ("waktu_mulai", "Waktu tes dimulai"),
            ("kode_pekerja", "Kode Pekerja"), ("nama_pekerja", "Nama Pekerja"),
        ]
        options += [(f"kode_produk_{n}", f"Kode produk {n} (Step {idx + 1})")
                    for n, idx in enumerate(barcode_steps, start=1)]
        options += [("lini", "Lini"), ("waktu_berakhir", "Waktu tes berakhir")]
        return options

    def rebuild_save_columns(self):
        # Column inclusion is configured separately for each CSV destination.
        # This section only allows editing header names shared by all destinations.
        while self.columns_layout.count():
            item = self.columns_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        configured_names = self.config.get("saving", {}).get("column_names", {})
        live_names = {key: field.text() for key, field in getattr(self, "column_names", {}).items()}
        self.column_checks = {}
        self.column_names = {}
        for key, default_label in self.get_column_options():
            row_widget = QWidget()
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.addWidget(QLabel(default_label))
            name_input = QLineEdit(live_names.get(key, configured_names.get(key, default_label)))
            name_input.setPlaceholderText("Nama kolom CSV")
            row_layout.addWidget(name_input, 1)
            self.columns_layout.addWidget(row_widget)
            self.column_names[key] = name_input

    def _collect_live_steps(self):
        """Capture the current Barcode-tab controls before rebuilding widgets."""
        collected = []
        for step in getattr(self, "group_widgets", []):
            groups = []
            for group in step.get("groups", []):
                groups.append({
                    "name": group["name"].text(),
                    "mode": group["mode"].currentData(),
                    "start": group["start"].value(),
                    "length": group["length"].value(),
                    "excel_enabled": group["excel_enabled"].isChecked(),
                    "xlsx_path": group["xlsx_path"].text().strip(),
                    "xlsx_column": group["xlsx_column"].text().strip().upper(),
                    "xlsx_match_chars": group["xlsx_match_chars"].value(),
                    "xlsx_sheet_mode": group["xlsx_sheet_mode"].currentData(),
                    "xlsx_sheet_name": group["xlsx_sheet_name"].currentText().strip(),
                    "excel_error": group["excel_error"].text(), "excel_popup": group["excel_popup"].isChecked(),
                    "excel_found_result": group["excel_found_result"].currentData(),
                    "char_enabled": group["char_enabled"].isChecked(),
                    "char_start": group["char_start"].value(), "char_length": group["char_length"].value(),
                    "char_expected": group["char_expected"].text(),
                    "char_error": group["char_error"].text(), "char_popup": group["char_popup"].isChecked(),
                    "unique": group["unique"].isChecked(),
                    "unique_error": group["unique_error"].text(), "unique_popup": group["unique_popup"].isChecked(),
                    "judge_ok": group["judge_ok"].text(),
                    "judge_ng": group["judge_ng"].text(),
                })
            collected.append({
                "name_before": step["before"].text(),
                "name_after": step["after"].text(),
                "message_before": step["message_before"].text(),
                "input_type": step["kind"].currentData(),
                "groups": groups,
            })
        return collected

    def rebuild_steps(self, _value=None):
        # Preserve unsaved edits when the number of steps changes.
        live_steps = self._collect_live_steps() if getattr(self, "group_widgets", None) else []
        old_steps = live_steps or self.config.get("steps", [])

        while self.step_area.count():
            item = self.step_area.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
        self.group_widgets = []
        count = self.step_count.value()

        for i in range(count):
            old = old_steps[i] if i < len(old_steps) else {
                "name_before": f"Step {i + 1}",
                "name_after": f"Step {i + 1} selesai",
                "input_type": "barcode",
                "groups": [],
            }
            box = QGroupBox(f"Step {i + 1}")
            form = QFormLayout(box)
            before = QLineEdit(old.get("name_before", ""))
            after = QLineEdit(old.get("name_after", ""))
            kind = QComboBox()
            kind.addItem("Operator ID", "operator")
            kind.addItem("Barcode produk", "barcode")
            idx = kind.findData(old.get("input_type", "barcode"))
            kind.setCurrentIndex(max(0, idx))
            form.addRow("Name Before", before)
            form.addRow("Name After", after)
            message_before = QLineEdit(old.get("message_before", ""))
            message_before.setPlaceholderText("Kosong = pesan scan default")
            form.addRow("Message Before", message_before)
            form.addRow("Jenis input", kind)

            copy_row = QWidget()
            copy_layout = QHBoxLayout(copy_row)
            copy_layout.setContentsMargins(0, 0, 0, 0)
            copy_button = QPushButton("Salin konfigurasi Step ini")
            paste_button = QPushButton("Tempel konfigurasi dari Step yang disalin")
            copy_layout.addWidget(copy_button)
            copy_layout.addWidget(paste_button)
            form.addRow("Duplikasi konfigurasi", copy_row)

            group_count = QSpinBox()
            group_count.setRange(0, 12)
            group_count.setValue(len(old.get("groups", [])))
            form.addRow("Jumlah kelompok data", group_count)
            group_layout = QVBoxLayout()
            form.addRow(group_layout)
            widgets = []
            original_groups = list(old.get("groups", []))

            def collect_groups_live(group_widgets=widgets):
                return [{
                    "name": item["name"].text(),
                    "mode": item["mode"].currentData(),
                    "start": item["start"].value(),
                    "length": item["length"].value(),
                    "excel_enabled": item["excel_enabled"].isChecked(),
                    "xlsx_path": item["xlsx_path"].text().strip(),
                    "xlsx_column": item["xlsx_column"].text().strip().upper(),
                    "xlsx_match_chars": item["xlsx_match_chars"].value(),
                    "xlsx_sheet_mode": item["xlsx_sheet_mode"].currentData(),
                    "xlsx_sheet_name": item["xlsx_sheet_name"].currentText().strip(),
                    "excel_error": item["excel_error"].text(),
                    "excel_popup": item["excel_popup"].isChecked(),
                    "excel_found_result": item["excel_found_result"].currentData(),
                    "char_enabled": item["char_enabled"].isChecked(),
                    "char_start": item["char_start"].value(),
                    "char_length": item["char_length"].value(),
                    "char_expected": item["char_expected"].text(),
                    "char_error": item["char_error"].text(),
                    "char_popup": item["char_popup"].isChecked(),
                    "unique": item["unique"].isChecked(),
                    "unique_error": item["unique_error"].text(),
                    "unique_popup": item["unique_popup"].isChecked(),
                    "judge_ok": item["judge_ok"].text(),
                    "judge_ng": item["judge_ng"].text(),
                } for item in group_widgets]

            def rebuild_groups(_value=None, layout=group_layout, group_widgets=widgets,
                               count_box=group_count, fallback_groups=original_groups,
                               kind_box=kind, collect=collect_groups_live):
                # Crucial: read the live controls before deleting them. Otherwise changing
                # the group count silently discards edited modes, names, and validation.
                current_groups = collect() if group_widgets else list(fallback_groups)
                while layout.count():
                    item = layout.takeAt(0)
                    widget = item.widget()
                    if widget:
                        widget.deleteLater()
                group_widgets.clear()

                for j in range(count_box.value()):
                    g = current_groups[j] if j < len(current_groups) else {
                        "name": f"Data {j + 1}", "start": 1, "length": 1,
                        "mode": "length", "excel_enabled": False, "xlsx_path": "", "xlsx_column": "G", "xlsx_match_chars": 8, "xlsx_sheet_mode": "specific", "xlsx_sheet_name": "Sheet1",
                        "excel_error": "Barcode sudah terdaftar di database Excel.", "excel_popup": False,
                        "char_enabled": False, "char_start": 1, "char_length": 3, "char_expected": "ABC",
                        "char_error": "Format karakter barcode tidak sesuai.", "char_popup": False,
                        "unique": False, "unique_error": "Barcode sudah pernah digunakan.", "unique_popup": False,
                        "judge_ok": "OK", "judge_ng": "NG",
                    }
                    group_box = QGroupBox(f"Data {j + 1}")
                    gf = QFormLayout(group_box)
                    name = QLineEdit(g.get("name", f"Data {j + 1}"))
                    mode = QComboBox()
                    mode.addItem("Jumlah karakter", "length")
                    mode.addItem("Seluruh barcode", "all")
                    if kind_box.currentData() == "operator":
                        mode.addItem("Nama operator", "operator_name")
                    requested_mode = g.get("mode", "length")
                    mode_index = mode.findData(requested_mode)
                    mode.setCurrentIndex(mode_index if mode_index >= 0 else 0)
                    start = QSpinBox()
                    start.setRange(1, 9999)
                    start.setValue(max(1, int(g.get("start", 1))))
                    length = QSpinBox()
                    length.setRange(1, 9999)
                    length.setValue(max(1, int(g.get("length", 1))))
                    # Validasi 1: blacklist Excel
                    excel_box = QGroupBox("1. Validasi Database Excel (barcode ditemukan = NG)")
                    ef = QFormLayout(excel_box)
                    excel_enabled = QCheckBox("Aktifkan validasi Excel")
                    excel_enabled.setChecked(bool(g.get("excel_enabled", g.get("validate", False) and g.get("source") == "excel")))
                    xlsx_path_row = QWidget()
                    xlsx_path_layout = QHBoxLayout(xlsx_path_row)
                    xlsx_path_layout.setContentsMargins(0, 0, 0, 0)
                    xlsx_path = QLineEdit(g.get("xlsx_path", ""))
                    xlsx_path.setPlaceholderText("Pilih file database .xlsx")
                    xlsx_browse = QPushButton("Pilih file...")
                    xlsx_path_layout.addWidget(xlsx_path, 1); xlsx_path_layout.addWidget(xlsx_browse)
                    xlsx_column = QLineEdit(g.get("xlsx_column", "G")); xlsx_column.setMaximumWidth(100)
                    xlsx_match_chars = QSpinBox(); xlsx_match_chars.setRange(1, 9999)
                    xlsx_match_chars.setValue(max(1, int(g.get("xlsx_match_chars", 8))))
                    xlsx_sheet_mode = QComboBox()
                    xlsx_sheet_mode.addItem("Sheet tertentu", "specific")
                    xlsx_sheet_mode.addItem("Semua sheet", "all")
                    sheet_mode_index = xlsx_sheet_mode.findData(g.get("xlsx_sheet_mode", "specific"))
                    xlsx_sheet_mode.setCurrentIndex(sheet_mode_index if sheet_mode_index >= 0 else 0)
                    xlsx_sheet_name = QComboBox(); xlsx_sheet_name.setEditable(True)
                    xlsx_sheet_name.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
                    xlsx_sheet_name.setMinimumWidth(180)
                    configured_sheet = g.get("xlsx_sheet_name", "Sheet1")
                    if xlsx_path.text().strip() and Path(xlsx_path.text().strip()).is_file():
                        try:
                            _wb = load_workbook(xlsx_path.text().strip(), read_only=True, data_only=True)
                            xlsx_sheet_name.addItems(_wb.sheetnames); _wb.close()
                        except Exception:
                            pass
                    if configured_sheet:
                        idx = xlsx_sheet_name.findText(configured_sheet)
                        if idx >= 0: xlsx_sheet_name.setCurrentIndex(idx)
                        else: xlsx_sheet_name.setEditText(configured_sheet)
                    # Hubungkan tombol setelah combo nama sheet dibuat.
                    xlsx_browse.clicked.connect(lambda _, field=xlsx_path, sheet=xlsx_sheet_name: self.browse_step_xlsx(field, sheet))
                    excel_found = QComboBox()
                    excel_found.addItem("NG (ditolak)", "ng")
                    excel_found.addItem("OK (diterima)", "ok")
                    found_index = excel_found.findData(g.get("excel_found_result", "ng"))
                    excel_found.setCurrentIndex(found_index if found_index >= 0 else 0)
                    default_excel_msg = EXCEL_MSG_NOTFOUND_NG if excel_found.currentData() == "ok" else EXCEL_MSG_FOUND_NG
                    excel_error = QLineEdit(g.get("excel_error", default_excel_msg))
                    excel_popup = QCheckBox("Tampilkan pop-up untuk error Excel")
                    excel_popup.setChecked(bool(g.get("excel_popup", False)))
                    ef.addRow(excel_enabled); ef.addRow("Jika barcode ditemukan di Excel, hasilnya", excel_found); ef.addRow("File database Excel", xlsx_path_row)
                    ef.addRow("Kolom Excel", xlsx_column); ef.addRow("Jumlah karakter awal untuk pencarian", xlsx_match_chars)
                    ef.addRow("Cakupan sheet", xlsx_sheet_mode); ef.addRow("Nama sheet", xlsx_sheet_name)
                    ef.addRow("Message jika NG", excel_error); ef.addRow(excel_popup)

                    # Validasi 2: karakter tertentu pada posisi tertentu
                    char_box = QGroupBox("2. Validasi Karakter / Format Barcode")
                    cf = QFormLayout(char_box)
                    legacy_expected = g.get("expected", g.get("xlsx_required_prefix", ""))
                    legacy_char_enabled = bool(g.get("validate", False) and g.get("source", "text") == "text" and legacy_expected)
                    legacy_prefix_enabled = bool(g.get("xlsx_required_prefix", ""))
                    char_enabled = QCheckBox("Aktifkan validasi karakter")
                    char_enabled.setChecked(bool(g.get("char_enabled", legacy_char_enabled or legacy_prefix_enabled)))
                    char_start = QSpinBox(); char_start.setRange(1, 9999); char_start.setValue(max(1, int(g.get("char_start", 1))))
                    char_length = QSpinBox(); char_length.setRange(1, 9999); char_length.setValue(max(1, int(g.get("char_length", len(legacy_expected) if legacy_expected else 3))))
                    char_expected = QLineEdit(g.get("char_expected", legacy_expected))
                    char_expected.setPlaceholderText("Contoh: ABC")
                    char_error = QLineEdit(g.get("char_error", "Format karakter barcode tidak sesuai."))
                    char_popup = QCheckBox("Tampilkan pop-up untuk error karakter")
                    char_popup.setChecked(bool(g.get("char_popup", False)))
                    cf.addRow(char_enabled); cf.addRow("Mulai dari karakter ke-", char_start)
                    cf.addRow("Jumlah karakter", char_length); cf.addRow("Karakter yang diwajibkan", char_expected)
                    cf.addRow("Message error karakter", char_error); cf.addRow(char_popup)

                    # Validasi 3: keunikan barcode
                    unique_box = QGroupBox("3. Validasi Data Unik")
                    uf = QFormLayout(unique_box)
                    unique = QCheckBox("Data harus unik (tidak boleh sama dengan scan sebelumnya dalam satu siklus)")
                    unique.setChecked(bool(g.get("unique", False)))
                    unique_error = QLineEdit(g.get("unique_error", "Barcode sudah pernah digunakan."))
                    unique_popup = QCheckBox("Tampilkan pop-up untuk error duplikat")
                    unique_popup.setChecked(bool(g.get("unique_popup", False)))
                    uf.addRow(unique); uf.addRow("Message error data unik", unique_error); uf.addRow(unique_popup)

                    judge_ok = QLineEdit(g.get("judge_ok", "OK")); judge_ng = QLineEdit(g.get("judge_ng", "NG"))
                    judge_ok.setPlaceholderText("Contoh: PASS / ACCEPTED"); judge_ng.setPlaceholderText("Contoh: FAIL / REJECT")
                    gf.addRow("Nama kolom", name); gf.addRow("Mode pengambilan", mode)
                    gf.addRow("Mulai karakter ke- (data output)", start); gf.addRow("Jumlah karakter (data output)", length)
                    gf.addRow(excel_box); gf.addRow(char_box); gf.addRow(unique_box)
                    gf.addRow("Tulisan Judge jika valid", judge_ok); gf.addRow("Tulisan Judge jika tidak valid", judge_ng)

                    def update_validation_controls(_value=None, ebox=excel_enabled, path=xlsx_path,
                                                   browse=xlsx_browse, col=xlsx_column, chars=xlsx_match_chars,
                                                   sheet_mode=xlsx_sheet_mode, sheet_name=xlsx_sheet_name,
                                                   cbox=char_enabled, cstart=char_start, clen=char_length,
                                                   cexpected=char_expected, eerr=excel_error, epop=excel_popup, efound=excel_found,
                                                   cerr=char_error, cpop=char_popup, ubox=unique,
                                                   uerr=unique_error, upop=unique_popup):
                        active = ebox.isChecked()
                        path.setEnabled(active); browse.setEnabled(active); col.setEnabled(active); chars.setEnabled(active)
                        sheet_mode.setEnabled(active); sheet_name.setEnabled(active and sheet_mode.currentData() == "specific")
                        eerr.setEnabled(active); epop.setEnabled(active); efound.setEnabled(active)
                        cactive = cbox.isChecked()
                        cstart.setEnabled(cactive); clen.setEnabled(cactive); cexpected.setEnabled(cactive)
                        cerr.setEnabled(cactive); cpop.setEnabled(cactive)
                        uerr.setEnabled(ubox.isChecked()); upop.setEnabled(ubox.isChecked())
                    xlsx_sheet_mode.currentIndexChanged.connect(update_validation_controls)
                    excel_enabled.toggled.connect(update_validation_controls)
                    char_enabled.toggled.connect(update_validation_controls)
                    unique.toggled.connect(update_validation_controls)
                    update_validation_controls()

                    def update_excel_found(_index=None, combo=excel_found, box=excel_box, err=excel_error):
                        ok_mode = combo.currentData() == "ok"
                        box.setTitle(
                            "1. Validasi Database Excel (barcode ditemukan = OK, tidak ditemukan = NG)" if ok_mode
                            else "1. Validasi Database Excel (barcode ditemukan = NG)"
                        )
                        # Ganti pesan bawaan agar sesuai mode; pesan buatan sendiri tidak diubah.
                        current = err.text().strip()
                        if ok_mode and current in ("", EXCEL_MSG_FOUND_NG):
                            err.setText(EXCEL_MSG_NOTFOUND_NG)
                        elif not ok_mode and current in ("", EXCEL_MSG_NOTFOUND_NG):
                            err.setText(EXCEL_MSG_FOUND_NG)
                    excel_found.currentIndexChanged.connect(update_excel_found)
                    update_excel_found()


                    def update_group_mode(_index=None, mode_box=mode,
                                          start_box=start, length_box=length):
                        selected_mode = mode_box.currentData()
                        # Only this group's character-range controls are disabled.
                        # Selecting "Seluruh barcode" must not disable other data groups.
                        uses_range = selected_mode == "length"
                        start_box.setEnabled(uses_range)
                        length_box.setEnabled(uses_range)

                    mode.currentIndexChanged.connect(update_group_mode)
                    update_group_mode()
                    layout.addWidget(group_box)
                    group_widgets.append({
                        "name": name, "mode": mode, "start": start,
                        "length": length, "excel_enabled": excel_enabled, "xlsx_path": xlsx_path,
                        "xlsx_column": xlsx_column, "xlsx_match_chars": xlsx_match_chars,
                        "xlsx_sheet_mode": xlsx_sheet_mode, "xlsx_sheet_name": xlsx_sheet_name,
                        "excel_error": excel_error, "excel_popup": excel_popup, "excel_found_result": excel_found,
                        "char_enabled": char_enabled, "char_start": char_start,
                        "char_length": char_length, "char_expected": char_expected,
                        "char_error": char_error, "char_popup": char_popup,
                        "unique": unique, "unique_error": unique_error, "unique_popup": unique_popup,
                        "judge_ok": judge_ok, "judge_ng": judge_ng,
                    })

            group_count.valueChanged.connect(rebuild_groups)
            rebuild_groups()

            def update_operator_modes(_index=None, kind_box=kind, group_widgets=widgets):
                is_operator = kind_box.currentData() == "operator"
                for group_widget in group_widgets:
                    mode_box = group_widget["mode"]
                    selected = mode_box.currentData()
                    has_operator_name = mode_box.findData("operator_name") >= 0
                    if is_operator and not has_operator_name:
                        mode_box.addItem("Nama operator", "operator_name")
                    elif not is_operator and has_operator_name:
                        remove_index = mode_box.findData("operator_name")
                        mode_box.removeItem(remove_index)
                    restored = mode_box.findData(selected)
                    if restored >= 0:
                        mode_box.setCurrentIndex(restored)
                    # Reapply enabled/disabled state after changing input type.
                    uses_range = mode_box.currentData() == "length"
                    group_widget["start"].setEnabled(uses_range)
                    group_widget["length"].setEnabled(uses_range)

            kind.currentIndexChanged.connect(update_operator_modes)
            update_operator_modes()
            self.step_area.addWidget(box)
            step_widget = {
                "before": before, "after": after, "kind": kind,
                "message_before": message_before,
                "group_count": group_count, "groups": widgets,
                "rebuild_groups": rebuild_groups,
            }
            self.group_widgets.append(step_widget)

            def copy_this_step(_checked=False, source_index=i):
                self._copied_step_config = self._collect_live_steps()[source_index]
                QMessageBox.information(self, "Konfigurasi disalin",
                    f"Konfigurasi Step {source_index + 1} berhasil disalin. Pilih Tempel pada Step tujuan.")

            def paste_to_this_step(_checked=False, target=step_widget):
                copied = getattr(self, "_copied_step_config", None)
                if not copied:
                    QMessageBox.warning(self, "Belum ada konfigurasi", "Salin konfigurasi dari Step sumber terlebih dahulu.")
                    return
                target["before"].setText(copied.get("name_before", ""))
                target["after"].setText(copied.get("name_after", ""))
                target["message_before"].setText(copied.get("message_before", ""))
                input_index = target["kind"].findData(copied.get("input_type", "barcode"))
                if input_index >= 0:
                    target["kind"].setCurrentIndex(input_index)
                target["group_count"].setValue(len(copied.get("groups", [])))
                # Rebuild uses current live values; inject copied group definitions as fallback.
                for current_group in target["groups"]:
                    current_group["name"].setText("")
                target["rebuild_groups"]()
                # Set controls from copied values after widget reconstruction.
                for control, data in zip(target["groups"], copied.get("groups", [])):
                    control["name"].setText(data.get("name", ""))
                    idx = control["mode"].findData(data.get("mode", "length"))
                    if idx >= 0: control["mode"].setCurrentIndex(idx)
                    control["start"].setValue(max(1, int(data.get("start", 1))))
                    control["length"].setValue(max(1, int(data.get("length", 1))))
                    control["excel_enabled"].setChecked(bool(data.get("excel_enabled", data.get("validate", False) and data.get("source") == "excel")))
                    control["xlsx_path"].setText(data.get("xlsx_path", ""))
                    control["xlsx_column"].setText(data.get("xlsx_column", "G"))
                    control["xlsx_match_chars"].setValue(max(1, int(data.get("xlsx_match_chars", 8))))
                    smi = control["xlsx_sheet_mode"].findData(data.get("xlsx_sheet_mode", "specific"))
                    if smi >= 0: control["xlsx_sheet_mode"].setCurrentIndex(smi)
                    control["xlsx_sheet_name"].setEditText(data.get("xlsx_sheet_name", "Sheet1"))
                    fi = control["excel_found_result"].findData(data.get("excel_found_result", "ng"))
                    control["excel_found_result"].setCurrentIndex(fi if fi >= 0 else 0)
                    control["excel_error"].setText(data.get(
                        "excel_error",
                        EXCEL_MSG_NOTFOUND_NG if data.get("excel_found_result") == "ok" else EXCEL_MSG_FOUND_NG
                    ))
                    control["excel_popup"].setChecked(bool(data.get("excel_popup", False)))
                    legacy_expected = data.get("expected", data.get("xlsx_required_prefix", ""))
                    legacy_char_enabled = bool(data.get("validate", False) and data.get("source", "text") == "text" and legacy_expected)
                    control["char_enabled"].setChecked(bool(data.get("char_enabled", legacy_char_enabled or bool(data.get("xlsx_required_prefix", "")))))
                    control["char_start"].setValue(max(1, int(data.get("char_start", 1))))
                    control["char_length"].setValue(max(1, int(data.get("char_length", len(legacy_expected) if legacy_expected else 3))))
                    control["char_expected"].setText(data.get("char_expected", legacy_expected))
                    control["char_error"].setText(data.get("char_error", "Format karakter barcode tidak sesuai."))
                    control["char_popup"].setChecked(bool(data.get("char_popup", False)))
                    control["unique"].setChecked(bool(data.get("unique", False)))
                    control["unique_error"].setText(data.get("unique_error", "Barcode sudah pernah digunakan."))
                    control["unique_popup"].setChecked(bool(data.get("unique_popup", False)))
                    control["judge_ok"].setText(data.get("judge_ok", "OK"))
                    control["judge_ng"].setText(data.get("judge_ng", "NG"))
                QMessageBox.information(self, "Konfigurasi ditempel", "Konfigurasi berhasil diterapkan ke Step ini.")

            copy_button.clicked.connect(copy_this_step)
            paste_button.clicked.connect(paste_to_this_step)

        self.rebuild_save_columns()
        # Refresh the filename data-source choices after step/group definitions change.
        if getattr(self, "target_widgets", None):
            self.rebuild_targets()

    def test_ip_connection(self):
        ok, reason = check_ip_connection(self.ip_host.text(), self.ip_port.value())
        if ok:
            QMessageBox.information(self, "Tes Koneksi", f"Tersambung ke {self.ip_host.text().strip()}.")
        else:
            QMessageBox.critical(
                self, "Tes Koneksi",
                f"Tidak tersambung ke {self.ip_host.text().strip() or '-'}.\n\n{reason}"
            )

    def get_config(self):
        steps = []
        for step in self.group_widgets:
            groups = []
            for g in step["groups"]:
                groups.append({
                    "name": g["name"].text().strip() or "Data", "mode": g["mode"].currentData(),
                    "start": g["start"].value(), "length": g["length"].value(),
                    "excel_enabled": g["excel_enabled"].isChecked(),
                    "xlsx_path": g["xlsx_path"].text().strip(),
                    "xlsx_column": g["xlsx_column"].text().strip().upper(),
                    "xlsx_match_chars": g["xlsx_match_chars"].value(),
                    "xlsx_sheet_mode": g["xlsx_sheet_mode"].currentData(),
                    "xlsx_sheet_name": g["xlsx_sheet_name"].currentText().strip(),
                    "excel_error": g["excel_error"].text(), "excel_popup": g["excel_popup"].isChecked(),
                    "excel_found_result": g["excel_found_result"].currentData() or "ng",
                    "char_enabled": g["char_enabled"].isChecked(),
                    "char_start": g["char_start"].value(), "char_length": g["char_length"].value(),
                    "char_expected": g["char_expected"].text(),
                    "char_error": g["char_error"].text(), "char_popup": g["char_popup"].isChecked(),
                    "unique": g["unique"].isChecked(),
                    "unique_error": g["unique_error"].text(), "unique_popup": g["unique_popup"].isChecked(),
                    "judge_ok": g["judge_ok"].text().strip() or "OK",
                    "judge_ng": g["judge_ng"].text().strip() or "NG",
                })
            steps.append({
                "name_before": step["before"].text(),
                "name_after": step["after"].text(),
                "message_before": step["message_before"].text(),
                "input_type": step["kind"].currentData(),
                "groups": groups,
            })
        self.config["steps"] = steps
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
        targets = []
        for target in self.target_widgets:
            targets.append({
                "destination": target["destination"].text().strip(),
                "filename": target["filename"].text().strip(),
                "filename_part1": target["filename_part1"].currentData(),
                "filename_part2": target["filename_part2"].currentData(),
                "filename_data1": target["filename_data1"].currentData(),
                "filename_data2": target["filename_data2"].currentData(),
                "filename_text": target["filename_text"].text().strip(),
                "date_prefix": target["date_prefix"].isChecked(),
                "filename_source": target["filename_source"].currentData(),
                "columns": [key for key, box in target.get("column_checks", {}).items() if box.isChecked()],
            })
        self.config["saving"] = {
            "enabled": self.save_enabled.isChecked(),
            "targets": targets,
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
        except OSError:
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
        return values

    def validate_xlsx_group(self, group, value):
        """Cek barcode ke database Excel. Default (excel_found_result = "ng"): ditemukan = NG,
        tidak ditemukan = OK. Dengan "ok": ditemukan = OK, tidak ditemukan = NG."""
        found_ok = group.get("excel_found_result", "ng") == "ok"
        file_path = group.get("xlsx_path", "").strip()
        column_letter = group.get("xlsx_column", "G").strip().upper()
        match_chars = max(1, int(group.get("xlsx_match_chars", 8)))
        sheet_mode = group.get("xlsx_sheet_mode", "specific")
        sheet_name = group.get("xlsx_sheet_name", "Sheet1")
        normalized = value.strip().upper()
        if not file_path:
            return False, "Pilih file database Excel untuk data ini."
        if len(normalized) < match_chars:
            return False, f"Data hanya {len(normalized)} karakter; dibutuhkan {match_chars} karakter untuk pemeriksaan Excel."
        values = self.get_xlsx_column_values(file_path, column_letter, sheet_mode, sheet_name)
        prefix = normalized[:match_chars]
        matches = [(name, row_number) for item, name, row_number in values
                   if len(item) >= match_chars and item[:match_chars] == prefix]
        if matches:
            sheet, row_number = matches[0]
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
                    start = max(0, int(group.get("start", 1)) - 1)
                    length = max(0, int(group.get("length", 1)))
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
                start = max(0, int(group.get("start", 1)) - 1)
                length = max(0, int(group.get("length", 1)))
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
                char_start = max(1, int(group.get("char_start", 1)))
                char_length = max(1, int(group.get("char_length", 1)))
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
            start = max(0, int(group.get("start", 1)) - 1)
            length = max(0, int(group.get("length", 1)))
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

        output_path = folder / f"{name}.dat"
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
            filename += free_text
            # Keep a usable filename if all three configurable components are empty.
            filename = filename or "hasil_pengujian"
            filename += ".csv"

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
        if password != "1234":
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
        if password != "checkerPE1234":
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
    window = BarcodeApp()
    window.show()
    sys.exit(app.exec())
