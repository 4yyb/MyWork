import sys
import json
import csv
from datetime import datetime
from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QMessageBox, QHeaderView, QDialog, QFormLayout, QSpinBox,
    QComboBox, QCheckBox, QGroupBox, QScrollArea, QInputDialog, QFileDialog,
    QTabWidget
)

APP_DIR = Path.home() / ".barcode_input_app"
APP_DIR.mkdir(exist_ok=True)

CONFIG_FILE = APP_DIR / "config.json"
SESSION_FILE = APP_DIR / "session.json"

EMPTY_CONFIG = {
    "operators": {},
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
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return default


def save_json(path, data):
    path.write_text(
        json.dumps(data, indent=2, ensure_ascii=False),
        encoding="utf-8"
    )


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

        # TAB 1: Operator (scrollable)
        operator_tab = QScrollArea()
        operator_tab.setWidgetResizable(True)
        operator_content = QWidget()
        operator_layout = QVBoxLayout(operator_content)
        operator_layout.addWidget(QLabel("Daftarkan barcode operator dan nama pekerja."))
        self.operator_table = QTableWidget()
        self.operator_table.setColumnCount(2)
        self.operator_table.setHorizontalHeaderLabels(["Barcode Operator", "Nama Pekerja"])
        self.operator_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.operator_table.setRowCount(0)
        for barcode, name in self.config.get("operators", {}).items():
            self.add_operator_row(barcode, name)
        operator_buttons = QHBoxLayout()
        add_operator_button = QPushButton("Tambah Operator")
        remove_operator_button = QPushButton("Hapus Baris Terpilih")
        add_operator_button.clicked.connect(lambda: self.add_operator_row())
        remove_operator_button.clicked.connect(self.remove_operator_row)
        operator_buttons.addWidget(add_operator_button)
        operator_buttons.addWidget(remove_operator_button)
        operator_buttons.addStretch()
        operator_layout.addWidget(self.operator_table)
        operator_layout.addLayout(operator_buttons)
        operator_layout.addStretch()
        operator_tab.setWidget(operator_content)
        self.tabs.addTab(operator_tab, "1. Operator")

        # TAB 2: Barcode / workflow steps
        barcode_tab = QWidget()
        barcode_layout = QVBoxLayout(barcode_tab)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QWidget()
        self.form = QVBoxLayout(content)
        self.step_count = QSpinBox()
        self.step_count.setRange(1, 20)
        self.step_count.setValue(len(self.config.get("steps", [])) or 1)
        self.form.addWidget(QLabel("Jumlah step"))
        self.form.addWidget(self.step_count)
        self.step_area = QVBoxLayout()
        self.form.addLayout(self.step_area)
        scroll.setWidget(content)
        barcode_layout.addWidget(scroll)
        self.tabs.addTab(barcode_tab, "2. Barcode")

        # TAB 3: Storage
        saving_tab = QWidget()
        saving_layout = QVBoxLayout(saving_tab)
        saving_form_widget = QWidget()
        saving_form = QFormLayout(saving_form_widget)
        saving = self.config.get("saving", {})
        self.title_input = QLineEdit(self.config.get("display_title", "BARCODE DATA INPUT"))
        self.title_input.setPlaceholderText("Judul pada tampilan utama")
        saving_form.addRow("Judul tampilan utama", self.title_input)
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
        self.target_count.setRange(1, 10)
        self.target_count.setValue(max(1, min(10, len(saved_targets))))
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

        # Build targets only after all widgets used by update_saving_controls exist.
        self.rebuild_targets()
        saving_scroll = QScrollArea()
        saving_scroll.setWidgetResizable(True)
        saving_scroll.setWidget(saving_form_widget)
        saving_layout.addWidget(saving_scroll)
        self.tabs.addTab(saving_tab, "3. Penyimpanan")

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

    def add_operator_row(self, barcode="", name=""):
        row = self.operator_table.rowCount()
        self.operator_table.insertRow(row)
        self.operator_table.setItem(row, 0, QTableWidgetItem(str(barcode)))
        self.operator_table.setItem(row, 1, QTableWidgetItem(str(name)))

    def remove_operator_row(self):
        rows = sorted({index.row() for index in self.operator_table.selectionModel().selectedRows()}, reverse=True)
        if not rows and self.operator_table.currentRow() >= 0:
            rows = [self.operator_table.currentRow()]
        for row in rows:
            self.operator_table.removeRow(row)

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
        operators = {}
        for row in range(self.operator_table.rowCount()):
            barcode_item = self.operator_table.item(row, 0)
            name_item = self.operator_table.item(row, 1)
            barcode = barcode_item.text().strip() if barcode_item else ""
            name = name_item.text().strip() if name_item else ""
            if not barcode and not name:
                continue
            if not barcode or not name:
                QMessageBox.warning(self, "Data operator belum lengkap", f"Lengkapi barcode dan nama operator pada baris {row + 1}.")
                self.tabs.setCurrentIndex(0)
                return
            if barcode in operators:
                QMessageBox.warning(self, "Barcode duplikat", f"Barcode operator {barcode} terdaftar lebih dari sekali.")
                self.tabs.setCurrentIndex(0)
                return
            operators[barcode] = name

        if self.save_enabled.isChecked():
            for index, target in enumerate(self.target_widgets, start=1):
                destination = Path(target["destination"].text().strip()).expanduser()
                filename_text = target["filename_text"].text().strip()
                if not destination.is_dir():
                    QMessageBox.warning(self, "Folder tidak valid", f"Folder tujuan file {index} tidak ditemukan. Pilih folder yang sudah ada; program tidak akan membuat folder baru.")
                    self.tabs.setCurrentIndex(2)
                    return
                if any(char in filename_text for char in '<>:/\\|?*'):
                    QMessageBox.warning(self, "Teks nama file tidak valid", f'Teks nama file tujuan {index} tidak boleh memuat karakter \\ / : * ? " < > |.')
                    self.tabs.setCurrentIndex(2)
                    return
                if not any(box.isChecked() for box in target.get("column_checks", {}).values()):
                    QMessageBox.warning(self, "Kolom belum dipilih", f"Pilih minimal satu kolom untuk file tujuan {index}.")
                    self.tabs.setCurrentIndex(2)
                    return
            if not self.result_input.text().strip():
                QMessageBox.warning(self, "Hasil Tes kosong", "Isi nilai yang akan disimpan pada kolom Hasil Tes.")
                self.tabs.setCurrentIndex(2)
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

    def rebuild_steps(self, _value=None):
        while self.step_area.count():
            item = self.step_area.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
        self.group_widgets = []
        old_steps = self.config.get("steps", [])
        count = self.step_count.value()
        for i in range(count):
            old = old_steps[i] if i < len(old_steps) else {
                "name_before": f"Step {i + 1}", "name_after": f"Step {i + 1} selesai",
                "input_type": "barcode", "groups": []
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
            form.addRow("Jenis input", kind)
            group_count = QSpinBox()
            group_count.setRange(0, 12)
            group_count.setValue(len(old.get("groups", [])))
            form.addRow("Jumlah kelompok data", group_count)
            group_layout = QVBoxLayout()
            form.addRow(group_layout)
            widgets = []

            def rebuild_groups(_value=None, layout=group_layout, widgets=widgets, count_box=group_count, old_groups=old.get("groups", [])):
                while layout.count():
                    item = layout.takeAt(0)
                    w = item.widget()
                    if w:
                        w.deleteLater()
                widgets.clear()
                for j in range(count_box.value()):
                    g = old_groups[j] if j < len(old_groups) else {
                        "name": f"Data {j + 1}", "start": 1, "length": 1,
                        "mode": "length", "validate": False, "expected": ""
                    }
                    group_box = QGroupBox(f"Kelompok {j + 1}")
                    gf = QFormLayout(group_box)
                    name = QLineEdit(g.get("name", f"Data {j + 1}"))
                    mode = QComboBox()
                    mode.addItem("Jumlah karakter", "length")
                    mode.addItem("Seluruh barcode", "all")
                    if kind.currentData() == "operator":
                        mode.addItem("Nama operator", "operator_name")
                    mode.setCurrentIndex(max(0, mode.findData(g.get("mode", "length"))))
                    start = QSpinBox(); start.setRange(1, 9999); start.setValue(g.get("start", 1))
                    length = QSpinBox(); length.setRange(1, 9999); length.setValue(max(1, g.get("length", 1)))
                    validate = QCheckBox("Aktifkan validasi teks"); validate.setChecked(g.get("validate", False))
                    expected = QLineEdit(g.get("expected", ""))
                    gf.addRow("Nama kolom", name)
                    gf.addRow("Mode pengambilan", mode)
                    gf.addRow("Mulai karakter ke-", start)
                    gf.addRow("Jumlah karakter", length)
                    gf.addRow(validate)
                    gf.addRow("Teks yang harus cocok", expected)
                    def update_mode(data, length=length, start=start):
                        all_mode = data == "all"
                        start.setEnabled(not all_mode); length.setEnabled(not all_mode)
                    mode.currentIndexChanged.connect(lambda _, m=mode: update_mode(m.currentData()))
                    update_mode(mode.currentData())
                    layout.addWidget(group_box)
                    widgets.append({"name": name, "mode": mode, "start": start, "length": length, "validate": validate, "expected": expected})
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

            kind.currentIndexChanged.connect(update_operator_modes)
            update_operator_modes()
            self.step_area.addWidget(box)
            self.group_widgets.append({"before": before, "after": after, "kind": kind,
                                       "group_count": group_count, "groups": widgets})
        self.rebuild_save_columns()
        # Refresh each destination's available column list after step changes.
        if getattr(self, "target_widgets", None):
            self.rebuild_targets()

    def get_config(self):
        steps = []
        for step in self.group_widgets:
            groups = []
            for g in step["groups"]:
                groups.append({
                    "name": g["name"].text().strip() or "Data", "mode": g["mode"].currentData(),
                    "start": g["start"].value(), "length": g["length"].value(),
                    "validate": g["validate"].isChecked(), "expected": g["expected"].text()
                })
            steps.append({"name_before": step["before"].text(), "name_after": step["after"].text(),
                          "input_type": step["kind"].currentData(),
                          "groups": groups})
        operators = {}
        for row in range(self.operator_table.rowCount()):
            barcode_item = self.operator_table.item(row, 0)
            name_item = self.operator_table.item(row, 1)
            barcode = barcode_item.text().strip() if barcode_item else ""
            name = name_item.text().strip() if name_item else ""
            if barcode and name:
                operators[barcode] = name
        self.config["operators"] = operators
        self.config["steps"] = steps
        self.config["display_title"] = self.title_input.text().strip() or "BARCODE DATA INPUT"
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
        return self.config


class BarcodeApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Barcode Data Input")
        self.resize(1200, 760)

        self.config = load_json(CONFIG_FILE, EMPTY_CONFIG)
        if not CONFIG_FILE.exists():
            save_json(CONFIG_FILE, self.config)

        old_session = load_json(SESSION_FILE, {"total": 0})
        self.total = max(0, int(old_session.get("total", 0)))

        self.current_step = 0
        self.current_operator = ""
        self.running = False
        self.last_barcode = ""
        self.cycle_started_at = None
        self.cycle_operator_code = ""
        self.cycle_operator_name = ""
        self.cycle_product_codes = {}

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
        self.total_label = QLabel()

        self.start_button.clicked.connect(self.start_program)
        self.stop_button.clicked.connect(self.stop_program)
        self.reset_button.clicked.connect(self.reset_total)
        self.edit_button.clicked.connect(self.edit_program)

        controls.addWidget(self.start_button)
        controls.addWidget(self.stop_button)
        controls.addWidget(self.reset_button)
        controls.addWidget(self.edit_button)
        controls.addStretch()
        controls.addWidget(self.total_label)

        root.addLayout(controls)
        self.setCentralWidget(central)

    def refresh_table(self):
        steps = self.config["steps"]

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

                if row == self.current_step:
                    item.setBackground(Qt.yellow)
                else:
                    item.setBackground(Qt.white)

    def show_step_message(self):
        if not self.config["steps"]:
            self.message.setText("Belum ada step. Buka Edit Program.")
            return

        step = self.config["steps"][self.current_step]
        if step["input_type"] == "operator":
            msg = "Silakan scan ID operator barcode"
        else:
            msg = "Silakan scan barcode produk"

        if not self.running:
            msg = "STOP — Tekan Start untuk melanjutkan. " + msg

        self.message.setText(msg)
        self.highlight_step()

    def update_controls(self):
        self.start_button.setEnabled(not self.running)
        self.stop_button.setEnabled(self.running)
        self.edit_button.setEnabled(not self.running)
        self.scan_input.setEnabled(self.running)

    def start_program(self):
        # Start always begins a fresh cycle from the first configured step.
        self.current_step = 0
        self.clear_cycle_data()
        self.current_operator = ""
        self.cycle_started_at = None
        self.cycle_operator_code = ""
        self.cycle_operator_name = ""
        self.cycle_product_codes = {}
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
        for row, step in enumerate(self.config["steps"]):
            # Kembalikan Name ke kondisi sebelum input
            self.put(row, 1, step.get("name_before", ""))

            # Bersihkan seluruh kolom data
            for col in range(2, self.table.columnCount() - 1):
                self.put(row, col, "")

            # Bersihkan kolom Judge
            self.put(row, self.table.columnCount() - 1, "")


    def process_scan(self):
        if not self.running:
            return

        code = self.scan_input.text().strip()
        self.scan_input.clear()

        if not code or not self.config.get("steps"):
            return

        steps = self.config["steps"]
        # Pastikan indeks step masih valid jika konfigurasi berubah
        self.current_step = min(self.current_step, len(steps) - 1)
        step = steps[self.current_step]
        remark_col = self.table.columnCount() - 1

        # STEP OPERATOR
        if step.get("input_type") == "operator":
            operator = self.config.get("operators", {}).get(code)

            if operator is None:
                self.put(self.current_step, remark_col, "NG")
                self.message.setText(
                    "ID operator tidak terdaftar. Silakan scan ulang."
                )
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
            self.advance_step()
            self.scan_input.setFocus()
            return

        # STEP BARCODE / INPUT DATA
        groups = step.get("groups", [])
        valid = True

        # Bersihkan data sebelumnya pada baris step aktif
        for col in range(2, remark_col):
            self.put(self.current_step, col, "")

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

            if group.get("validate", False):
                if value != group.get("expected", ""):
                    valid = False

        if not valid:
            self.put(self.current_step, remark_col, "NG")
            self.message.setText(
                "Input NG. Periksa data dan silakan scan ulang."
            )
            self.scan_input.setFocus()
            return

        self.put(self.current_step, remark_col, "OK")
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

        self.current_step = 0
        self.current_operator = ""
        self.cycle_started_at = None
        self.cycle_operator_code = ""
        self.cycle_operator_name = ""
        self.cycle_product_codes = {}
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
            # Build filename from two selectable parts plus a free-text part.
            filename_parts = []
            today_name = datetime.now().strftime("%d%m%Y")
            for part_number, part_key in enumerate(("filename_part1", "filename_part2"), start=1):
                part_type = target.get(part_key, "none")
                if part_type == "date":
                    filename_parts.append(today_name)
                elif part_type == "data":
                    source = target.get(f"filename_data{part_number}", "")
                    source_value = self.get_filename_data_value(source)
                    if source_value:
                        import re
                        safe_value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(source_value)).strip(" ._")
                        if safe_value:
                            filename_parts.append(safe_value)
            free_text = (target.get("filename_text") or "").strip()
            if not free_text:
                free_text = Path(target.get("filename", "hasil_pengujian.csv")).stem or "hasil_pengujian"
            if free_text:
                filename_parts.append(free_text)
            filename = "_".join(part for part in filename_parts if part) or "hasil_pengujian"
            filename += ".csv"

            # Legacy filename source is intentionally ignored; placement is configured above.
            source = target.get("filename_source", "")
            source_value = ""
            if source.startswith("barcode_step_"):
                try:
                    source_step = int(source.rsplit("_", 1)[1]) - 1
                    source_value = self.cycle_product_codes.get(source_step, "")
                except (ValueError, IndexError):
                    source_value = ""
            elif source.startswith("data_step_"):
                try:
                    _, _, step_part, _, group_part = source.split("_")
                    source_step = int(step_part) - 1
                    group_index = int(group_part) - 1
                    scanned_code = self.cycle_product_codes.get(source_step, "")
                    group = self.config.get("steps", [])[source_step].get("groups", [])[group_index]
                    if group.get("mode") == "operator_name":
                        source_value = self.cycle_operator_name
                    elif group.get("mode") == "all":
                        source_value = scanned_code
                    else:
                        start = max(0, int(group.get("start", 1)) - 1)
                        length = max(0, int(group.get("length", 1)))
                        source_value = scanned_code[start:start + length]
                except (ValueError, IndexError, KeyError):
                    source_value = ""
            # Legacy suffix/date-prefix fields are ignored to avoid duplicate parts.
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
                        QMessageBox.critical(
                            self, "Kolom CSV berbeda",
                            f"Susunan kolom konfigurasi berbeda dari header file:\n{output_path}\n"
                            "Gunakan nama file baru atau kembalikan susunan kolom sebelumnya."
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
        if dialog.exec_() == QDialog.Accepted:
            self.config = dialog.get_config()
            save_json(CONFIG_FILE, self.config)
            self.title_label.setText(self.config.get("display_title", "BARCODE DATA INPUT"))
            self.current_step = 0
            self.current_operator = ""
            self.cycle_started_at = None
            self.cycle_operator_code = ""
            self.cycle_operator_name = ""
            self.cycle_product_codes = {}
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
    sys.exit(app.exec_())
