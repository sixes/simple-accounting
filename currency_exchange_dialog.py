from PySide6.QtWidgets import QDialog, QVBoxLayout, QFormLayout, QComboBox, QLineEdit, QLabel, QPushButton, QHBoxLayout, QMessageBox, QDateEdit, QTableWidgetItem
from PySide6.QtCore import QDate, Qt
import uuid

class CurrencyExchangePLDialog(QDialog):
    def __init__(self, parent=None, all_sheets=None, from_sheet=None):
        super().__init__(parent)
        self.setWindowTitle("Add Currency Exchange")
        self.all_sheets = all_sheets or []
        self.from_sheet = from_sheet
        self.to_sheet = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)
        form.setFormAlignment(Qt.AlignHCenter | Qt.AlignTop)
        form.setHorizontalSpacing(20)
        form.setVerticalSpacing(12)

        # Filter bank-type sheets
        self.bank_sheets = [s for s in self.all_sheets if getattr(s, 'type', None) == 'bank']
        self.sheet_names = [getattr(s, 'name', '') for s in self.bank_sheets]

        # From section
        from_section = QHBoxLayout()
        from_label = QLabel("<b>From</b>")
        from_label.setFixedWidth(40)
        from_section.addWidget(from_label)
        self.from_date = QDateEdit(QDate.currentDate())
        self.from_date.setCalendarPopup(True)
        self.from_date.setDisplayFormat('yyyy/MM/dd')
        self.from_date.setFixedWidth(110)
        from_section.addWidget(self.from_date)
        self.from_bank_combo = QComboBox()
        self.from_bank_combo.addItems(self.sheet_names)
        self.from_bank_combo.setFixedWidth(160)
        if self.from_sheet and self.from_sheet.name in self.sheet_names:
            self.from_bank_combo.setCurrentText(self.from_sheet.name)
        from_section.addWidget(self.from_bank_combo)
        self.from_amount_label = QLabel("Amount")
        self.from_amount_label.setFixedWidth(50)
        from_section.addWidget(self.from_amount_label)
        self.from_amount = QLineEdit()
        self.from_amount.setPlaceholderText("Amount")
        self.from_amount.setFixedWidth(100)
        from_section.addWidget(self.from_amount)
        from_section.addStretch(1)
        form.addRow(from_section)

        # To section
        to_section = QHBoxLayout()
        to_label = QLabel("<b>To</b>")
        to_label.setFixedWidth(40)
        to_section.addWidget(to_label)
        self.to_date = QDateEdit()
        self.to_date.setReadOnly(True)
        self.to_date.setDisplayFormat('yyyy/MM/dd')
        self.to_date.setDate(self.from_date.date())
        self.to_date.setFixedWidth(110)
        to_section.addWidget(self.to_date)
        self.to_bank_combo = QComboBox()
        self.to_bank_combo.addItems(self.sheet_names)
        self.to_bank_combo.setFixedWidth(160)
        to_section.addWidget(self.to_bank_combo)
        self.to_amount_label = QLabel("Amount")
        self.to_amount_label.setFixedWidth(50)
        to_section.addWidget(self.to_amount_label)
        self.to_amount = QLineEdit()
        self.to_amount.setPlaceholderText("Amount")
        self.to_amount.setFixedWidth(100)
        to_section.addWidget(self.to_amount)
        to_section.addStretch(1)
        form.addRow(to_section)

        layout.addLayout(form)

        # Sync to_date with from_date
        self.from_date.dateChanged.connect(self.to_date.setDate)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch(1)
        self.add_btn = QPushButton("Add")
        self.add_btn.setDefault(True)
        self.cancel_btn = QPushButton("Cancel")
        btn_layout.addWidget(self.add_btn)
        btn_layout.addWidget(self.cancel_btn)
        btn_layout.addStretch(1)
        layout.addLayout(btn_layout)

        self.add_btn.clicked.connect(self.on_add)
        self.cancel_btn.clicked.connect(self.reject)

    def on_add(self):
        from_idx = self.from_bank_combo.currentIndex()
        to_idx = self.to_bank_combo.currentIndex()
        if from_idx == to_idx:
            QMessageBox.warning(self, "Invalid Selection", "From and To bank sheets must be different.")
            return
        try:
            from_amt = float(self.from_amount.text())
            to_amt = float(self.to_amount.text())
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Amounts must be valid numbers.")
            return
        from_sheet = self.bank_sheets[from_idx]
        to_sheet = self.bank_sheets[to_idx]
        date_str = self.from_date.date().toString('yyyy-MM-dd')
        unique_str = f"CurrencyEx-{uuid.uuid4().hex[:8]}"
        # Add row to from_sheet (debit)
        self.add_bank_row(from_sheet, date_str, from_amt, to_sheet.name, unique_str, is_debit=True)
        # Add row to to_sheet (credit)
        self.add_bank_row(to_sheet, date_str, to_amt, from_sheet.name, unique_str, is_debit=False)
        QMessageBox.information(self, "Success", "Currency exchange rows added.")
        self.accept()

    def add_bank_row(self, sheet, date, amount, other_sheet_name, unique_str, is_debit=True):
        # Find columns
        headers = [sheet.horizontalHeaderItem(j).text() for j in range(sheet.columnCount())]
        idx_date = next((i for i, h in enumerate(headers) if '日期' in h), None)
        idx_debit = next((i for i, h in enumerate(headers) if '借方' in h), None)
        idx_credit = next((i for i, h in enumerate(headers) if '貸方' in h or '贷方' in h), None)
        idx_duifang = next((i for i, h in enumerate(headers) if '对方科目' in h or '對方科目' in h), None)
        idx_zike = next((i for i, h in enumerate(headers) if '子科目' in h), None)
        idx_zhaiyao = next((i for i, h in enumerate(headers) if '摘要' in h), None)
        # Find the first empty data row (all key columns empty), else append before pinned rows
        def is_empty_row(r):
            cols = [idx_date, idx_duifang, idx_zike, idx_debit, idx_credit]
            for c in cols:
                if c is not None:
                    item = sheet.item(r, c)
                    if item and item.text().strip():
                        return False
            return True
        # Exclude pinned rows (last 2 rows)
        data_row_count = sheet.rowCount() - 2 if sheet.rowCount() > 2 else sheet.rowCount()
        row = data_row_count  # Default to append before pinned rows
        for r in range(data_row_count):
            if is_empty_row(r):
                row = r
                break
        sheet.insertRow(row)
        # Format date as yyyy/MM/dd (e.g., 2025/08/23)
        from datetime import datetime
        try:
            dt = datetime.strptime(date, "%Y-%m-%d")
            date_str_fmt = dt.strftime("%Y/%m/%d")
        except Exception:
            date_str_fmt = date.replace("-", "/")  # fallback
        if idx_date is not None:
            sheet.setItem(row, idx_date, QLineEditItem(date_str_fmt))
        if is_debit and idx_debit is not None:
            sheet.setItem(row, idx_debit, QLineEditItem(f"{amount:.2f}"))
        if not is_debit and idx_credit is not None:
            sheet.setItem(row, idx_credit, QLineEditItem(f"{amount:.2f}"))
        if idx_duifang is not None:
            sheet.setItem(row, idx_duifang, QLineEditItem(other_sheet_name))
        if idx_zike is not None:
            sheet.setItem(row, idx_zike, QLineEditItem("中转"))
        if idx_zhaiyao is not None:
            sheet.setItem(row, idx_zhaiyao, QLineEditItem(unique_str))
        sheet.viewport().update()

class QLineEditItem(QTableWidgetItem):
    def __init__(self, text):
        super().__init__(text)
        self.setFlags(self.flags() | Qt.ItemIsEditable)
