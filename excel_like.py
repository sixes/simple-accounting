import pickle
from PySide6.QtWidgets import (
    QMainWindow, QTabWidget, QLineEdit, QLabel, QHBoxLayout, QVBoxLayout, QWidget, QFileDialog, QInputDialog, QComboBox, QDialog, QDialogButtonBox, QFormLayout, QDoubleSpinBox
)
from PySide6.QtGui import QAction
from excel_table import ExcelTable

class AddSheetDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add Sheet")
        self.layout = QFormLayout(self)
        self.type_combo = QComboBox()
        self.type_combo.addItems(["Bank Sheet", "商業登記證書", "秘書費", "工資", "審計費"])
        self.layout.addRow("Sheet Type:", self.type_combo)
        self.bank_name = QLineEdit()
        self.currency = QLineEdit()
        self.layout.addRow("Bank Name:", self.bank_name)
        self.layout.addRow("Currency:", self.currency)
        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.layout.addWidget(self.button_box)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.type_combo.currentIndexChanged.connect(self.update_fields)
        self.update_fields()

    def update_fields(self):
        is_bank = self.type_combo.currentText() == "Bank Sheet"
        self.bank_name.setEnabled(is_bank)
        self.currency.setEnabled(is_bank)

    def get_result(self):
        if self.type_combo.currentText() == "Bank Sheet":
            return f"{self.bank_name.text().strip()}-{self.currency.text().strip()}", "bank"
        else:
            return self.type_combo.currentText(), "other"

class ExcelLike(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("深圳好景商务做账软件")
        self.central = QWidget()
        self.setCentralWidget(self.central)
        self.layout = QVBoxLayout(self.central)

        # Top bar for company name and period
        self.top_bar = QHBoxLayout()
        self.company_label = QLabel("Company Name:")
        self.company_input = QLineEdit()
        self.company_input.setText("cmpName")
        self.company_input.setPlaceholderText("Enter company name...")
        self.period_label = QLabel("Period:")
        self.period_input = QLineEdit()
        self.period_input.setPlaceholderText("Enter period...")
        self.top_bar.addWidget(self.company_label)
        self.top_bar.addWidget(self.company_input)
        self.top_bar.addWidget(self.period_label)
        self.top_bar.addWidget(self.period_input)
        self.top_bar.addStretch()
        self.layout.addLayout(self.top_bar)

        self.tabs = QTabWidget()
        self.layout.addWidget(self.tabs)
        self.sheets = []
        self.company_input.textChanged.connect(self.auto_save)
        self.period_input.textChanged.connect(self.auto_save)
        # Always add a default bank sheet named HSBC-USD
        self.add_sheet("HSBC-USD", is_bank=True)

        self.menu = self.menuBar()
        file_menu = self.menu.addMenu("File")
        new_file = QAction("New", self)
        add_sheet = QAction("Add Sheet", self)
        del_sheet = QAction("Delete Sheet", self)
        save = QAction("Save", self)
        load = QAction("Load", self)
        file_menu.addAction(new_file)
        file_menu.addSeparator()
        file_menu.addAction(add_sheet)
        file_menu.addAction(del_sheet)
        file_menu.addSeparator()
        file_menu.addAction(save)
        file_menu.addAction(load)

        new_file.triggered.connect(self.new_file)
        add_sheet.triggered.connect(self.add_sheet_dialog)
        del_sheet.triggered.connect(self.delete_sheet)
        save.triggered.connect(self.save_file)
        load.triggered.connect(self.load_file)

        self.tabs.setTabsClosable(True)
        self.tabs.tabCloseRequested.connect(self.close_tab)

    def add_sheet(self, name=None, is_bank=False):
        if not name:
            name, ok = QInputDialog.getText(self, "Sheet Name", "Enter sheet name:")
            if not ok or not name:
                return
        if is_bank or (name and ("-" in name)):
            columns = ["序 號", "日  期", "摘   要", "借     方", "貸     方", "餘    額", "發票號碼"]
            table = ExcelTable(auto_save_callback=self.auto_save)
            table.setColumnCount(len(columns))
            table.setHorizontalHeaderLabels(columns)
            self.tabs.addTab(table, name)
            self.tabs.setCurrentWidget(table)
            self.sheets.append(table)
            # --- Add exchange rate input ---
            rate_input = QDoubleSpinBox()
            rate_input.setPrefix("Exchange Rate: ")
            rate_input.setValue(1.0)
            rate_input.setDecimals(4)
            rate_input.valueChanged.connect(lambda v, t=table: t.set_exchange_rate(v))
            self.layout.addWidget(rate_input)
            table.exchange_rate_input = rate_input
            table.set_exchange_rate(1.0)
        else:
            columns = ["序 號", "日  期", "摘   要", "借     方", "貸     方", "借或貸", "餘    額", "發票號碼"]
            table = ExcelTable(auto_save_callback=self.auto_save)
            table.setColumnCount(len(columns))
            table.setHorizontalHeaderLabels(columns)
            self.tabs.addTab(table, name)
            self.tabs.setCurrentWidget(table)
            self.sheets.append(table)

    def add_sheet_dialog(self):
        dlg = AddSheetDialog(self)
        if dlg.exec() == QDialog.Accepted:
            name, sheet_type = dlg.get_result()
            if not name:
                return
            if sheet_type == "bank":
                columns = ["序 號", "日  期", "摘   要", "借     方", "貸     方", "餘    額", "發票號碼"]
                table = ExcelTable(auto_save_callback=self.auto_save)
                table.setColumnCount(len(columns))
                table.setHorizontalHeaderLabels(columns)
                self.tabs.addTab(table, name)
                self.tabs.setCurrentWidget(table)
                self.sheets.append(table)
                # --- Add exchange rate input ---
                rate_input = QDoubleSpinBox()
                rate_input.setPrefix("Exchange Rate: ")
                rate_input.setValue(1.0)
                rate_input.setDecimals(4)
                rate_input.valueChanged.connect(lambda v, t=table: t.set_exchange_rate(v))
                self.layout.addWidget(rate_input)
                table.exchange_rate_input = rate_input
                table.set_exchange_rate(1.0)
            else:
                columns = ["序 號", "日  期", "摘   要", "借     方", "貸     方", "借或貸", "餘    額", "發票號碼"]
                table = ExcelTable(auto_save_callback=self.auto_save)
                table.setColumnCount(len(columns))
                table.setHorizontalHeaderLabels(columns)
                self.tabs.addTab(table, name)
                self.tabs.setCurrentWidget(table)
                self.sheets.append(table)

    def delete_sheet(self):
        idx = self.tabs.currentIndex()
        if self.tabs.count() > 1:
            self.tabs.removeTab(idx)
            del self.sheets[idx]
            self.auto_save()

    def close_tab(self, idx):
        if self.tabs.count() > 1:
            self.tabs.removeTab(idx)
            del self.sheets[idx]
            self.auto_save()

    def new_file(self):
        self.tabs.clear()
        self.sheets = []
        self.company_input.setText("")
        self.period_input.setText("")
        # Always add a default bank sheet named HSBC-USD
        self.add_sheet("HSBC-USD", is_bank=True)
        self.auto_save()

    def save_file(self):
        fname = self.company_input.text().strip() or "untitled"
        path, _ = QFileDialog.getSaveFileName(self, "Save File", f"{fname}.exl", "ExcelLike (*.exl)")
        if not path:
            return
        self._save_to_path(path)

    def _save_to_path(self, path):
        data = {
            "company": self.company_input.text(),
            "period": self.period_input.text(),
            "sheets": []
        }
        for i in range(self.tabs.count()):
            sheet = self.tabs.widget(i)
            sheet_data = sheet.data()
            # Save custom headers if they exist
            if hasattr(sheet, '_custom_headers') and sheet._custom_headers:
                sheet_data["headers"] = sheet._custom_headers
            data["sheets"].append({"name": self.tabs.tabText(i), "data": sheet_data})
        with open(path, "wb") as f:
            pickle.dump(data, f)

    def load_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open File", "", "ExcelLike (*.exl)")
        if not path:
            return
        with open(path, "rb") as f:
            data = pickle.load(f)
        self.tabs.clear()
        self.sheets = []
        # Clear existing exchange rate inputs
        for i in reversed(range(self.layout.count())):
            item = self.layout.itemAt(i)
            if item and hasattr(item.widget(), 'setPrefix'):  # Check if it's an exchange rate input
                item.widget().deleteLater()
                self.layout.removeItem(item)
        self.company_input.setText(data.get("company", ""))
        self.period_input.setText(data.get("period", ""))
        for sheet in data.get("sheets", []):
            table = ExcelTable(auto_save_callback=self.auto_save)
            sheet_name = sheet["name"]
            # Check if it's a bank sheet (contains "-" in name)
            if "-" in sheet_name:
                columns = ["序 號", "日  期", "摘   要", "借     方", "貸     方", "餘    額", "發票號碼"]
                table.setColumnCount(len(columns))
                table.setHorizontalHeaderLabels(columns)
                # Add exchange rate input for bank sheets
                rate_input = QDoubleSpinBox()
                rate_input.setPrefix("Exchange Rate: ")
                rate_input.setValue(1.0)
                rate_input.setDecimals(4)
                rate_input.valueChanged.connect(lambda v, t=table: t.set_exchange_rate(v))
                self.layout.addWidget(rate_input)
                table.exchange_rate_input = rate_input
                table.set_exchange_rate(1.0)
            else:
                columns = ["序 號", "日  期", "摘   要", "借     方", "貸     方", "借或貸", "餘    額", "發票號碼"]
                table.setColumnCount(len(columns))
                table.setHorizontalHeaderLabels(columns)
            table.load_data(sheet["data"])
            self.tabs.addTab(table, sheet_name)
            self.sheets.append(table)

    def auto_save(self):
        fname = self.company_input.text().strip() or "untitled"
        path = f"{fname}.exl"
        try:
            # Only auto-save if we have sheets to save
            if self.tabs.count() > 0:
                self._save_to_path(path)
                print(f"Auto-saved to: {path}")
        except Exception as e:
            print(f"Auto-save failed: {e}")