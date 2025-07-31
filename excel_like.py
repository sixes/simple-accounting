import pickle
from PySide6.QtWidgets import (
    QMainWindow, QTabWidget, QLineEdit, QLabel, QHBoxLayout, QVBoxLayout,
    QWidget, QFileDialog, QInputDialog, QComboBox, QDialog, QDialogButtonBox,
    QFormLayout, QDoubleSpinBox, QMessageBox, QStatusBar, QTableWidgetItem
)
from PySide6.QtGui import QAction
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from excel_table import ExcelTable
from datetime import datetime  # Add this import

class AddSheetDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add Sheet")
        self.layout = QFormLayout(self)

        # Sheet type selection
        self.type_combo = QComboBox()
        self.type_combo.addItems(["Bank Sheet", "商業登記證書", "秘書費", "工資", "審計費"])
        self.layout.addRow("Sheet Type:", self.type_combo)

        # Bank-specific fields
        self.bank_name = QLineEdit()
        self.currency = QLineEdit()
        self.layout.addRow("Bank Name:", self.bank_name)
        self.layout.addRow("Currency:", self.currency)

        # Dialog buttons
        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.layout.addWidget(self.button_box)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        # Connect signals
        self.type_combo.currentIndexChanged.connect(self.update_fields)
        self.update_fields()


    def update_fields(self):
        """Show/hide bank-specific fields based on sheet type"""
        is_bank = self.type_combo.currentText() == "Bank Sheet"
        self.bank_name.setEnabled(is_bank)
        self.currency.setEnabled(is_bank)

    def get_result(self):
        """Returns (sheet_name, sheet_type) tuple"""
        if self.type_combo.currentText() == "Bank Sheet":
            return f"{self.bank_name.text().strip()}-{self.currency.text().strip()}", "bank"
        return self.type_combo.currentText(), "other"

class ExcelLike(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("bankNote")
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

        # Tab widget for sheets
        self.tabs = QTabWidget()
        self.tabs.setTabsClosable(True)
        self.layout.addWidget(self.tabs)

        # Data storage
        self.sheets = []
        self.sales_sheet = None
        self.cost_sheet = None
        self.user_added_rows = None
        self.company_input.textChanged.connect(self.auto_save)
        self.period_input.textChanged.connect(self.auto_save)

        # Add default sheet
        self._create_bank_sheet("HSBC-USD")

        # Menu bar setup
        self._setup_menu_bar()

        # Connect tab signals
        self.tabs.tabCloseRequested.connect(self.close_tab)
        self.tabs.currentChanged.connect(self._on_tab_changed)  # Add this line

    def _on_tab_changed(self, index):
        if index >= 0:
            for sheet in self.sheets:
                if hasattr(sheet, 'exchange_rate_input'):
                    sheet.exchange_rate_input.setVisible(False)

            # Show exchange rate control for current bank sheet if it exists
            current_tab = self.tabs.widget(index)
            if hasattr(current_tab, 'exchange_rate_input'):
                current_tab.exchange_rate_input.setVisible(True)

            tab_name = self.tabs.tabText(index)
            if tab_name == "銷售收入":
                self._refresh_aggregate_sheet("销售收入", "借     方")
            elif tab_name == "銷售成本":
                self._refresh_aggregate_sheet("销售成本", "貸     方")
            elif tab_name == "銀行費用":
                self._refresh_aggregate_sheet("银行费用", "貸     方")
            elif tab_name == "利息收入":
                self._refresh_aggregate_sheet("利息收入", "借     方")
            elif tab_name == "應付費用":
                self._refresh_aggregate_sheet("董事往来", "貸     方")
            elif tab_name == "董事往來":
                current_tab.user_added_rows = getattr(current_tab, 'user_added_rows', set())
                self._refresh_aggregate_sheet("董事往来", "貸     方")

    def _refresh_aggregate_sheet(self, subject_filter, column_title):
        """Refresh the specified aggregate sheet"""
        current_tab = self.tabs.currentWidget()
        if current_tab:
        # First preserve any user-added rows for director sheet
            user_data = []
            if hasattr(current_tab, 'user_added_rows') and current_tab.user_added_rows:
                for row in current_tab.user_added_rows:
                    row_data = []
                    for col in range(current_tab.columnCount()):
                        item = current_tab.item(row, col)
                        row_data.append(item.text() if item else "")
                    user_data.append(row_data)

            # Clear existing data
            for row in range(current_tab.rowCount() - 2):
                for col in range(current_tab.columnCount()):
                    current_tab.setItem(row, col, None)

            # Repopulate data
            data_rows = []
            amount_col_index = 5  # Column index for amount (debit/credit)

            # Collect all matching rows from bank sheets
            for sheet in self.sheets:
                if "-" in sheet.name:  # Bank sheet
                    print(f"looking at sheet: {sheet.name} subject_filter: {subject_filter} column_title: {column_title}")
                    for row in range(sheet.rowCount() - 2):  # Exclude pinned rows
                        subject_item = sheet.item(row, 2)  # 對方科目 column
                        if subject_item and subject_filter in subject_item.text():
                            date_item = sheet.item(row, 1)  # 日期
                            desc_item = sheet.item(row, 3)  # 摘要
                            invoice_item = sheet.item(row, 7)  # 發票號碼
                            amount_item = sheet.item(row, 4 if "借" in column_title else 5)  # Debit/Credit column
                            print(f"sheet: {sheet.name} amount: {amount_item.text()}")

                            date_str = date_item.text() if date_item else ""
                            try:
                                date_obj = datetime.strptime(date_str, "%Y/%m/%d").date()
                            except (ValueError, AttributeError):
                                date_obj = datetime.min.date()

                            data_rows.append({
                                'date_str': date_str,
                                'date_obj': date_obj,
                                'subject': subject_item.text(),
                                'desc': desc_item.text() if desc_item else "",
                                'invoice': invoice_item.text() if invoice_item else "",
                                'amount': amount_item.text() if amount_item else "",
                                'source_row': row,
                                'source_sheet': sheet
                            })

            # Sort by date
            data_rows.sort(key=lambda x: x['date_obj'])

            for i, row_data in enumerate(data_rows):
                for col, value in [
                    (0, str(i+1)),  # 序號
                    (1, row_data['date_str']),
                    (2, row_data['subject']),
                    (3, row_data['desc']),
                    (4, row_data['invoice']),
                    (amount_col_index, row_data['amount']),
                    (6, "")  # Balance column
                ]:
                    item = QTableWidgetItem(value)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # Make read-only
                    item.setBackground(QColor(200, 255, 200))
                    current_tab.setItem(i, col, item)

            if hasattr(current_tab, 'user_added_rows') and user_data:
                start_row = len(data_rows)
                current_tab.user_added_rows = set()
                for i, row_data in enumerate(user_data):
                    row = start_row + i
                    current_tab.user_added_rows.add(row)
                    for col, text in enumerate(row_data):
                        item = QTableWidgetItem(text)
                        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
                        current_tab.setItem(row, col, item)

    def _setup_menu_bar(self):
        """Initialize the menu bar with actions"""
        self.menu = self.menuBar()
        file_menu = self.menu.addMenu("File")

        # File actions
        actions = [
            ("New", self.new_file),
            ("Add Sheet", self.add_sheet_dialog),
            ("Delete Sheet", self.delete_sheet),
            ("Save", self.save_file),
            ("Load", self.load_file)
        ]

        for text, callback in actions:
            action = QAction(text, self)
            action.triggered.connect(callback)
            file_menu.addAction(action)

            if text == "Delete Sheet":
                file_menu.addSeparator()

    def _create_bank_sheet(self, name):
        """Create a bank sheet with exchange rate control"""
        columns = ["序 號", "日  期", "對方科目", "摘   要", "借     方", "貸     方", "餘    額", "發票號碼"]
        table = ExcelTable(auto_save_callback=self.auto_save, name=name)
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)

        # Add exchange rate control
        rate_input = QDoubleSpinBox()
        rate_input.setPrefix("Exchange Rate: ")
        rate_input.setValue(1.0)
        rate_input.setDecimals(2)
        rate_input.valueChanged.connect(lambda v, t=table: t.set_exchange_rate(v))
        rate_input.setVisible(False)  # Start hidden
        self.layout.addWidget(rate_input)
        table.exchange_rate_input = rate_input
        table.set_exchange_rate(1.0)

        self.tabs.addTab(table, name)
        self.sheets.append(table)
        return table

    def _create_regular_sheet(self, name):
        """Create a regular sheet"""
        columns = ["序 號", "日  期", "對方科目", "摘   要", "借     方", "貸     方", "借或貸", "餘    額", "發票號碼"]
        table = ExcelTable(auto_save_callback=self.auto_save, name=name)
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)
        self.tabs.addTab(table, name)
        self.sheets.append(table)
        return table

    def add_sheet(self, name=None, is_bank=False):
        """Add a new sheet with optional name and type"""
        if not name:
            name, ok = QInputDialog.getText(self, "Sheet Name", "Enter sheet name:")
            if not ok or not name:
                return

        if is_bank or ("-" in name):
            self._create_bank_sheet(name)
        else:
            self._create_regular_sheet(name)

    def add_sheet_dialog(self):
        """Show dialog to add a new sheet"""
        dlg = AddSheetDialog(self)
        if dlg.exec() == QDialog.Accepted:
            name, sheet_type = dlg.get_result()
            if not name:
                return
            if sheet_type == "bank":
                self._create_bank_sheet(name)
            else:
                self._create_regular_sheet(name)

    def delete_sheet(self):
        """Delete the current sheet"""
        idx = self.tabs.currentIndex()
        if self.tabs.count() > 1:
            self.tabs.removeTab(idx)
            del self.sheets[idx]
            self.auto_save()

    def close_tab(self, idx):
        """Close tab at given index"""
        if self.tabs.count() > 1:
            self.tabs.removeTab(idx)
            del self.sheets[idx]
            self.auto_save()

    def new_file(self):
        """Create a new file with default sheet"""
        self.tabs.clear()
        self.sheets = []
        self.company_input.setText("")
        self.period_input.setText("")
        self._create_bank_sheet("HSBC-USD")
        self.sales_sheet = self._create_sales_sheet()
        self.cost_sheet = self._create_cost_sheet()
        self._create_bank_fee_sheet()
        self._create_interest_sheet()
        self._create_payable_sheet()
        self._create_director_sheet()
        self.auto_save()

    def _create_director_sheet(self):
        """Create a director sheet"""
        return self._create_aggregate_sheet(
            sheet_name="董事往來",
            subject_filter="董事往来",
            column_title="貸     方"
        )

    def _create_payable_sheet(self):
        """Create a payable sheet"""
        return self._create_aggregate_sheet(
            sheet_name="應付費用",
            subject_filter="董事往来",
            column_title="貸     方"
        )

    def _create_interest_sheet(self):
        """Create an interest income sheet"""
        return self._create_aggregate_sheet(
            sheet_name="利息收入",
            subject_filter="利息收入",
            column_title="貸     方"
        )

    def _create_sales_sheet(self):
        return self._create_aggregate_sheet(
            sheet_name="銷售收入",
            subject_filter="销售收入",
            column_title="貸     方"
        )

    def _create_cost_sheet(self):
        return self._create_aggregate_sheet(
            sheet_name="銷售成本",
            subject_filter="销售成本",
            column_title="借     方"
        )

    def _create_bank_fee_sheet(self):
        """Create a bank fee sheet"""
        return self._create_aggregate_sheet(
            sheet_name="銀行費用",
            subject_filter="银行费用",
            column_title="借     方"
        )


    def _create_aggregate_sheet(self, sheet_name, subject_filter, column_title):
        """Common method to create aggregate sheets (sales/cost)"""
        columns = ["序 號", "日  期", "對方科目", "摘  要", "發票號碼", column_title, "餘    額"]
        table = ExcelTable(auto_save_callback=self.auto_save, name=sheet_name)
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)

        # Add pinned rows for totals
        table.setRowCount(100 + 2)  # Regular rows + 2 pinned rows
        if sheet_name == "董事往來":
            table.user_added_rows = set()

        # Populate data from bank sheets
        self._populate_aggregate_data(table, subject_filter, column_title)

        self.tabs.addTab(table, sheet_name)
        self.sheets.append(table)
        return table

    def _populate_aggregate_data(self, table, subject_filter, amount_column_title):
        user_data = []
        if hasattr(table, 'user_added_rows'):
            for row in table.user_added_rows:
                row_data = []
                for col in range(table.columnCount()):
                    item = table.item(row, col)
                    row_data.append(item.text() if item else "")
                user_data.append(row_data)

        """Common method to populate aggregate sheets"""
        for row in range(table.rowCount() - 2):
            for col in range(table.columnCount()):
                table.setItem(row, col, None)

        data_rows = []
        amount_col_index = 5  # Column index for amount (debit/credit)

        # Collect all matching rows from bank sheets
        for sheet in self.sheets:
            if "-" in sheet.name:  # Bank sheet
                for row in range(sheet.rowCount() - 2):  # Exclude pinned rows
                    subject_item = sheet.item(row, 2)  # 對方科目 column
                    if subject_item and subject_filter in subject_item.text():
                        date_item = sheet.item(row, 1)  # 日期
                        desc_item = sheet.item(row, 3)  # 摘要
                        invoice_item = sheet.item(row, 7)  # 發票號碼
                        amount_item = sheet.item(row, 4 if "借" in amount_column_title else 5)  # Debit/Credit column
                        print(f"sheet: {sheet} amount: {amount_item}")
                        date_str = date_item.text() if date_item else ""
                        try:
                            date_obj = datetime.strptime(date_str, "%Y/%m/%d").date()
                        except (ValueError, AttributeError):
                            date_obj = datetime.min.date()

                        data_rows.append({
                            'date_str': date_str,
                            'date_obj': date_obj,
                            'subject': subject_item.text(),
                            'desc': desc_item.text() if desc_item else "",
                            'invoice': invoice_item.text() if invoice_item else "",
                            'amount': amount_item.text() if amount_item else "",
                            'source_row': row,
                            'source_sheet': sheet
                        })

        # Sort by date
        data_rows.sort(key=lambda x: x['date_obj'])

        bank_data_color = QColor(200, 255, 200)
        # Add to table
        for i, row_data in enumerate(data_rows):
            for col, value in [
                (0, str(i+1)),  # 序號
                (1, row_data['date_str']),
                (2, row_data['subject']),
                (3, row_data['desc']),
                (4, row_data['invoice']),
                (amount_col_index, row_data['amount']),
                (6, "")  # Balance column
            ]:
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # Make read-only
                item.setBackground(bank_data_color)  # Set background color
                table.setItem(i, col, item)


        if hasattr(table, 'user_added_rows'):
            start_row = len(data_rows)
            table.user_added_rows = set()
            for i, row_data in enumerate(user_data):
                row = start_row + i
                table.user_added_rows.add(row)
                for col, text in enumerate(row_data):
                    item = QTableWidgetItem(text)
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
                    table.setItem(row, col, item)

    def save_file(self):
        """Save current file"""
        fname = self.company_input.text().strip() or "untitled"
        path, _ = QFileDialog.getSaveFileName(
            self, "Save File", f"{fname}.exl", "ExcelLike (*.exl)"
        )
        if path:
            try:
                self._save_to_path(path)
            except Exception as e:
                QMessageBox.warning(self, "Save Error", f"Failed to save file: {str(e)}")

    def _save_to_path(self, path):
        """Save data to specified path"""
        data = {
            "version": "1.0",
            "company": self.company_input.text(),
            "period": self.period_input.text(),
            "sheets": []
        }

        # Save all sheets
        for i in range(self.tabs.count()):
            sheet = self.tabs.widget(i)
            try:
                sheet_data = sheet.data()
                if hasattr(sheet, '_custom_headers'):
                    sheet_data["headers"] = sheet._custom_headers

                # Get exchange rate if available
                exchange_rate = getattr(sheet, "exchange_rate", 1.0)

                # Get sheet type
                tab_text = self.tabs.tabText(i)
                sheet_type = "bank" if "-" in tab_text else "regular"

                # Save user_added_rows if it exists
                user_added_data = None
                if hasattr(sheet, 'user_added_rows'):
                    user_added_data = list(sheet.user_added_rows) if sheet.user_added_rows else None

                sheet_info = {
                    "name": tab_text,
                    "type": sheet_type,
                    "data": sheet_data,
                    "exchange_rate": exchange_rate,
                    "currency": getattr(sheet, "currency", ""),
                    "user_added_rows": user_added_data
                }

                data["sheets"].append(sheet_info)
            except Exception as e:
                print(f"Error saving sheet {self.tabs.tabText(i)}: {e}")
                continue

        try:
            with open(path, "wb") as f:
                pickle.dump(data, f)
        except Exception as e:
            raise Exception(f"Failed to write file: {str(e)}")

    def load_file(self):
        """Load file from disk"""
        path, _ = QFileDialog.getOpenFileName(
            self, "Open File", "", "ExcelLike (*.exl)"
        )
        if not path:
            return

        try:
            with open(path, "rb") as f:
                data = pickle.load(f)
        except Exception as e:
            QMessageBox.warning(self, "Load Error", f"Failed to load file: {str(e)}")
            return

        self.tabs.clear()
        self.user_added_rows = None
        self.sheets = []

        # Clear exchange rate inputs
        for i in reversed(range(self.layout.count())):
            item = self.layout.itemAt(i)
            if item and hasattr(item.widget(), 'setPrefix'):
                item.widget().deleteLater()
                self.layout.removeItem(item)

        self.company_input.setText(data.get("company", ""))
        self.period_input.setText(data.get("period", ""))

        # Load all sheets
        for sheet_info in data.get("sheets", []):
            try:
                if sheet_info["type"] == "bank":
                    table = self._create_bank_sheet(sheet_info["name"])
                    # Set currency if available
                    if "currency" in sheet_info:
                        table.currency = sheet_info["currency"]
                else:
                    table = self._create_regular_sheet(sheet_info["name"])

                table.load_data(sheet_info["data"])

                # Restore user_added_rows if it exists
                if "user_added_rows" in sheet_info and sheet_info["user_added_rows"]:
                    table.user_added_rows = set(sheet_info["user_added_rows"])

                if "exchange_rate" in sheet_info:
                    table.set_exchange_rate(sheet_info["exchange_rate"])
                    if hasattr(table, "exchange_rate_input"):
                        table.exchange_rate_input.setValue(sheet_info["exchange_rate"])
            except Exception as e:
                print(f"Error loading sheet {sheet_info.get('name', 'unknown')}: {e}")
                continue

    def auto_save(self):
        """Auto-save current state"""
        fname = self.company_input.text().strip() or "untitled"
        path = f"{fname}.exl"
        try:
            if self.tabs.count() > 0:
                self._save_to_path(path)
        except Exception as e:
            print(f"Auto-save failed: {e}")
