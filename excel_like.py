import pickle
from PySide6.QtWidgets import (
    QMainWindow, QTabWidget, QLineEdit, QLabel, QHBoxLayout, QVBoxLayout,
    QWidget, QFileDialog, QInputDialog, QComboBox, QDialog, QDialogButtonBox,
    QFormLayout, QDoubleSpinBox, QMessageBox, QStatusBar, QTableWidgetItem,
    QDateEdit
)
from PySide6.QtGui import QAction
from PySide6.QtCore import Qt, QDate
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
        self.type_combo.addItems([
            "Bank Sheet",
            "銷售收入",
            "銷售成本",
            "銀行費用",
            "利息收入",
            "應付費用",
            "董事往來",
            "商業登記證",
            "秘書費",
            "工資",
            "審計費"
        ])
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
        sheet_type = self.type_combo.currentText()

        if sheet_type == "Bank Sheet":
            return f"{self.bank_name.text().strip()}-{self.currency.text().strip()}", "bank"
        elif sheet_type in ["銷售收入", "銷售成本", "銀行費用", "利息收入", "應付費用", "董事往來", "工資", "商業登記證書", "秘書費", "審計費"]:
            return sheet_type, "aggregate"
        else:
            return sheet_type, "other"

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
        self.company_input.setText("company_name")
        self.company_input.setPlaceholderText("Enter company name...")

        # Period date selectors
        self.period_from_label = QLabel("Period From:")
        self.period_from_input = QDateEdit()
        self.period_from_input.setDate(QDate.currentDate().addMonths(-1))  # Default to last month
        self.period_from_input.setDisplayFormat("yyyy/MM/dd")
        self.period_from_input.setCalendarPopup(True)

        self.period_to_label = QLabel("To:")
        self.period_to_input = QDateEdit()
        self.period_to_input.setDate(QDate.currentDate())  # Default to current date
        self.period_to_input.setDisplayFormat("yyyy/MM/dd")
        self.period_to_input.setCalendarPopup(True)

        self.top_bar.addWidget(self.company_label)
        self.top_bar.addWidget(self.company_input)
        self.top_bar.addWidget(self.period_from_label)
        self.top_bar.addWidget(self.period_from_input)
        self.top_bar.addWidget(self.period_to_label)
        self.top_bar.addWidget(self.period_to_input)
        self.top_bar.addStretch()
        self.layout.addLayout(self.top_bar)

        # Tab widget for sheets
        self.tabs = QTabWidget()
        self.tabs.setTabsClosable(True)
        self.tabs.setMovable(True)  # Add this line to enable tab dragging
        self.layout.addWidget(self.tabs)

        # Data storage
        self.sheets = []
        self.sales_sheet = None
        self.cost_sheet = None
        self.user_added_rows = None
        self.company_input.textChanged.connect(self.auto_save)
        self.period_from_input.dateChanged.connect(self.auto_save)
        self.period_to_input.dateChanged.connect(self.auto_save)

        # Add default sheet
        #self._create_bank_sheet("HSBC-USD")

        # Menu bar setup
        self._setup_menu_bar()

        # Connect tab signals
        self.tabs.tabCloseRequested.connect(self.close_tab)
        self.tabs.currentChanged.connect(self._on_tab_changed)  # Add this line
        self.tabs.tabBar().tabMoved.connect(self._on_tab_moved)  # Add this line

        # Try to auto-load the default company file
        self._auto_load_company_file()

    def _on_tab_moved(self, from_index, to_index):
        """Handle tab reordering to keep sheets list in sync"""
        self.sheets.insert(to_index, self.sheets.pop(from_index))
        self.auto_save()

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
                print("bank feeeeeeeeeeee")
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
            for row in range(current_tab.rowCount()):
                for col in range(current_tab.columnCount()):
                    current_tab.setItem(row, col, None)

            # Repopulate data
            data_rows = []
            amount_col_index = 5  # Column index for amount (debit/credit)

            # Collect all matching rows from bank sheets
            print(f"sheets: {self.sheets}")
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
                            #print(f"sheet: {sheet.name} amount: {amount_item.text()}")

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
                                'source_sheet': sheet,
                                'source_label': f"{sheet.name}:{row+1}"  # Add source label
                            })

            # Sort by date
            data_rows.sort(key=lambda x: x['date_obj'])
            print(f"data_rows: {data_rows}")
            bank_data_color = QColor(200, 255, 200)  # Light green for bank data
            for i, row_data in enumerate(data_rows):
                for col, value in [
                    (0, str(i+1)),  # 序號
                    (1, row_data['date_str']),
                    (2, row_data['subject']),
                    (3, row_data['desc']),
                    (4, row_data['invoice']),
                    (amount_col_index, row_data['amount']),
                    (6, ""),  # Balance column
                    (7, row_data['source_label'])  # Source column
                ]:
                    item = QTableWidgetItem(value)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # Make read-only
                    # Use setData to force background color
                    item.setData(Qt.BackgroundRole, bank_data_color)
                    current_tab.setItem(i, col, item)

            if hasattr(current_tab, 'user_added_rows') and user_data:
                #print(f"user add rows: {user_data}")
                start_row = len(data_rows)
                current_tab.user_added_rows = set()
                for i, row_data in enumerate(user_data):
                    row = start_row + i
                    current_tab.user_added_rows.add(row)
                    for col, text in enumerate(row_data):
                        item = QTableWidgetItem(text)
                        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
                        current_tab.setItem(row, col, item)

    def update_tab_name(self, old_name, new_name):
        """Update the tab text when sheet is renamed"""
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == old_name:
                self.tabs.setTabText(i, new_name)
                break

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
        currency = name.split("-")[1] if "-" in name else ""
        rate_input.setPrefix(f"{currency}:HKD = 1:")
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
        print(f"DEBUG CREATE: Creating regular sheet '{name}'")
        columns = ["序 號", "日  期", "對方科目", "摘   要", "借     方", "貸     方", "借或貸", "餘    額", "發票號碼"]
        table = ExcelTable(auto_save_callback=self.auto_save, name=name)
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)
        self.tabs.addTab(table, name)
        self.sheets.append(table)
        print(f"DEBUG CREATE: Regular sheet '{name}' created, total sheets: {len(self.sheets)}, total tabs: {self.tabs.count()}")
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
        print(f"DEBUG ADD: Starting add sheet dialog, current tabs: {self.tabs.count()}")
        dlg = AddSheetDialog(self)
        if dlg.exec() == QDialog.Accepted:
            name, sheet_type = dlg.get_result()
            print(f"DEBUG ADD: Dialog accepted, name='{name}', type='{sheet_type}'")
            if not name:
                print(f"DEBUG ADD: No name provided, aborting")
                return

            new_sheet = None
            if sheet_type == "bank":
                print(f"DEBUG ADD: Creating bank sheet: {name}")
                new_sheet = self._create_bank_sheet(name)
            elif sheet_type == "aggregate":
                print(f"DEBUG ADD: Creating aggregate sheet: {name}")
                # Create the appropriate aggregate sheet
                if name == "銷售收入":
                    new_sheet = self._create_sales_sheet()
                elif name == "銷售成本":
                    new_sheet = self._create_cost_sheet()
                elif name == "銀行費用":
                    new_sheet = self._create_bank_fee_sheet()
                elif name == "利息收入":
                    new_sheet = self._create_interest_sheet()
                elif name == "應付費用":
                    new_sheet = self._create_payable_sheet()
                elif name == "董事往來":
                    new_sheet = self._create_director_sheet()
                else:
                    # Create a regular sheet for other aggregate types like 工資, 商業登記證書, etc.
                    print(f"DEBUG ADD: Creating regular sheet for aggregate type: {name}")
                    new_sheet = self._create_regular_sheet(name)
            else:
                print(f"DEBUG ADD: Creating regular sheet: {name}")
                new_sheet = self._create_regular_sheet(name)

            print(f"DEBUG ADD: Sheet created successfully, total tabs now: {self.tabs.count()}")

            # Switch to the newly added sheet
            if new_sheet:
                self.tabs.setCurrentWidget(new_sheet)
                print(f"DEBUG ADD: Switched to new sheet")

            # Force auto-save after adding sheet
            print(f"DEBUG ADD: Triggering auto-save...")
            self.auto_save()
            print(f"DEBUG ADD: Add sheet complete")
        else:
            print(f"DEBUG ADD: Dialog cancelled")

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
        self.period_from_input.setDate(QDate.currentDate().addMonths(-1))
        self.period_to_input.setDate(QDate.currentDate())
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
        columns = ["序 號", "日  期", "對方科目", "摘  要", "發票號碼", column_title, "餘    額", "來源"]
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
                        #print(f"sheet: {sheet} amount: {amount_item}")
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
                            'source_sheet': sheet,
                            'source_label': f"{sheet.name}:{row+1}"  # Add source label
                        })

        # Sort by date
        data_rows.sort(key=lambda x: x['date_obj'])

        bank_data_color = QColor(200, 255, 200)  # Light green for bank data
        # Add to table
        for i, row_data in enumerate(data_rows):
            for col, value in [
                (0, str(i+1)),  # 序號
                (1, row_data['date_str']),
                (2, row_data['subject']),
                (3, row_data['desc']),
                (4, row_data['invoice']),
                (amount_col_index, row_data['amount']),
                (6, ""),  # Balance column
                (7, row_data['source_label'])  # Source column
            ]:
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # Make read-only
                # Use setData to force background color
                item.setData(Qt.BackgroundRole, bank_data_color)
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
            "period_from": self.period_from_input.date().toString("yyyy/MM/dd"),
            "period_to": self.period_to_input.date().toString("yyyy/MM/dd"),
            "sheets": [],
            "tab_order": [self.tabs.tabText(i) for i in range(self.tabs.count())]  # Add this line
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
                tab_text =  getattr(sheet, 'name', self.tabs.tabText(i))
                if tab_text in ["銷售收入", "銷售成本", "銀行費用", "利息收入", "應付費用"]:
                    sheet_data = {"cells": {}, "spans": []}
                # Get sheet type
                if "-" in tab_text:
                    sheet_type = "bank"
                elif tab_text in ["銷售收入", "銷售成本", "銀行費用", "利息收入", "應付費用", "董事往來", "工資", "商業登記證書", "秘書費", "審計費"]:
                    sheet_type = "aggregate"
                else:
                    sheet_type = "regular"

                #print(f"DEBUG SAVE: Sheet {i}: name='{tab_text}', type='{sheet_type}', cells={len(sheet_data.get('cells', {}))}")

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
                print(f"DEBUG SAVE: Successfully added sheet '{sheet_info}' to save data")
            except Exception as e:
                print(f"ERROR SAVE: Error saving sheet {self.tabs.tabText(i)}: {e}")
                import traceback
                traceback.print_exc()
                continue

        #print(f"DEBUG SAVE: Total sheets to save: {len(data['sheets'])}")

        try:
            #print(f"saving:{data}")
            with open(path, "wb") as f:
                pickle.dump(data, f)
            print(f"DEBUG SAVE: Successfully saved to {path}")
        except Exception as e:
            raise Exception(f"Failed to write file: {str(e)}")

    def _load_data_from_dict(self, data):
        """Common method to load data from a dictionary (used by both auto-load and manual load)"""
        print(f"DEBUG LOAD: Starting data load, found {len(data.get('sheets', []))} sheets")
        #print(f"loading: {data}")
        self.tabs.clear()
        self.user_added_rows = None
        self.sheets = []

        # Store sheets temporarily to reorder them
        temp_sheets = {}

        # Clear exchange rate inputs
        for i in reversed(range(self.layout.count())):
            item = self.layout.itemAt(i)
            if item and hasattr(item.widget(), 'setPrefix'):
                item.widget().deleteLater()
                self.layout.removeItem(item)

        # Set company name if it exists in data
        company_name = data.get("company", "").strip()
        self.company_input.setText(company_name)
        print(f"DEBUG LOAD: Set company name to: '{company_name}'")

        # Load period dates with backward compatibility
        if "period_from" in data and "period_to" in data:
            from_date = QDate.fromString(data.get("period_from", ""), "yyyy/MM/dd")
            to_date = QDate.fromString(data.get("period_to", ""), "yyyy/MM/dd")
            if from_date.isValid():
                self.period_from_input.setDate(from_date)
            if to_date.isValid():
                self.period_to_input.setDate(to_date)
            print(f"DEBUG LOAD: Set period to {from_date.toString()} - {to_date.toString()}")

        # First pass: create all sheets and store them in temp_sheets
        for sheet_info in data.get("sheets", []):
            try:
                sheet_type = sheet_info.get("type", "regular")
                sheet_name = sheet_info["name"]
                print(f"DEBUG LOAD: Creating sheet: name='{sheet_name}', type='{sheet_type}'")

                if sheet_type == "bank":
                    table = self._create_bank_sheet(sheet_name)
                elif sheet_type == "aggregate":
                    if sheet_name == "銷售收入":
                        table = self._create_sales_sheet()
                    elif sheet_name == "銷售成本":
                        table = self._create_cost_sheet()
                    elif sheet_name == "銀行費用":
                        table = self._create_bank_fee_sheet()
                    elif sheet_name == "利息收入":
                        table = self._create_interest_sheet()
                    elif sheet_name == "應付費用":
                        table = self._create_payable_sheet()
                    elif sheet_name == "董事往來":
                        table = self._create_director_sheet()
                    else:
                        table = self._create_regular_sheet(sheet_name)
                else:
                    table = self._create_regular_sheet(sheet_name)

                table.name = sheet_name
                temp_sheets[sheet_name] = table

            except Exception as e:
                print(f"ERROR LOAD: Error creating sheet {sheet_info.get('name', 'unknown')}: {e}")
                import traceback
                traceback.print_exc()
                raise e


        # Second pass: load data and add sheets in correct order
        tab_order = data.get("tab_order", [sheet["name"] for sheet in data.get("sheets", [])])
        print(f"tab_order: {tab_order}")
        for sheet_name in tab_order:
            if sheet_name in temp_sheets:
                sheet_info = next((s for s in data["sheets"] if s["name"] == sheet_name), None)
                if sheet_info:
                    table = temp_sheets[sheet_name]
                    try:
                        print(f"DEBUG LOAD: Loading data for sheet '{sheet_name}'")
                        table.load_data(sheet_info["data"])

                        if "user_added_rows" in sheet_info and sheet_info["user_added_rows"]:
                            table.user_added_rows = set(sheet_info["user_added_rows"])

                        if "exchange_rate" in sheet_info:
                            table.set_exchange_rate(sheet_info["exchange_rate"])
                            if hasattr(table, "exchange_rate_input"):
                                table.exchange_rate_input.setValue(sheet_info["exchange_rate"])

                        if "currency" in sheet_info:
                            table.currency = sheet_info["currency"]

                        self.tabs.addTab(table, sheet_name)
                        print(f"loading {sheet_info}")

                    except Exception as e:
                        print(f"ERROR LOAD: Error loading data for sheet {sheet_name}: {e}")
                        import traceback
                        traceback.print_exc()
                        raise e


        print("DEBUG LOAD: Data loading completed successfully")

    def load_file(self):
        """Load file from disk"""
        path, _ = QFileDialog.getOpenFileName(
            self, "Open File", "", "ExcelLike (*.exl)"
        )
        if not path:
            print("DEBUG LOAD: No file selected")
            return

        print(f"DEBUG LOAD: Attempting to load file {path}")
        try:
            with open(path, "rb") as f:
                data = pickle.load(f)
            self._load_data_from_dict(data)
        except Exception as e:
            print(f"ERROR LOAD: Failed to load file: {str(e)}")
            QMessageBox.warning(self, "Load Error", f"Failed to load file: {str(e)}")

    def _auto_load_company_file(self):
        """Try to automatically load the company file on startup"""
        company_name = self.company_input.text().strip()
        print(f"DEBUG AUTO_LOAD: Starting auto-load, company_name: '{company_name}'")

        if company_name:
            file_path = f"{company_name}.exl"
            print(f"DEBUG AUTO_LOAD: Looking for file: {file_path}")

            try:
                import os
                if os.path.exists(file_path):
                    print(f"DEBUG AUTO_LOAD: File exists, loading...")
                    with open(file_path, "rb") as f:
                        data = pickle.load(f)
                    self._load_data_from_dict(data)
                    print(f"DEBUG AUTO_LOAD: Auto-loaded company file: {file_path}")
                else:
                    self.new_file()
            except Exception as e:
                print(f"ERROR AUTO_LOAD: Failed to auto-load company file: {e}")
                # If loading fails, keep the default sheet that was already created
        else:
            print(f"DEBUG AUTO_LOAD: No company name, keeping default sheet")


    def auto_save(self):
        """Auto-save current state"""
        fname = self.company_input.text().strip() or "untitled"
        path = f"{fname}.exl"
        #print(f"DEBUG AUTO_SAVE: Starting auto-save, tabs count: {self.tabs.count()}")
        try:
            if self.tabs.count() > 0:
                self._save_to_path(path)
                #print(f"DEBUG AUTO_SAVE: Auto-save completed successfully")
            else:
                print(f"DEBUG AUTO_SAVE: No tabs to save")
        except Exception as e:
            print(f"DEBUG AUTO_SAVE: Auto-save failed: {e}")
            import traceback
            traceback.print_exc()
