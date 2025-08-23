from PySide6.QtWidgets import (
    QMainWindow, QTabWidget, QLineEdit, QLabel, QHBoxLayout, QVBoxLayout,
    QWidget, QInputDialog, QDateEdit, QDialog, QMenu, QMessageBox, QDoubleSpinBox,
    QToolButton, QTabBar, QApplication, QPushButton, QTableWidgetItem
)
from PySide6.QtGui import QAction, QPalette
from PySide6.QtCore import Qt, QDate, qInstallMessageHandler
from dialogs import AddSheetDialog
from sheet_manager import SheetManager
from file_manager import FileManager
import platform
import time

def qt_message_handler(mode, context, message):
    if "single cell span won't be added" in message:
        return  # Ignore this specific warning
    # Handle other messages as needed

qInstallMessageHandler(qt_message_handler)

class ExcelLike(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("bankNote")

        # Force light theme for better readability
        if platform.system() == "Darwin":
            from datetime import datetime
            if datetime.now().hour >= 19:
                self.set_light_theme()

        self.central = QWidget()
        self.setCentralWidget(self.central)
        self.layout = QVBoxLayout(self.central)

        # Initialize managers
        self.sheet_manager = SheetManager(self)
        self.file_manager = FileManager(self)

        # Top bar for company name and period
        self.setup_top_bar()

        # Tab widget for sheets
        self.tabs = QTabWidget()
        self.tabs.setTabsClosable(True)
        self.tabs.setMovable(True)
        self.layout.addWidget(self.tabs)
        # Create initial sheets
        self.sheets = []
        self.sales_sheet = None
        self.cost_sheet = None
        self.user_added_rows = None

        self._add_plus_tab()
        self.tabs.currentChanged.connect(self._on_tab_or_plus_clicked)

        # Connect signals for auto-save
        self.company_input.editingFinished.connect(self.auto_save)
        self.period_from_input.dateChanged.connect(self.auto_save)
        self.period_to_input.dateChanged.connect(self.auto_save)
        self.update_button.clicked.connect(self.on_update_clicked)  # Connect to a handler method

        # Connect tab signals
        self.tabs.tabCloseRequested.connect(self.close_tab)
        self.tabs.tabBar().tabMoved.connect(self._on_tab_moved)

        # Menu bar setup
        self.setup_menu_bar()

        # Try to auto-load the default company file
        self.file_manager.auto_load_company_file()
        # Ensure company name is set in the UI after auto-load
        if hasattr(self.file_manager, 'last_loaded_company_name'):
            print(f"DEBUG UI: Setting company name to '{self.file_manager.last_loaded_company_name}' after auto-load")
            self.company_input.setText(self.file_manager.last_loaded_company_name)
        else:
            print("DEBUG UI: No company name loaded, keeping default")
        # Always ensure the plus tab is present after all startup logic
        self._add_plus_tab()

    def on_update_clicked(self):
        # 1. Collect data from all sheets
        bank_data = []
        non_bank_data = []
        non_bank_header = None
        # Find all sheets and collect data
        for i, sheet in enumerate(self.sheets):
            if getattr(sheet, 'type', None) == 'bank':
                headers = [sheet.horizontalHeaderItem(j).text() for j in range(sheet.columnCount())]
                idx_duifang = headers.index("对方科目") if "对方科目" in headers else -1
                idx_zike = headers.index("子科目") if "子科目" in headers else -1
                idx_debit = headers.index("借方") if "借方" in headers else -1
                idx_credit = headers.index("贷方") if "贷方" in headers else -1
                idx_balance = headers.index("余额") if "余额" in headers else -1
                for row in range(sheet.rowCount()):
                    key = None
                    if idx_duifang >= 0 and idx_zike >= 0:
                        duifang = sheet.item(row, idx_duifang).text() if sheet.item(row, idx_duifang) else ""
                        zike = sheet.item(row, idx_zike).text() if sheet.item(row, idx_zike) else ""
                        key = duifang
                        if zike != "":
                            key = duifang + "-" + zike
                    debit_text = sheet.item(row, idx_debit).text().strip().replace(",", "") if idx_debit >= 0 and sheet.item(row, idx_debit) and sheet.item(row, idx_debit).text() else ""
                    credit_text = sheet.item(row, idx_credit).text().strip().replace(",", "") if idx_credit >= 0 and sheet.item(row, idx_credit) and sheet.item(row, idx_credit).text() else ""
                    try:
                        debit_val = float(debit_text) if debit_text else 0
                    except ValueError:
                        debit_val = 0
                    try:
                        credit_val = float(credit_text) if credit_text else 0
                    except ValueError:
                        credit_val = 0
                    if (debit_val != 0 or credit_val != 0) and key:
                        row_dict = {}
                        for c, h in enumerate(headers):
                            if c == idx_balance:
                                continue
                            val = sheet.item(row, c).text() if sheet.item(row, c) else ""
                            row_dict[h] = val
                        bank_data.append({
                            "row_dict": row_dict,
                            "currency": getattr(sheet, "currency", ""),
                            "debit": debit_val,
                            "credit": credit_val,
                            "key": key,
                            "sheet_name": getattr(sheet, "name", ""),
                            "row_number": row + 1
                        })
            elif getattr(sheet, 'type', None) == 'non_bank':
                if not non_bank_header:
                    non_bank_header = [sheet.horizontalHeaderItem(j).text() for j in range(sheet.columnCount())]
                headers = [sheet.horizontalHeaderItem(j).text() for j in range(sheet.columnCount())]
                currency_cols = [(j, h) for j, h in enumerate(headers) if "借方(" in h or "贷方(" in h]
                idx_duifang = headers.index("借方科目") if "借方科目" in headers else -1
                idx_zike = headers.index("子科目") if "子科目" in headers else -1
                idx_daifang = headers.index("贷方科目") if "贷方科目" in headers else -1
                for row in range(sheet.rowCount()):
                    key = None
                    if idx_daifang >= 0 and sheet.item(row, idx_daifang) and sheet.item(row, idx_daifang).text():
                        daifang = sheet.item(row, idx_daifang).text()
                        zike = sheet.item(row, idx_zike).text() if idx_zike >= 0 and sheet.item(row, idx_zike) else ""
                        key = daifang
                        if zike != "":
                            key = daifang + "-" + zike
                    elif idx_duifang >= 0 and sheet.item(row, idx_duifang) and sheet.item(row, idx_duifang).text():
                        jiefang = sheet.item(row, idx_duifang).text()
                        zike = sheet.item(row, idx_zike).text() if idx_zike >= 0 and sheet.item(row, idx_zike) else ""
                        key = jiefang
                        if zike != "":
                            key = jiefang + "-" + zike
                    for col, h in currency_cols:
                        val = sheet.item(row, col).text().strip() if sheet.item(row, col) else ""
                        try:
                            fval = float(val.replace(",", "")) if val else 0
                        except Exception:
                            fval = 0
                        if fval != 0 and key:
                            if "(" in h and ")" in h:
                                currency = h.split("(")[1].split(")")[0]
                            else:
                                currency = ""
                            row_dict = {}
                            for c, hh in enumerate(headers):
                                row_dict[hh] = sheet.item(row, c).text() if sheet.item(row, c) else ""
                            non_bank_data.append({
                                "row_dict": row_dict,
                                "currency": currency,
                                "col": col,
                                "value": fval,
                                "key": key,
                                "sheet_name": getattr(sheet, "name", ""),
                                "row_number": row + 1
                            })
        # 2. For each key, create a payable detail sheet if not exists
        all_data = bank_data + non_bank_data
        keys = set(item["key"] for item in all_data)
        print(f"[DEBUG] Start payable detail update for {len(keys)} keys at", time.time())
        for key in keys:
            payable_sheet_name = key
            payable_sheet = None
            for s in self.sheets:
                if getattr(s, 'name', None) == payable_sheet_name:
                    payable_sheet = s
                    break
            if not payable_sheet:
                if non_bank_header:
                    print(f"[DEBUG] Creating payable detail sheet: {payable_sheet_name} at", time.time())
                    payable_sheet = self.sheet_manager.create_payable_detail_sheet(payable_sheet_name)
                else:
                    QMessageBox.warning(self, "Error", "No non-bank sheet found to create payable detail sheet header.")
                    continue
            else:
                print(f"[DEBUG] Erasing data in payable sheet: {payable_sheet_name} at", time.time())
                payable_sheet.clearContents()
                print(f"[DEBUG] Finished erasing data in {payable_sheet_name} at", time.time())
            print(f"[DEBUG] Building headers and mapping for {payable_sheet_name} at", time.time())
            headers = [payable_sheet.horizontalHeaderItem(j).text() for j in range(payable_sheet.columnCount())]
            mapping = {"debit": {}, "credit": {}}
            for idx, h in enumerate(headers):
                if "借方(" in h:
                    currency = h.split("(")[1].split(")")[0]
                    mapping["debit"][currency] = idx
                elif "贷方(" in h:
                    currency = h.split("(")[1].split(")")[0]
                    mapping["credit"][currency] = idx
            print(f"[DEBUG] Filtering and sorting data for {payable_sheet_name} at", time.time())
            filtered_bank_data = [item for item in bank_data if item["key"] == key]
            filtered_non_bank_data = [item for item in non_bank_data if item["key"] == key]
            all_rows = []
            for item in filtered_bank_data:
                row_dict = item["row_dict"]
                date_val = row_dict.get("日期", "")
                all_rows.append((date_val, 'bank', item))
            for item in filtered_non_bank_data:
                row_dict = item["row_dict"]
                date_val = row_dict.get("日期", "")
                all_rows.append((date_val, 'non_bank', item))
            def date_key(x):
                from datetime import datetime
                for fmt in ("%Y/%m/%d", "%m/%d/%y", "%Y-%m-%d", "%m-%d-%y"):
                    try:
                        return datetime.strptime(x[0], fmt)
                    except Exception:
                        continue
                return datetime(1900, 1, 1)
            all_rows.sort(key=date_key)
            print(f"[DEBUG] Finished sorting data for {payable_sheet_name} at", time.time())
            payable_sheet.setRowCount(max(100, len(all_rows) + 10))
            row_idx = 0
            print(f"[DEBUG] Writing data to payable sheet {payable_sheet_name} at", time.time())
            # Determine if special handling is needed for this sheet name
            is_creditor = payable_sheet_name in ["销售收入", "銷售收入", "利息收入"]
            is_debit = payable_sheet_name in ["销售成本", "銷售成本", "银行费用", "銀行費用"]
            for date_val, typ, item in all_rows:
                row_dict = item["row_dict"]
                currency = item["currency"]
                source_col_idx = None
                for idx, h in enumerate(headers):
                    if h == "来源":
                        source_col_idx = idx
                        break
                if typ == 'bank':
                    for h, v in row_dict.items():
                        if h in headers and not ("余额" in h):
                            col_idx = headers.index(h)
                            payable_sheet.setItem(row_idx, col_idx, QTableWidgetItem(v))
                    # Special handling for payable detail sheets
                    if is_creditor and currency in mapping["credit"]:
                        payable_sheet.setItem(row_idx, mapping["credit"][currency], QTableWidgetItem(str(item["debit"] + item["credit"])))
                    elif is_debit and currency in mapping["debit"]:
                        payable_sheet.setItem(row_idx, mapping["debit"][currency], QTableWidgetItem(str(item["debit"] + item["credit"])))
                    else:
                        if item["debit"] != 0 and currency in mapping["debit"]:
                            payable_sheet.setItem(row_idx, mapping["debit"][currency], QTableWidgetItem(str(item["debit"])))
                        elif item["credit"] != 0 and currency in mapping["credit"]:
                            payable_sheet.setItem(row_idx, mapping["credit"][currency], QTableWidgetItem(str(item["credit"])))
                    if source_col_idx is not None:
                        payable_sheet.setItem(row_idx, source_col_idx, QTableWidgetItem(f"{item.get('sheet_name', '')}:{item.get('row_number', '')}"))
                else:
                    for h, v in row_dict.items():
                        if h in headers and not ("余额" in h or "借方(" in h or "贷方(" in h):
                            col_idx = headers.index(h)
                            payable_sheet.setItem(row_idx, col_idx, QTableWidgetItem(v))
                row_idx += 1
            print(f"[DEBUG] Finished writing data to {payable_sheet_name} at", time.time())
        print(f"[DEBUG] Finished payable detail update at", time.time())
        self._add_plus_tab()

    def setup_top_bar(self):
        """Setup the top bar with company name, exchange rate, and period inputs"""
        self.top_bar = QHBoxLayout()
        self.company_label = QLabel("Company Name:")
        self.company_input = QLineEdit()
        self.company_input.setText("company_name")
        self.company_input.setPlaceholderText("Enter company name...")

        # Exchange rate input
        self.exchange_rate_label = QLabel("Exchange Rate:")
        self.exchange_rate_input = QDoubleSpinBox()
        self.exchange_rate_input.setDecimals(2)
        self.exchange_rate_input.setMinimum(0.01)
        self.exchange_rate_input.setMaximum(99)
        self.exchange_rate_input.setValue(1.0)
        self.exchange_rate_input.setSingleStep(0.01)
        self.exchange_rate_input.setEnabled(False)  # Disabled by default

        # Period date selectors
        self.period_from_label = QLabel("Period From:")
        self.period_from_input = QDateEdit()
        self.period_from_input.setDate(QDate.currentDate().addMonths(-1))
        self.period_from_input.setDisplayFormat("yyyy/MM/dd")
        self.period_from_input.setCalendarPopup(True)

        self.period_to_label = QLabel("To:")
        self.period_to_input = QDateEdit()
        self.period_to_input.setDate(QDate.currentDate())
        self.period_to_input.setDisplayFormat("yyyy/MM/dd")
        self.period_to_input.setCalendarPopup(True)

        self.update_button = QPushButton("Update")

        self.top_bar.addWidget(self.company_label)
        self.top_bar.addWidget(self.company_input)
        self.top_bar.addWidget(self.exchange_rate_label)
        self.top_bar.addWidget(self.exchange_rate_input)
        self.top_bar.addWidget(self.period_from_label)
        self.top_bar.addWidget(self.period_from_input)
        self.top_bar.addWidget(self.period_to_label)
        self.top_bar.addWidget(self.period_to_input)
        self.top_bar.addWidget(self.update_button)  # Add the button
        self.top_bar.addStretch()
        self.layout.addLayout(self.top_bar)

        self.exchange_rate_input.valueChanged.connect(self.on_exchange_rate_changed)

    def on_exchange_rate_changed(self, value):
        index = self.tabs.currentIndex()
        if index >= 0:
            current_tab = self.tabs.widget(index)
            if hasattr(current_tab, 'name') and '-' in current_tab.name:
                current_tab.exchange_rate = value
                if hasattr(current_tab, 'set_exchange_rate'):
                    current_tab.set_exchange_rate(value)
                self.auto_save()

    def on_tab_changed(self, index):
        """Handle tab change events"""
        if index >= 0:
            # Hide all exchange rate controls in sheets
            for sheet in self.sheets:
                if hasattr(sheet, 'exchange_rate_input'):
                    sheet.exchange_rate_input.setVisible(False)
            current_tab = self.tabs.widget(index)
            # Avoid AttributeError for the plus tab (QWidget)
            if not hasattr(current_tab, 'type'):
                self.exchange_rate_input.setEnabled(False)
                self.exchange_rate_input.setValue(1.0)
                return
            # Enable/disable main exchange rate input
            if current_tab.type == "bank":
                self.exchange_rate_input.setEnabled(True)
                if hasattr(current_tab, 'exchange_rate'):
                    self.exchange_rate_input.setValue(current_tab.exchange_rate)
                else:
                    self.exchange_rate_input.setValue(1.0)
                    current_tab.exchange_rate = 1.0
            else:
                self.exchange_rate_input.setEnabled(False)
                self.exchange_rate_input.setValue(1.0)


    def update_tab_name(self, old_name, new_name):
        """Update the tab text when sheet is renamed, with bank/non-bank name validation"""
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == old_name:
                tab = self.tabs.widget(i)
                is_bank = getattr(tab, 'type', None) == 'bank'
                if is_bank and '-' not in new_name:
                    QMessageBox.warning(self, "Invalid Name", "Bank sheet names must contain a '-'!")
                    return
                if not is_bank and '-' in new_name:
                    QMessageBox.warning(self, "Invalid Name", "Non-bank sheet names must NOT contain a '-'!")
                    return
                self.tabs.setTabText(i, new_name)
                if hasattr(tab, 'name'):
                    tab.name = new_name
                break

    def setup_menu_bar(self):
        """Initialize the menu bar with actions"""
        self.menu = self.menuBar()
        file_menu = self.menu.addMenu("File")

        # File actions
        actions = [
            ("New", self.new_file),
            ("Add Sheet", self.add_sheet_dialog),
            ("Delete Sheet", self.delete_sheet),
            ("Save", self.file_manager.save_file),
            ("Load", self.file_manager.load_file)
        ]

        for text, callback in actions:
            action = QAction(text, self)
            action.triggered.connect(callback)
            file_menu.addAction(action)

            if text == "Delete Sheet":
                file_menu.addSeparator()

        # Add tab switcher to menu
        navigate_menu = self.menu.addMenu("Navigate")
        switch_tab_action = QAction("Switch Tab...", self)
        switch_tab_action.setShortcut("Ctrl+K")
        switch_tab_action.triggered.connect(self.show_tab_switcher)
        navigate_menu.addAction(switch_tab_action)

        # Add right-click context menu for tab switching
        self.tabs.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tabs.customContextMenuRequested.connect(self.show_tab_context_menu)
        # Make tab bar scrollable
        self.tabs.setTabBarAutoHide(False)
        self.tabs.setUsesScrollButtons(True)
        self.tabs.tabBar().setElideMode(Qt.ElideRight)
        # Set tooltips for all tabs
        self.tabs.currentChanged.connect(self.update_tab_tooltips)
        self.update_tab_tooltips()
        # Set tab width to fit sheet name
        self.adjust_tab_widths()

    def adjust_tab_widths(self):
        tab_bar = self.tabs.tabBar()
        font_metrics = tab_bar.fontMetrics()
        for i in range(tab_bar.count()):
            text = tab_bar.tabText(i)
            tab_bar.setTabToolTip(i, text)  # Set tooltip for full name

    def update_tab_tooltips(self):
        for i in range(self.tabs.count()):
            name = self.tabs.tabText(i)
            self.tabs.setTabToolTip(i, name)
        self.adjust_tab_widths()

    def show_tab_context_menu(self, pos):
        menu = QMenu(self)
        from PySide6.QtGui import QActionGroup
        tab_names = [self.tabs.tabText(i) for i in range(self.tabs.count()) if self.tabs.tabText(i) != "+"]
        current_index = self.tabs.currentIndex()
        group = QActionGroup(menu)
        group.setExclusive(True)
        for i, name in enumerate(tab_names):
            action = QAction(name, menu)
            action.setCheckable(True)
            if i == current_index:
                action.setChecked(True)
            action.triggered.connect(lambda checked, idx=i: self.tabs.setCurrentIndex(idx))
            group.addAction(action)
            menu.addAction(action)
        menu.addSeparator()
        switcher_action = QAction("Switch Tab...", menu)
        switcher_action.triggered.connect(self.show_tab_switcher)
        menu.addAction(switcher_action)
        menu.exec(self.tabs.mapToGlobal(pos))

    def add_sheet(self, name=None, is_bank=False):
        """Add a new sheet with optional name and type"""
        if not name:
            name, ok = QInputDialog.getText(self, "Sheet Name", "Enter sheet name:")
            if not ok or not name:
                return

        if is_bank or ("-" in name):
            sheet = self.sheet_manager.create_bank_sheet(name)
        else:
            sheet = self.sheet_manager.create_regular_sheet(name)

    def add_sheet_dialog(self):
        """Show dialog to add a new sheet"""
        print(f"DEBUG ADD: Starting add sheet dialog, current tabs: {self.tabs.count()}")
        dlg = AddSheetDialog(self)
        if dlg.exec() == QDialog.Accepted:
            result = dlg.get_result()
            if len(result) == 3:
                name, sheet_type, currency = result
            else:
                name, sheet_type = result
                currency = None
            print(f"DEBUG ADD: Dialog accepted, name='{name}', type='{sheet_type}', currency='{currency}'")
            if not name:
                print(f"DEBUG ADD: No name provided, aborting")
                return
            # For bank, generate tab name as '{name}-{currency}'
            if sheet_type == "bank":
                if not currency:
                    QMessageBox.warning(self, "Input Error", "Please select a currency.")
                    return
                tab_name = f"{name}-{currency}"
            else:
                tab_name = name
            # Prevent duplicate sheet names
            for i in range(self.tabs.count()):
                if self.tabs.tabText(i) == tab_name:
                    QMessageBox.warning(self, "Duplicate Name", f"A sheet named '{tab_name}' already exists.")
                    return
            new_sheet = None
            if sheet_type == "bank":
                print(f"DEBUG ADD: Creating bank sheet: {tab_name}")
                try:
                    new_sheet = self.sheet_manager.create_bank_sheet(tab_name, currency)
                except TypeError:
                    raise
            elif sheet_type == "非银行交易":
                new_sheet = self.sheet_manager.create_non_bank_sheet()
            else:
                print(f"DEBUG ADD: Creating regular sheet: {tab_name}")
            print(f"DEBUG ADD: Sheet created successfully, total tabs now: {self.tabs.count()}")
            # Remove the blank '+' tab before adding the new sheet
            plus_index = self.tabs.count() - 1
            if self.tabs.tabText(plus_index) == "+":
                self.tabs.removeTab(plus_index)
            # Add the new sheet as a tab
            if new_sheet:
                self.tabs.addTab(new_sheet, tab_name)
                self.tabs.setCurrentWidget(new_sheet)
                print(f"DEBUG ADD: Switched to new sheet")
            # Force auto-save after adding sheet
            print(f"DEBUG ADD: Triggering auto-save...")
            self.auto_save()
            print(f"DEBUG ADD: Add sheet complete")
        else:
            print(f"DEBUG ADD: Dialog cancelled")
        # After adding a sheet, always ensure the plus tab is present
        self._add_plus_tab()

    def delete_sheet(self):
        """Delete the current sheet"""
        idx = self.tabs.currentIndex()
        # Prevent deleting the plus tab
        if self.tabs.tabText(idx) == "+":
            return
        if self.tabs.count() > 1:
            self._suppress_plus_tab = True  # Suppress add sheet dialog after delete
            sheet_name = self.tabs.tabText(idx)
            sheet_to_delete = self.sheets[idx]
            reply = QMessageBox.question(
                self,
                "Delete Sheet",
                f"Are you sure you want to delete the sheet '{sheet_name}'?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                # Check if this is a bank sheet (has currency)
                is_bank_sheet = hasattr(sheet_to_delete, 'type') and sheet_to_delete.type == "bank"

                self.tabs.removeTab(idx)
                del self.sheets[idx]
                self._add_plus_tab()

                self.auto_save()
            self._suppress_plus_tab = False

    def close_tab(self, idx):
        """Close tab at given index"""
        # Prevent closing the plus tab
        if self.tabs.tabText(idx) == "+":
            return
        if self.tabs.count() > 1:
            self._suppress_plus_tab = True  # Suppress add sheet dialog after close
            sheet_name = self.tabs.tabText(idx)
            reply = QMessageBox.question(
                self,
                "Close Sheet",
                f"Are you sure you want to close the sheet '{sheet_name}'?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.tabs.removeTab(idx)
                del self.sheets[idx]
                self._add_plus_tab()
                self.auto_save()
            self._suppress_plus_tab = False

    def new_file(self):
        """Create a new file with default sheet"""
        self.tabs.clear()
        self.sheets = []
        # Set default company name if empty
        self.company_input.setText(self.company_input.text() or "company_name")
        self.period_from_input.setDate(QDate.currentDate().addMonths(-1))
        self.period_to_input.setDate(QDate.currentDate())
        self.sheet_manager.create_bank_sheet("HSBC-USD")
        self.sheet_manager.create_bank_sheet("HSBC-RMB")
        self.sheet_manager.create_bank_sheet("HSBC-EUR")
        self.sheet_manager.create_bank_sheet("HSBC-JPY")
        self.sheet_manager.create_bank_sheet("HSBC-GBP")
        self.sheet_manager.create_non_bank_sheet()
        # Always add '+' tab at the end (even if no other tabs)
        self._add_plus_tab()
        self.auto_save()

    def auto_save(self):
        """Auto-save current state"""
        # Before saving, update exchange rate in current tab if bank sheet
        if True:
            return
        index = self.tabs.currentIndex()
        if index >= 0:
            current_tab = self.tabs.widget(index)
            if hasattr(current_tab, 'name') and '-' in current_tab.name:
                current_tab.exchange_rate = self.exchange_rate_input.value()
        self.file_manager.auto_save()

    def save_file(self):
        """Save current file"""
        self.file_manager.save_file()

    def load_file(self):
        """Load file from disk"""
        self.file_manager.load_file()
        # After loading, set company name to the loaded value
        if hasattr(self, 'company_input') and hasattr(self.file_manager, 'last_loaded_company_name'):
            self.company_input.setText(self.file_manager.last_loaded_company_name)
        # Always ensure the plus tab is present after loading
        self._add_plus_tab()

    def _add_plus_tab(self):
        # Remove all existing '+' tabs first
        for i in reversed(range(self.tabs.count())):
            if self.tabs.tabText(i) == "+":
                self.tabs.removeTab(i)
        # Add a single '+' tab at the end
        plus_index = self.tabs.addTab(QWidget(), "+")
        self.tabs.tabBar().setTabButton(self.tabs.count()-1, QTabBar.RightSide, None)
        self.tabs.tabBar().setTabButton(plus_index, QTabBar.LeftSide, None)
        self.tabs.tabBar().setTabButton(plus_index, QTabBar.RightSide, None)

    def _on_tab_or_plus_clicked(self, index):
        # If last tab (the plus tab) is clicked, open add sheet dialog and revert to previous tab immediately
        if getattr(self, '_suppress_plus_tab', False):
            return
        if index == self.tabs.count() - 1 and self.tabs.tabText(index) == "+":
            prev_index = getattr(self, '_prev_tab_index', None)
            if prev_index is None:
                prev_index = 0
            if self.tabs.count() > 1:
                self.tabs.setCurrentIndex(prev_index)
            dlg = AddSheetDialog(self)
            if dlg.exec() == QDialog.Accepted:
                result = dlg.get_result()
                if len(result) == 3:
                    name, sheet_type, currency = result
                else:
                    name, sheet_type = result
                    currency = None
                if not name:
                    return
                if sheet_type == "bank":
                    if not currency:
                        QMessageBox.warning(self, "Input Error", "Please select a currency.")
                        return
                    tab_name = f"{name}-{currency}"
                else:
                    tab_name = name
                for i in range(self.tabs.count()):
                    if self.tabs.tabText(i) == tab_name:
                        QMessageBox.warning(self, "Duplicate Name", f"A sheet named '{tab_name}' already exists.")
                        return
                new_sheet = None
                if sheet_type == "bank":
                    try:
                        new_sheet = self.sheet_manager.create_bank_sheet(tab_name, currency)
                    except TypeError:
                        new_sheet = self.sheet_manager.create_bank_sheet(tab_name)
                elif sheet_type == "非银行交易":
                    new_sheet = self.sheet_manager.create_non_bank_sheet(tab_name)
                else:
                    print("ERROR: Unknown sheet type from dialog")
                plus_index = self.tabs.count() - 1
                if self.tabs.tabText(plus_index) == "+":
                    self.tabs.removeTab(plus_index)
                if new_sheet:
                    self.tabs.addTab(new_sheet, tab_name)
                    self.tabs.setCurrentWidget(new_sheet)
                self.auto_save()
            else:
                if self.tabs.count() > 1:
                    self.tabs.setCurrentIndex(prev_index)
            self._add_plus_tab()
        else:
            self._prev_tab_index = index
            self.on_tab_changed(index)

    def _on_tab_moved(self, from_index, to_index):
        # Always move the plus tab to the rightmost position after any move
        plus_tab_index = None
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == "+":
                plus_tab_index = i
                break
        if plus_tab_index is not None and plus_tab_index != self.tabs.count() - 1:
            # Remove and re-add the plus tab at the end
            plus_widget = self.tabs.widget(plus_tab_index)
            self.tabs.removeTab(plus_tab_index)
            new_index = self.tabs.addTab(plus_widget, "+")
            self.tabs.tabBar().setTabButton(new_index, QTabBar.RightSide, None)
            self.tabs.tabBar().setTabButton(new_index, QTabBar.LeftSide, None)
            self.tabs.setCurrentIndex(new_index)

        # Only call reorder_sheets for real sheet tabs
        plus_index = self.tabs.count() - 1
        if hasattr(self.sheet_manager, 'reorder_sheets') and from_index < plus_index and to_index < plus_index:
            self.sheet_manager.reorder_sheets(from_index, to_index)

    def set_light_theme(self):
        """Force light theme for better readability regardless of system settings"""
        app = QApplication.instance()
        if app:
            # Create a light palette
            palette = QPalette()

            # Set light colors for all elements
            palette.setColor(QPalette.Window, Qt.white)
            palette.setColor(QPalette.WindowText, Qt.black)
            palette.setColor(QPalette.Base, Qt.white)
            palette.setColor(QPalette.AlternateBase, Qt.lightGray)
            palette.setColor(QPalette.ToolTipBase, Qt.white)
            palette.setColor(QPalette.ToolTipText, Qt.black)
            palette.setColor(QPalette.Text, Qt.black)
            palette.setColor(QPalette.Button, Qt.lightGray)
            palette.setColor(QPalette.ButtonText, Qt.black)
            palette.setColor(QPalette.BrightText, Qt.red)
            palette.setColor(QPalette.Link, Qt.blue)
            palette.setColor(QPalette.Highlight, Qt.blue)
            palette.setColor(QPalette.HighlightedText, Qt.white)

            # Apply the palette to the application
            app.setPalette(palette)

            # Also set stylesheet for tables to ensure white background
            app.setStyleSheet("""
                QTableWidget {
                    background-color: white;
                    alternate-background-color: #f0f0f0;
                    color: black;
                    gridline-color: #d0d0d0;
                }
                QTableWidget::item {
                    background-color: white;
                    color: black;
                }
                QTableWidget::item:selected {
                    background-color: #3daee9;
                    color: white;
                }
                QHeaderView::section {
                    background-color: #e0e0e0;
                    color: black;
                    border: 1px solid #d0d0d0;
                }
                QTabWidget::pane {
                    background-color: white;
                    border: 1px solid #d0d0d0;
                }
                QTabBar::tab {
                    background-color: #e0e0e0;
                    color: black;
                    border: 1px solid #d0d0d0;
                    padding: 4px 8px;
                }
                QTabBar::tab:selected {
                    background-color: white;
                    color: black;
                }
            """)
    def show_tab_switcher(self):
        """Show a dropdown dialog to quickly jump to any tab by name."""
        from PySide6.QtWidgets import QInputDialog
        tab_names = [self.tabs.tabText(i) for i in range(self.tabs.count()) if self.tabs.tabText(i) != "+"]
        if not tab_names:
            return
        name, ok = QInputDialog.getItem(self, "Switch Tab", "Select a sheet:", tab_names, 0, False)
        if ok and name:
            for i in range(self.tabs.count()):
                if self.tabs.tabText(i) == name:
                    self.tabs.setCurrentIndex(i)
                    break
