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
        if platform.system() != "Windows":
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
        """Scan all sheets for 應付賬款 entries and create payable sheets"""
        self.scan_and_create_payable_sheets()

    def scan_and_create_payable_sheets(self):
        """Scan all bank and non-bank sheets for 應付賬款 entries and create corresponding payable sheets"""
        payable_data = {}  # Dictionary to group entries by 摘要 (description)
        
        # Scan all sheets for 應付賬款 entries
        for i in range(self.tabs.count()):
            tab_name = self.tabs.tabText(i)
            if tab_name == "+" or tab_name.startswith("應付賬款-"):
                continue  # Skip plus tab and existing payable sheets
                
            tab = self.tabs.widget(i)
            if not hasattr(tab, 'type'):
                continue
                
            # Only scan bank and non_bank sheets
            if tab.type not in ["bank", "non_bank"]:
                continue
                
            print(f"DEBUG: Scanning sheet '{tab_name}' for 應付賬款 entries")
            
            # Find the column indices
            col_count = tab.columnCount()
            对方科目_col = None
            子科目_col = None
            
            for col in range(col_count):
                header_item = tab.horizontalHeaderItem(col)
                if header_item:
                    header_text = header_item.text().replace(" ", "")
                    if "對方科目" in header_text:
                        对方科目_col = col
                    elif "子科目" in header_text:
                        子科目_col = col
            
            if 对方科目_col is None or 子科目_col is None:
                print(f"DEBUG: Could not find required columns in sheet '{tab_name}'")
                continue
                
            # Scan rows for 應付賬款 entries
            for row in range(tab.rowCount()):
                对方科目_item = tab.item(row, 对方科目_col)
                子科目_item = tab.item(row, 子科目_col)
                
                if (对方科目_item and 对方科目_item.text().strip() == "應付賬款" and
                    子科目_item and 子科目_item.text().strip()):
                    
                    子科目_value = 子科目_item.text().strip()
                    print(f"DEBUG: Found 應付賬款 entry with 子科目: '{子科目_value}' in sheet '{tab_name}', row {row}")
                    
                    # Initialize the payable group if not exists
                    if 子科目_value not in payable_data:
                        payable_data[子科目_value] = []
                    
                    # Copy the entire row data
                    row_data = []
                    for col in range(col_count):
                        item = tab.item(row, col)
                        row_data.append(item.text() if item else "")
                    
                    # Store row data with source sheet info
                    payable_data[子科目_value].append({
                        'source_sheet': tab_name,
                        'row_data': row_data,
                        'source_headers': [tab.horizontalHeaderItem(col).text() if tab.horizontalHeaderItem(col) else f"Col{col}" for col in range(col_count)]
                    })
        
        # Create payable sheets for each 子科目 group
        for 子科目_value, entries in payable_data.items():
            sheet_name = f"應付賬款-{子科目_value}"
            
            # Check if sheet already exists
            sheet_exists = False
            for i in range(self.tabs.count()):
                if self.tabs.tabText(i) == sheet_name:
                    sheet_exists = True
                    existing_sheet = self.tabs.widget(i)
                    print(f"DEBUG: Sheet '{sheet_name}' already exists, updating it")
                    break
            
            if not sheet_exists:
                print(f"DEBUG: Creating new payable sheet: '{sheet_name}'")
                new_sheet = self.sheet_manager.create_payable_detail_sheet(sheet_name)
                
                # Remove plus tab temporarily
                plus_index = self.tabs.count() - 1
                if self.tabs.tabText(plus_index) == "+":
                    self.tabs.removeTab(plus_index)
                
                # Add the new sheet
                self.tabs.addTab(new_sheet, sheet_name)
                self._add_plus_tab()  # Re-add plus tab
                existing_sheet = new_sheet
            
            # Populate the sheet with data
            self._populate_payable_sheet(existing_sheet, entries)
        
        if payable_data:
            print(f"DEBUG: Created/updated {len(payable_data)} payable sheets")
            QMessageBox.information(self, "Update Complete", 
                                   f"Created/updated {len(payable_data)} payable sheets for 應付賬款 entries.")
        else:
            QMessageBox.information(self, "Update Complete", 
                                   "No 應付賬款 entries found in bank or non-bank sheets.")
    
    def _populate_payable_sheet(self, sheet, entries):
        """Populate a payable sheet with the collected entries"""
        print(f"DEBUG POPULATE: Starting to populate sheet with {len(entries)} entries")
        print(f"DEBUG POPULATE: Sheet has {sheet.columnCount()} columns and {sheet.rowCount()} rows")
        
        # Check headers
        for col in range(min(sheet.columnCount(), 16)):
            header_item = sheet.horizontalHeaderItem(col)
            if header_item:
                print(f"DEBUG POPULATE: Column {col} header: '{header_item.text()}'")
            else:
                print(f"DEBUG POPULATE: Column {col} has no header item")
        
        # Clear existing data (keep headers)
        for row in range(2, sheet.rowCount()):
            for col in range(sheet.columnCount()):
                sheet.setItem(row, col, None)
        # Set row count to accommodate all entries plus some extra
        sheet.setRowCount(max(len(entries) + 10, 50))
        # Populate data starting from row 2 (after headers)
        for row_idx, entry in enumerate(entries):
            target_row = row_idx + 2
            source_data = entry['row_data']
            source_headers = entry['source_headers']
            source_sheet = entry.get('source_sheet', '')
            currency = self._extract_currency_from_sheet(source_sheet)
            print(f"DEBUG PAYABLE: Processing entry from sheet '{source_sheet}' with currency '{currency}'")
            # Copy all columns except 貸方, which goes to 借方(currency)
            for src_col, src_value in enumerate(source_data):
                if src_col < len(source_headers):
                    src_header = source_headers[src_col].replace(" ", "")
                    print(f"DEBUG PAYABLE: src_header='{src_header}', src_value='{src_value}'")
                    # 貸方 value from bank sheet goes to 借方(currency) in payable sheet
                    if src_header == "貸方" and currency and src_value and src_value.strip():
                        print(f"DEBUG PAYABLE: Mapping 貸方 value '{src_value}' to 借方({currency}) column")
                        # Find 借方(currency) column in payable sheet
                        for col in range(sheet.columnCount()):
                            # Get currency header from row 1, col
                            header_item = sheet.item(1, col)
                            if header_item and header_item.text().replace(" ", "") == f"原币({currency})":
                                print(f"DEBUG PAYABLE: Setting value '{src_value}' at row {target_row}, col {col}")
                                sheet.setItem(target_row, col, QTableWidgetItem(str(src_value)))
                                break
                        continue  # skip default mapping for 貸方
                    # Other columns: copy by header match
                    target_col = None
                    for col in range(sheet.columnCount()):
                        target_header_item = sheet.item(0, col)
                        if target_header_item:
                            target_header = target_header_item.text().replace(" ", "")
                            if src_header == target_header or src_header in target_header:
                                target_col = col
                                break
                    if target_col is not None:
                        print(f"DEBUG PAYABLE: Copying value '{src_value}' to col {target_col} (header '{target_header_item.text()}')")
                        sheet.setItem(target_row, target_col, QTableWidgetItem(str(src_value)))
        print(f"DEBUG: Populated payable sheet with {len(entries)} entries")
    
    def _extract_currency_from_sheet(self, sheet_name):
        """Extract currency from sheet name (e.g., 'HSBC-USD' -> 'USD')"""
        if '-' in sheet_name:
            return sheet_name.split('-')[-1]
        return None
    
    def _find_currency_column(self, sheet, column_type, currency):
        """Find the appropriate currency column in the payable sheet"""
        # Look for columns with the specific currency
        for col in range(sheet.columnCount()):
            header_item = sheet.horizontalHeaderItem(col)
            if header_item:
                header_text = header_item.text()
                if f"原币({currency})" in header_text:
                    # Check if it's the right type (借方 or 貸方)
                    if "借方" in column_type and col >= 4 and col <= 8:  # Debit currency columns (back to original range)
                        return col
                    elif "貸方" in column_type and col >= 9 and col <= 13:  # Credit currency columns (back to original range)
                        return col
        return None

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

            tab_name = self.tabs.tabText(index)
            if tab_name == "銷售收入":
                self.sheet_manager.refresh_aggregate_sheet("销售收入", "貸     方")
            elif tab_name == "銷售成本":
                self.sheet_manager.refresh_aggregate_sheet("销售成本", "借     方")
            elif tab_name == "銀行費用":
                self.sheet_manager.refresh_aggregate_sheet("银行费用", "借     方")
            elif tab_name == "利息收入":
                self.sheet_manager.refresh_aggregate_sheet("利息收入", "貸     方")
            elif tab_name == "應付費用":
                # do nothing, payable data edit or input by user
                pass
            elif tab_name == "董事往來":
                # Always refresh to get latest bank data at top (not saved to file)
                self.sheet_manager.refresh_aggregate_sheet("董事往来", "貸     方")

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
            elif sheet_type == "銷售收入":
                new_sheet = self.sheet_manager.create_sales_sheet()
            elif sheet_type == "銷售成本":
                new_sheet = self.sheet_manager.create_cost_sheet()
            elif sheet_type == "銀行費用":
                new_sheet = self.sheet_manager.create_bank_fee_sheet()
            elif sheet_type == "利息收入":
                new_sheet = self.sheet_manager.create_interest_sheet()
            elif sheet_type == "應付費用":
                new_sheet = self.sheet_manager.create_payable_sheet()
            elif sheet_type == "董事往來":
                new_sheet = self.sheet_manager.create_director_sheet()
            elif sheet_type == "資":
                new_sheet = self.sheet_manager.create_salary_sheet()
            elif sheet_type == "非银行交易":
                new_sheet = self.sheet_manager.create_non_bank_sheet()
            else:
                print(f"DEBUG ADD: Creating regular sheet: {tab_name}")
                new_sheet = self.sheet_manager.create_regular_sheet(tab_name)
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
                
                # If we deleted a bank sheet, refresh all aggregate sheets to remove the currency column
                if is_bank_sheet:
                    print(f"DEBUG: Deleted bank sheet '{sheet_name}', refreshing all aggregate sheets")
                    self._refresh_all_aggregate_sheets()
                
                self.auto_save()
            self._suppress_plus_tab = False

    def _refresh_all_aggregate_sheets(self):
        """Refresh all aggregate sheets to update currency columns after bank sheet deletion"""
        print("DEBUG: Refreshing all aggregate sheets after bank sheet deletion")
        
        # Save current tab to restore later
        current_index = self.tabs.currentIndex()
        
        # Find and refresh each aggregate sheet
        for i in range(self.tabs.count()):
            tab_name = self.tabs.tabText(i)
            if tab_name in ["銷售收入", "銷售成本", "銀行費用", "利息收入", "董事往來"]:
                print(f"DEBUG: Refreshing aggregate sheet '{tab_name}'")
                # Switch to the tab to make it current for refresh
                self.tabs.setCurrentIndex(i)
                
                # Refresh based on sheet type
                if tab_name == "銷售收入":
                    self.sheet_manager.refresh_aggregate_sheet("销售收入", "貸     方")
                elif tab_name == "銷售成本":
                    self.sheet_manager.refresh_aggregate_sheet("销售成本", "借     方")
                elif tab_name == "銀行費用":
                    self.sheet_manager.refresh_aggregate_sheet("银行费用", "借     方")
                elif tab_name == "利息收入":
                    self.sheet_manager.refresh_aggregate_sheet("利息收入", "貸     方")
                elif tab_name == "董事往來":
                    self.sheet_manager.refresh_aggregate_sheet("董事往来", "貸     方")
        
        # Restore original tab selection
        if current_index < self.tabs.count():
            self.tabs.setCurrentIndex(current_index)

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
        self.sheet_manager.create_non_bank_sheet()
        self.sales_sheet = self.sheet_manager.create_sales_sheet()
        self.cost_sheet = self.sheet_manager.create_cost_sheet()
        self.sheet_manager.create_bank_fee_sheet()
        self.sheet_manager.create_interest_sheet()
        self.sheet_manager.create_payable_sheet()
        self.sheet_manager.create_director_sheet()
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
                break
        # Add a single '+' tab at the end
        plus_index = self.tabs.addTab(QWidget(), "+")
        self.tabs.tabBar().setTabButton(self.tabs.count()-1, QTabBar.RightSide, None)
        # Remove close button from the plus tab
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
                elif sheet_type == "銷售收入":
                    new_sheet = self.sheet_manager.create_sales_sheet()
                elif sheet_type == "銷售成本":
                    new_sheet = self.sheet_manager.create_cost_sheet()
                elif sheet_type == "銀行費用":
                    new_sheet = self.sheet_manager.create_bank_fee_sheet()
                elif sheet_type == "利息收入":
                    new_sheet = self.sheet_manager.create_interest_sheet()
                elif sheet_type == "應付費用":
                    new_sheet = self.sheet_manager.create_payable_sheet()
                elif sheet_type == "董事往來":
                    new_sheet = self.sheet_manager.create_director_sheet()
                elif sheet_type == "工資":
                    new_sheet = self.sheet_manager.create_salary_sheet()
                elif sheet_type == "非银行交易":
                    new_sheet = self.sheet_manager.create_non_bank_sheet(tab_name)
                else:
                    new_sheet = self.sheet_manager.create_regular_sheet(tab_name)
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
        plus_index = self.tabs.count() - 1
        current_index = self.tabs.currentIndex()
        # Only re-add the plus tab if it was moved
        if from_index == plus_index or to_index == plus_index:
            # Remove all existing '+' tabs first
            for i in reversed(range(self.tabs.count())):
                if self.tabs.tabText(i) == "+":
                    self.tabs.removeTab(i)
                    break
            # Add a single '+' tab at the end
            self.tabs.addTab(QWidget(), "+")
            self.tabs.tabBar().setTabButton(self.tabs.count()-1, QTabBar.RightSide, None)
            # Remove close button from the plus tab
            self.tabs.tabBar().setTabButton(plus_index, QTabBar.LeftSide, None)
            self.tabs.tabBar().setTabButton(plus_index, QTabBar.RightSide, None)
            # Restore focus
            if current_index < self.tabs.count():
                self.tabs.setCurrentIndex(current_index)
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