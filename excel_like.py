from PySide6.QtWidgets import (
    QMainWindow, QTabWidget, QLineEdit, QLabel, QHBoxLayout, QVBoxLayout,
    QWidget, QInputDialog, QDateEdit, QDialog, QMenu, QMessageBox, QDoubleSpinBox,
    QToolButton, QTabBar
)
from PySide6.QtGui import QAction
from PySide6.QtCore import Qt, QDate
from dialogs import AddSheetDialog
from sheet_manager import SheetManager
from file_manager import FileManager

class ExcelLike(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("bankNote")
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
        self.sheet_manager.create_bank_sheet("HSBC-USD")
        self.sales_sheet = self.sheet_manager.create_sales_sheet()
        self.cost_sheet = self.sheet_manager.create_cost_sheet()
        self.sheet_manager.create_bank_fee_sheet()
        self.sheet_manager.create_interest_sheet()
        self.sheet_manager.create_payable_sheet()
        self.sheet_manager.create_director_sheet()
        # Always add '+' tab at the end (even if no other tabs)
        self._add_plus_tab()
        self.tabs.currentChanged.connect(self._on_tab_or_plus_clicked)

        # Connect signals for auto-save
        self.company_input.editingFinished.connect(self.auto_save)
        self.period_from_input.dateChanged.connect(self.auto_save)
        self.period_to_input.dateChanged.connect(self.auto_save)

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

        self.top_bar.addWidget(self.company_label)
        self.top_bar.addWidget(self.company_input)
        self.top_bar.addWidget(self.exchange_rate_label)
        self.top_bar.addWidget(self.exchange_rate_input)
        self.top_bar.addWidget(self.period_from_label)
        self.top_bar.addWidget(self.period_from_input)
        self.top_bar.addWidget(self.period_to_label)
        self.top_bar.addWidget(self.period_to_input)
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
            if hasattr(current_tab, 'name') and '-' in current_tab.name:
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
                print("bank feeeeeeeeeeee")
                self.sheet_manager.refresh_aggregate_sheet("销售收入", "貸     方")
            elif tab_name == "銷售成本":
                self.sheet_manager.refresh_aggregate_sheet("销售成本", "借     方")
            elif tab_name == "銀行費用":
                self.sheet_manager.refresh_aggregate_sheet("银行费用", "借     方")
            elif tab_name == "利息收入":
                self.sheet_manager.refresh_aggregate_sheet("利息收入", "貸     方")
            elif tab_name == "應付費用":
                self.sheet_manager.refresh_aggregate_sheet("董事往来", "貸     方")
            elif tab_name == "董事往來":
                current_tab.user_added_rows = getattr(current_tab, 'user_added_rows', set())
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
            sheet.type = "bank"
        else:
            sheet = self.sheet_manager.create_regular_sheet(name)
            sheet.type = "regular"

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
                    new_sheet = self.sheet_manager.create_bank_sheet(tab_name)
                new_sheet.type = "bank"
                if currency:
                    new_sheet.currency = currency
            elif sheet_type == "銷售收入":
                new_sheet = self.sheet_manager.create_sales_sheet()
                new_sheet.type = "aggregate"
            elif sheet_type == "銷售成本":
                new_sheet = self.sheet_manager.create_cost_sheet()
                new_sheet.type = "aggregate"
            elif sheet_type == "銀行費用":
                new_sheet = self.sheet_manager.create_bank_fee_sheet()
                new_sheet.type = "aggregate"
            elif sheet_type == "利息收入":
                new_sheet = self.sheet_manager.create_interest_sheet()
                new_sheet.type = "aggregate"
            elif sheet_type == "應付費用":
                new_sheet = self.sheet_manager.create_payable_sheet()
                new_sheet.type = "aggregate"
            elif sheet_type == "董事往來":
                new_sheet = self.sheet_manager.create_director_sheet()
                new_sheet.type = "aggregate"
            elif sheet_type == "工資":
                new_sheet = self.sheet_manager.create_salary_sheet()
                new_sheet.type = "aggregate"
            else:
                print(f"DEBUG ADD: Creating regular sheet: {tab_name}")
                new_sheet = self.sheet_manager.create_regular_sheet(tab_name)
                new_sheet.type = "regular"
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
            reply = QMessageBox.question(
                self,
                "Delete Sheet",
                f"Are you sure you want to delete the sheet '{sheet_name}'?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
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
                prev_index = max(0, index - 1)
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
                    new_sheet.type = "bank"
                    if currency:
                        new_sheet.currency = currency
                elif sheet_type == "銷售收入":
                    new_sheet = self.sheet_manager.create_sales_sheet()
                    new_sheet.type = "aggregate"
                elif sheet_type == "銷售成本":
                    new_sheet = self.sheet_manager.create_cost_sheet()
                    new_sheet.type = "aggregate"
                elif sheet_type == "銀行費用":
                    new_sheet = self.sheet_manager.create_bank_fee_sheet()
                    new_sheet.type = "aggregate"
                elif sheet_type == "利息收入":
                    new_sheet = self.sheet_manager.create_interest_sheet()
                    new_sheet.type = "aggregate"
                elif sheet_type == "應付費用":
                    new_sheet = self.sheet_manager.create_payable_sheet()
                    new_sheet.type = "aggregate"
                elif sheet_type == "董事往來":
                    new_sheet = self.sheet_manager.create_director_sheet()
                    new_sheet.type = "aggregate"
                elif sheet_type == "工資":
                    new_sheet = self.sheet_manager.create_salary_sheet()
                    new_sheet.type = "aggregate"
                else:
                    new_sheet = self.sheet_manager.create_regular_sheet(tab_name)
                    new_sheet.type = "regular"
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
