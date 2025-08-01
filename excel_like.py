from PySide6.QtWidgets import (
    QMainWindow, QTabWidget, QLineEdit, QLabel, QHBoxLayout, QVBoxLayout,
    QWidget, QInputDialog, QDateEdit, QDialog, QMenu
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

        # Data storage
        self.sheets = []
        self.sales_sheet = None
        self.cost_sheet = None
        self.user_added_rows = None

        # Connect signals for auto-save
        self.company_input.textChanged.connect(self.auto_save)
        self.period_from_input.dateChanged.connect(self.auto_save)
        self.period_to_input.dateChanged.connect(self.auto_save)

        # Connect tab signals
        self.tabs.tabCloseRequested.connect(self.close_tab)
        self.tabs.currentChanged.connect(self.on_tab_changed)
        self.tabs.tabBar().tabMoved.connect(self.sheet_manager.reorder_sheets)

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

    def setup_top_bar(self):
        """Setup the top bar with company name and period inputs"""
        self.top_bar = QHBoxLayout()
        self.company_label = QLabel("Company Name:")
        self.company_input = QLineEdit()
        self.company_input.setText("company_name")
        self.company_input.setPlaceholderText("Enter company name...")

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
        self.top_bar.addWidget(self.period_from_label)
        self.top_bar.addWidget(self.period_from_input)
        self.top_bar.addWidget(self.period_to_label)
        self.top_bar.addWidget(self.period_to_input)
        self.top_bar.addStretch()
        self.layout.addLayout(self.top_bar)

    def on_tab_changed(self, index):
        """Handle tab change events"""
        if index >= 0:
            # Hide all exchange rate controls
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
                self.sheet_manager.refresh_aggregate_sheet("销售收入", "借     方")
            elif tab_name == "銷售成本":
                self.sheet_manager.refresh_aggregate_sheet("销售成本", "貸     方")
            elif tab_name == "銀行費用":
                self.sheet_manager.refresh_aggregate_sheet("银行费用", "貸     方")
            elif tab_name == "利息收入":
                self.sheet_manager.refresh_aggregate_sheet("利息收入", "借     方")
            elif tab_name == "應付費用":
                self.sheet_manager.refresh_aggregate_sheet("董事往来", "貸     方")
            elif tab_name == "董事往來":
                current_tab.user_added_rows = getattr(current_tab, 'user_added_rows', set())
                self.sheet_manager.refresh_aggregate_sheet("董事往来", "貸     方")

    def update_tab_name(self, old_name, new_name):
        """Update the tab text when sheet is renamed"""
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == old_name:
                self.tabs.setTabText(i, new_name)
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
            self.sheet_manager.create_bank_sheet(name)
        else:
            self.sheet_manager.create_regular_sheet(name)

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
                new_sheet = self.sheet_manager.create_bank_sheet(name)
            elif sheet_type == "aggregate":
                print(f"DEBUG ADD: Creating aggregate sheet: {name}")
                # Create the appropriate aggregate sheet
                if name == "銷售收入":
                    new_sheet = self.sheet_manager.create_sales_sheet()
                elif name == "銷售成本":
                    new_sheet = self.sheet_manager.create_cost_sheet()
                elif name == "銀行費用":
                    new_sheet = self.sheet_manager.create_bank_fee_sheet()
                elif name == "利息收入":
                    new_sheet = self.sheet_manager.create_interest_sheet()
                elif name == "應付費用":
                    new_sheet = self.sheet_manager.create_payable_sheet()
                elif name == "董事往來":
                    new_sheet = self.sheet_manager.create_director_sheet()
                else:
                    # Create a regular sheet for other aggregate types like 工資, 商業登記證書, etc.
                    print(f"DEBUG ADD: Creating regular sheet for aggregate type: {name}")
                    new_sheet = self.sheet_manager.create_regular_sheet(name)
            else:
                print(f"DEBUG ADD: Creating regular sheet: {name}")
                new_sheet = self.sheet_manager.create_regular_sheet(name)

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
        self.auto_save()

    def auto_save(self):
        """Auto-save current state"""
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
