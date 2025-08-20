from PySide6.QtWidgets import QDoubleSpinBox, QTableWidgetItem, QMenu
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QColor
from excel_table import ExcelTable
from datetime import datetime
import logging

logger = logging.getLogger(__name__)

class SheetManager:
    def __init__(self, main_window):
        self.main_window = main_window

    def create_bank_sheet(self, name, currency=None):
        """Create a bank sheet with exchange rate control"""
        columns = ["序 號", "日  期", "對方科目", "摘   要", "借     方", "貸     方", "餘    額", "發票號碼"]
        table = ExcelTable(auto_save_callback=self.main_window.auto_save, name=name, type="bank")
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)

        # Add exchange rate control
        rate_input = QDoubleSpinBox()
        # Use currency argument if provided, else try to parse from name
        currency_str = currency if currency else (name.split("-")[1] if "-" in name else "")
        rate_input.setPrefix(f"{currency_str}:HKD = 1:")
        rate_input.setValue(1.0)
        rate_input.setDecimals(2)
        rate_input.valueChanged.connect(lambda v, t=table: t.set_exchange_rate(v))
        rate_input.setVisible(False)  # Start hidden
        self.main_window.layout.addWidget(rate_input)
        table.exchange_rate_input = rate_input
        table.set_exchange_rate(1.0)

        self.main_window.tabs.addTab(table, name)
        self.main_window.sheets.append(table)
        return table

    def create_non_bank_sheet(self, name):
        """Create a regular sheet"""
        columns = ["序 號", "日  期", "對方科目", "子科目", "摘   要", "借方(USD)", "借方(EUR)", "借方(JPY)", "借方(GBP)", "借方(CHF)", "借方(CAD)", "借方(AUD)", "借方(CNY)", "借方(HKD)", "借方(NZD)", "贷方(USD)", "贷方(EUR)", "贷方(JPY)", "贷方(GBP)", "贷方(CHF)", "贷方(CAD)", "贷方(AUD)", "贷方(CNY)", "贷方(HKD)", "贷方(NZD)", "备注"]
        table = ExcelTable(auto_save_callback=self.main_window.auto_save, name=name, type="non_bank")
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)

        self.main_window.tabs.addTab(table, name)
        self.main_window.sheets.append(table)
        return table
    
    def create_regular_sheet(self, name):
        """Create a regular sheet"""
        columns = ["序 號", "日  期", "對方科目", "摘   要", "借     方", "貸     方", "借或貸", "餘    額", "發票號碼"]
        table = ExcelTable(auto_save_callback=self.main_window.auto_save, name=name, type="regular")
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)

        self.main_window.tabs.addTab(table, name)
        self.main_window.sheets.append(table)
        return table

    def create_aggregate_sheet(self, sheet_name, subject_filter, column_title):
        """Common method to create aggregate sheets (sales/cost)"""
        columns = ["序 號", "日  期", "對方科目", "摘  要", "發票號碼", column_title, "餘    額", "來源"]
        table = ExcelTable(auto_save_callback=self.main_window.auto_save, name=sheet_name, type="aggregate")
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)
        table.setRowCount(2)  # Just header rows, data will resize it
        self.main_window.tabs.addTab(table, sheet_name)
        self.main_window.sheets.append(table)
        return table

    def create_sales_sheet(self):
        return self.create_aggregate_sheet("銷售收入", "销售收入", "貸     方")

    def create_cost_sheet(self):
        return self.create_aggregate_sheet("銷售成本", "销售成本", "借     方")

    def create_bank_fee_sheet(self):
        """Create a bank fee sheet"""
        return self.create_aggregate_sheet("銀行費用", "银行费用", "借     方")

    def create_director_sheet(self):
        """Create a director sheet - works like other aggregate sheets, no user input"""
        table = self.create_aggregate_sheet("董事往來", "董事往来", "借     方")
        # No special edit triggers - director sheet is read-only like other aggregate sheets
        return table

    def create_payable_sheet(self):
        """Create a payable sheet"""
        # Modified columns - removed invoice, added debit before credit
        columns = ["序 號", "日  期", "對方科目", "摘  要", "借     方", "貸     方", "餘    額", "來源"]
        table = ExcelTable("regular", auto_save_callback=self.main_window.auto_save, name="應付費用")
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)

        table.setEditTriggers(table.EditTrigger.DoubleClicked |
                    table.EditTrigger.EditKeyPressed |
                    table.EditTrigger.AnyKeyPressed)

        self.main_window.tabs.addTab(table, "應付費用")
        self.main_window.sheets.append(table)
        return table

    def create_interest_sheet(self):
        """Create an interest income sheet"""
        return self.create_aggregate_sheet("利息收入", "利息收入", "貸     方")

    def create_salary_sheet(self):
        """Create a salary sheet"""
        return self.create_aggregate_sheet("工資", "工資", "借     方")

    def _setup_currency_sheet_structure(self, table, amount_column_title):
        """Common method to set up sales sheet structure with multi-currency columns"""
        currency_set = set()
        for sheet in self.main_window.sheets:
            if sheet.type == "bank":
                print(f"DEBUG SETUP: Found bank sheet '{sheet.name}' with currency '{sheet.currency}'")
                currency_set.add(sheet.currency)

        currency_list = sorted(currency_set)
        print(f"DEBUG SETUP: Final currency list: {currency_list}")
        
        # Set up columns
        main_headers = ["序 號", "日  期", "對方科目", "摘  要", "發票號碼"]
        sub_headers = ["", "", "", "", ""]
        
        credit_col_start = len(main_headers)
        credit_col_count = len(currency_list)
        print(f"DEBUG SETUP: Credit columns start at {credit_col_start}, count: {credit_col_count}")
        
        # Add currency columns
        for currency in currency_list:
            main_headers.append(amount_column_title)
            sub_headers.append(f"原币({currency})")
        
        # Add final columns
        main_headers.extend(["餘    額", "來源"])
        sub_headers.extend(["", ""])
        
        # Set up merged ranges for amount columns (if more than one currency)
        merged_ranges = []
        if credit_col_count > 1:
            merged_ranges.append((credit_col_start, credit_col_start + credit_col_count - 1))
        
        # Use 2-row headers for aggregate sheets
        table.setup_two_row_headers(main_headers, sub_headers, merged_ranges)
        # Start data from row 0 instead of row 2
        if table.rowCount() < 1:
            table.setRowCount(1)

        return currency_list, credit_col_start, credit_col_count

    def _populate_currency_sheet_data(self, table, subject_filter, currency_list, amount_col_start, amount_col_count, is_debit):
        """Common method to populate sales sheet data rows"""
        # For aggregate sheets with 2-row table headers, start data from row 2
        # For other sheets, also start from row 2
        start_data_row = 2
        
        # Only clear data rows, NOT header rows
        for row in range(start_data_row, table.rowCount()):
            for col in range(table.columnCount()):
                table.setItem(row, col, None)

        # Collect bank rows
        bank_rows = []
        print(f"DEBUG DATA: Looking for subject_filter '{subject_filter}' in bank sheets")
        for sheet in self.main_window.sheets:
            if sheet.type == "bank":
                print(f"DEBUG DATA: Checking bank sheet '{sheet.name}' (currency: {sheet.currency}) with {sheet.rowCount()} rows")
                for row in range(sheet.rowCount() - 2):
                    subject_item = sheet.item(row, 2)
                    if subject_item and subject_filter in subject_item.text():
                        print(f"DEBUG DATA: Found match in {sheet.name} row {row}: '{subject_item.text()}'")
                        bank_rows.append((sheet, row, sheet.currency))

        # Sort bank rows by date
        def get_date_obj(bank_row):
            sheet, row, currency = bank_row
            date_item = sheet.item(row, 1)
            date_str = date_item.text() if date_item else ""
            
            # Try multiple date formats
            date_formats = ["%Y/%m/%d"]
            for fmt in date_formats:
                try:
                    return datetime.strptime(date_str, fmt).date()
                except (ValueError, AttributeError) as e:
                    continue 
            # If no format works, return minimum date so it appears first
            print(f"DEBUG DATE: Could not parse date '{date_str}' from {sheet.name} row {row}")
            return datetime.min.date()

        bank_rows.sort(key=get_date_obj)
        print(f"DEBUG DATA: Total bank rows found: {len(bank_rows)}")

        # Calculate needed rows based on sheet type
        needed_rows = start_data_row + len(bank_rows)
        table.setRowCount(needed_rows)

        # Write data rows starting from the appropriate row
        row_idx = start_data_row
        for i, (sheet, row, currency) in enumerate(bank_rows):
            for col in range(table.columnCount()):
                if col < amount_col_start:
                    if col == 0:
                        value = str(i+1)
                    elif col == 1:
                        item = sheet.item(row, 1)
                        value = item.text() if item else ""
                    elif col == 2:
                        item = sheet.item(row, 2)
                        value = item.text() if item else ""
                    elif col == 3:
                        item = sheet.item(row, 3)
                        value = item.text() if item else ""
                    elif col == 4:
                        item = sheet.item(row, 7)
                        value = item.text() if item else ""
                    else:
                        value = ""
                    item = QTableWidgetItem(value)
                    # Make all generated data uneditable for director sheet (and all aggregate sheets)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    item.setData(Qt.BackgroundRole, QColor(200, 255, 200))
                    table.setItem(row_idx, col, item)
                elif amount_col_start <= col < amount_col_start + amount_col_count:
                    cur = currency_list[col - amount_col_start]
                    if cur == currency:
                        # Get DEBIT value from bank sheet (column 4) and write to CREDIT column in sales sheet
                        # For interest income, we should actually read from credit column since interest is received
                        amount_item = sheet.item(row, 5 if is_debit else 4)  # Original logic for others
                        value = amount_item.text() if amount_item else ""
                        item = QTableWidgetItem(value)
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                        item.setData(Qt.BackgroundRole, QColor(200, 255, 200))
                        table.setItem(row_idx, col, item)
                    else:
                        item = QTableWidgetItem("")
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                        item.setData(Qt.BackgroundRole, QColor(200, 255, 200))
                        table.setItem(row_idx, col, item)
                else:
                    # Balance and source columns
                    if col == table.columnCount() - 1:  # Source column
                        value = f"{sheet.name}:{row+1}"
                    else:
                        value = ""
                    item = QTableWidgetItem(value)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    item.setData(Qt.BackgroundRole, QColor(200, 255, 200))
                    table.setItem(row_idx, col, item)
            row_idx += 1


    def refresh_aggregate_sheet(self, subject_filter, column_title):
        """Refresh the specified aggregate sheet"""
        current_tab = self.main_window.tabs.currentWidget()
        if current_tab:
            # Special handling for sales sheet refresh
            if subject_filter in ["销售收入", "销售成本", "银行费用", "利息收入", "董事往来", "董事往來"]:
                print(f"DEBUG REFRESH: Refreshing aggregate sheet for '{subject_filter}' with column '{column_title}'")
                
                # Check if table has correct multi-currency structure
                # Interest income should be credit (like sales), director should be debit (like costs)
                is_debit = subject_filter in ["销售成本", "银行费用", "董事往来", "董事往來"]
                
                # Get current currencies from remaining bank sheets
                currency_set = set()
                for sheet in self.main_window.sheets:
                    if sheet.type == "bank":
                        currency_set.add(sheet.currency)
                        print(f"DEBUG REFRESH: Found bank sheet '{sheet.name}' with currency '{sheet.currency}'")

                print(f"DEBUG REFRESH: Current currencies: {sorted(currency_set)}")
                
                # CRITICAL: Force complete rebuild of the table structure
                currency_list, credit_col_start, credit_col_count = self._setup_currency_sheet_structure(current_tab, column_title)
                
                # Always populate data (director sheet now works like other aggregate sheets)
                self._populate_currency_sheet_data(current_tab, subject_filter, currency_list, credit_col_start, credit_col_count, is_debit)
                
                print(f"DEBUG REFRESH: Completed refresh for '{subject_filter}', columns: {current_tab.columnCount()}")
                return

    def reorder_sheets(self, from_index, to_index):
        """Handle tab reordering to keep sheets list in sync"""
        self.main_window.sheets.insert(to_index, self.main_window.sheets.pop(from_index))
        self.main_window.auto_save()