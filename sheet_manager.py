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

    def create_regular_sheet(self, name):
        """Create a regular sheet"""
        logger.info(f"DEBUG CREATE: Creating regular sheet '{name}'")
        columns = ["序 號", "日  期", "對方科目", "摘   要", "借     方", "貸     方", "借或貸", "餘    額", "發票號碼"]
        table = ExcelTable(auto_save_callback=self.main_window.auto_save, name=name, type="regular")
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)

        self.main_window.tabs.addTab(table, name)
        self.main_window.sheets.append(table)
        logger.info(f"DEBUG CREATE: Regular sheet '{name}' created, total sheets: {len(self.main_window.sheets)}, total tabs: {self.main_window.tabs.count()}")
        return table

    def create_aggregate_sheet(self, sheet_name, subject_filter, column_title):
        """Common method to create aggregate sheets (sales/cost)"""
        columns = ["序 號", "日  期", "對方科目", "摘  要", "發票號碼", column_title, "餘    額", "來源"]
        table = ExcelTable(auto_save_callback=self.main_window.auto_save, name=sheet_name, type="aggregate")
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)

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
        """Create a director sheet"""
        table = self.create_aggregate_sheet("董事往來", "董事往来", "借     方")
        # Make empty rows editable, but bank data rows remain uneditable
        table.setEditTriggers(table.EditTrigger.DoubleClicked |
                           table.EditTrigger.EditKeyPressed |
                           table.EditTrigger.AnyKeyPressed)

        # Set up cell change tracking for director sheet
        def on_cell_changed(row, col):
            if row >= 2:  # Skip header rows
                item = table.item(row, col)
                # Only add to user_added_rows if it's not bank data (no green background)
                if not (item and item.data(Qt.BackgroundRole)):
                    table.user_added_rows.add(row)
                    # Trigger auto-save when user edits director sheet
                    self.main_window.auto_save()

        # Disconnect any existing connection first
        try:
            table.cellChanged.disconnect()
        except:
            pass

        table.cellChanged.connect(on_cell_changed)

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
                currency_set.add(sheet.currency)

        currency_list = sorted(currency_set)
        columns = ["序 號", "日  期", "對方科目", "摘  要", "發票號碼"]
        credit_col_start = len(columns)
        credit_col_count = len(currency_list)
        # Add "貸     方" columns for each currency
        columns += [amount_column_title] * credit_col_count
        columns += ["餘    額", "來源"]
        excel_headers = [chr(ord('A') + i) for i in range(len(columns))]
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(excel_headers)

        # Preserve existing row count if it's already set properly, otherwise set to 2 for headers
        if table.rowCount() < 102:
            table.setRowCount(max(102, table.rowCount()))  # Ensure minimum rows for director sheet

        # Set title row (row 0) - each credit column gets "貸     方"
        for col, title in enumerate(columns):
            item = QTableWidgetItem(title)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            table.setItem(0, col, item)

        # Clear all existing spans first
        for row in range(2):
            for col in range(table.columnCount()):
                table.setSpan(row, col, 1, 1)

        # Set currency sub-header row (row 1) BEFORE merging
        for col in range(table.columnCount()):
            if credit_col_start <= col < credit_col_start + credit_col_count:
                cur = currency_list[col - credit_col_start]
                item = QTableWidgetItem(f"原币({cur})")
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                table.setItem(1, col, item)

        # Now do the merging AFTER all cells are set
        # Merge all "貸     方" columns horizontally in the title row ONLY
        if credit_col_count > 1:
            table.setSpan(0, credit_col_start, 1, credit_col_count)  # Only row 0
            header_item = table.item(0, credit_col_start)
            header_item.setTextAlignment(Qt.AlignCenter)

        # For non-credit columns, merge vertically (span 2 rows)
        for col in range(table.columnCount()):
            if not (credit_col_start <= col < credit_col_start + credit_col_count):
                table.setSpan(0, col, 2, 1)

        return currency_list, credit_col_start, credit_col_count

    def _populate_currency_sheet_data(self, table, subject_filter, currency_list, amount_col_start, amount_col_count, is_debit):
        """Common method to populate sales sheet data rows"""
        for row in range(2, table.rowCount()):
            for col in range(table.columnCount()):
                table.setItem(row, col, None)

        # Collect bank rows
        bank_rows = []
        for sheet in self.main_window.sheets:
            if "-" in sheet.name:
                currency = getattr(sheet, "currency", None)
                if not currency:
                    parts = sheet.name.split("-")
                    currency = parts[1] if len(parts) > 1 else ""
                for row in range(sheet.rowCount() - 2):
                    subject_item = sheet.item(row, 2)
                    if subject_item and subject_filter in subject_item.text():
                        bank_rows.append((sheet, row, currency))

        # Sort bank rows by date
        def get_date_obj(bank_row):
            sheet, row, currency = bank_row
            date_item = sheet.item(row, 1)
            date_str = date_item.text() if date_item else ""
            try:
                return datetime.strptime(date_str, "%Y/%m/%d").date()
            except (ValueError, AttributeError):
                return datetime.min.date()

        bank_rows.sort(key=get_date_obj)

        needed_rows = max(102, 2 + len(bank_rows))  # Ensure minimum 102 rows for director sheet
        table.setRowCount(needed_rows)

        # Write data rows starting from row 2
        row_idx = 2
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
                    # Mark bank data with green background for director sheet
                    if subject_filter == "董事往来":
                        item.setData(Qt.BackgroundRole, QColor(200, 255, 200))
                    table.setItem(row_idx, col, item)
                elif amount_col_start <= col < amount_col_start + amount_col_count:
                    cur = currency_list[col - amount_col_start]
                    if cur == currency:
                        # Get DEBIT value from bank sheet (column 4) and write to CREDIT column in sales sheet
                        amount_item = sheet.item(row, 5 if is_debit else 4)  # Debit column in bank sheet
                        value = amount_item.text() if amount_item else ""
                        logger.info(f"DEBUG: Currency match! cur={cur}, currency={currency}, debit_value='{value}', sheet={sheet.name}, row={row}")
                        item = QTableWidgetItem(value)
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                        # Mark bank data with green background for director sheet
                        if subject_filter == "董事往来":
                            item.setData(Qt.BackgroundRole, QColor(200, 255, 200))
                        table.setItem(row_idx, col, item)
                    else:
                        logger.info(f"DEBUG: Currency mismatch! cur={cur}, currency={currency}")
                        item = QTableWidgetItem("")
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                        # Mark bank data with green background for director sheet
                        if subject_filter == "董事往来":
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
                    # Mark bank data with green background for director sheet
                    if subject_filter == "董事往来":
                        item.setData(Qt.BackgroundRole, QColor(200, 255, 200))
                    table.setItem(row_idx, col, item)
            row_idx += 1

    def refresh_aggregate_sheet(self, subject_filter, column_title):
        """Refresh the specified aggregate sheet"""
        current_tab = self.main_window.tabs.currentWidget()
        if current_tab:
            # Special handling for sales sheet refresh
            if subject_filter in ["销售收入", "销售成本", "银行费用", "利息收入", "董事往来"]:
                # Check if table has correct multi-currency structure
                is_debit = subject_filter in ["销售成本", "银行费用", "董事往来"]
                currency_set = set()
                for sheet in self.main_window.sheets:
                    if sheet.type == "bank":
                        print(f"name: {sheet.name} sheet.currency {sheet.currency}")
                        currency_set.add(sheet.currency)

                currency_list, credit_col_start, credit_col_count = self._setup_currency_sheet_structure(current_tab, column_title)
                currency_list = sorted(currency_set)
                credit_col_start = 5
                credit_col_count = len(currency_list)

                user_data = []
                if subject_filter == "董事往來":
                    # Preserve user data with original row positions for director sheet
                    user_data = self.preserve_director_user_data_with_positions(current_tab)
                    print(f"user data {user_data}")

                # Always populate data (either after structure setup or just refresh)
                self._populate_currency_sheet_data(current_tab, subject_filter, currency_list, credit_col_start, credit_col_count, is_debit)

                # Restore user data after refresh
                if subject_filter == "董事往來":
                    # Restore user data to original positions
                    self.restore_director_user_data_to_positions(current_tab, user_data)
                return

    def reorder_sheets(self, from_index, to_index):
        """Handle tab reordering to keep sheets list in sync"""
        self.main_window.sheets.insert(to_index, self.main_window.sheets.pop(from_index))
        self.main_window.auto_save()

    def preserve_director_user_data(self, table):
        """Preserve user data in director sheet before refresh"""
        if not hasattr(table, 'user_added_rows') or not table.user_added_rows:
            return []

        user_data = []
        for row in table.user_added_rows:
            if row < table.rowCount():
                row_data = []
                has_data = False
                for col in range(table.columnCount()):
                    item = table.item(row, col)
                    text = item.text() if item else ""
                    row_data.append(text)
                    if text.strip():
                        has_data = True

                # Only preserve rows that have actual data
                if has_data:
                    user_data.append((row, row_data))

        return user_data

    def restore_director_user_data(self, table, user_data, bank_data_end=0):
        """Restore user data in director sheet after refresh"""
        if not user_data:
            return

        table.user_added_rows = set()

        # Position user data after bank data
        for i, (original_row, row_data) in enumerate(user_data):
            new_row = bank_data_end + i

            # Ensure we have enough rows
            if new_row >= table.rowCount():
                table.setRowCount(new_row + 50)

            table.user_added_rows.add(new_row)

            # Restore the data - handle both old and new column structures
            for col, text in enumerate(row_data):
                if col < table.columnCount() and text:  # Only restore non-empty cells
                    item = QTableWidgetItem(text)
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
                    table.setItem(new_row, col, item)

    def preserve_director_user_data_with_positions(self, table):
        user_data = []
        for row in table.user_added_rows:
            if row < table.rowCount():
                row_data = []
                has_data = False
                for col in range(table.columnCount()):
                    item = table.item(row, col)
                    text = item.text() if item else ""
                    if text.strip():
                        has_data = True
                        row_data.append(text)

                # Only preserve rows that have actual data, save with original row position
                if has_data:
                    user_data.append((row, row_data))

        return user_data

    def restore_director_user_data_to_positions(self, table, user_data):
        """Restore user data to their original row positions in director sheet"""
        if not user_data:
            return

        # Find the maximum row number needed
        max_row_needed = 0
        for original_row, row_data in user_data:
            max_row_needed = max(max_row_needed, original_row)

        # Ensure table has enough rows
        if max_row_needed >= table.rowCount():
            table.setRowCount(max(102, max_row_needed + 50))  # Ensure minimum 102 rows

        # Restore data to original positions
        for original_row, row_data in user_data:
            table.user_added_rows.add(original_row)

            # Restore the data to the exact original row position
            for col, text in enumerate(row_data):
                if col < table.columnCount() and text:  # Only restore non-empty cells
                    item = QTableWidgetItem(text)
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
                    table.setItem(original_row, col, item)

    def scan_director_user_data(self, table):
        """Scan director sheet to identify user-added rows automatically"""
        if not hasattr(table, 'user_added_rows'):
            table.user_added_rows = set()

        # Find the end of bank-generated data (rows with green background)
        bank_data_end = 2  # Start after headers
        for row in range(2, table.rowCount()):
            item = table.item(row, 0)
            if item and item.data(Qt.BackgroundRole):
                # This is bank data (has background color)
                bank_data_end = row + 1
            elif item and item.text().strip():
                # This is user data (no background color but has content)
                table.user_added_rows.add(row)
            elif not (item and item.text().strip()):
                # Empty row, check if any column has data
                has_data = False
                for col in range(table.columnCount()):
                    cell_item = table.item(row, col)
                    if cell_item and cell_item.text().strip():
                        has_data = True
                        break
                if has_data:
                    table.user_added_rows.add(row)

