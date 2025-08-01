from PySide6.QtWidgets import QDoubleSpinBox, QTableWidgetItem, QMenu
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QColor
from excel_table import ExcelTable
from datetime import datetime

class SheetManager:
    def __init__(self, main_window):
        self.main_window = main_window

    def create_bank_sheet(self, name):
        """Create a bank sheet with exchange rate control"""
        columns = ["序 號", "日  期", "對方科目", "摘   要", "借     方", "貸     方", "餘    額", "發票號碼"]
        table = ExcelTable(auto_save_callback=self.main_window.auto_save, name=name)
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)

        # Set the load/save callbacks that ExcelTable's context menu should use
        def debug_load_wrapper():
            print("DEBUG WRAPPER: load_callback called!")
            return self.main_window.file_manager.load_file()
            
        def debug_save_wrapper():
            print("DEBUG WRAPPER: save_callback called!")
            return self.main_window.file_manager.save_file()
            
        table.load_callback = debug_load_wrapper
        table.save_callback = debug_save_wrapper
        print(f"DEBUG: Set callbacks on table {name}")
        
        # Try to find and override ExcelTable's actual context menu actions
        print(f"DEBUG: Table attributes: {[attr for attr in dir(table) if 'load' in attr.lower() or 'save' in attr.lower() or 'action' in attr.lower()]}")
        
        # Override the contextMenuEvent method to intercept context menu
        original_context_menu = table.contextMenuEvent
        def custom_context_menu(event):
            print("DEBUG: Custom context menu called!")
            # Call original to get the menu
            original_context_menu(event)
            
        table.contextMenuEvent = custom_context_menu

        # Add exchange rate control
        rate_input = QDoubleSpinBox()
        currency = name.split("-")[1] if "-" in name else ""
        rate_input.setPrefix(f"{currency}:HKD = 1:")
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
        print(f"DEBUG CREATE: Creating regular sheet '{name}'")
        columns = ["序 號", "日  期", "對方科目", "摘   要", "借     方", "貸     方", "借或貸", "餘    額", "發票號碼"]
        table = ExcelTable(auto_save_callback=self.main_window.auto_save, name=name)
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)
        
        # Set the load/save callbacks that ExcelTable's context menu should use
        table.load_callback = self.main_window.file_manager.load_file
        table.save_callback = self.main_window.file_manager.save_file
        
        self.main_window.tabs.addTab(table, name)
        self.main_window.sheets.append(table)
        print(f"DEBUG CREATE: Regular sheet '{name}' created, total sheets: {len(self.main_window.sheets)}, total tabs: {self.main_window.tabs.count()}")
        return table

    def create_aggregate_sheet(self, sheet_name, subject_filter, column_title):
        """Common method to create aggregate sheets (sales/cost)"""
        columns = ["序 號", "日  期", "對方科目", "摘  要", "發票號碼", column_title, "餘    額", "來源"]
        table = ExcelTable(auto_save_callback=self.main_window.auto_save, name=sheet_name)
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)

        # Set the load/save callbacks that ExcelTable's context menu should use
        table.load_callback = self.main_window.file_manager.load_file
        table.save_callback = self.main_window.file_manager.save_file

        # Add pinned rows for totals
        table.setRowCount(100 + 2)  # Regular rows + 2 pinned rows
        if sheet_name == "董事往來":
            table.user_added_rows = set()

        # Populate data from bank sheets
        self.populate_aggregate_data(table, subject_filter, column_title)

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
        return self.create_aggregate_sheet("董事往來", "董事往来", "貸     方")

    def create_payable_sheet(self):
        """Create a payable sheet"""
        return self.create_aggregate_sheet("應付費用", "董事往来", "貸     方")

    def create_interest_sheet(self):
        """Create an interest income sheet"""
        return self.create_aggregate_sheet("利息收入", "利息收入", "貸     方")

    def populate_aggregate_data(self, table, subject_filter, amount_column_title):
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
        for sheet in self.main_window.sheets:
            if "-" in sheet.name:  # Bank sheet
                for row in range(sheet.rowCount() - 2):  # Exclude pinned rows
                    subject_item = sheet.item(row, 2)  # 對方科目 column
                    if subject_item and subject_filter in subject_item.text():
                        date_item = sheet.item(row, 1)  # 日期
                        desc_item = sheet.item(row, 3)  # 摘要
                        invoice_item = sheet.item(row, 7)  # 發票號碼
                        amount_item = sheet.item(row, 4 if "借" in amount_column_title else 5)  # Debit/Credit column

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

    def refresh_aggregate_sheet(self, subject_filter, column_title):
        """Refresh the specified aggregate sheet"""
        current_tab = self.main_window.tabs.currentWidget()
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
            print(f"sheets: {self.main_window.sheets}")
            for sheet in self.main_window.sheets:
                if "-" in sheet.name:  # Bank sheet
                    print(f"looking at sheet: {sheet.name} subject_filter: {subject_filter} column_title: {column_title}")
                    for row in range(sheet.rowCount() - 2):  # Exclude pinned rows
                        subject_item = sheet.item(row, 2)  # 對方科目 column
                        if subject_item and subject_filter in subject_item.text():
                            date_item = sheet.item(row, 1)  # 日期
                            desc_item = sheet.item(row, 3)  # 摘要
                            invoice_item = sheet.item(row, 7)  # 發票號碼
                            amount_item = sheet.item(row, 4 if "借" in column_title else 5)  # Debit/Credit column

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
                start_row = len(data_rows)
                current_tab.user_added_rows = set()
                for i, row_data in enumerate(user_data):
                    row = start_row + i
                    current_tab.user_added_rows.add(row)
                    for col, text in enumerate(row_data):
                        item = QTableWidgetItem(text)
                        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
                        current_tab.setItem(row, col, item)

    def reorder_sheets(self, from_index, to_index):
        """Handle tab reordering to keep sheets list in sync"""
        self.main_window.sheets.insert(to_index, self.main_window.sheets.pop(from_index))
        self.main_window.auto_save()