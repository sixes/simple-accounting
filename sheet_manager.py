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
        columns = ["序號", "日期", "對方科目", "摘要", "借方", "貸方", "餘額", "發票號碼"]
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

    def create_non_bank_sheet(self, name="非银行交易"):
        """Create a regular sheet"""
        columns = ["序 號", "日  期", "對方科目", "子科目", "摘   要", "借方(USD)", "借方(EUR)", "借方(JPY)", "借方(GBP)", "借方(CHF)", "借方(CAD)", "借方(AUD)", "借方(CNY)", "借方(HKD)", "借方(NZD)", "贷方(USD)", "贷方(EUR)", "贷方(JPY)", "贷方(GBP)", "贷方(CHF)", "贷方(CAD)", "贷方(AUD)", "贷方(CNY)", "贷方(HKD)", "贷方(NZD)", "备注"]
        table = ExcelTable(auto_save_callback=self.main_window.auto_save, name=name, type="non_bank")
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)

        self.main_window.tabs.addTab(table, name)
        self.main_window.sheets.append(table)
        return table
    
    def create_payable_detail_sheet(self, sheet_name):
        """Create a payable sheet with exactly the same header structure as the sales sheet"""
        # Use the same main and sub headers, merged ranges, and formatting as sales sheet
        main_headers = [
            "序 號", "日  期", "對方科目", "子科目", "發票號碼",
            "借方(USD)", "借方(EUR)", "借方(JPY)", "借方(GBP)", "借方(CHF)", "借方(CAD)", "借方(AUD)", "借方(CNY)", "借方(HKD)", "借方(NZD)",
            "贷方(USD)", "贷方(EUR)", "贷方(JPY)", "贷方(GBP)", "贷方(CHF)", "贷方(CAD)", "贷方(AUD)", "贷方(CNY)", "贷方(HKD)", "贷方(NZD)",
            "餘額", "摘  要", "來源"
        ]
        table = ExcelTable("payable_detail", auto_save_callback=self.main_window.auto_save, name=sheet_name)
        table.setColumnCount(len(main_headers))
        table.setHorizontalHeaderLabels(columns)
        table.setRowCount(300)
        self.main_window.sheets.append(table)
        self.main_window.sheets.append(table)
        return table

    def reorder_sheets(self, from_index, to_index):
        """Handle tab reordering to keep sheets list in sync"""
        self.main_window.sheets.insert(to_index, self.main_window.sheets.pop(from_index))
        self.main_window.auto_save()