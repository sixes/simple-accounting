import pickle
import logging
from PySide6.QtWidgets import (
    QTableWidget, QTableWidgetItem, QMenu, QApplication
)
from PySide6.QtGui import QColor
from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QKeySequence
from PySide6.QtGui import QPainter, QFont


logger = logging.getLogger(__name__)


def excel_column_name(n):
    name = ""
    while n >= 0:
        name = chr(n % 26 + 65) + name
        n = n // 26 - 1
    return name

class ExcelTable(QTableWidget):
    def __init__(self, type, rows=100, cols=20, name="", auto_save_callback=None):
        super().__init__(rows, cols)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.context_menu)
        self.copied_range = None
        self.name = name
        assert type != None
        self.type = type
        self.currency = name.split("-")[1] if "-" in self.name else ""
        self.auto_save_callback = auto_save_callback
        self._custom_headers = None  # Track custom headers
        self.update_headers()
        self.horizontalHeader().setStretchLastSection(True)
        self.verticalHeader().setDefaultSectionSize(24)
        self.horizontalHeader().setDefaultSectionSize(80)
        self.setSizeAdjustPolicy(QTableWidget.AdjustToContents)
        self.itemChanged.connect(self._on_item_changed)
        self.cellChanged.connect(self._on_cell_changed)
        self.user_added_rows = set()  # Track user-added rows
        self._last_paint_pos = -1
        # Enable smooth scrolling and proper updates
        self.setVerticalScrollMode(QTableWidget.ScrollPerPixel)
        self.viewport().setAttribute(Qt.WA_OpaquePaintEvent, False)
        
        # Connect scroll events to update viewport - fixes Windows duplicate pinned rows issue
        self.verticalScrollBar().valueChanged.connect(self._on_scroll)
        self.horizontalScrollBar().valueChanged.connect(self._on_scroll)
        if False:
            self.setStyleSheet("""
                QTableWidget {
                    background: #ffffff;
                    gridline-color: #e0e0e0;
                    selection-background-color: #cce2ff;
                    selection-color: #000;
                    font-size: 14px;
                }
                QHeaderView::section {
                    background: #e9ecef;
                    color: #222;
                    border: 1px solid #e0e0e0;
                    font-weight: bold;
                    font-size: 15px;
                    padding: 4px;
                }
                QTableWidget::item {
                    border: 0.5px solid #e0e0e0;
                    font-size: 14px;
                    background: #ffffff;
                    color: #000;
                }
                QTableWidget::item:selected {
                    border: 1.5px solid #0078d4;
                    background: #cce2ff;
                    color: #000;
                }
            """)

    def paintEvent(self, event):
        # First call the parent paintEvent to draw the table contents
        super().paintEvent(event)

        # Always paint pinned rows - remove optimization that causes Windows issues
        painter = QPainter(self.viewport())
        try:
            viewport = self.viewport()
            visible_rect = viewport.rect()
            row_height = self.rowHeight(0)
            col_count = self.columnCount()

            # Calculate positions for pinned rows (always at bottom)
            y1 = visible_rect.height() - 2 * row_height
            y2 = visible_rect.height() - row_height

            # Clear the area where pinned rows will be drawn - more robust clearing for Windows
            painter.setCompositionMode(QPainter.CompositionMode_Source)
            # Use white background instead of palette color for better Windows compatibility
            painter.fillRect(0, y1, visible_rect.width(), 2 * row_height, QColor(255, 255, 255))
            painter.setCompositionMode(QPainter.CompositionMode_SourceOver)
            
            # Add antialiasing for better rendering on Windows
            painter.setRenderHint(QPainter.Antialiasing, False)

            # Calculate sums and balances
            debit_sum, credit_sum = self.sum_columns()
            balance = debit_sum - credit_sum
            rate = getattr(self, "exchange_rate", 1.0)
            hkd_balance = balance * rate

            # Find column positions
            debit_col = None
            credit_col = None
            balance_col = None
            for col in range(col_count):
                header = self.horizontalHeaderItem(col).text().replace(" ", "")
                if "借方" in header:
                    debit_col = col
                if "貸方" in header:
                    credit_col = col
                if "餘額" in header:
                    balance_col = col

            # Draw first pinned row (sheet currency)
            painter.fillRect(0, y1, visible_rect.width(), row_height, QColor(240, 240, 240))
            font = painter.font()
            font.setBold(True)
            painter.setFont(font)

            for col in range(col_count):
                x = self.columnViewportPosition(col)
                w = self.columnWidth(col)
                painter.drawRect(x, y1, w, row_height)

                text = ""
                if col == debit_col:
                    text = self._format_number(debit_sum)
                elif col == credit_col:
                    text = self._format_number(credit_sum)
                elif col == balance_col:
                    text = self._format_number(balance)

                painter.drawText(x + 6, y1 + row_height//2 + 5, text)

            # Draw second pinned row (HKD)
            painter.fillRect(0, y2, visible_rect.width(), row_height, QColor(220, 220, 220))
            for col in range(col_count):
                x = self.columnViewportPosition(col)
                w = self.columnWidth(col)
                painter.drawRect(x, y2, w, row_height)

                text = ""
                if col == debit_col:
                    text = self._format_number(debit_sum * rate)
                elif col == credit_col:
                    text = self._format_number(credit_sum * rate)
                elif col == balance_col:
                    text = self._format_number(hkd_balance)

                painter.drawText(x + 6, y2 + row_height//2 + 5, text)

        finally:
            painter.end()

    def _format_number(self, value):
        """Helper method to format numbers consistently"""
        abs_value = abs(value)
        formatted = f"{abs_value:,.2f}" if abs_value >= 1000 else f"{abs_value:.2f}"
        return f"({formatted})" if value < 0 else formatted

    def resizeEvent(self, event):
        """Handle resize events to ensure proper viewport updates on Windows"""
        super().resizeEvent(event)
        # Force viewport update after resize to prevent rendering artifacts
        self.viewport().update()

    def showEvent(self, event):
        """Handle show events to ensure proper initial painting on Windows"""
        super().showEvent(event)
        # Force initial viewport update when widget becomes visible
        self.viewport().update()

    def _on_scroll(self):
        """Handle scroll events to properly update viewport on Windows"""
        # Force viewport update to prevent duplicate pinned rows on Windows
        self.viewport().update()

    def _on_cell_changed(self, row, column):
        """Track user edits by adding row to user_added_rows"""
        if row not in self.user_added_rows:
            item = self.item(row, column)
            if item and (item.flags() & Qt.ItemIsEditable):
                self.user_added_rows.add(row)
                #print(f"add user data:{row}")

    def _on_item_changed(self, item):
        # Only recalculate if debit, credit, or balance in first row changes
        balance_col = None
        debit_col = None
        credit_col = None
        for col in range(self.columnCount()):
            header = self.horizontalHeaderItem(col).text().replace(" ", "")
            if "餘額" in header:
                balance_col = col
            if "借方" in header:
                debit_col = col
            if "貸方" in header:
                credit_col = col
        if balance_col is None or debit_col is None or credit_col is None:
            return
        row = item.row()
        col = item.column()
        # Only allow editing balance in first row
        if col == balance_col and row > 0:
            self.blockSignals(True)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            self.blockSignals(False)
            return
        # Recalculate balances for all rows except the first
        if col in (debit_col, credit_col, balance_col) or (row == 0 and col == balance_col):
            self.blockSignals(True)
            for r in range(1, self.rowCount()):
                prev_balance = self.item(r-1, balance_col)
                debit = self.item(r, debit_col)
                credit = self.item(r, credit_col)
                try:
                    prev_text = prev_balance.text().replace(',', '') if prev_balance and prev_balance.text() else '0'
                    prev_val = float(prev_text) if prev_text else 0.0
                except Exception:
                    prev_val = 0.0
                try:
                    debit_text = debit.text().replace(',', '') if debit and debit.text() else '0'
                    debit_val = float(debit_text) if debit_text else 0.0
                except Exception:
                    debit_val = 0.0
                try:
                    credit_text = credit.text().replace(',', '') if credit and credit.text() else '0'
                    credit_val = float(credit_text) if credit_text else 0.0
                except Exception:
                    credit_val = 0.0
                bal = prev_val + debit_val - credit_val
                bal_item = self.item(r, balance_col)
                if not bal_item:
                    bal_item = QTableWidgetItem()
                    self.setItem(r, balance_col, bal_item)
                # Format balance: no minus sign, comma separation for >1000
                abs_bal = abs(bal)
                if abs_bal >= 1000:
                    bal_text = f"{abs_bal:,.2f}"
                else:
                    bal_text = f"{abs_bal:.2f}"
                bal_item.setText(bal_text)
                bal_item.setFlags(bal_item.flags() & ~Qt.ItemIsEditable)
            self.blockSignals(False)
        self.update_pinned_rows()
        self._auto_save()

    def update_pinned_rows(self):
        self.blockSignals(True)
        debit_sum, credit_sum = self.sum_columns()
        balance = debit_sum - credit_sum
        rate = getattr(self, "exchange_rate", 1.0)
        hkd_balance = balance * rate
        balance_col = None
        debit_col = None
        credit_col = None
        for col in range(self.columnCount()):
            header = self.horizontalHeaderItem(col).text().replace(" ", "")
            if "餘額" in header:
                balance_col = col
            if "借方" in header:
                debit_col = col
            if "貸方" in header:
                credit_col = col
        last_row = self.rowCount() - 2
        last_row2 = self.rowCount() - 1
        def format_number(value):
            abs_value = abs(value)
            formatted = f"{abs_value:,.2f}" if abs_value >= 1000 else f"{abs_value:.2f}"
            return f"({formatted})" if value < 0 else formatted
        # Sheet currency row
        for col in range(self.columnCount()):
            item = self.item(last_row, col)
            if not item:
                item = QTableWidgetItem()
                self.setItem(last_row, col, item)
            # Only set flags if not already read-only
            if item.flags() & Qt.ItemIsEditable:
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            item.setBackground(QColor(240, 240, 240))
            if col == debit_col:
                item.setText(format_number(debit_sum))
            elif col == credit_col:
                item.setText(format_number(credit_sum))
            elif col == balance_col:
                item.setText(format_number(balance))
            else:
                item.setText("")
        # HKD row
        for col in range(self.columnCount()):
            item = self.item(last_row2, col)
            if not item:
                item = QTableWidgetItem()
                self.setItem(last_row2, col, item)
            if item.flags() & Qt.ItemIsEditable:
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            item.setBackground(QColor(220, 220, 220))
            if col == debit_col:
                item.setText(format_number(debit_sum * rate))
            elif col == credit_col:
                item.setText(format_number(credit_sum * rate))
            elif col == balance_col:
                item.setText(format_number(hkd_balance))
            else:
                item.setText("")
        self.blockSignals(False)

    def setHorizontalHeaderLabels(self, labels):
        super().setHorizontalHeaderLabels(labels)
        self._custom_headers = list(labels)
        # Set only the first row's balance cell editable, others not
        balance_col = None
        for col, label in enumerate(labels):
            if "餘" in label:
                balance_col = col
                break
        if balance_col is not None:
            for r in range(self.rowCount()):
                item = self.item(r, balance_col)
                if not item:
                    item = QTableWidgetItem()
                    self.setItem(r, balance_col, item)
                if r == 0:
                    item.setFlags(item.flags() | Qt.ItemIsEditable)
                else:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)

    def update_headers(self):
        if self._custom_headers:
            for col, label in enumerate(self._custom_headers):
                if col < self.columnCount():
                    self.setHorizontalHeaderItem(col, QTableWidgetItem(label))
        else:
            self.setColumnCount(self.columnCount())
            for col in range(self.columnCount()):
                self.setHorizontalHeaderItem(col, QTableWidgetItem(excel_column_name(col)))

    def insertColumn(self, col):
        super().insertColumn(col)
        if self._custom_headers:
            # Use Excel-style column name for the new column
            self._custom_headers.insert(col, excel_column_name(col))
            self.setHorizontalHeaderLabels(self._custom_headers)
        else:
            self.update_headers()
        self._auto_save()

    def removeColumn(self, col):
        super().removeColumn(col)
        if self._custom_headers and col < len(self._custom_headers):
            del self._custom_headers[col]
            self.setHorizontalHeaderLabels(self._custom_headers)
        else:
            self.update_headers()
        self._auto_save()

    def insertRow(self, row):
        # For aggregate sheets, prevent inserting between title rows (0 and 1)
        if hasattr(self, 'name') and self.name in ["銷售收入", "銷售成本", "銀行費用", "利息收入", "應付費用", "董事往來"]:
            if row <= 1:  # Can't insert between or before title rows
                row = 2  # Insert after title rows instead
        super().insertRow(row)
        self._auto_save()

    def removeRow(self, row):
        super().removeRow(row)
        self._auto_save()

    def context_menu(self, pos):
        menu = QMenu(self)

        # Table operations
        add_row = QAction("Add Row", self)
        add_col = QAction("Add Column", self)
        del_row = QAction("Delete Row", self)
        del_col = QAction("Delete Column", self)
        copy = QAction("Copy", self)
        paste = QAction("Paste", self)
        merge = QAction("Merge Cells", self)
        split = QAction("Unmerge Cells", self)

        menu.addAction(add_row)
        menu.addAction(add_col)
        menu.addAction(del_row)
        menu.addAction(del_col)
        menu.addSeparator()
        menu.addAction(copy)
        menu.addAction(paste)
        menu.addSeparator()
        menu.addAction(merge)
        menu.addAction(split)
        menu.addSeparator()

        # File operations
        new_file = QAction("New", self)
        add_sheet = QAction("Add Sheet", self)
        rename_sheet = QAction("Rename Sheet", self)  # ADDED THIS LINE
        delete_sheet = QAction("Delete Sheet", self)
        save_file = QAction("Save", self)
        load_file = QAction("Load", self)

        menu.addAction(new_file)
        menu.addAction(add_sheet)
        menu.addAction(rename_sheet)  # ADDED THIS LINE
        menu.addAction(delete_sheet)
        menu.addSeparator()
        menu.addAction(save_file)
        menu.addAction(load_file)

        # Connect table actions
        add_row.triggered.connect(lambda: self.insertRow(self.currentRow() + 1))
        add_col.triggered.connect(lambda: self.insertColumn(self.currentColumn() + 1))
        del_row.triggered.connect(lambda: self.removeRow(self.currentRow()))
        del_col.triggered.connect(lambda: self.removeColumn(self.currentColumn()))
        copy.triggered.connect(self.copy_cells)
        paste.triggered.connect(self.paste_cells)
        merge.triggered.connect(self.merge_cells)
        split.triggered.connect(self.unmerge_cells)
        rename_sheet.triggered.connect(self.rename_sheet)

        # Connect file actions to parent window
        parent_window = self.window()
        if hasattr(parent_window, 'new_file'):
            new_file.triggered.connect(parent_window.new_file)
        if hasattr(parent_window, 'add_sheet_dialog'):
            add_sheet.triggered.connect(parent_window.add_sheet_dialog)
        if hasattr(parent_window, 'delete_sheet'):
            delete_sheet.triggered.connect(parent_window.delete_sheet)
        if hasattr(parent_window, 'save_file'):
            save_file.triggered.connect(parent_window.save_file)
        if hasattr(parent_window, 'load_file'):
            load_file.triggered.connect(parent_window.load_file)

        menu.exec(self.viewport().mapToGlobal(pos))

    def rename_sheet(self):
        """Rename the current sheet"""
        from PySide6.QtWidgets import QInputDialog
        new_name, ok = QInputDialog.getText(
            self,
            "Rename Sheet",
            "Enter new sheet name:",
            text=self.name
        )
        if ok and new_name and new_name != self.name:
            old_name = self.name
            self.name = new_name
            if "-" in new_name:
                self.currency = new_name.split("-")[1]
            else:
                self.currency = ""

            # Update tab name if parent has update_tab_name method
            if hasattr(self.window(), 'update_tab_name'):
                self.window().update_tab_name(old_name, new_name)

            # Force save and refresh
            self._auto_save()
            self.viewport().update()

    def copy_cells(self):
        sel = self.selectedRanges()
        if sel:
            r = sel[0]
            self.copied_range = [
                [self.item(r.topRow() + i, r.leftColumn() + j).text() if self.item(r.topRow() + i, r.leftColumn() + j) else ""
                 for j in range(r.columnCount())]
                for i in range(r.rowCount())
            ]

    def paste_cells(self):
        clipboard = QApplication.clipboard()
        text = clipboard.text()
        if not text:
            return
        rows = text.split('\n')
        start_row = self.currentRow()
        start_col = self.currentColumn()
        for i, row_data in enumerate(rows):
            if not row_data.strip():
                continue
            columns = row_data.split('\t')
            for j, val in enumerate(columns):
                r = start_row + i
                c = start_col + j
                if r < self.rowCount() and c < self.columnCount():
                    # For director sheet, check if target cell is bank data
                    if hasattr(self, 'name') and self.name == "董事往來":
                        item = self.item(r, c)
                        # Skip this cell if it's bank data (has green background)
                        if item and item.data(Qt.BackgroundRole):
                            continue

                    # Create new item or get existing
                    new_item = QTableWidgetItem(val)

                    # If it's user data in director sheet, mark the row
                    if hasattr(self, 'name') and self.name == "董事往來":
                        if r > 1:  # Skip header rows
                            self.user_added_rows.add(r)

                    self.setItem(r, c, new_item)

    def merge_cells(self):
        sel = self.selectedRanges()
        if sel:
            r = sel[0]
            self.setSpan(r.topRow(), r.leftColumn(), r.rowCount(), r.columnCount())

    def unmerge_cells(self):
        sel = self.selectedRanges()
        if sel:
            r = sel[0]
            self.setSpan(r.topRow(), r.leftColumn(), 1, 1)

    def data(self):
        cells = {}
        for row in range(self.rowCount()):
            for col in range(self.columnCount()):
                item = self.item(row, col)
                if item and item.text().strip():  # Only save non-empty cells
                    # For director sheet, don't save bank-generated data (cells with green background)
                    if hasattr(self, 'name') and self.name == "董事往來":
                        # Only save cells that don't have green background (user data)
                        if not item.data(Qt.BackgroundRole) or item.data(Qt.BackgroundRole) != QColor(200, 255, 200):
                            cells[(row, col)] = item.text()
                    else:
                        # For other sheets, save all data normally
                        cells[(row, col)] = item.text()
        spans = []
        for row in range(self.rowCount()):
            for col in range(self.columnCount()):
                rs, cs = self.rowSpan(row, col), self.columnSpan(row, col)
                if rs > 1 or cs > 1:
                    spans.append((row, col, rs, cs))
        # Always return data structure even if empty
        return {"cells": cells, "spans": spans, "rows": self.rowCount(), "cols": self.columnCount(), "name": self.name}

    def load_data(self, data):
        """Load data into the table with error handling"""
        try:
            # Set row and column counts only if they exist
            if "rows" in data:
                self.setRowCount(data["rows"])
            if "cols" in data:
                self.setColumnCount(data["cols"])

            if "name" in data:
                self.name = data["name"]
                if "-" in self.name:
                    self.currency = self.name.split("-")[1]
                else:
                    self.currency = ""

            # Restore custom headers if they exist
            if "headers" in data and data["headers"]:
                self._custom_headers = data["headers"]
                self.setHorizontalHeaderLabels(self._custom_headers)
            else:
                self.update_headers()

            # Load cell data if it exists
            if "cells" in data:
                for (row, col), text in data["cells"].items():
                    self.setItem(row, col, QTableWidgetItem(text))

            # Load cell spans if they exist
            if "spans" in data:
                for row, col, rs, cs in data["spans"]:
                    if rs > 1 or cs > 1:
                        self.setSpan(row, col, rs, cs)

        except Exception as e:
            logger.error(f"ERROR LOAD DATA: Failed to load data: {e}")
            import traceback
            traceback.print_exc()

    def keyPressEvent(self, event):
        # Handle Ctrl+V for paste
        if event.key() == Qt.Key_Delete:
            for item in self.selectedItems():
                # For director sheet, also check for background color (bank data)
                if hasattr(self, 'name') and self.name == "董事往來":
                    if item.flags() & Qt.ItemIsEditable and not item.data(Qt.BackgroundRole):
                        item.setText("")
                else:
                    if item.flags() & Qt.ItemIsEditable:
                        item.setText("")
            if self.auto_save_callback:
                self.auto_save_callback()
            return
        if event.matches(QKeySequence.Paste):
            # For director sheet, prevent pasting into bank data rows
            if hasattr(self, 'name') and self.name == "董事往來":
                # Check if any selected cell is bank data
                for item in self.selectedItems():
                    if item and item.data(Qt.BackgroundRole):
                        return  # Abort paste if trying to paste into bank data
            self.paste_cells()
            return
        super().keyPressEvent(event)

    def set_exchange_rate(self, rate):
        self.exchange_rate = rate
        self.viewport().update()

    def sum_columns(self):
        debit_sum = 0.0
        credit_sum = 0.0
        debit_col = None
        credit_col = None

        for col in range(self.columnCount()):
            header = self.horizontalHeaderItem(col).text().replace(" ", "")
            if "借方" in header:
                debit_col = col
            if "貸方" in header:
                credit_col = col
        # Calculate totals (exclude last two pinned rows)
        for row in range(self.rowCount() - 2):
            if credit_col is not None:
                credit_item = self.item(row, credit_col)
                try:
                    credit_text = credit_item.text().replace(',', '') if credit_item and credit_item.text() else '0'
                    credit_sum += float(credit_text) if credit_text else 0.0
                except Exception:
                    pass
            if debit_col is not None:
                debit_item = self.item(row, debit_col)
                try:
                    debit_text = debit_item.text().replace(',', '') if debit_item and debit_item.text() else '0'
                    debit_sum += float(debit_text) if debit_text else 0.0
                except Exception:
                    pass

        return debit_sum, credit_sum

    def _auto_save(self, *_):
        if self.auto_save_callback:
            self.auto_save_callback()
