import logging
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QAction, QColor, QKeySequence, QPainter
from PySide6.QtWidgets import QApplication, QMenu, QTableWidget, QTableWidgetItem


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
        self.name = name
        self.type = type
        self.currency = name.split("-")[1] if "-" in self.name else ""
        self.auto_save_callback = auto_save_callback
        self._custom_headers = None  # Track custom headers
        
        # For aggregate sheets, set up 2-row horizontal header
        if self.type == "aggregate":
            self.horizontalHeader().setMinimumSectionSize(60)
            self.horizontalHeader().setDefaultSectionSize(80)
            # We'll set this up in setup_two_row_headers method
            # Store pinned height for scroll limiting
            self._pinned_height = 48  # 2 rows * 24px
        else:
            self.update_headers()
            # For bank sheets, set up bottom margins for pinned rows
            row_height = 24  # Default row height
            pinned_height = row_height * 2
            self.setViewportMargins(0, 0, 0, pinned_height)
            
            # Store pinned height for scroll limiting
            self._pinned_height = pinned_height
            
        self.horizontalHeader().setStretchLastSection(True)
        self.verticalHeader().setDefaultSectionSize(24)
        self.horizontalHeader().setDefaultSectionSize(80)
        self.setSizeAdjustPolicy(QTableWidget.AdjustToContents)
        self.itemChanged.connect(self._on_item_changed)
        self.user_added_rows = set()  # Track user-added rows
        self._last_paint_pos = -1
        # Enable smooth scrolling and proper updates
        self.setVerticalScrollMode(QTableWidget.ScrollPerPixel)
        self.viewport().setAttribute(Qt.WA_OpaquePaintEvent, False)
        
        # Connect scroll events to update viewport - fixes Windows duplicate pinned rows issue
        self.verticalScrollBar().valueChanged.connect(self._on_scroll)
        self.horizontalScrollBar().valueChanged.connect(self._on_scroll)
        
        # Set up a timer to constantly enforce scroll limits
#        self._scroll_limit_timer = QTimer()
#        self._scroll_limit_timer.timeout.connect(self._enforce_scroll_limits)
#        self._scroll_limit_timer.start(100)  # Check every 100ms (reduced frequency for less spam)


    def paintEvent(self, event):
        # First call the parent paintEvent to draw the table contents
        super().paintEvent(event)

        painter = QPainter(self.viewport())
        try:
            viewport = self.viewport()
            visible_rect = viewport.rect()
            row_height = self.rowHeight(0)
            col_count = self.columnCount()

            # Paint frozen rows if they exist (for aggregate sheets with 2-row headers)
            if hasattr(self, '_frozen_row_count') and self._frozen_row_count > 0:
                self._paint_frozen_rows(painter, visible_rect)

            # Always paint pinned rows at the absolute bottom of the viewport
            # The viewport margins ensure these don't overlap with scrollable content
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
            
            # Get currency-specific sums for aggregate sheets
            currency_sums = self.sum_currency_columns()

            # Find column positions - handle both regular and multi-currency sheets
            debit_col = None
            credit_col = None
            balance_col = None
            currency_cols = []
            
            # Check if this is a multi-currency aggregate sheet
            is_aggregate_sheet = (hasattr(self, 'name') and 
                                 self.name in ["銷售收入", "銷售成本", "銀行費用", "利息收入", "董事往來"])
            
            if is_aggregate_sheet:
                # For aggregate sheets with 2-row headers, use sub_headers to find currency columns
                if hasattr(self, '_sub_headers'):
                    for col, sub_header in enumerate(self._sub_headers):
                        if sub_header and "原币(" in sub_header:
                            currency_cols.append(col)
                        # Find balance column by checking main headers
                        if col < len(self._main_headers) and "餘" in self._main_headers[col]:
                            balance_col = col
                else:
                    # Fallback for aggregate sheets without new headers
                    for col in range(col_count):
                        if col < self.columnCount():
                            row1_item = self.item(1, col)
                            if row1_item and "原币(" in row1_item.text():
                                currency_cols.append(col)
                            # Find balance column by checking row 0
                            row0_item = self.item(0, col)
                            if row0_item and "餘" in row0_item.text():
                                balance_col = col
                # For aggregate sheets, treat all currency columns as credit columns
                credit_col = currency_cols[0] if currency_cols else None
            else:
                # Original logic for regular bank sheets
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
            
            # Set darker pen for frames and text
            painter.setPen(QColor(80, 80, 80))

            # Draw merged cell background for first 3 columns
            if col_count >= 3:
                merge_width = self.columnWidth(0) + self.columnWidth(1) + self.columnWidth(2)
                painter.fillRect(0, y1, merge_width, row_height, QColor(240, 240, 240))
                painter.drawRect(0, y1, merge_width, row_height)
                
                # Draw the merged cell text
                if self.type == "bank":
                    text = f"本币 TOTAL: {self.currency}"
                else:
                    text = "本币种"
                painter.drawText(6, y1 + row_height//2 + 5, text)
                
                # Start drawing individual cells from column 3
                start_col = 3
            else:
                start_col = 0

            # Draw remaining individual cells
            for col in range(start_col, col_count):
                x = self.columnViewportPosition(col)
                w = self.columnWidth(col)
                painter.drawRect(x, y1, w, row_height)

                text = ""
                if is_aggregate_sheet:
                    # For aggregate sheets, only show currency column sums, no balance
                    if col in currency_sums:
                        currency, column_sum = currency_sums[col]
                        text = self._format_number(column_sum)
                else:
                    # For bank sheets, show debit/credit/balance as before
                    if col == debit_col:
                        text = self._format_number(debit_sum)
                    elif col == credit_col:
                        text = self._format_number(credit_sum)
                    elif col == balance_col:
                        text = self._format_number(balance)

                # Use darker text color for better visibility
                painter.setPen(QColor(40, 40, 40))
                painter.drawText(x + 6, y1 + row_height//2 + 5, text)

            # Draw second pinned row (HKD)
            painter.fillRect(0, y2, visible_rect.width(), row_height, QColor(220, 220, 220))
            
            # Set darker pen for frames
            painter.setPen(QColor(80, 80, 80))
            
            # Draw merged cell background for first 3 columns
            if col_count >= 3:
                merge_width = self.columnWidth(0) + self.columnWidth(1) + self.columnWidth(2)
                painter.fillRect(0, y2, merge_width, row_height, QColor(220, 220, 220))
                painter.drawRect(0, y2, merge_width, row_height)
                
                # Draw the merged cell text
                text = "本期 TOTAL: HKD"
                painter.drawText(6, y2 + row_height//2 + 5, text)
                
                # Start drawing individual cells from column 3
                start_col = 3
            else:
                start_col = 0

            # Draw remaining individual cells
            for col in range(start_col, col_count):
                x = self.columnViewportPosition(col)
                w = self.columnWidth(col)
                painter.drawRect(x, y2, w, row_height)

                text = ""
                if is_aggregate_sheet:
                    # For aggregate sheets, only show HKD equivalent for currency columns, no balance
                    if col in currency_sums:
                        currency, column_sum = currency_sums[col]
                        text = self._format_number(column_sum * rate)
                else:
                    # For bank sheets, show debit/credit/balance as before
                    if col == debit_col:
                        text = self._format_number(debit_sum * rate)
                    elif col == credit_col:
                        text = self._format_number(credit_sum * rate)
                    elif col == balance_col:
                        text = self._format_number(hkd_balance)

                # Use darker text color for better visibility
                painter.setPen(QColor(40, 40, 40))
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
        
        # Also enforce scroll limits
        if hasattr(self, '_pinned_height') and self._pinned_height > 0:
            self._enforce_scroll_limits()

    def wheelEvent(self, event):
        """Override wheel events to prevent scrolling into pinned row area"""
        # Call parent first to get normal wheel handling
        super().wheelEvent(event)
        
        # Then enforce scroll limits after the scroll happens
        if hasattr(self, '_pinned_height') and self._pinned_height > 0:
            self._enforce_scroll_limits()

    def scrollContentsBy(self, dx, dy):
        """Override scroll behavior to prevent scrolling into pinned row area"""
        # Call parent first to get normal scroll handling
        super().scrollContentsBy(dx, dy)
        
        # Then enforce scroll limits after the scroll happens
        if hasattr(self, '_pinned_height') and self._pinned_height > 0:
            self._enforce_scroll_limits()
    
    def _enforce_scroll_limits(self):
        """Enforce scroll limits to prevent scrolling into pinned area"""
        scrollbar = self.verticalScrollBar()
        current_scroll = scrollbar.value()
        # For aggregate sheets, we need to account for frozen rows at the top
        if hasattr(self, '_frozen_row_count') and self._frozen_row_count > 0:
            # Calculate actual data rows (total - frozen headers - pinned summary)
            # Check if the last 2 rows are actually pinned rows by looking at background color
            last_row_item = self.item(self.rowCount() - 1, 0) if self.rowCount() > 0 else None
            second_last_row_item = self.item(self.rowCount() - 2, 0) if self.rowCount() > 1 else None
            
            has_pinned_rows = False
            if (last_row_item and last_row_item.background() == QColor(220, 220, 220) and
                second_last_row_item and second_last_row_item.background() == QColor(240, 240, 240)):
                has_pinned_rows = True
                
            if has_pinned_rows:
                data_row_count = self.rowCount() - self._frozen_row_count - 2  # Exclude frozen headers and pinned rows
            else:
                data_row_count = self.rowCount() - self._frozen_row_count  # Only exclude frozen headers
                
            total_content_height = self.rowCount() * self.rowHeight(0)  # All rows including frozen
            
            # Get viewport height
            viewport_height = self.viewport().height()
            
            # Maximum scroll should allow seeing all data rows
            # The frozen rows are handled by viewport margins, pinned rows painted at bottom
            frozen_height = self._frozen_row_count * self.rowHeight(0)
            
            # Ensure user can scroll to see all data rows without being blocked by pinned rows
            if has_pinned_rows:
                # Reserve space for pinned rows at bottom
                max_scroll = max(0, total_content_height - viewport_height + frozen_height + self._pinned_height)
            else:
                max_scroll = max(0, total_content_height - viewport_height + frozen_height)
        else:
            # For regular bank sheets - they never have actual pinned row content
            # We always use painted pinned rows at the bottom, so treat all table rows as data
            data_row_count = self.rowCount()  # All rows are data rows for bank sheets
            total_content_height = self.rowCount() * self.rowHeight(0)
            
            # Get viewport height
            viewport_height = self.viewport().height()
            
            # Maximum scroll should allow seeing all data rows without being blocked by painted pinned rows
            # Reserve space for painted pinned rows at bottom
            max_scroll = max(0, total_content_height - viewport_height + self._pinned_height)
        
        # CRITICAL FIX: Set the scrollbar's maximum to our calculated value
        # This ensures Qt allows us to scroll to see all data rows
        scrollbar.setMaximum(max_scroll)
        
        # If current scroll exceeds the limit, force it back
        if current_scroll > max_scroll:
            scrollbar.setValue(max_scroll)

    def _on_item_changed(self, item):
        # Only recalculate if debit, credit, or balance in first row changes
        balance_col = None
        debit_col = None
        credit_col = None
        
        # Skip balance calculation for aggregate sheets (they don't use traditional debit/credit structure)
        if self.type == "aggregate":
            self._auto_save()
            return
            
        for col in range(self.columnCount()):
            header_item = self.horizontalHeaderItem(col)
            if header_item is None:
                continue  # Skip columns without header items
            header = header_item.text().replace(" ", "")
            if "餘額" in header:
                balance_col = col
            if "借方" in header:
                debit_col = col
            if "貸方" in header:
                credit_col = col
        if balance_col is None or debit_col is None or credit_col is None:
            self._auto_save()
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
        
        # Get currency-specific sums for aggregate sheets
        currency_sums = self.sum_currency_columns()
        
        # Check if this is a multi-currency aggregate sheet
        is_aggregate_sheet = (hasattr(self, 'name') and 
                             self.name in ["銷售收入", "銷售成本", "銀行費用", "利息收入", "董事往來"])
        
        # For bank sheets, don't create table content in last two rows
        # We only use painted pinned rows at the bottom viewport
        if not is_aggregate_sheet:
            self.blockSignals(False)
            return
        
        balance_col = None
        debit_col = None
        credit_col = None
        currency_cols = []
        
        if is_aggregate_sheet:
            # For aggregate sheets with 2-row headers, use sub_headers to find currency columns
            if hasattr(self, '_sub_headers'):
                for col, sub_header in enumerate(self._sub_headers):
                    if sub_header and "原币(" in sub_header:
                        currency_cols.append(col)
                    # Find balance column by checking main headers
                    if col < len(self._main_headers) and "餘" in self._main_headers[col]:
                        balance_col = col
            else:
                # Fallback for aggregate sheets without new headers
                for col in range(self.columnCount()):
                    # Check row 1 for currency indicators
                    row1_item = self.item(1, col)
                    if row1_item and "原币(" in row1_item.text():
                        currency_cols.append(col)
                    # Find balance column by checking row 0
                    row0_item = self.item(0, col)
                    if row0_item and "餘" in row0_item.text():
                        balance_col = col
        else:
            # Original logic for regular bank sheets
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
            
            if col == 0:
                # First column shows the currency label
                if self.type == "bank":
                    item.setText(f"本币 TOTAL: {self.currency}")
                else:
                    item.setText("本币种")
            elif col == 1 or col == 2:
                # Clear columns 1 and 2 as they will be merged with column 0
                item.setText("")
            elif is_aggregate_sheet:
                # For aggregate sheets, only show currency column sums, no balance
                if col in currency_sums:
                    currency, column_sum = currency_sums[col]
                    item.setText(format_number(column_sum))
                else:
                    item.setText("")
            else:
                # For bank sheets, show debit/credit/balance as before
                if col == debit_col:
                    item.setText(format_number(debit_sum))
                elif col == credit_col:
                    item.setText(format_number(credit_sum))
                elif col == balance_col:
                    item.setText(format_number(balance))
                else:
                    item.setText("")
        
        # Merge first 3 columns in the currency row
        if self.columnCount() >= 3:
            self.setSpan(last_row, 0, 1, 3)
                
        # HKD row
        for col in range(self.columnCount()):
            item = self.item(last_row2, col)
            if not item:
                item = QTableWidgetItem()
                self.setItem(last_row2, col, item)
            if item.flags() & Qt.ItemIsEditable:
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            item.setBackground(QColor(220, 220, 220))
            
            if col == 0:
                # First column shows HKD label for both bank and aggregate sheets
                item.setText("本期TOTAL:HKD")
            elif col == 1 or col == 2:
                # Clear columns 1 and 2 as they will be merged with column 0
                item.setText("")
            elif is_aggregate_sheet:
                # For aggregate sheets, only show HKD equivalent for currency columns, no balance
                if col in currency_sums:
                    currency, column_sum = currency_sums[col]
                    item.setText(format_number(column_sum * rate))
                else:
                    item.setText("")
            else:
                # For bank sheets, show debit/credit/balance as before
                if col == debit_col:
                    item.setText(format_number(debit_sum * rate))
                elif col == credit_col:
                    item.setText(format_number(credit_sum * rate))
                elif col == balance_col:
                    item.setText(format_number(hkd_balance))
                else:
                    item.setText("")
        
        # Merge first 3 columns in the HKD row
        if self.columnCount() >= 3:
            self.setSpan(last_row2, 0, 1, 3)
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

    def setup_two_row_headers(self, main_headers, sub_headers, merged_ranges=None):
        """Set up 2-row horizontal headers for aggregate sheets"""
        if self.type != "aggregate":
            return
            
        print(f"DEBUG SETUP: Setting up headers for {len(main_headers)} columns")
        
        # CRITICAL: Clear all existing spans first to avoid conflicts
        for row in range(min(10, self.rowCount())):  # Clear spans in first 10 rows  
            for col in range(min(20, self.columnCount())):  # Clear spans in first 20 columns
                self.setSpan(row, col, 1, 1)
        
        self.setColumnCount(len(main_headers))
        print(f"DEBUG SETUP: Column count set to {len(main_headers)}")
        
        self._main_headers = main_headers
        self._sub_headers = sub_headers
        self._merged_ranges = merged_ranges or []
        
        if self.rowCount() < 2:
            self.setRowCount(2)
        
        for col, header in enumerate(main_headers):
            item = QTableWidgetItem(header)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            item.setBackground(QColor(220, 220, 220))
            self.setItem(0, col, item)
        
        for col, sub_header in enumerate(sub_headers):
            item = QTableWidgetItem(sub_header if sub_header else "")
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            item.setBackground(QColor(240, 240, 240))
            self.setItem(1, col, item)
        
        # Identify currency columns to decide which columns to merge vertically
        currency_columns = {i for i, h in enumerate(sub_headers) if h and "原币(" in h}

        # Apply horizontal merged ranges to main headers (row 0)
        for start_col, end_col in self._merged_ranges:
            if start_col < len(main_headers) and end_col < len(main_headers):
                self.setSpan(0, start_col, 1, end_col - start_col + 1)
                print(f"DEBUG: Horizontally merged credit header at row 0, from col {start_col} to {end_col}")

        # Vertically merge non-currency columns
        for col in range(len(main_headers)):
            if col not in currency_columns:
                self.setSpan(0, col, 2, 1)
                print(f"DEBUG: Vertically merged non-currency col {col}")

        self._setup_frozen_rows(2)
        self.horizontalHeader().setVisible(False)
        
        print(f"DEBUG SETUP: Headers setup complete, final column count: {self.columnCount()}")

    def _setup_frozen_rows(self, freeze_count):
        """Setup frozen rows that stay at the top when scrolling"""
        self._frozen_row_count = freeze_count
        
        # Calculate frozen height
        row_height = self.rowHeight(0)
        frozen_height = row_height * freeze_count
        
        # Set viewport margins to reserve space for frozen headers at top and pinned rows at bottom
        pinned_height = row_height * 2
        self.setViewportMargins(0, frozen_height, 0, pinned_height)
        
        # Store pinned height for scroll limiting (ensure it's set for aggregate sheets too)
        self._pinned_height = pinned_height
        
        # Force repaint
        self.viewport().update()
    
    def _paint_frozen_rows(self, painter, visible_rect):
        """Paint frozen rows in the reserved margin area at the top"""
        if not hasattr(self, '_frozen_row_count') or self._frozen_row_count <= 0:
            return
            
        # 1. Count currencies from parent window tabs for debug purposes
        parent_window = self.window()
        currencies = []
        if hasattr(parent_window, 'tabs'):
            tabs = parent_window.tabs
            for i in range(tabs.count()):
                tab_text = tabs.tabText(i)
                if "-" in tab_text and tab_text != "+":
                    try:
                        currencies.append(tab_text.split('-')[-1])
                    except IndexError:
                        pass # Should not happen with the check
        print(f"DEBUG: Found {len(currencies)} bank sheets with currencies: {currencies}")

        # 2. Identify currency columns from sub-headers
        currency_columns = []
        if hasattr(self, '_sub_headers'):
            for col, sub_header in enumerate(self._sub_headers):
                if sub_header and "原币(" in sub_header:
                    currency_columns.append(col)
        print(f"DEBUG: Currency columns found: {currency_columns}")

        row_height = self.rowHeight(0)
        frozen_height = row_height * self._frozen_row_count
        
        margin_rect = self.contentsMargins()
        frozen_y_start = -margin_rect.top()
        
        # Clear the frozen area
        painter.fillRect(0, frozen_y_start, visible_rect.width(), frozen_height, QColor(255, 255, 255))
        
        painted_cells = set() # Track top-left cell of already painted spans

        for row in range(self._frozen_row_count):
            # Calculate fixed Y position for frozen rows (they don't scroll)
            row_y = frozen_y_start + (row * row_height)
            
            for col in range(self.columnCount()):
                if self.isColumnHidden(col) or (row, col) in painted_cells:
                    continue

                x = self.columnViewportPosition(col)
                w = self.columnWidth(col)
                
                if x + w < 0 or x > visible_rect.width():
                    continue
                
                row_span = self.rowSpan(row, col)
                col_span = self.columnSpan(row, col)

                # Simple span detection - if current cell is not top-left of its span, skip
                # Check if we're the top-left cell of a span by looking at the span values
                if row_span > 1 or col_span > 1:
                    # This is a merged cell, check if it's the top-left
                    is_top_left = True
                    for check_row in range(max(0, row - row_span + 1), row + 1):
                        for check_col in range(max(0, col - col_span + 1), col + 1):
                            if (check_row, check_col) != (row, col):
                                if self.rowSpan(check_row, check_col) == row_span and self.columnSpan(check_row, check_col) == col_span:
                                    is_top_left = False
                                    break
                        if not is_top_left:
                            break
                    
                    if not is_top_left:
                        continue

                current_item = self.item(row, col)
                if not current_item:
                    continue

                merge_width = sum(self.columnWidth(col + i) for i in range(col_span))
                merge_height = row_height * row_span

                # Determine background color from the row
                bg_color = QColor(220, 220, 220) if row == 0 else QColor(240, 240, 240)
                
                painter.fillRect(x, row_y, merge_width, merge_height, bg_color)
                painter.setPen(QColor(80, 80, 80))
                painter.drawRect(x, row_y, merge_width, merge_height)

                text = current_item.text()
                if text:
                    painter.setPen(QColor(40, 40, 40))
                    font = painter.font()
                    font.setBold(True)
                    painter.setFont(font)
                    text_y = row_y + merge_height // 2 + 5
                    painter.drawText(x + 6, text_y, text)

                # Mark the top-left cell of the span as painted
                painted_cells.add((row, col))
                if row_span > 1 or col_span > 1:
                    print(f"DEBUG: Painted merged cell at ({row},{col}) with span ({row_span},{col_span})")

    def update_headers(self):
        if self.type == "aggregate":
            # For aggregate sheets, don't use standard header update
            return
            
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

        # Check if this is an aggregate sheet (uneditable)
        is_aggregate_sheet = (hasattr(self, 'name') and 
                             self.name in ["銷售收入", "銷售成本", "銀行費用", "利息收入", "董事往來"])

        # Table operations
        add_row = QAction("Add Row", self)
        add_col = QAction("Add Column", self)
        del_row = QAction("Delete Row", self)
        del_col = QAction("Delete Column", self)
        copy = QAction("Copy", self)
        paste = QAction("Paste", self)
        clear_content = QAction("Clear Content", self) 
        merge = QAction("Merge Cells", self)
        split = QAction("Unmerge Cells", self)

        if is_aggregate_sheet:
            add_row.setEnabled(False)
            add_col.setEnabled(False)
            del_row.setEnabled(False)
            del_col.setEnabled(False)
            paste.setEnabled(False)
            clear_content.setEnabled(False)
            merge.setEnabled(False)
            split.setEnabled(False)

        menu.addAction(add_row)
        menu.addAction(add_col)
        menu.addAction(del_row)
        menu.addAction(del_col)
        menu.addSeparator()
        menu.addAction(copy)
        menu.addAction(paste)
        menu.addAction(clear_content)
        menu.addSeparator()
        menu.addAction(merge)
        menu.addAction(split)
        menu.addSeparator()

        # File operations
        new_file = QAction("New", self)
        add_sheet = QAction("Add Sheet", self)
        rename_sheet = QAction("Rename Sheet", self)
        delete_sheet = QAction("Delete Sheet", self)
        save_file = QAction("Save", self)
        load_file = QAction("Load", self)

        menu.addAction(new_file)
        menu.addAction(add_sheet)
        menu.addAction(rename_sheet)
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
        clear_content.triggered.connect(self.clear_cell_contents)
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

    def clear_cell_contents(self):
        """Clear content from selected cells while preserving formatting"""
        selected_indexes = self.selectedIndexes()
        
        if not selected_indexes:
            return
            
        for index in selected_indexes:
            if index.isValid():
                row, col = index.row(), index.column()
                item = self.item(row, col)
                
                if item and item.flags() & Qt.ItemIsEditable:
                    item.setText("")

    def rename_sheet(self):
        """Rename the current sheet"""
        from PySide6.QtWidgets import (QInputDialog, QMessageBox, QDialog, 
                                    QVBoxLayout, QFormLayout, QLineEdit, 
                                    QComboBox, QDialogButtonBox)

        if self.type == "bank":
            # Custom dialog for bank sheets
            dialog = QDialog(self)
            dialog.setWindowTitle("Rename Bank Sheet")
            layout = QVBoxLayout(dialog)

            # Pre-fill bank name by removing current currency suffix
            current_bank_name = self.name
            if self.currency and self.name.endswith('-' + self.currency):
                current_bank_name = self.name[:-len(self.currency)-1]

            # Widget setup
            bank_name_edit = QLineEdit(current_bank_name)
            currency_combo = QComboBox()

            # Common currencies list
            common_currencies = ['USD', 'CAD', 'EUR', 'GBP', 'JPY', 'CNY', 
                                'AUD', 'CHF', 'HKD', 'SGD', 'INR', 'MXN']
            currency_combo.addItems(common_currencies)
            
            # Select current currency or add if missing
            if self.currency:
                index = currency_combo.findText(self.currency)
                if index >= 0:
                    currency_combo.setCurrentIndex(index)
                else:
                    currency_combo.addItem(self.currency)
                    currency_combo.setCurrentIndex(currency_combo.count() - 1)
            else:
                currency_combo.setCurrentIndex(0)

            # Form layout
            form_layout = QFormLayout()
            form_layout.addRow("Bank Name:", bank_name_edit)
            form_layout.addRow("Currency:", currency_combo)
            layout.addLayout(form_layout)

            # Dialog buttons
            button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            button_box.accepted.connect(dialog.accept)
            button_box.rejected.connect(dialog.reject)
            layout.addWidget(button_box)

            # Show dialog and process input
            if dialog.exec() != QDialog.Accepted:
                return
            
            new_bank_name = bank_name_edit.text().strip()
            if not new_bank_name:
                QMessageBox.critical(
                    self,
                    "Invalid Name",
                    "Bank name cannot be empty.",
                    QMessageBox.Ok
                )
                return
            
            new_currency = currency_combo.currentText().strip()
            new_name = f"{new_bank_name}-{new_currency}"
        
        else:
            # Standard input for non-bank sheets
            new_name, ok = QInputDialog.getText(
                self,
                "Rename Sheet",
                "Enter new sheet name:",
                text=self.name
            )
            if not ok or not new_name or new_name == self.name:
                return

        # Exit if name unchanged (prevents unnecessary updates)
        if new_name == self.name:
            return

        # Update name and currency
        old_name = self.name
        self.name = new_name
        
        if self.type == "bank":
            self.currency = new_currency  # Set from combo box
        else:
            self.currency = ""  # Clear currency for non-bank

        # Update UI and save
        if hasattr(self.window(), 'update_tab_name'):
            self.window().update_tab_name(old_name, new_name)
        
        self._auto_save()
        self.viewport().update()

    def copy_cells(self):
        sel = self.selectedRanges()
        if not sel:
            return
        r = sel[0]
        # Build 2D list of cell contents
        rows = []
        for i in range(r.rowCount()):
            row = []
            for j in range(r.columnCount()):
                item = self.item(r.topRow() + i, r.leftColumn() + j)
                row.append(item.text() if item else "")
            rows.append(row)
        
        clipboard_text = "\n".join("\t".join(row) for row in rows)
        QApplication.clipboard().setText(clipboard_text)

    def paste_cells(self):
        clipboard = QApplication.clipboard()
        text = clipboard.text()
        if not text:
            return
            
        # Split clipboard text into rows
        rows = text.split('\n')
        
        # Remove any trailing empty rows
        if rows and not rows[-1].strip():
            rows = rows[:-1]
            
        # Determine paste start position based on selection
        selected_ranges = self.selectedRanges()
        if selected_ranges:
            # Use top-left of first selected range
            sel_range = selected_ranges[0]
            start_row = sel_range.topRow()
            start_col = sel_range.leftColumn()
        else:
            # No selection - use current cell
            start_row = self.currentRow()
            start_col = self.currentColumn()
            
        max_row = self.rowCount()   # Entire table boundaries
        max_col = self.columnCount()
        
        # Paste clipboard content starting at (start_row, start_col)
        for i, row_data in enumerate(rows):
            r = start_row + i
            # Stop if we exceed table row boundary
            if r >= max_row:
                break
                
            columns = row_data.split('\t')
            for j, content in enumerate(columns):
                c = start_col + j
                # Stop if we exceed table column boundary
                if c >= max_col:
                    break
                    
                # Only paste within table limits
                if r < self.rowCount() and c < self.columnCount():
                    item = self.item(r, c)
                    if item is None:
                        # Create new item if none exists
                        new_item = QTableWidgetItem(content)
                        self.setItem(r, c, new_item)
                    elif item.flags() & Qt.ItemIsEditable:
                        # Update existing editable item
                        item.setText(content)
                        
        self.viewport().update()


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
        # Check if this is an aggregate sheet (uneditable)
        is_aggregate_sheet = (hasattr(self, 'name') and 
                             self.name in ["銷售收入", "銷售成本", "銀行費用", "利息收入", "董事往來"])
        
        # Handle Delete key - disabled for aggregate sheets
        if event.key() == Qt.Key_Delete:
            if is_aggregate_sheet:
                return  # Ignore delete key for aggregate sheets
                
            for item in self.selectedItems():
                if item.flags() & Qt.ItemIsEditable:
                    item.setText("")
            if self.auto_save_callback:
                self.auto_save_callback()
            return
        
        if event.matches(QKeySequence.Copy):
            self.copy_cells()  # Call copy function
            return
        # Handle Ctrl+V for paste - disabled for aggregate sheets
        if event.matches(QKeySequence.Paste):
            if is_aggregate_sheet:
                return  # Ignore paste for aggregate sheets
                
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

        # Check if this is a multi-currency aggregate sheet by looking at the data structure
        is_aggregate_sheet = (hasattr(self, 'name') and 
                             self.name in ["銷售收入", "銷售成本", "銀行費用", "利息收入", "董事往來"])
        
        if is_aggregate_sheet and self.rowCount() >= 1:
            # For aggregate sheets with 2-row table headers, use sub_headers to find currency columns
            if self.type == "aggregate" and hasattr(self, '_sub_headers'):
                for col, sub_header in enumerate(self._sub_headers):
                    if sub_header and "原币(" in sub_header:
                        # Sum this currency column (starting from row 2 for aggregate sheets with table headers)
                        start_row = 2
                        for row in range(start_row, self.rowCount()):
                            item = self.item(row, col)
                            if item and item.text():
                                try:
                                    value_text = item.text().replace(',', '')
                                    value = float(value_text) if value_text else 0.0
                                    # For aggregate sheets, all currency columns represent credit amounts
                                    credit_sum += value
                                except Exception:
                                    pass
            else:
                # Fallback to old method
                for col in range(self.columnCount()):
                    # Check row 1 for currency indicators
                    row1_item = self.item(1, col)
                    if row1_item and "原币(" in row1_item.text():
                        # Sum this currency column (starting from row 2 to exclude headers)
                        for row in range(2, self.rowCount()):
                            item = self.item(row, col)
                            if item and item.text():
                                try:
                                    value_text = item.text().replace(',', '')
                                    value = float(value_text) if value_text else 0.0
                                    # For aggregate sheets, all currency columns represent credit amounts
                                    credit_sum += value
                                except Exception:
                                    pass
        else:
            # Original logic for regular bank sheets
            for col in range(self.columnCount()):
                header = self.horizontalHeaderItem(col).text().replace(" ", "")
                if "借方" in header:
                    debit_col = col
                if "貸方" in header:
                    credit_col = col
                    
            # Calculate totals (exclude pinned rows from sum - they are NOT data rows)
            # The pinned rows are artificial summary rows created by update_pinned_rows()
            # We need to exclude them from calculation but still show all actual data rows
            effective_row_count = self.rowCount()
            
            # Check if the last two rows are pinned rows (have special background colors)
            last_row_item = self.item(self.rowCount() - 1, 0) if self.rowCount() > 0 else None
            second_last_row_item = self.item(self.rowCount() - 2, 0) if self.rowCount() > 1 else None
            
            has_pinned_rows = False
            if (last_row_item and last_row_item.background() == QColor(220, 220, 220) and
                second_last_row_item and second_last_row_item.background() == QColor(240, 240, 240)):
                has_pinned_rows = True
                effective_row_count = self.rowCount() - 2
            
            for row in range(effective_row_count):
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

    def sum_currency_columns(self):
        """Get sum for each currency column in aggregate sheets"""
        currency_sums = {}
        
        if not (hasattr(self, 'name') and 
                self.name in ["銷售收入", "銷售成本", "銀行費用", "利息收入", "董事往來"]):
            return currency_sums
            
        if self.rowCount() < 1:
            return currency_sums
        
        # Check if we have pinned rows
        last_row_item = self.item(self.rowCount() - 1, 0) if self.rowCount() > 0 else None
        second_last_row_item = self.item(self.rowCount() - 2, 0) if self.rowCount() > 1 else None
        has_pinned_rows = (last_row_item and last_row_item.background() == QColor(220, 220, 220) and
                          second_last_row_item and second_last_row_item.background() == QColor(240, 240, 240))
        
        # For aggregate sheets with 2-row headers, look at sub_headers instead of row 1
        if self.type == "aggregate" and hasattr(self, '_sub_headers'):
            # Find currency columns using stored sub_headers
            for col, sub_header in enumerate(self._sub_headers):
                if sub_header and "原币(" in sub_header:
                    # Extract currency from text like "原币(USD)"
                    currency = sub_header.split("(")[1].split(")")[0]
                    column_sum = 0.0
                    
                    # Sum this currency column (starting from row 2 for aggregate sheets with table headers)
                    start_row = 2 if self.type == "aggregate" else 2
                    end_row = self.rowCount() - 2 if has_pinned_rows else self.rowCount()  # Exclude pinned rows only if they exist
                    
                    for row in range(start_row, end_row):
                        item = self.item(row, col)
                        if item and item.text():
                            try:
                                value_text = item.text().replace(',', '').strip()
                                value = float(value_text) if value_text else 0.0
                                column_sum += value
                            except Exception:
                                pass
                                
                    currency_sums[col] = (currency, round(column_sum, 2))
                    # Optional: Keep minimal logging for debugging if needed - comment out to reduce spam
                    # if column_sum > 0:
                    #     print(f"CURRENCY SUM: {self.name} - {currency}: {column_sum}")
        else:
            # Fallback to old method for non-aggregate sheets or sheets without new headers
            # Find currency columns and their currencies
            for col in range(self.columnCount()):
                row1_item = self.item(1, col)
                if row1_item and "原币(" in row1_item.text():
                    # Extract currency from text like "原币(USD)"
                    currency = row1_item.text().split("(")[1].split(")")[0]
                    column_sum = 0.0
                    
                    # Sum this currency column (starting from row 2 to exclude headers)
                    end_row = self.rowCount() - 2 if has_pinned_rows else self.rowCount()  # Exclude pinned rows only if they exist
                    
                    for row in range(2, end_row):
                        item = self.item(row, col)
                        if item and item.text():
                            try:
                                value_text = item.text().replace(',', '').strip()
                                value = float(value_text) if value_text else 0.0
                                column_sum += value
                            except Exception:
                                pass
                                
                    currency_sums[col] = (currency, round(column_sum, 2))
                    # Optional: Keep minimal logging for debugging if needed - comment out to reduce spam  
                    # if column_sum > 0:
                    #     print(f"CURRENCY SUM: {self.name} - {currency}: {column_sum}")
                
        return currency_sums

    def _auto_save(self, *_):
        if self.auto_save_callback:
            self.auto_save_callback()