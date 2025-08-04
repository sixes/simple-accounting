import pickle
import os
import logging
from PySide6.QtWidgets import QFileDialog, QMessageBox, QTableWidgetItem
from PySide6.QtCore import QDate, Qt

logger = logging.getLogger(__name__)

class FileManager:
    def __init__(self, main_window):
        self.main_window = main_window

    def save_file(self):
        """Save current file"""
        logger.info("save_file() called from context menu!")
        fname = self.main_window.company_input.text().strip() or "untitled"
        path, _ = QFileDialog.getSaveFileName(
            self.main_window, "Save File", f"{fname}.exl", "ExcelLike (*.exl)"
        )
        if path:
            try:
                self.save_to_path(path)
                # Update company name input to saved file name (without extension)
                base = os.path.basename(path)
                name = os.path.splitext(base)[0]
                self.main_window.company_input.setText(name)
                logger.info(f"Updated company name to '{name}' after saving")
            except Exception as e:
                logger.error(f"Failed to save file: {str(e)}")
                QMessageBox.warning(self.main_window, "Save Error", f"Failed to save file: {str(e)}")

    def save_to_path(self, path):
        """Save data to specified path"""
        data = {
            "version": "1.0",
            "company": self.main_window.company_input.text(),
            "period_from": self.main_window.period_from_input.date().toString("yyyy/MM/dd"),
            "period_to": self.main_window.period_to_input.date().toString("yyyy/MM/dd"),
            "sheets": [],
            "tab_order": [self.main_window.tabs.tabText(i) for i in range(self.main_window.tabs.count())]
        }

        # Save all sheets
        for i in range(self.main_window.tabs.count()):
            tab = self.main_window.tabs.widget(i)
            tab_name = self.main_window.tabs.tabText(i)
            logger.info(f"save_to_path begin to process {tab_name}")
            if tab_name == "+":  # Skip the plus tab
                continue
            # Only call .data() on real sheet tabs
            sheet_data = tab.data()
            try:
                if hasattr(tab, '_custom_headers'):
                    sheet_data["headers"] = tab._custom_headers

                # Get exchange rate if available
                exchange_rate = getattr(tab, "exchange_rate", 1.0)
                # Only save user-added rows for aggregate sheets (no generated data)
                if tab_name in [
                    "銷售收入", "銷售成本", "銀行費用", "利息收入", "董事往來",
                ]:
                    sheet_data = {"cells": {}, "spans": []}
                # Get sheet type
                if "-" in tab_name:
                    sheet_type = "bank"
                elif tab_name in [
                    "銷售收入", "銷售成本", "銀行費用", "利息收入", "應付費用", "董事往來", "工資", "商業登記證書", "秘書費", "審計費"
                ]:
                    sheet_type = "aggregate"
                else:
                    sheet_type = "regular"

                # Save user_added_rows if it exists
                user_added_data = None
                if hasattr(tab, 'user_added_rows') and tab.user_added_rows:
                    # For director sheet, save the actual user data with row positions
                    if sheet_type == "aggregate" and tab_name == "董事往來":
                        # Use the preserve method to get user data with content and row positions
                        if hasattr(self.main_window.sheet_manager, 'preserve_director_user_data_with_positions'):
                            user_added_data = self.main_window.sheet_manager.preserve_director_user_data_with_positions(tab)
                        else:
                            # Fallback method - save with actual row positions
                            user_added_data = []
                            for row in tab.user_added_rows:
                                if row < tab.rowCount():
                                    row_data = []
                                    has_data = False
                                    for col in range(tab.columnCount()):
                                        item = tab.item(row, col)
                                        text = item.text() if item else ""
                                        row_data.append(text)
                                        if text.strip():
                                            has_data = True
                                    if has_data:
                                        # Save as (actual_row_number, row_data) to preserve position
                                        user_added_data.append((row, row_data))
                    else:
                        # For other sheets, just save the row numbers
                        user_added_data = list(tab.user_added_rows)

                sheet_info = {
                    "name": tab_name,
                    "type": sheet_type,
                    "data": sheet_data,
                    "exchange_rate": exchange_rate,
                    "currency": getattr(tab, "currency", ""),
                    "user_added_rows": user_added_data
                }

                data["sheets"].append(sheet_info)
                logger.info(f"Successfully added sheet '{tab_name}' to save data")
            except Exception as e:
                logger.error(f"Error saving sheet {self.main_window.tabs.tabText(i)}: {e}")
                import traceback
                traceback.print_exc()
                continue

        try:
            with open(path, "wb") as f:
                pickle.dump(data, f)
            logger.info(f"Successfully saved to {path}")
        except Exception as e:
            logger.error(f"Failed to write file: {str(e)}")
            raise Exception(f"Failed to write file: {str(e)}")

    def load_file(self):
        """Load file from disk"""
        logger.info("load_file() called from context menu!")
        path, _ = QFileDialog.getOpenFileName(
            self.main_window, "Open File", "", "ExcelLike (*.exl)"
        )
        if not path:
            logger.info("No file selected")
            return

        logger.info(f"Attempting to load file {path}")
        try:
            with open(path, "rb") as f:
                data = pickle.load(f)
            self.load_data_from_dict(data)
        except Exception as e:
            logger.error(f"Failed to load file: {str(e)}")
            QMessageBox.warning(self.main_window, "Load Error", f"Failed to load file: {str(e)}")

    def load_data_from_dict(self, data):
        """Common method to load data from a dictionary (used by both auto-load and manual load)"""
        logger.info(f"Starting data load, found {len(data.get('sheets', []))} sheets")
        self.main_window.tabs.clear()
        self.main_window.user_added_rows = None
        self.main_window.sheets = []

        # Store sheets temporarily to reorder them
        temp_sheets = {}

        # Clear exchange rate inputs
        for i in reversed(range(self.main_window.layout.count())):
            item = self.main_window.layout.itemAt(i)
            if item and hasattr(item.widget(), 'setPrefix'):
                item.widget().deleteLater()
                self.main_window.layout.removeItem(item)

        # Set company name if it exists in data
        company_name = data.get("company", "").strip()
        self.main_window.company_input.setText(company_name)
        logger.info(f"Set company name to: '{company_name}'")

        # Load period dates with backward compatibility
        if "period_from" in data and "period_to" in data:
            from_date = QDate.fromString(data.get("period_from", ""), "yyyy/MM/dd")
            to_date = QDate.fromString(data.get("period_to", ""), "yyyy/MM/dd")
            if from_date.isValid():
                self.main_window.period_from_input.setDate(from_date)
            if to_date.isValid():
                self.main_window.period_to_input.setDate(to_date)
            logger.info(f"Set period to {from_date.toString()} - {to_date.toString()}")

        # First pass: create all sheets and store them in temp_sheets
        for sheet_info in data.get("sheets", []):
            try:
                sheet_type = sheet_info.get("type", "regular")
                sheet_name = sheet_info["name"]
                logger.info(f"Creating sheet: name='{sheet_name}', type='{sheet_type}'")

                if sheet_type == "bank":
                    table = self.main_window.sheet_manager.create_bank_sheet(sheet_name)
                elif sheet_type == "aggregate":
                    if sheet_name == "銷售收入":
                        table = self.main_window.sheet_manager.create_sales_sheet()
                    elif sheet_name == "銷售成本":
                        table = self.main_window.sheet_manager.create_cost_sheet()
                    elif sheet_name == "銀行費用":
                        table = self.main_window.sheet_manager.create_bank_fee_sheet()
                    elif sheet_name == "利息收入":
                        table = self.main_window.sheet_manager.create_interest_sheet()
                    elif sheet_name == "應付費用":
                        table = self.main_window.sheet_manager.create_payable_sheet()
                    elif sheet_name == "董事往來":
                        table = self.main_window.sheet_manager.create_director_sheet()
                    else:
                        table = self.main_window.sheet_manager.create_regular_sheet(sheet_name)
                else:
                    table = self.main_window.sheet_manager.create_regular_sheet(sheet_name)

                table.name = sheet_name
                temp_sheets[sheet_name] = table

            except Exception as e:
                logger.error(f"Error creating sheet {sheet_info.get('name', 'unknown')}: {e}")
                import traceback
                traceback.print_exc()
                raise e

        # Second pass: load data and add sheets in correct order
        tab_order = data.get("tab_order", [sheet["name"] for sheet in data.get("sheets", [])])
        logger.info(f"tab_order: {tab_order}")
        aggregate_names = [
            "銷售收入", "銷售成本", "銀行費用", "利息收入", "應付費用", "董事往來", "工資"
        ]
        for sheet_name in tab_order:
            if sheet_name in temp_sheets:
                sheet_info = next((s for s in data["sheets"] if s["name"] == sheet_name), None)
                if sheet_info:
                    table = temp_sheets[sheet_name]
                    try:
                        logger.info(f"Loading data for sheet '{sheet_name}'")
                        if sheet_name in aggregate_names:
                            # Do NOT call table.load_data for any aggregate sheet
                            if sheet_name == "董事往來" and "user_added_rows" in sheet_info and sheet_info["user_added_rows"]:
                                # For director sheet, store user data temporarily for restoration after refresh
                                user_data = sheet_info["user_added_rows"]
                                table._pending_user_data = user_data
                        else:
                            table.load_data(sheet_info["data"])
                            if "user_added_rows" in sheet_info and sheet_info["user_added_rows"]:
                                table.user_added_rows = set(sheet_info["user_added_rows"])
                        if "exchange_rate" in sheet_info:
                            table.set_exchange_rate(sheet_info["exchange_rate"])
                            if hasattr(table, "exchange_rate_input"):
                                table.exchange_rate_input.setValue(sheet_info["exchange_rate"])
                        if "currency" in sheet_info:
                            table.currency = sheet_info["currency"]
                        self.main_window.tabs.addTab(table, sheet_name)
                        logger.info(f"loading {sheet_info}")
                    except Exception as e:
                        logger.error(f"Error loading data for sheet {sheet_name}: {e}")
                        import traceback
                        traceback.print_exc()
                        raise e

        logger.info("Data loading completed successfully")
        self.last_loaded_company_name = data.get("company", "")
        logger.info(f"last_loaded_company_name set to '{self.last_loaded_company_name}'")
        
        # After all sheets are loaded, restore director sheet user data
        for i in range(self.main_window.tabs.count()):
            tab = self.main_window.tabs.widget(i)
            tab_name = self.main_window.tabs.tabText(i)
            
            if tab_name == "董事往來" and hasattr(tab, '_pending_user_data'):
                logger.info(f"Restoring user data for director sheet")
                user_data = tab._pending_user_data
                delattr(tab, '_pending_user_data')  # Clean up temporary storage
                
                if user_data and isinstance(user_data[0], tuple):
                    # New format: list of (row, row_data) tuples - restore to original positions
                    if hasattr(self.main_window.sheet_manager, 'restore_director_user_data_to_positions'):
                        self.main_window.sheet_manager.restore_director_user_data_to_positions(tab, user_data)
                    else:
                        # Fallback: restore to original row positions
                        tab.user_added_rows = set()
                        max_row_needed = 0
                        for original_row, row_data in user_data:
                            max_row_needed = max(max_row_needed, original_row)
                        
                        # Ensure table has enough rows
                        if max_row_needed >= tab.rowCount():
                            tab.setRowCount(max_row_needed + 50)
                        
                        # Restore data to original positions
                        for original_row, row_data in user_data:
                            tab.user_added_rows.add(original_row)
                            for col, text in enumerate(row_data):
                                if col < tab.columnCount() and text:
                                    item = QTableWidgetItem(text)
                                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
                                    tab.setItem(original_row, col, item)
                elif user_data:
                    # Old format: just row numbers - convert to new format with empty data
                    tab.user_added_rows = set(user_data)

    def auto_save(self):
        """Auto-save current state"""
        fname = self.main_window.company_input.text().strip() or "untitled"
        path = f"{fname}.exl"
        try:
            if self.main_window.tabs.count() > 0:
                self.save_to_path(path)
            else:
                logger.info("No tabs to save")
        except Exception as e:
            logger.error(f"Auto-save failed: {e}")
            import traceback
            traceback.print_exc()

    def auto_load_company_file(self):
        """Try to automatically load the company file on startup"""
        company_name = self.main_window.company_input.text().strip()
        logger.info(f"Starting auto-load, company_name: '{company_name}'")

        if company_name:
            file_path = f"{company_name}.exl"
            logger.info(f"Looking for file: {file_path}")

            try:
                if os.path.exists(file_path):
                    logger.info(f"File exists, loading...")
                    with open(file_path, "rb") as f:
                        data = pickle.load(f)
                    self.load_data_from_dict(data)
                    logger.info(f"Auto-loaded company file: {file_path}")
                else:
                    self.main_window.new_file()
            except Exception as e:
                logger.error(f"Failed to auto-load company file: {e}")
                # If loading fails, keep the default sheet that was already created
        else:
            logger.info("No company name, keeping default sheet")

