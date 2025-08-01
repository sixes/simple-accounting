import pickle
import os
from PySide6.QtWidgets import QFileDialog, QMessageBox
from PySide6.QtCore import QDate

class FileManager:
    def __init__(self, main_window):
        self.main_window = main_window

    def save_file(self):
        """Save current file"""
        print("DEBUG FILE_MANAGER: save_file() called from context menu!")
        fname = self.main_window.company_input.text().strip() or "untitled"
        path, _ = QFileDialog.getSaveFileName(
            self.main_window, "Save File", f"{fname}.exl", "ExcelLike (*.exl)"
        )
        if path:
            try:
                self.save_to_path(path)
                # Update company name input to saved file name (without extension)
                import os
                base = os.path.basename(path)
                name = os.path.splitext(base)[0]
                self.main_window.company_input.setText(name)
                print(f"DEBUG SAVE: Updated company name to '{name}' after saving")
            except Exception as e:
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
            sheet = self.main_window.tabs.widget(i)
            try:
                sheet_data = sheet.data()
                if hasattr(sheet, '_custom_headers'):
                    sheet_data["headers"] = sheet._custom_headers

                # Get exchange rate if available
                exchange_rate = getattr(sheet, "exchange_rate", 1.0)
                tab_text = getattr(sheet, 'name', self.main_window.tabs.tabText(i))
                if tab_text in ["銷售收入", "銷售成本", "銀行費用", "利息收入", "應付費用"]:
                    sheet_data = {"cells": {}, "spans": []}
                # Get sheet type
                if "-" in tab_text:
                    sheet_type = "bank"
                elif tab_text in ["銷售收入", "銷售成本", "銀行費用", "利息收入", "應付費用", "董事往來", "工資", "商業登記證書", "秘書費", "審計費"]:
                    sheet_type = "aggregate"
                else:
                    sheet_type = "regular"

                # Save user_added_rows if it exists
                user_added_data = None
                if hasattr(sheet, 'user_added_rows'):
                    user_added_data = list(sheet.user_added_rows) if sheet.user_added_rows else None

                sheet_info = {
                    "name": tab_text,
                    "type": sheet_type,
                    "data": sheet_data,
                    "exchange_rate": exchange_rate,
                    "currency": getattr(sheet, "currency", ""),
                    "user_added_rows": user_added_data
                }

                data["sheets"].append(sheet_info)
                print(f"DEBUG SAVE: Successfully added sheet '{tab_text}' to save data")
            except Exception as e:
                print(f"ERROR SAVE: Error saving sheet {self.main_window.tabs.tabText(i)}: {e}")
                import traceback
                traceback.print_exc()
                continue

        try:
            with open(path, "wb") as f:
                pickle.dump(data, f)
            print(f"DEBUG SAVE: Successfully saved to {path}")
        except Exception as e:
            raise Exception(f"Failed to write file: {str(e)}")

    def load_file(self):
        """Load file from disk"""
        print("DEBUG FILE_MANAGER: load_file() called from context menu!")
        path, _ = QFileDialog.getOpenFileName(
            self.main_window, "Open File", "", "ExcelLike (*.exl)"
        )
        if not path:
            print("DEBUG FILE_MANAGER: No file selected")
            return

        print(f"DEBUG FILE_MANAGER: Attempting to load file {path}")
        try:
            with open(path, "rb") as f:
                data = pickle.load(f)
            self.load_data_from_dict(data)
        except Exception as e:
            print(f"ERROR LOAD: Failed to load file: {str(e)}")
            QMessageBox.warning(self.main_window, "Load Error", f"Failed to load file: {str(e)}")

    def load_data_from_dict(self, data):
        """Common method to load data from a dictionary (used by both auto-load and manual load)"""
        print(f"DEBUG LOAD: Starting data load, found {len(data.get('sheets', []))} sheets")
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
        print(f"DEBUG LOAD: Set company name to: '{company_name}'")

        # Load period dates with backward compatibility
        if "period_from" in data and "period_to" in data:
            from_date = QDate.fromString(data.get("period_from", ""), "yyyy/MM/dd")
            to_date = QDate.fromString(data.get("period_to", ""), "yyyy/MM/dd")
            if from_date.isValid():
                self.main_window.period_from_input.setDate(from_date)
            if to_date.isValid():
                self.main_window.period_to_input.setDate(to_date)
            print(f"DEBUG LOAD: Set period to {from_date.toString()} - {to_date.toString()}")

        # First pass: create all sheets and store them in temp_sheets
        for sheet_info in data.get("sheets", []):
            try:
                sheet_type = sheet_info.get("type", "regular")
                sheet_name = sheet_info["name"]
                print(f"DEBUG LOAD: Creating sheet: name='{sheet_name}', type='{sheet_type}'")

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
                print(f"ERROR LOAD: Error creating sheet {sheet_info.get('name', 'unknown')}: {e}")
                import traceback
                traceback.print_exc()
                raise e

        # Second pass: load data and add sheets in correct order
        tab_order = data.get("tab_order", [sheet["name"] for sheet in data.get("sheets", [])])
        print(f"tab_order: {tab_order}")
        for sheet_name in tab_order:
            if sheet_name in temp_sheets:
                sheet_info = next((s for s in data["sheets"] if s["name"] == sheet_name), None)
                if sheet_info:
                    table = temp_sheets[sheet_name]
                    try:
                        print(f"DEBUG LOAD: Loading data for sheet '{sheet_name}'")
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
                        print(f"loading {sheet_info}")

                    except Exception as e:
                        print(f"ERROR LOAD: Error loading data for sheet {sheet_name}: {e}")
                        import traceback
                        traceback.print_exc()
                        raise e

        print("DEBUG LOAD: Data loading completed successfully")
        self.last_loaded_company_name = data.get("company", "")
        print(f"DEBUG LOAD: last_loaded_company_name set to '{self.last_loaded_company_name}'")

    def auto_save(self):
        """Auto-save current state"""
        fname = self.main_window.company_input.text().strip() or "untitled"
        path = f"{fname}.exl"
        try:
            if self.main_window.tabs.count() > 0:
                self.save_to_path(path)
            else:
                print(f"DEBUG AUTO_SAVE: No tabs to save")
        except Exception as e:
            print(f"DEBUG AUTO_SAVE: Auto-save failed: {e}")
            import traceback
            traceback.print_exc()

    def auto_load_company_file(self):
        """Try to automatically load the company file on startup"""
        company_name = self.main_window.company_input.text().strip()
        print(f"DEBUG AUTO_LOAD: Starting auto-load, company_name: '{company_name}'")

        if company_name:
            file_path = f"{company_name}.exl"
            print(f"DEBUG AUTO_LOAD: Looking for file: {file_path}")

            try:
                if os.path.exists(file_path):
                    print(f"DEBUG AUTO_LOAD: File exists, loading...")
                    with open(file_path, "rb") as f:
                        data = pickle.load(f)
                    self.load_data_from_dict(data)
                    print(f"DEBUG AUTO_LOAD: Auto-loaded company file: {file_path}")
                else:
                    self.main_window.new_file()
            except Exception as e:
                print(f"ERROR AUTO_LOAD: Failed to auto-load company file: {e}")
                # If loading fails, keep the default sheet that was already created
        else:
            print(f"DEBUG AUTO_LOAD: No company name, keeping default sheet")