BankNote - Excel-like Accounting Application
==========================================

Project Description:
BankNote is a desktop accounting application built with Python and PySide6 that provides an Excel-like interface for managing banking transactions, multi-currency accounting, and financial reporting.

System Requirements:
- Python 3.7+
- PySide6 (Qt6 framework for Python)
- Operating System: Windows, macOS, or Linux

Main Features:
1. Multi-sheet workbook interface with tabbed navigation
2. Multiple sheet types for different accounting purposes:
   - Bank sheets (with multi-currency support)
   - Sales revenue tracking
   - Cost accounting
   - Bank fees management
   - Interest income tracking
   - Accounts payable
   - Director transactions
   - Salary/payroll records
   - Non-bank transactions
3. Exchange rate management for foreign currencies
4. Automatic aggregation and calculation of totals
5. Excel-like cell editing, copy/paste, merge/unmerge
6. Context menus with extensive functionality
7. Auto-save functionality
8. File save/load in custom .exl format
9. Multi-currency column support with automatic conversions
10. Pinned summary rows showing totals and balances

Core Architecture:
The application follows a modular design with clear separation of concerns:

1. main.py - Application entry point
   - Sets up logging to banknote.log
   - Enables fault handler for debugging
   - Initializes the main application window

2. excel_like.py - Main application window (ExcelLike class)
   - Inherits from QMainWindow
   - Manages the tabbed interface
   - Handles company information and date periods
   - Controls exchange rate inputs
   - Coordinates between different managers
   - Implements light theme for better readability

3. excel_table.py - Core spreadsheet widget (ExcelTable class)
   - Inherits from QTableWidget
   - Supports different sheet types (bank, regular, aggregate, non_bank)
   - Implements Excel-like functionality:
     * Cell editing and navigation
     * Copy/paste operations
     * Cell merging/unmerging
     * Context menus
     * Keyboard shortcuts
   - Handles pinned rows for totals display
   - Manages multi-currency calculations
   - Custom paint events for visual enhancements

4. sheet_manager.py - Sheet creation and management (SheetManager class)
   - Creates different types of sheets with appropriate columns
   - Manages aggregate sheet refreshing
   - Handles multi-currency sheet structure
   - Populates data from bank sheets into aggregate views
   - Sheet reordering functionality

5. file_manager.py - File operations (FileManager class)
   - Save/load functionality using pickle format
   - Auto-save capabilities
   - Auto-load company files on startup
   - Data serialization and deserialization
   - Company name management

6. dialogs.py - User interface dialogs (AddSheetDialog class)
   - Sheet creation dialog with type selection
   - Currency selection for bank sheets
   - Input validation
   - Dynamic UI updates based on selection

Sheet Types and Their Purpose:

1. Bank Sheets (name format: "BankName-CURRENCY")
   - Track individual bank account transactions
   - Support multiple currencies (USD, EUR, JPY, GBP, CHF, CAD, AUD, CNY, HKD, NZD)
   - Exchange rate management
   - Columns: 序號, 日期, 對方科目, 摘要, 借方, 貸方, 餘額, 發票號碼

2. Aggregate Sheets (auto-populated from bank data):
   - 銷售收入 (Sales Revenue) - Credit side transactions
   - 銷售成本 (Sales Costs) - Debit side transactions  
   - 銀行費用 (Bank Fees) - Debit side transactions
   - 利息收入 (Interest Income) - Credit side transactions
   - 董事往來 (Director Transactions) - Credit side transactions

3. Manual Entry Sheets:
   - 應付費用 (Accounts Payable) - Manual debit/credit entry
   - 工資 (Salary/Payroll) - Manual entry for payroll records

4. Non-Bank Sheets:
   - Support for non-banking transactions with multiple currency columns
   - Comprehensive currency support for both debit and credit sides

File Structure:
bankNotePy/
├── main.py                 # Application entry point
├── excel_like.py          # Main window and UI coordination
├── excel_table.py         # Core spreadsheet widget implementation
├── sheet_manager.py       # Sheet creation and management
├── file_manager.py        # File save/load operations
├── dialogs.py            # UI dialogs for user input
├── bankNote.spec         # PyInstaller configuration for executable
├── banknote.log          # Application log file
├── traceback.log         # Error tracking log
└── *.exl                 # Saved workbook files

Key Features in Detail:

Exchange Rate Management:
- Each bank sheet can have its own exchange rate
- Automatic HKD conversion for reporting
- Real-time rate updates across all calculations

Data Aggregation:
- Bank sheet data automatically populates aggregate sheets
- Filtering by transaction description (摘要 field)
- Multi-currency column support in aggregate views
- Automatic refresh when source data changes

User Interface:
- Excel-like keyboard navigation (Arrow keys, Tab, Enter)
- Right-click context menus for all operations
- Copy/paste with proper formatting
- Cell merging and formatting
- Resizable columns and rows
- Tab-based sheet navigation with + button for new sheets

File Management:
- Custom .exl format using Python pickle
- Preserves all formatting and structure
- Auto-save prevents data loss
- Company name and period tracking
- Automatic file loading on startup

Multi-Currency Support:
- 10 supported currencies with automatic detection
- Currency-specific totals in pinned rows
- Exchange rate conversion to HKD base currency
- Multi-currency aggregate reporting

Installation and Usage:

1. Install Python 3.7 or higher
2. Install required dependencies:
   pip install PySide6

3. Run the application:
   python main.py

4. Or build executable using PyInstaller:
   pyinstaller bankNote.spec

Operating Instructions:

1. Starting the Application:
   - Run main.py to launch the application
   - Default company name is "company_name"
   - Period dates default to current month

2. Creating Sheets:
   - Click the "+" tab to add new sheets
   - Select sheet type and currency (for bank sheets)
   - Enter descriptive name for custom sheets

3. Data Entry:
   - Double-click cells to edit
   - Use Tab/Enter to navigate
   - Right-click for context menu options
   - Bank sheets auto-populate aggregate sheets

4. File Operations:
   - Use File menu for Save/Load operations
   - Files are saved in custom .exl format
   - Auto-save prevents data loss

5. Multi-Currency:
   - Set exchange rates in top bar
   - Rates apply to currently selected bank sheet
   - HKD conversions shown in pinned rows

Technical Notes:

- Built with PySide6 (Qt6) for cross-platform compatibility
- Uses custom painting for pinned rows and visual effects
- Pickle serialization for data persistence
- Comprehensive error handling and logging
- Memory-efficient design for large datasets
- Platform-specific optimizations (light theme forcing on non-Windows)

Logging and Debugging:
- All operations logged to banknote.log
- Traceback logging for error diagnosis
- Fault handler enabled for crash debugging
- Debug output for sheet operations and data flow

This application provides a comprehensive solution for small to medium business accounting needs, with particular strength in multi-currency transaction management and automated reporting.
