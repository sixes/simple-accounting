from PySide6.QtWidgets import QDialog, QComboBox, QLineEdit, QDialogButtonBox, QFormLayout

class AddSheetDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add Sheet")
        self.layout = QFormLayout(self)

        # Sheet type selection
        self.type_combo = QComboBox()
        self.type_combo.addItems([
            "Bank Sheet",
            "銷售收入",
            "銷售成本",
            "銀行費用",
            "利息收入",
            "應付費用",
            "董事往來",
            "商業登記證",
            "秘書費",
            "工資",
            "審計費"
        ])
        self.layout.addRow("Sheet Type:", self.type_combo)

        # Bank-specific fields
        self.bank_name = QLineEdit()
        self.currency = QLineEdit()
        self.layout.addRow("Bank Name:", self.bank_name)
        self.layout.addRow("Currency:", self.currency)

        # Dialog buttons
        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.layout.addWidget(self.button_box)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        # Connect signals
        self.type_combo.currentIndexChanged.connect(self.update_fields)
        self.update_fields()

    def update_fields(self):
        """Show/hide bank-specific fields based on sheet type"""
        is_bank = self.type_combo.currentText() == "Bank Sheet"
        self.bank_name.setEnabled(is_bank)
        self.currency.setEnabled(is_bank)

    def get_result(self):
        """Returns (sheet_name, sheet_type) tuple"""
        sheet_type = self.type_combo.currentText()
        if sheet_type == "Bank Sheet":
            return f"{self.bank_name.text().strip()}-{self.currency.text().strip()}", "bank"
        elif sheet_type in ["銷售收入", "銷售成本", "銀行費用", "利息收入", "應付費用", "董事往來", "工資", "商業登記證書", "秘書費", "審計費"]:
            return sheet_type, "aggregate"
        else:
            return sheet_type, "other"