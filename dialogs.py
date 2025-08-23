from PySide6.QtWidgets import QDialog, QVBoxLayout, QLineEdit, QComboBox, QDialogButtonBox, QLabel, QRadioButton, QButtonGroup, QMessageBox

class AddSheetDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add Sheet")
        layout = QVBoxLayout(self)

        self.name_label = QLabel("Sheet Name:")
        self.name_input = QLineEdit()
        layout.addWidget(self.name_label)
        layout.addWidget(self.name_input)

        self.type_group = QButtonGroup(self)
        self.bank_radio = QRadioButton("Bank")
        self.non_bank_radio = QRadioButton("Non-Bank")
        self.type_group.addButton(self.bank_radio)
        self.type_group.addButton(self.non_bank_radio)
        layout.addWidget(self.bank_radio)
        layout.addWidget(self.non_bank_radio)
        self.bank_radio.setChecked(True)

        self.currency_label = QLabel("Currency:")
        self.currency_combo = QComboBox()
        self.currency_combo.addItems(["USD", "EUR", "JPY", "GBP", "AUD", "CAD", "CHF", "CNY", "HKD", "NZD"])
        layout.addWidget(self.currency_label)
        layout.addWidget(self.currency_combo)
        self.currency_combo.setEnabled(False)
        self.bank_radio.toggled.connect(self._on_type_changed)
        self.non_bank_radio.toggled.connect(self._on_type_changed)

        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

        self._on_type_changed()

    def _on_type_changed(self, checked=None):
        # If bank, allow editing name and enable currency
        if self.bank_radio.isChecked():
            self.name_input.setReadOnly(False)
            self.name_input.clear()
            self.currency_combo.setEnabled(True)
        elif self.non_bank_radio.isChecked():
            self.name_input.setText("非银行交易")
            self.name_input.setReadOnly(True)
            self.currency_combo.setEnabled(False)

    def accept(self):
        name = self.name_input.text().strip()
        if self.bank_radio.isChecked():
            if not name:
                QMessageBox.warning(self, "Input Error", "Please enter a sheet name.")
                return
            if not self.currency_combo.currentText():
                QMessageBox.warning(self, "Input Error", "Please select a currency.")
                return
        elif not name:
            QMessageBox.warning(self, "Input Error", "Please enter a sheet name.")
            return
        super().accept()

    def get_result(self):
        name = self.name_input.text().strip()
        if self.bank_radio.isChecked():
            return (name, "bank", self.currency_combo.currentText())
        elif self.non_bank_radio.isChecked():
            return (name, "非银行交易")
        else:
            print("Error: No type selected")
            return None