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
        self.sales_radio = QRadioButton("銷售收入")
        self.cost_radio = QRadioButton("銷售成本")
        self.bank_fee_radio = QRadioButton("銀行費用")
        self.interest_radio = QRadioButton("利息收入")
        self.payable_radio = QRadioButton("應付費用")
        self.director_radio = QRadioButton("董事往來")
        self.salary_radio = QRadioButton("工資")
        self.type_group.addButton(self.bank_radio)
        self.type_group.addButton(self.sales_radio)
        self.type_group.addButton(self.cost_radio)
        self.type_group.addButton(self.bank_fee_radio)
        self.type_group.addButton(self.interest_radio)
        self.type_group.addButton(self.payable_radio)
        self.type_group.addButton(self.director_radio)
        self.type_group.addButton(self.salary_radio)
        layout.addWidget(self.bank_radio)
        layout.addWidget(self.sales_radio)
        layout.addWidget(self.cost_radio)
        layout.addWidget(self.bank_fee_radio)
        layout.addWidget(self.interest_radio)
        layout.addWidget(self.payable_radio)
        layout.addWidget(self.director_radio)
        layout.addWidget(self.salary_radio)
        self.sales_radio.setChecked(True)

        self.currency_label = QLabel("Currency:")
        self.currency_combo = QComboBox()
        self.currency_combo.addItems(["USD", "CNY", "HKD", "EUR", "JPY", "GBP", "AUD", "CAD", "SGD", "TWD"])
        layout.addWidget(self.currency_label)
        layout.addWidget(self.currency_combo)
        self.currency_combo.setEnabled(False)
        self.bank_radio.toggled.connect(self._on_type_changed)
        self.sales_radio.toggled.connect(self._on_type_changed)
        self.cost_radio.toggled.connect(self._on_type_changed)
        self.bank_fee_radio.toggled.connect(self._on_type_changed)
        self.interest_radio.toggled.connect(self._on_type_changed)
        self.payable_radio.toggled.connect(self._on_type_changed)
        self.director_radio.toggled.connect(self._on_type_changed)
        self.salary_radio.toggled.connect(self._on_type_changed)

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
        else:
            # Set name to selected type and make read-only
            if self.sales_radio.isChecked():
                self.name_input.setText("銷售收入")
            elif self.cost_radio.isChecked():
                self.name_input.setText("銷售成本")
            elif self.bank_fee_radio.isChecked():
                self.name_input.setText("銀行費用")
            elif self.interest_radio.isChecked():
                self.name_input.setText("利息收入")
            elif self.payable_radio.isChecked():
                self.name_input.setText("應付費用")
            elif self.director_radio.isChecked():
                self.name_input.setText("董事往來")
            elif self.salary_radio.isChecked():
                self.name_input.setText("工資")
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
        elif self.sales_radio.isChecked():
            return (name, "銷售收入")
        elif self.cost_radio.isChecked():
            return (name, "銷售成本")
        elif self.bank_fee_radio.isChecked():
            return (name, "銀行費用")
        elif self.interest_radio.isChecked():
            return (name, "利息收入")
        elif self.payable_radio.isChecked():
            return (name, "應付費用")
        elif self.director_radio.isChecked():
            return (name, "董事往來")
        elif self.salary_radio.isChecked():
            return (name, "工資")