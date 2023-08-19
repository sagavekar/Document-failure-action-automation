import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QRadioButton, QFileDialog, QLabel, QVBoxLayout, QWidget, QMessageBox
Assignee = {
    "AY": "Aditya Yadav",
    "OS": "Omkar Sagavekar",
    "MH": "Masroor Hafiz",
}

def Inbound_auto(assignee_key, file_path):
    print(f"{assignee_key}  {file_path}  Auto function called")

def Outbound_auto(assignee_key, file_path):
    print(f"{assignee_key}  {file_path} Outbound Auto function called")

class SampleGUI(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Sample GUI")
        self.setGeometry(100, 100, 400, 400)

        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()

        self.browse_button = QPushButton("Browse", self)
        self.browse_button.setStyleSheet("background-color: #333; color: white; padding: 10px;")
        self.browse_button.clicked.connect(self.browse_file)
        self.layout.addWidget(self.browse_button)

        self.file_label = QLabel("Selected File: ", self)
        self.file_label.setStyleSheet("color: white;")
        self.layout.addWidget(self.file_label)

        self.radio_buttons1 = []
        self.selected_radio_button = None  # To track selected radio button

        for key in Assignee:
            radio = QRadioButton(Assignee[key], self)
            radio.setStyleSheet("color: white;")
            radio.toggled.connect(lambda checked, button=radio, key=key: self.radio_button_selected(key, button))
            self.radio_buttons1.append(radio)
            self.layout.addWidget(radio)

        self.inbound_button = QPushButton("Inbound Auto", self)
        self.inbound_button.setStyleSheet("background-color: #555; color: white; padding: 10px;")
        self.inbound_button.clicked.connect(self.handle_inbound_auto)
        self.layout.addWidget(self.inbound_button)

        self.outbound_button = QPushButton("Outbound Auto", self)
        self.outbound_button.setStyleSheet("background-color: #555; color: white; padding: 10px;")
        self.outbound_button.clicked.connect(self.handle_outbound_auto)
        self.layout.addWidget(self.outbound_button)

        self.central_widget.setLayout(self.layout)
        self.central_widget.setStyleSheet("background-color: #222;")

    def browse_file(self):
        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter("Excel files (*.xlsx)")
        file_dialog.setFileMode(QFileDialog.ExistingFile)  # Restrict to existing files
        file_path, _ = file_dialog.getOpenFileName()
        if file_path:
            self.file_label.setText("Selected File: " + file_path)

    def radio_button_selected(self, assignee_key, button):
        if button.isChecked():
            self.selected_radio_button = button
            self.selected_assignee_key = assignee_key
        else:
            self.selected_radio_button = None
            self.selected_assignee_key = None


    def handle_inbound_auto(self):
        if self.selected_radio_button and self.file_label.text() != "Selected File: ": # this to avoid calling belo function if button is not select and file path / file lable in not selected
            
            assignee_key = next(key for key, value in Assignee.items() if value == self.selected_radio_button.text())
            file_path = self.file_label.text().replace("Selected File: ", "")
            Inbound_auto(assignee_key, file_path)
            QMessageBox.information(self, "Alert", "Inbound Auto completed")

    def handle_outbound_auto(self):
        if self.selected_radio_button and self.file_label.text() != "Selected File: ": # this to avoid calling belo function if button is not select and file path / file lable in not selected
            assignee_key = next(key for key, value in Assignee.items() if value == self.selected_radio_button.text())
            file_path = self.file_label.text().replace("Selected File: ", "")
            Outbound_auto(assignee_key, file_path)
            QMessageBox.information(self, "Alert", "Out Auto completed")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SampleGUI()
    window.show()
    sys.exit(app.exec_())


