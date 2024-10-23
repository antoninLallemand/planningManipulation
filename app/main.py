import sys
import importlib
import settings
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QMessageBox, QLineEdit, QLabel, QWidget, QVBoxLayout, QDialog, QToolButton, QHBoxLayout, QComboBox
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import Qt


## COLOR SETTINGS!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

class ConfigOverlay(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Configuration")
        self.setFixedSize(300, 200)

        #READ JSON
        self.config = settings.settingsHandler.retrieve(self)
        self.selectedRole = self.config['role']

        # Layout for the config overlay
        self.v_layout = QVBoxLayout()

        # Add a label and a text input in the overlay
        self.name_label = QLabel(f"name : {self.config['name']}", self)
        self.v_layout.addWidget(self.name_label)

        self.name_input = QLineEdit(self)
        self.name_input.setPlaceholderText("type new name")
        self.name_input.hide()
        self.v_layout.addWidget(self.name_input)

        # ROLE
        self.role_label = QLabel(f"role : {self.roleMeaning(self.config['role'])}", self)
        self.v_layout.addWidget(self.role_label)

        self.role_input = QComboBox(self)
        self.role_input.addItem("pharmacist")
        self.role_input.addItem("pharmacy technician")
        self.role_input.addItem("other")
        self.role_input.hide()
        self.role_input.setCurrentIndex(self.config['role']-1)
        # Connect the combo box selection change event to a handler
        self.role_input.currentIndexChanged.connect(self.onRoleChange)
        self.v_layout.addWidget(self.role_input)

        # Edit button
        self.edit_button = QPushButton("Edit", self)
        self.edit_button.clicked.connect(self.onEdit)
        self.v_layout.addWidget(self.edit_button)

        # Save button
        self.save_button = QPushButton("Save", self)
        self.save_button.clicked.connect(self.onSave)
        self.save_button.hide()
        self.v_layout.addWidget(self.save_button)

        # Cancel button
        self.cancel_button = QPushButton("Cancel", self)
        self.cancel_button.clicked.connect(self.close)  # Close the overlay on cancel
        self.v_layout.addWidget(self.cancel_button)

        self.setLayout(self.v_layout)

    def onEdit(self):
        self.name_input.show()
        self.role_input.show()
        self.save_button.show()
        self.edit_button.hide()

    def onRoleChange(self):
        self.selectedRole = self.role_input.currentText()  # Get the selected text
        print(self.selectedRole)


    def onSave(self):
        name = self.name_input.text()
        if name == "":
            name = self.config['name']

        settings.settingsHandler.save(self, name, self.setRole(self.selectedRole), self.config['colors'])
        self.config = settings.settingsHandler.retrieve(self)
        self.name_label.setText(f"name : {self.config['name']}")
        self.role_label.setText(f"role : {self.roleMeaning(self.config['role'])}")

        self.name_input.hide()
        self.save_button.hide()
        self.role_input.hide()
        self.edit_button.show()
        # QMessageBox.warning(self, "Bad entry", "Please retry.")

    def roleMeaning(self, role): 
        if(role == 1):
            return "pharmacist"
        elif(role == 2):
            return "pharmacy technician"
        elif(role == 3):
            return "other"
        else:
            return "unknow"
        
    def setRole(self, strRole):
        if(strRole == "pharmacist"):
            return 1
        elif strRole == "pharmacy technician":
            return 2
        elif strRole == "other":
            return 3
        else:
            return 4


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Planning manipulation")
        self.setGeometry(100, 100, 600, 400)  # (x, y, width, height)

        # Create the central widget and set the layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        # Create a QVBoxLayout (vertical box layout)
        layout = QVBoxLayout()

        # Add a static label for app description/title
        self.app_title_label = QLabel("Planning manipulation App", self)
        self.app_title_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(self.app_title_label)

        # Create a button
        self.explorer_button = QPushButton("Select a planning Excel file", self)
        # Connect button click to the file open method
        self.explorer_button.clicked.connect(self.open_file_dialog)
        layout.addWidget(self.explorer_button)
        # Instance variable to store the selected file name
        self.selected_file = ""
        # Create a label for the file name output
        self.selected_file_label = QLabel("", self)
        layout.addWidget(self.selected_file_label)

        self.central_widget.setLayout(layout)




        # Horizontal layout for top-right config button
        top_layout = QHBoxLayout()

        # Add a spacer to push the button to the right
        top_layout.addStretch()

        # Create a tool button for the config (gear icon)
        self.config_button = QToolButton(self)
        self.config_button.setIcon(QIcon("gear_icon.png"))  # Use a gear icon here
        self.config_button.setStyleSheet("width: 50; height: auto;")
        self.config_button.clicked.connect(self.show_config_overlay)
        top_layout.addWidget(self.config_button)

        # Add top layout above the main layout
        layout.insertLayout(0, top_layout)

        # Set layout to the central widget
        self.central_widget.setLayout(layout)

        # Instance variable to store the selected file name
        self.selected_file = ""

        # Create the config overlay dialog (but keep it hidden initially)
        self.config_overlay = ConfigOverlay(self)

    def open_file_dialog(self):
        # Open the file dialog to select .xls files
        file_name, _ = QFileDialog.getOpenFileName(self, 
                                                   "Open Excel File", 
                                                   "", 
                                                   "Excel Files (*.xls *.xlsx *.xlsm);;All Files (*)")
        
        # If a file is selected, display the file path
        if file_name:
            self.selected_file = file_name
            self.selected_file_label.setText(f"Selected File: {self.selected_file}")
        else:
            QMessageBox.warning(self, "No File Selected", "Please select a file.")

    def show_config_overlay(self):
        # Position the config overlay in the middle of the window
        self.config_overlay.move(self.geometry().center() - self.config_overlay.rect().center())
        self.config_overlay.exec()



# Run the application
if __name__ == "__main__":
    importlib.reload(settings)
    app = QApplication(sys.argv)
    window = MainWindow()

    # Show the main window
    window.show()

    # Start the event loop
    sys.exit(app.exec())