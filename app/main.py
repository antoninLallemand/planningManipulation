import sys
import importlib
import settings
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QMessageBox, QLineEdit, QLabel, QWidget, QVBoxLayout

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

        # Create a label for the name field
        self.name_label = QLabel("Enter your name:", self)
        layout.addWidget(self.name_label)

        # Create a QLineEdit widget for name entry
        self.name_input = QLineEdit(self)
        layout.addWidget(self.name_input)

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

        # Create a button to retrieve the name entered
        self.infos_button = QPushButton("Show Name and File", self)
        self.infos_button.clicked.connect(self.show_infos)
        layout.addWidget(self.infos_button)


        self.central_widget.setLayout(layout)

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


    def show_infos(self):
        # Retrieve the text entered in the QLineEdit
        name = self.name_input.text()
        file = self.selected_file
        # Display the name entered using a message box
        if name:
            QMessageBox.information(self, "Entered Name and file", f"Hello, {name} and {file}")
        else:
            QMessageBox.warning(self, "No Name Entered", "Please enter your name.")



# Run the application
if __name__ == "__main__":
    importlib.reload(settings)
    app = QApplication(sys.argv)
    window = MainWindow()

    # Show the main window
    window.show()

    # Start the event loop
    sys.exit(app.exec())