import sys
import importlib
import settings
import planning_generation
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QMessageBox, QLineEdit, QLabel, QWidget, QVBoxLayout, QDialog, QToolButton, QHBoxLayout, QComboBox, QColorDialog, QStackedWidget
from PyQt5.QtGui import QIcon, QColor, QMovie, QPalette
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import copy


## SAVE PNG NAME + ROLE SETTINGS + QSS + .EXE

class ConfigOverlay(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Configuration")
        self.setFixedSize(300, 500)

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

        self.config_colors = copy.deepcopy(self.config['colors'])
        self.selected_colors = copy.deepcopy(self.config['colors'])

        # COLORS
        #OFF
        # Create a QLabel to display the selected color
        self.off_color_label = QLabel("Planning off Color", self)
        self.off_color_label.setStyleSheet(f"color: black; background-color: {self.config_colors['off']};")  # Initial color is white
        self.v_layout.addWidget(self.off_color_label)
        # Create a QPushButton to open the color picker dialog
        self.off_color_button = QPushButton("Pick a off Color", self)
        self.off_color_button.clicked.connect(lambda: self.open_color_dialog(self.off_color_label, 'off'))  # Connect button to color dialog
        self.v_layout.addWidget(self.off_color_button)
        self.off_color_button.hide()

        #WORK
        # Create a QLabel to display the selected color
        self.work_color_label = QLabel("Planning work Color", self)
        self.work_color_label.setStyleSheet(f"color: black; background-color: {self.config_colors['work']};")  # Initial color is white
        self.v_layout.addWidget(self.work_color_label)
        # Create a QPushButton to open the color picker dialog
        self.work_color_button = QPushButton("Pick a work Color", self)
        self.work_color_button.clicked.connect(lambda: self.open_color_dialog(self.work_color_label, 'work'))  # Connect button to color dialog
        self.v_layout.addWidget(self.work_color_button)
        self.work_color_button.hide()

        #SICK
        # Create a QLabel to display the selected color
        self.sick_color_label = QLabel("Planning sick Color", self)
        self.sick_color_label.setStyleSheet(f"color: black; background-color: {self.config_colors['sick']};")  # Initial color is white
        self.v_layout.addWidget(self.sick_color_label)
        # Create a QPushButton to open the color picker dialog
        self.sick_color_button = QPushButton("Pick a sick Color", self)
        self.sick_color_button.clicked.connect(lambda: self.open_color_dialog(self.sick_color_label, 'sick'))  # Connect button to color dialog
        self.v_layout.addWidget(self.sick_color_button)
        self.sick_color_button.hide()

        #VACATION
        # Create a QLabel to display the selected color
        self.vacation_color_label = QLabel("Planning vacation Color", self)
        self.vacation_color_label.setStyleSheet(f"color: black; background-color: {self.config_colors['vacation']};")  # Initial color is white
        self.v_layout.addWidget(self.vacation_color_label)
        # Create a QPushButton to open the color picker dialog
        self.vacation_color_button = QPushButton("Pick a vacation Color", self)
        self.vacation_color_button.clicked.connect(lambda: self.open_color_dialog(self.vacation_color_label, 'vacation'))  # Connect button to color dialog
        self.v_layout.addWidget(self.vacation_color_button)
        self.vacation_color_button.hide()

        #VACATION
        # Create a QLabel to display the selected color
        self.undefined_color_label = QLabel("Planning undefined Color", self)
        self.undefined_color_label.setStyleSheet(f"color: black; background-color: {self.config_colors['undefined']};")  # Initial color is white
        self.v_layout.addWidget(self.undefined_color_label)
        # Create a QPushButton to open the color picker dialog
        self.undefined_color_button = QPushButton("Pick a undefined Color", self)
        self.undefined_color_button.clicked.connect(lambda: self.open_color_dialog(self.undefined_color_label, 'undefined'))  # Connect button to color dialog
        self.v_layout.addWidget(self.undefined_color_button)
        self.undefined_color_button.hide()

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
        self.cancel_button.clicked.connect(self.onCancel)  # Close the overlay on cancel
        self.cancel_button.hide()
        self.v_layout.addWidget(self.cancel_button)
        self.selected_colors = self.config['colors']

        self.setLayout(self.v_layout)

    def onEdit(self):
        self.name_input.show()
        self.role_input.show()
        self.off_color_button.show()
        self.work_color_button.show()
        self.sick_color_button.show()
        self.vacation_color_button.show()
        self.undefined_color_button.show()
        self.save_button.show()
        self.cancel_button.show()
        self.edit_button.hide()

    def onCancel(self):
        self.name_input.hide()
        self.save_button.hide()
        self.off_color_button.hide()
        self.off_color_label.setStyleSheet(f"color: black; background-color: {self.config_colors['off']};")
        self.work_color_button.hide()
        self.work_color_label.setStyleSheet(f"color: black; background-color: {self.config_colors['work']};")
        self.sick_color_button.hide()
        self.sick_color_label.setStyleSheet(f"color: black; background-color: {self.config_colors['sick']};")
        self.vacation_color_button.hide()
        self.vacation_color_label.setStyleSheet(f"color: black; background-color: {self.config_colors['vacation']};")
        self.undefined_color_button.hide()
        self.undefined_color_label.setStyleSheet(f"color: black; background-color: {self.config_colors['undefined']};")
        self.role_input.hide()
        self.edit_button.show()
        self.cancel_button.hide()

    def onRoleChange(self):
        self.selectedRole = self.role_input.currentText()  # Get the selected text
        print(self.selectedRole)


    def onSave(self):
        name = self.name_input.text()
        if name == "":
            name = self.config['name']

        settings.settingsHandler.save(self, name, self.setRole(self.selectedRole), self.selected_colors)
        self.config = settings.settingsHandler.retrieve(self)
        self.name_label.setText(f"name : {self.config['name']}")
        self.role_label.setText(f"role : {self.roleMeaning(self.config['role'])}")

        self.name_input.hide()
        self.save_button.hide()
        self.off_color_button.hide()
        self.work_color_button.hide()
        self.sick_color_button.hide()
        self.vacation_color_button.hide()
        self.undefined_color_button.hide()
        self.role_input.hide()
        self.edit_button.show()
        self.cancel_button.hide()
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

    def open_color_dialog(self, color_label: QLabel, color_type):
        """Open the color picker dialog and set the selected color."""
        color = QColorDialog.getColor(QColor(self.selected_colors[color_type]), self)

        if color.isValid():  # Check if a valid color is selected
            self.selected_colors[color_type] = color.name()
            color_label.setStyleSheet(f"color: black; background-color: {color.name()};")


'''------------------------- PLANNING WINDOW -------------------------'''
class PlanningOverlay(QDialog):
    def __init__(self, figures, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.figures = figures
        self.setWindowTitle("Plannings")

        # Carousel widget to hold figures
        self.carousel = QStackedWidget()

        # Convert each Plotly figure to HTML and add it to QStackedWidget
        for fig in figures:
            html = fig.to_html(include_plotlyjs="cdn")
            web_view = QWebEngineView()
            web_view.setHtml(html)
            self.carousel.addWidget(web_view)

         # Navigation buttons
        self.prev_button = QPushButton("Précédent")
        self.next_button = QPushButton("Suivant")
        self.prev_button.clicked.connect(self.show_previous_image)
        self.next_button.clicked.connect(self.show_next_image)

        # Set up the layout
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.prev_button)
        button_layout.addWidget(self.next_button)
        
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.carousel)
        main_layout.addLayout(button_layout)

        self.setLayout(main_layout)

        # Initial button state check
        self.update_button_states()

    def update_button_states(self):
        """Enable/disable buttons based on the current index in the carousel."""
        current_index = self.carousel.currentIndex()
        total_images = self.carousel.count()
        
        # Disable "Previous" if at the first image
        self.prev_button.setEnabled(current_index > 0)

        # Disable "Next" if at the last image
        self.next_button.setEnabled(current_index < total_images - 1)

    def show_previous_image(self):
        # Show the previous image in the carousel
        current_index = self.carousel.currentIndex()
        if current_index > 0:
            self.carousel.setCurrentIndex(current_index - 1)
            self.update_button_states()

    def show_next_image(self):
        # Show the next image in the carousel
        current_index = self.carousel.currentIndex()
        if current_index < self.carousel.count() - 1:
            self.carousel.setCurrentIndex(current_index + 1)
            self.update_button_states()


"""---------------- LOADING OVERLAY ------------------"""
class LoadingOverlay(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_NoSystemBackground)

         # Set overlay to cover the entire parent window
        self.resize(parent.size())
        self.move(parent.pos())

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Label to hold the animated GIF
        self.spinner_label = QLabel(self)
        layout.addWidget(self.spinner_label)

        # Load an animated GIF as a spinner
        self.spinner_movie = QMovie("./loader.gif")
        self.spinner_label.setMovie(self.spinner_movie)
        self.spinner_movie.start()

        # self.setLayout(layout)



"""--------- MANAGING GENERATION FUNCTION THREAD ----------"""

class GeneratePlanningThread(QThread):
    finished = pyqtSignal(list)

    def __init__(self, selected_file, config):
        super().__init__()
        self.selected_file = selected_file
        self.config = config

    def run(self):
        # Perform the long-running function
        figures = planning_generation.generate_planning(self.selected_file, self.config)
        self.finished.emit(figures)  # Emit the results when done

'''-------------------------- MAIN WINDOW ----------------------------'''

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

        # Planning file selsction
        self.explorer_button = QPushButton("Select a planning Excel file", self)
        # Connect button click to the file open method
        self.explorer_button.clicked.connect(self.open_file_dialog)
        layout.addWidget(self.explorer_button)
        # Instance variable to store the selected file name
        self.selected_file = ""
        # Create a label for the file name output
        self.selected_file_label = QLabel("", self)
        layout.addWidget(self.selected_file_label)

        # CALCULATION
        self.calculation_button = QPushButton("Generate planning", self)
        self.calculation_button.clicked.connect(self.calculate_planning)
        layout.addWidget(self.calculation_button)

        self.central_widget.setLayout(layout)


        # Horizontal layout for top-right config button
        top_layout = QHBoxLayout()

        # Add a spacer to push the button to the right
        top_layout.addStretch()

        # Create a tool button for the config (gear icon)
        self.config_button = QPushButton(self)
        self.config_button.setFixedSize(50, 50)
        self.config_button.clicked.connect(self.show_config_overlay)
        top_layout.addWidget(self.config_button)

        # Create QMovie and QLabel for GIF
        self.config_gif = QMovie("./gear.gif")  # Replace with your GIF file path
        self.config_gif.setScaledSize(self.config_button.size())  # Scale GIF to button size

        # Set the QMovie as an icon on the button
        self.label = QLabel(self.config_button)  # QLabel to hold GIF inside the button
        self.label.setMovie(self.config_gif)
        self.label.setGeometry(0, 0, self.config_button.width(), self.config_button.height())  # Center GIF in button

        # Start the animation
        self.config_gif.start()

        # Add top layout above the main layout
        layout.insertLayout(0, top_layout)

        # Set layout to the central widget
        self.central_widget.setLayout(layout)

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

    def calculate_planning(self):
        if self.selected_file == "":
            QMessageBox.warning(self, "No File Selected", "Please select a file nefore running generation.")
        else:
            self.config = settings.settingsHandler.retrieve(self)
             # Show loading overlay
            self.loading_overlay = LoadingOverlay(self)
            self.loading_overlay.show()

            # Run generate_planning in a separate thread
            self.thread = GeneratePlanningThread(self.selected_file, self.config)
            self.thread.finished.connect(self.on_planning_generated)
            self.thread.start()

    def on_planning_generated(self, figures):
        # Close the loading overlay when done
        self.loading_overlay.close()

        # Show planning result overlay with generated figures
        self.planning_overlay = PlanningOverlay(figures)
        self.planning_overlay.showMaximized()

# Run the application
if __name__ == "__main__":
    importlib.reload(settings)
    importlib.reload(planning_generation)
    app = QApplication(sys.argv)
    window = MainWindow()

    # Show the main window
    window.show()

    # Start the event loop
    sys.exit(app.exec())