from PyQt5 import QtCore, QtGui, QtWidgets, QtWebEngineWidgets
from PyQt5.uic import loadUi
import sys, os
import settings
import importlib

"""
TODO :
    - SETTINGS
    - QSS
    - ERROR MANAGEMENT
"""

class PlanningThread(QtCore.QThread):
    # Signal to notify when the planning generation is complete
    finished = QtCore.pyqtSignal(list)

    def __init__(self, file, config):
        super().__init__()
        self.file = file
        self.config = config

    def run(self):
        # Perform the heavy task here
        import planning_generation  # Import inside the thread to avoid unnecessary overhead
        figures = planning_generation.generate_planning(self.file, self.config)
        self.finished.emit(figures)  # Emit the result when done


class MainUI(QtWidgets.QMainWindow):
    def __init__(self):
        super(MainUI, self).__init__()

        loadUi("planning.ui", self)

        self.config = settings.settingsHandler.retrieve(self)
        self.selected_file = ""

        # Access UI elements here (e.g., buttons, labels, etc.)
        self.browseXlsButton = self.findChild(QtWidgets.QPushButton, 'browseXlsButton')
        self.browsedXlsLine = self.findChild(QtWidgets.QLineEdit, 'browsedXlsLine')
        self.generatePlanningButton = self.findChild(QtWidgets.QPushButton, 'generatePlanningButton')
        self.mainTabWidget = self.findChild(QtWidgets.QTabWidget, 'mainTabWidget')
        self.mainTabWidget.setCurrentIndex(0)
        self.mainTabWidget.setTabEnabled(1, False)

        self.loader = QtGui.QMovie("./loader.gif")
        self.original_text = self.generatePlanningButton.text()

        # Connect signals to slots
        self.browseXlsButton.clicked.connect(self.browse_xls)
        self.generatePlanningButton.clicked.connect(self.generate_planning)


        self.resultCarousel = self.findChild(QtWidgets.QStackedWidget, 'resultCarousel')
        self.carouselPreviousButton = self.findChild(QtWidgets.QPushButton, 'carouselPreviousButton')
        self.carouselNextButton = self.findChild(QtWidgets.QPushButton, 'carouselNextButton')

        self.carouselPreviousButton.clicked.connect(self.show_previous_image)
        self.carouselNextButton.clicked.connect(self.show_next_image)


    def browse_xls(self):
        file_name, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xls *.xlsm *.xlsx)")
        if file_name:
            self.browsedXlsLine.setText(file_name)
            self.selected_file = file_name

    def update_icon(self):
        self.generatePlanningButton.setIcon(QtGui.QIcon(self.loader.currentPixmap()))

    def generate_planning(self):
        # Implement your logic for generating the planning
        print("Generate planning clicked")
        if self.selected_file == "":
            QtWidgets.QMessageBox.warning(self, "No File Selected", "Please select a file nefore running generation.")
        else:
            self.generatePlanningButton.setText("")
            self.generatePlanningButton.setEnabled(False)
            self.generatePlanningButton.setIcon(QtGui.QIcon(self.loader.currentPixmap()))
            self.loader.frameChanged.connect(self.update_icon)
            self.loader.start()

            # Start the background thread for planning generation
            self.thread = PlanningThread(self.selected_file, self.config)
            self.thread.finished.connect(self.on_generation_complete)
            self.thread.start()

    def on_generation_complete(self, figures):
        # Handle the results and stop the loader animation
        self.loader.stop()
        self.generatePlanningButton.setIcon(QtGui.QIcon())  # Clear the icon
        self.generatePlanningButton.setText(self.original_text)
        self.generatePlanningButton.setEnabled(True)

        # Process the generated figures
        self.mainTabWidget.setTabEnabled(1, True)
        self.mainTabWidget.setCurrentIndex(1)
        self.clear_all_pages()
        for fig in figures:
            html = fig.to_html(include_plotlyjs="cdn")
            web_view = QtWebEngineWidgets.QWebEngineView()
            web_view.setHtml(html)
            self.resultCarousel.addWidget(web_view)

    def update_button_states(self):
        """Enable/disable buttons based on the current index in the carousel."""
        current_index = self.resultCarousel.currentIndex()
        total_images = self.resultCarousel.count()
        # Disable "Previous" if at the first image
        self.carouselPreviousButton.setEnabled(current_index > 0)
        # Disable "Next" if at the last image
        self.carouselNextButton.setEnabled(current_index < total_images - 1)

    def show_previous_image(self):
        # Show the previous image in the carousel
        current_index = self.resultCarousel.currentIndex()
        if current_index > 0:
            self.resultCarousel.setCurrentIndex(current_index - 1)
            self.update_button_states()

    def show_next_image(self):
        # Show the next image in the carousel
        current_index = self.resultCarousel.currentIndex()
        if current_index < self.resultCarousel.count() - 1:
            self.resultCarousel.setCurrentIndex(current_index + 1)
            self.update_button_states()

    def clear_all_pages(self):
        # Remove all pages from the QStackedWidget
        while self.resultCarousel.count() > 0:
            widget_to_remove = self.resultCarousel.widget(0)  # Always get the first widget
            self.resultCarousel.removeWidget(widget_to_remove)
            widget_to_remove.deleteLater()  # Safely delete the widget
        print("All pages removed successfully")



if __name__ == "__main__":
    importlib.reload(settings)

    app = QtWidgets.QApplication(sys.argv)

    QtCore.QDir.addSearchPath('Assets', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Assets'))
    app.setWindowIcon(QtGui.QIcon("Assets:Logo.png"))
    file = QtCore.QFile('Assets:style.qss')
    file.open(QtCore.QFile.ReadOnly | QtCore.QFile.Text)
    app.setStyleSheet(str(file.readAll(), 'utf-8'))

    ui = MainUI()
    ui.show()
    sys.exit(app.exec_())
    # app.exec_()