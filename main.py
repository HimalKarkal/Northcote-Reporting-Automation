import sys
import os
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QPushButton,
    QLabel,
    QFileDialog,
    QMessageBox,
)
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtCore import Qt

# Import the AllContracts class from your utils.py file
from utils import AllContracts


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class MainWindow(QMainWindow):
    """Main application window."""

    def __init__(self):
        super().__init__()

        # --- Asset Paths ---
        logo_path = resource_path("logo.jpg")

        # --- Window Configuration ---
        self.setWindowTitle("Northcote Reporting Automation")
        self.setWindowIcon(QIcon(logo_path))  # Set window icon
        self.setGeometry(100, 100, 500, 250)  # Increased height for logo

        # --- Class Variables ---
        self.selected_file_path = ""
        self.contracts_data = None

        # --- UI Elements ---
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        # --- Logo ---
        self.logo_label = QLabel()
        pixmap = QPixmap(logo_path)
        self.logo_label.setPixmap(
            pixmap.scaled(
                64,
                64,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
        )
        self.logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # --- Other Widgets ---
        self.info_label = QLabel("Please select an Excel file to begin.")
        self.info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.file_path_label = QLabel("No file selected.")
        self.file_path_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.file_path_label.setStyleSheet("color: grey;")

        self.select_file_button = QPushButton("Select Excel File")
        self.process_button = QPushButton("Process and Save CSV")
        self.process_button.setEnabled(False)  # Disabled by default

        # --- Layout ---
        self.layout.addWidget(self.logo_label)  # Add logo to top
        self.layout.addWidget(self.info_label)
        self.layout.addWidget(self.select_file_button)
        self.layout.addWidget(self.file_path_label)
        self.layout.addWidget(self.process_button)

        # --- Connections ---
        self.select_file_button.clicked.connect(self.open_file_dialog)
        self.process_button.clicked.connect(self.process_and_save_file)

    def open_file_dialog(self):
        """Opens a dialog to select the input Excel file."""
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Select Contract Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_name:
            self.selected_file_path = file_name
            self.file_path_label.setText(f"Selected: ...{file_name[-40:]}")
            self.file_path_label.setStyleSheet("color: black;")
            self.process_button.setEnabled(True)
            self.info_label.setText("File selected. Ready to process.")

    def process_and_save_file(self):
        """Processes the selected file and opens a dialog to save the output CSV."""
        try:
            self.contracts_data = AllContracts(self.selected_file_path)
            result_df = self.contracts_data.payment_plan_name_counts()
            save_file_name, _ = QFileDialog.getSaveFileName(
                self,
                "Save Processed Data",
                "payment_plan_counts.csv",
                "CSV Files (*.csv)",
            )
            if save_file_name:
                result_df.to_csv(save_file_name)
                QMessageBox.information(
                    self, "Success", f"File saved successfully to:\n{save_file_name}"
                )
                self.info_label.setText("Processing complete. Select a new file.")
                self.process_button.setEnabled(False)
                self.file_path_label.setText("No file selected.")
                self.file_path_label.setStyleSheet("color: grey;")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred:\n{str(e)}")
            self.process_button.setEnabled(False)
            self.file_path_label.setText("No file selected.")
            self.file_path_label.setStyleSheet("color: grey;")
            self.info_label.setText("An error occurred. Please select a valid file.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
