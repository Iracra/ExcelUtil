import sys
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QLabel,
    QPushButton,
    QLineEdit,
    QFileDialog,
    QMessageBox,
    QProgressBar,
)
from PyQt6.QtCore import Qt, QTimer
import os

# Importing the province mapping from constants.py
from constants import PROVINCE


class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Elaborazione File Excel")
        self.setGeometry(100, 100, 800, 300)

        # Main widget and layout
        self.main_widget = QWidget()
        self.setCentralWidget(self.main_widget)
        self.layout = QVBoxLayout(self.main_widget)

        # Label for file selection
        self.label = QLabel("Seleziona i file Excel da elaborare")
        self.label.setStyleSheet("font-size: 12pt;")
        self.layout.addWidget(self.label)

        # Button to select files
        self.select_button = QPushButton("Seleziona File")
        self.select_button.clicked.connect(self.select_files)
        self.layout.addWidget(self.select_button)

        # Label to display selected files
        self.selected_files_label = QLabel("Nessun file selezionato")
        self.selected_files_label.setStyleSheet("color: gray;")
        self.layout.addWidget(self.selected_files_label)

        # Button to select output folder
        self.output_button = QPushButton("Seleziona Cartella")
        self.output_button.clicked.connect(self.select_output)
        self.layout.addWidget(self.output_button)

        # Label to display selected output folder
        self.selected_output_label = QLabel("Nessuna cartella selezionata")
        self.selected_output_label.setStyleSheet("color: gray;")
        self.layout.addWidget(self.selected_output_label)

        # Label and input for header row
        self.header_label = QLabel("Riga di intestazione (Excel):")
        self.layout.addWidget(self.header_label)

        self.header_entry = QLineEdit()
        self.header_entry.setText("6")  # Default value (row 6 in Excel = index 5 in Python)
        self.layout.addWidget(self.header_entry)

        # Button to start processing
        self.process_button = QPushButton("Avvia Elaborazione")
        self.process_button.clicked.connect(self.process_excel)
        self.layout.addWidget(self.process_button)

        # Variables for paths
        self.input_paths = []
        self.output_path = ""

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Seleziona i file Excel da elaborare", "", "Excel files (*.xlsx *.xls)"
        )
        if files:
            self.input_paths = files
            self.selected_files_label.setText(f"{len(files)} file selezionati")
            self.selected_files_label.setStyleSheet("color: green;")

    def select_output(self):
        folder = QFileDialog.getExistingDirectory(self, "Seleziona cartella di output")
        if folder:
            self.output_path = folder
            self.selected_output_label.setText(f"Output: {folder}")
            self.selected_output_label.setStyleSheet("color: green;")

    def show_progress_popup(self):
        """Show a progress popup with a progress bar."""
        self.progress_popup = QMessageBox(self)
        self.progress_popup.setWindowTitle("Avanzamento")
        self.progress_popup.setText("Elaborazione in corso...")

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_popup.layout().addWidget(self.progress_bar, 1, 1)

        self.progress_popup.show()

    def update_progress(self, value):
        """Update the progress bar."""
        self.progress_bar.setValue(int(value))
        QApplication.processEvents()  # Force UI update

    def close_progress_popup(self):
        """Close the progress popup."""
        self.progress_popup.close()

    def process_excel(self):
        if not self.input_paths:
            QMessageBox.critical(self, "Errore", "Seleziona i file Excel da elaborare!")
            return
        if not self.output_path:
            QMessageBox.critical(self, "Errore", "Seleziona la cartella di output!")
            return

        try:
            # Show progress popup
            self.show_progress_popup()

            # Get header row from user input
            header_row = int(self.header_entry.text()) - 1  # Convert from Excel row to Python index

            # Total file count
            total_files_processed = 0

            # Process files sequentially
            total_files = len(self.input_paths)
            for i, path in enumerate(self.input_paths):
                # Update progress
                progress_value = (i + 1) / total_files * 100
                self.update_progress(progress_value)

                df = pd.read_excel(path, header=header_row, na_values=['', 'N/A', 'NaN', 'nan', 'None'], keep_default_na=False)

                # Find 'RegioneResidenza' column
                regione_col = next((col for col in df.columns if 'regione' in col.lower()), None)
                if not regione_col:
                    print(f"Attenzione: colonna 'RegioneResidenza' non trovata in {os.path.basename(path)}. File saltato.")
                    continue

                # Extract region and normalize
                df['Regione'] = df[regione_col].str.strip().str.capitalize()

                # Find 'Provincia' column
                provincia_col = next((col for col in df.columns if 'provincia' in col.lower()), None)
                if not provincia_col:
                    print(f"Attenzione: colonna 'Provincia' non trovata in {os.path.basename(path)}. File saltato.")
                    continue

                # Process data by region and province
                for regione, group in df.groupby('Regione'):
                    region_folder = os.path.join(self.output_path, regione)
                    os.makedirs(region_folder, exist_ok=True)

                    for provincia, sub_group in group.groupby(provincia_col):
                        # Use the PROVINCE mapping to get the full name of the province
                        provincia_full_name = PROVINCE.get(provincia, provincia)  # Fallback to provincia if not found

                        # Prepare the file name
                        safe_name = f"{provincia_full_name}.xlsx"

                        # Save the file in the respective region folder (without subfolders for provinces)
                        sub_group.to_excel(
                            os.path.join(region_folder, safe_name),
                            index=False,
                            header=False,  # Include header in the output file
                            engine='openpyxl'
                        )

                        # Update total files processed
                        total_files_processed += 1

            # Close progress popup
            self.close_progress_popup()

            # Show completion message
            QMessageBox.information(
                self,
                "Completato",
                f"Elaborazione completata!\nFile salvati in: {self.output_path}\n"
                f"Totale file elaborati: {total_files_processed}",
            )

        except ValueError:
            QMessageBox.critical(self, "Errore", "Inserisci un numero valido per la riga di intestazione!")
            self.close_progress_popup()
        except Exception as e:
            QMessageBox.critical(self, "Errore", f"Errore durante l'elaborazione: {str(e)}")
            self.close_progress_popup()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec())
