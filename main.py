import sys
import os
import pandas as pd
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QMessageBox
from PySide6.QtCore import Qt

class CSVConverterApp(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()
        self.csv_file = None

    def initUI(self):
        self.setWindowTitle('CSV to Excel Converter')

        layout = QVBoxLayout()

        self.select_button = QPushButton('Выбрать отчет для конвертации')
        self.select_button.clicked.connect(self.select_file)
        layout.addWidget(self.select_button)

        self.file_label = QLabel('Файл не выбран')
        layout.addWidget(self.file_label)

        self.convert_button = QPushButton('Пуск')
        self.convert_button.clicked.connect(self.convert_file)
        self.convert_button.setEnabled(False)
        layout.addWidget(self.convert_button)

        self.setLayout(layout)

    def select_file(self):
        file_dialog = QFileDialog()
        file_dialog.setNameFilter("CSV Files (*.csv)")
        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            if selected_files:
                self.csv_file = selected_files[0]
                self.file_label.setText(f'Выбранный файл: {os.path.basename(self.csv_file)}')
                self.convert_button.setEnabled(True)

    def convert_file(self):
        if not self.csv_file:
            return

        try:
            df = pd.read_csv(self.csv_file, sep=";", encoding="utf-8", quotechar='"')
            df = df.dropna(axis=1, how="all")
            for column in df.columns:
                try:
                    df[column] = pd.to_numeric(df[column])
                except ValueError:
                    pass

            output_folder = os.path.dirname(self.csv_file)
            excel_file = os.path.join(output_folder,
                                      os.path.splitext(os.path.basename(self.csv_file))[0] + ".xlsx")
            df.to_excel(excel_file, index=False, engine="openpyxl")

            QMessageBox.information(self, "Успех", f"Файл успешно сконвертирован в {excel_file}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    converter_app = CSVConverterApp()
    converter_app.show()
    sys.exit(app.exec())