import sys
import os
import serial
import serial.tools.list_ports
import win32print
import pandas as pd
from barcode import Code39
from barcode.writer import ImageWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import inch
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QFormLayout, QLineEdit, QPushButton,
    QComboBox, QLabel, QCheckBox, QFileDialog, QSpinBox, QDialog,
    QDialogButtonBox
)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt

class Application(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Carton Sticker [Print Module]")
        self.setGeometry(100, 100, 1000, 800)

        # ✅ Apply white background + red buttons
        self.setStyleSheet("""
            QWidget {
                background-color: white;
            }
            QPushButton {
                background-color: red;
                color: white;
                padding: 6px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #cc0000;
            }
        """)

        self.pdf_folder = ""
        self.excel_data = None

        self.layout = QVBoxLayout()
        self.create_widgets()
        self.setLayout(self.layout)

    def create_widgets(self):
        # Logo placeholder
        self.logo_label = QLabel()
        logo_path = os.path.abspath("E:/Hassan projects/Automation-Project-of-carton-Stickers-main/Logo/assets/logo file.png")
        if os.path.exists(logo_path):
            logo_pixmap = QPixmap(logo_path)
            self.logo_label.setPixmap(logo_pixmap)
        else:
            self.logo_label.setText("Logo not found.")
        self.layout.addWidget(self.logo_label)

        # Printer selection
        self.printer_select_label = QLabel("Select Printer:")
        self.printer_select_combo = QComboBox(self)
        self.layout.addWidget(self.printer_select_label)
        self.layout.addWidget(self.printer_select_combo)

        # Input form
        self.form_layout = QFormLayout()

        self.weight_scale_port_input = QComboBox(self)
        self.weight_scale_port_input.addItem("Select Weight Scale Port")
        self.form_layout.addRow("Weight Scale Port:", self.weight_scale_port_input)

        self.so_number_input = QLineEdit(self)
        self.so_number_input.textChanged.connect(self.on_so_number_change)
        self.form_layout.addRow("SO Number:", self.so_number_input)

        self.job_number_input = QLineEdit(self)
        self.form_layout.addRow("Job Number:", self.job_number_input)

        self.rbo_input = QLineEdit(self)
        self.form_layout.addRow("RBO:", self.rbo_input)

        self.weight_scale_weight_input = QLineEdit(self)
        self.form_layout.addRow("Weight:", self.weight_scale_weight_input)

        self.item_input = QLineEdit(self)
        self.form_layout.addRow("Item:", self.item_input)

        self.order_qty_input = QSpinBox(self)
        self.order_qty_input.setMaximum(1000000)
        self.form_layout.addRow("Order Qty:", self.order_qty_input)

        self.po_number_input = QLineEdit(self)
        self.form_layout.addRow("PO Number:", self.po_number_input)

        self.customer_input = QLineEdit(self)
        self.form_layout.addRow("Customer:", self.customer_input)

        # Print options
        self.print_to_pdf_checkbox = QCheckBox("Print to PDF")
        self.print_to_pdf_checkbox.stateChanged.connect(self.on_print_to_pdf_checkbox_changed)
        self.form_layout.addRow(self.print_to_pdf_checkbox)

        self.show_total_weight_checkbox = QCheckBox("Show Total Weight")
        self.form_layout.addRow(self.show_total_weight_checkbox)

        self.auto_weight_checkbox = QCheckBox("Auto Weight")
        self.form_layout.addRow(self.auto_weight_checkbox)

        self.layout.addLayout(self.form_layout)

        # Buttons
        self.load_button = QPushButton("Load Excel Data")
        self.load_button.clicked.connect(self.load_excel_file)
        self.layout.addWidget(self.load_button)

        self.print_button = QPushButton("Print")
        self.print_button.clicked.connect(self.print_preview)
        self.layout.addWidget(self.print_button)

        self.clear_button = QPushButton("Clear")
        self.clear_button.clicked.connect(self.clear_fields)
        self.layout.addWidget(self.clear_button)

        # Populate serial ports and printers
        self.populate_ports()
        self.populate_printers()

    def load_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx)")
        if file_path:
            self.load_excel_data(file_path)

    def load_excel_data(self, file_path):
        try:
            self.excel_data = pd.read_excel(file_path, engine='openpyxl')
            self.excel_data.columns = self.excel_data.columns.str.strip()

            if 'SO Number' not in self.excel_data.columns:
                self.show_message("Error", "'SO Number' column is missing in the Excel file.")
            else:
                self.show_message("Info", "Excel data loaded successfully.")
        except Exception as e:
            self.show_message("Error", f"An error occurred while loading the Excel file: {e}")

    def on_so_number_change(self):
        try:
            so_number = self.so_number_input.text().strip()
            if self.excel_data is not None and so_number:
                self.excel_data['SO Number'] = self.excel_data['SO Number'].astype(str)
                data_row = self.excel_data.loc[self.excel_data['SO Number'].str.strip() == so_number]

                if not data_row.empty:
                    data_row = data_row.iloc[0]
                    self.job_number_input.setText(str(data_row.get('Job Number', '')))
                    self.rbo_input.setText(str(data_row.get('RBO', '')))
                    self.weight_scale_weight_input.setText(str(data_row.get('Weight', '')))
                    self.item_input.setText(str(data_row.get('Item', '')))
                    self.order_qty_input.setValue(int(data_row.get('Order Qty', 0)))
                    self.po_number_input.setText(str(data_row.get('PO Number', '')))
                    self.customer_input.setText(str(data_row.get('Customer', '')))
                else:
                    self.show_message("Info", f"SO Number '{so_number}' not found in the Excel data.")
        except Exception as e:
            self.show_message("Error", f"An error occurred: {e}")

    def print_preview(self):
        text_data = {
            "SO Number": self.so_number_input.text(),
            "Job Number": self.job_number_input.text(),
            "RBO": self.rbo_input.text(),
            "Weight": self.weight_scale_weight_input.text(),
            "Item": self.item_input.text(),
            "Order Qty": str(self.order_qty_input.value()),
            "PO Number": self.po_number_input.text(),
            "Customer": self.customer_input.text()
        }

        if self.print_to_pdf_checkbox.isChecked():
            if not self.pdf_folder:
                self.show_message("Error", "PDF folder is not selected.")
                return
            file_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "", "PDF Files (*.pdf)")
            if file_path:
                self.save_to_pdf(text_data, file_path)
        else:
            # Logic for physical printer goes here
            pass

    def save_to_pdf(self, text_data, file_path):
        try:
            page_width = 4 * inch
            page_height = 4 * inch
            c = canvas.Canvas(file_path, pagesize=(page_width, page_height))
            c.setFont("Helvetica-Bold", 10)
            y_position = page_height - 0.3 * inch
            x_position = 0.3 * inch

            for key, value in text_data.items():
                c.drawString(x_position, y_position, f"{key}: {value}")
                y_position -= 0.3 * inch

            so_number = self.so_number_input.text()
            barcode = Code39(so_number, writer=ImageWriter())
            barcode.save("barcode")
            barcode_image = os.path.join(os.getcwd(), "barcode.png")

            c.drawImage(barcode_image, x_position, y_position - 0.8 * inch, width=2.5 * inch, height=0.8 * inch)
            c.showPage()
            c.save()

            self.show_message("Info", f"PDF saved successfully to {file_path}")
        except Exception as e:
            self.show_message("Error", f"An error occurred while saving to PDF: {e}")

    def on_print_to_pdf_checkbox_changed(self, state):
        if state == Qt.Checked:
            folder = QFileDialog.getExistingDirectory(self, "Select PDF Folder")
            if folder:
                self.pdf_folder = folder
            else:
                self.print_to_pdf_checkbox.setChecked(False)
                self.show_message("Error", "No folder selected for saving PDFs.")

    def populate_ports(self):
        try:
            ports = serial.tools.list_ports.comports()
            port_names = [port.device for port in ports]
            self.weight_scale_port_input.clear()
            self.weight_scale_port_input.addItems(port_names)
        except Exception as e:
            self.show_message("Error", f"Failed to retrieve serial ports: {e}")

    def populate_printers(self):
        try:
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
            printer_names = [printer[2] for printer in printers]
            self.printer_select_combo.addItems(printer_names)
        except Exception as e:
            self.show_message("Error", f"Failed to retrieve printers: {e}")

    def show_message(self, title, message):
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setModal(True)
        layout = QVBoxLayout()
        label = QLabel(message)
        layout.addWidget(label)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.setLayout(layout)
        dialog.exec_()

    def clear_fields(self):
        self.printer_select_combo.setCurrentIndex(0)
        self.weight_scale_port_input.setCurrentIndex(0)
        self.so_number_input.clear()
        self.job_number_input.clear()
        self.rbo_input.clear()
        self.weight_scale_weight_input.clear()
        self.item_input.clear()
        self.order_qty_input.setValue(0)
        self.po_number_input.clear()
        self.customer_input.clear()
        self.print_to_pdf_checkbox.setChecked(False)
        self.show_total_weight_checkbox.setChecked(False)
        self.auto_weight_checkbox.setChecked(False)
        self.pdf_folder = ""

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Application()
    window.show()
    sys.exit(app.exec_())
