import sys
import base64
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QPixmap, QIcon
from PIL import Image
import pytesseract
import docx
import io

class ImageToWordConverter(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.set_styles()  # Set the styles for the UI
        self.set_window_icon()  # Set the window icon

    def init_ui(self):
        self.setWindowTitle('Image to Word Converter')

        # Input field for image file
        self.input_label = QLabel('Input Image:')
        self.input_entry = QLineEdit(self)
        self.input_entry.setPlaceholderText("Drag & Drop or Browse Image")
        self.input_entry.setAcceptDrops(True)

        # Enable drag and drop
        self.setAcceptDrops(True)

        self.input_button = QPushButton('Browse')
        self.input_button.clicked.connect(self.select_input_file)

        # Output field for Word file path
        self.output_label = QLabel('Output Directory:')
        self.output_entry = QLineEdit(self)
        self.output_button = QPushButton('Browse')
        self.output_button.clicked.connect(self.select_output_directory)

        # Convert button
        self.convert_button = QPushButton('Convert')
        self.convert_button.clicked.connect(self.convert_image_to_word)

        # Status label
        self.status_label = QLabel('')

        # Layout
        layout = QVBoxLayout()
        layout.addWidget(self.input_label)
        layout.addWidget(self.input_entry)
        layout.addWidget(self.input_button)
        layout.addWidget(self.output_label)
        layout.addWidget(self.output_entry)
        layout.addWidget(self.output_button)
        layout.addWidget(self.convert_button)
        layout.addWidget(self.status_label)
        self.setLayout(layout)

        # Set window dimensions
        self.setGeometry(300, 300, 400, 300)

    def set_styles(self):
        self.setStyleSheet("""
            QWidget {
                background-color: #2C2F33;  /* Dark background */
                color: #FFFFFF;  /* White text */
                font-family: Arial, sans-serif;
            }
            QLabel {
                font-size: 16px;
                margin-bottom: 5px;
            }
            QLineEdit {
                background-color: #23272A;  /* Dark input field */
                color: #FFFFFF;
                border: 1px solid #7289DA;  /* Light blue border */
                padding: 5px;
                border-radius: 5px;
            }
            QPushButton {
                background-color: #7289DA;  /* Blue button */
                color: #FFFFFF;
                border: none;
                padding: 10px;
                border-radius: 5px;
                font-size: 14px;
                margin-top: 10px;
            }
            QPushButton:hover {
                background-color: #5B6EAE;  /* Darker blue on hover */
            }
            QLabel#status_label {
                margin-top: 20px;
                font-weight: bold;
            }
        """)

    def set_window_icon(self):
        # Base64 encoded icon image (replace with your actual base64 string)
        base64_icon = """
        iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAAA7AAAAOwBeShxvQAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAAPrSURBVFiF5ZVtTBxFHMZ/e7vXu+VMBBSaaltSAxXElGpLX8BPNZwtLUSFKFbaCE0vIvEFJGhi2iZGkxqNtFZqS2NMfKlf/KC1lQbO2lp5rQiIVFuqTbBXqTFVhOP2ure7flCOXbyjNN4ZE59kkp35z87z/J/5z4yAFUXAbcQXg8DhyY5gjsgJLrW6bpvktMeP/eUXdgSCwWBCRAEJruvU3uExKVGOn4C0ZFlRFCXMYIsf1exgFWAYhvlze30N2WmpuPOX8e3gwKwWDIVCNDXuYv/rDaiqetX5koXf9H308Ed8/90gPw4P4/V6qav2cOR4x1XJPRVl+Pw+RLtI2xfHefOdD7DboxfVtC2YcuCi7wLLly/D6XSSn5/PT74LM5JrmsZjW8u5pF5iY2MVZXseZeL6IJWbSmZ0wiLAMFlQWHwf7x18n6dqalhXWEj5Fk845h8fZ3T0N2vmjzzIiDrCA7s82CQRQRAofr6cMdk/o4ioDsy76WY+OdFJ2uJs6re/SO2z2wDo7mhjZU46q5Yu5q2mxrDtPxu/UPLKFmzi1JKCTeD+nRWMyxN4KsoiCrAcQ6dTVvp9E45ox7C7o43KzaVs3FdN4rxkDlbt5crlAHNzFlA6jdySlm4w5P2G+rLnuD0l1XIMLQIcTmfga1/AGUmAmXzBkkUAaGqIgeZTLFm/Miq5GbJNpjbj4WBQUZyTY9JMP0ziVGc7lZtLeWhvVZgcQLRLLC1ePZslAAjoAXRDMx+2yAIMpqzp7+2hYlOJJfNYQjR36nbszLkjNy9b00E3IKRD8zkvC9dkk7IoDVVhVi2kYIQUW0hTbKoelALahD1gBOaMo8z5PfGGlJ7BE11vmx24BbgXsJ89PXDjgd0vhQUNjZ9j1D8GwGlv37UkJgD2v5rTHHBckYLAM4AKfCg4HI6Lq3Lz5jodjn/1XQgoAb3ry84Rm6ZpSe41bpssx/EJnAZZlrnn7rU2TdOSw0VoGAafnTzGE48/GVfy1/bsptC9PtyXRFH8taO7PXX1ijzx2Oef4tnqwW63c/78D7S2tsaEtKCggPT0DAD27X+D3DtX0N7VrkmidFngzyLMA951uVxqb0+flJSUyJmzZ2hpaYmJALfbTeatmQDMXzh/8iYsB9osN6HL5VL7vuqXkpISY0IcCSYBwAw34aGPD1H7dG1MSBtebaBoQ1HE2H/XgaGhIZqPNseEdN3adWRkZESMRb18QlooJuQAmq5FjUV1ICszi6zMrJiJuGYB/5sijFgDhmFEGo4L/rYFhmHg908gSWKk+f8YgiBYkrYI0HWdpgNNcSE2CbBsuzAtvgG4K64K4CRwZLLzB4HglqPCtzY+AAAAAElFTkSuQmCC
        """  # Example base64 image (a tiny 1x1 pixel)

        # Decode the base64 string and create a QPixmap
        icon_data = base64.b64decode(base64_icon)
        icon_pixmap = QPixmap()
        icon_pixmap.loadFromData(icon_data)

        # Set the window icon directly from the QPixmap
        self.setWindowIcon(QIcon(icon_pixmap))

    # Enable drag and drop for the input entry
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            file_url = event.mimeData().urls()[0]
            file_path = file_url.toLocalFile()
            self.input_entry.setText(file_path)

    def select_input_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, 'Select Image File', '', 'Image files (*.jpg *.jpeg *.png *.bmp)')
        if file_path:
            self.input_entry.setText(file_path)

    def select_output_directory(self):
        dir_path = QFileDialog.getExistingDirectory(self, 'Select Output Directory')
        if dir_path:
            self.output_entry.setText(dir_path)

    def convert_image_to_word(self):
        input_file = self.input_entry.text()
        output_directory = self.output_entry.text()

        if not input_file or not output_directory:
            self.status_label.setText("Please provide both input and output directories.")
            return

        # Create the output file path
        output_file = os.path.join(output_directory, 'output.docx')

        try:
            img = Image.open(input_file)
            text = pytesseract.image_to_string(img)

            doc = docx.Document()
            doc.add_paragraph(text)
            doc.save(output_file)

            self.status_label.setText(f"Conversion successful! Saved to {output_file}")
        except Exception as e:
            self.status_label.setText(f"Error: {str(e)}")

def main():
    app = QApplication(sys.argv)
    window = ImageToWordConverter()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
