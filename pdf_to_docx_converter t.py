import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog
from docx import Document
from PyPDF2 import PdfReader


class PDFtoDOCXConverter(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle('PDF to DOCX Converter')
        self.setGeometry(100, 100, 400, 150)

        self.layout = QVBoxLayout()

        self.label_pdf = QLabel('Select a PDF file:')
        self.layout.addWidget(self.label_pdf)

        self.btn_browse_pdf = QPushButton('Browse')
        self.btn_browse_pdf.clicked.connect(self.browsePDF)
        self.layout.addWidget(self.btn_browse_pdf)

        self.label_save = QLabel('Select a location to save DOCX file:')
        self.layout.addWidget(self.label_save)

        self.btn_browse_save = QPushButton('Select Location')
        self.btn_browse_save.clicked.connect(self.browseSaveLocation)
        self.layout.addWidget(self.btn_browse_save)

        self.btn_convert = QPushButton('Convert to DOCX')
        self.btn_convert.clicked.connect(self.convertToDOCX)
        self.layout.addWidget(self.btn_convert)

        self.setLayout(self.layout)

    def browsePDF(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, 'Select PDF File', '', 'PDF Files (*.pdf)')
        self.label_pdf.setText(f'Selected PDF: {file_path}')
        self.pdf_path = file_path

    def browseSaveLocation(self):
        file_dialog = QFileDialog()
        save_location = file_dialog.getExistingDirectory(self, 'Select Save Location')
        self.label_save.setText(f'Selected Save Location: {save_location}')
        self.save_location = save_location

    def convertToDOCX(self):
        if hasattr(self, 'pdf_path') and hasattr(self, 'save_location'):
            pdf_reader = PdfReader(self.pdf_path)

            doc = Document()
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                doc.add_paragraph(page.extract_text())

            docx_filename = f'{self.save_location}/converted_document.docx'
            doc.save(docx_filename)
            self.label_save.setText(f'Conversion complete! DOCX saved as: {docx_filename}')
        else:
            self.label_save.setText('Please select a PDF file and a save location first.')



def main():
    app = QApplication(sys.argv)
    converter = PDFtoDOCXConverter()
    converter.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
