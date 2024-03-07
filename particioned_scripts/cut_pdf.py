import os
from PyPDF2 import PdfFileReader, PdfFileWriter
import shutil
from pathlib import Path
from time import sleep

#Digite o caminho completo do PDF
pdf_file_path = 'Caminho completo do PDF.pdf'
file_base_name = pdf_file_path.replace('.pdf', '')
output_folder_path = os.path.join(os.getcwd(), 'Output')

pdf = PdfFileReader(pdf_file_path)

for page_num in range(pdf.numPages):
    pdfWriter = PdfFileWriter()
    pdfWriter.addPage(pdf.getPage(page_num))

    with open(os.path.join(output_folder_path, '{0}_Almox{1}.pdf'.format(file_base_name, page_num + 1)), 'wb') as f:
        pdfWriter.write(f)
        f.close()

print('PDF Cortado com sucesso!')
print()


