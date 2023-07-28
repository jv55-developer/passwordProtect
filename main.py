import pandas as pd
from docx2pdf import convert
from PyPDF2 import PdfFileWriter, PdfFileReader

# Load the Excel file with passwords
df = pd.read_excel('passwords.xlsx')

# Loop through each row of the DataFrame (each file and password)
for index, row in df.iterrows():
    file = row['filename']  # assumes there's a 'filename' column
    password = str(row['password'])  # assumes there's a 'password' column

    # Convert the Word file to PDF
    convert(file)

    # Now we're going to encrypt the PDF
    pdf_file = file.replace('.docx', '.pdf')
    output_pdf_file = file.replace('.docx', '_protected.pdf')

    pdf_reader = PdfFileReader(pdf_file)

    # Read the PDF content
    pdf_writer = PdfFileWriter()
    for page_num in range(pdf_reader.getNumPages()):
        page = pdf_reader.getPage(page_num)
        pdf_writer.addPage(page)

    with open(output_pdf_file, 'wb') as output_pdf:
        pdf_writer.encrypt(password)
        pdf_writer.write(output_pdf)
