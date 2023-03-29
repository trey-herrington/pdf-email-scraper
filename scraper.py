import re
import PyPDF2
from openpyxl import Workbook

pdf_file = "12_Municipal_Government.pdf"
excel_file = "munincipal_government_contacts.xlsx"

# create a regular expression to match email addresses
email_regex = r'[\w\s\.-]+@[a-z0-9\.\s-]+'

# create a PyPDF2 reader object and open the PDF file
pdf_reader = PyPDF2.PdfReader(open(pdf_file, 'rb'))

# create an Excel workbook and select the active worksheet
workbook = Workbook()
worksheet = workbook.active


# loop through each page of the PDF file
for page in range(len(pdf_reader.pages)):
    # extract the text from the page
    page_text = pdf_reader.pages[page].extract_text()
    # search for email addresses in the text
    email_list = re.findall(email_regex, page_text)
    # write each email address to the Excel worksheet
    for email in email_list:
        worksheet.append([email])

# save the Excel workbook
workbook.save(excel_file)