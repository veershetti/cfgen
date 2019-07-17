from PyPDF2 import PdfFileWriter, PdfFileReader
import csv
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import stringWidth
from xlrd import open_workbook
import xlrd

file_Location = "VNR.xlsx"
workbook = xlrd.open_workbook(file_Location)
sheet = workbook.sheet_by_name('Sheet1')
num_rows = sheet.nrows - 1
curr_row = 0
s= 0
while curr_row < num_rows:
    curr_row += 1
    row = sheet.cell(curr_row,1)   #this is print only the cells selected (Index Start from 0).
    print row.value

    packet = io.BytesIO()
    x = 0

# create a new PDF with Reportlab
    can = canvas.Canvas(packet, pagesize=letter)    
    textWidth = stringWidth(row.value, 'Times-Italic', 26) 
    x +=(780-textWidth)/2
    le=(textWidth + x)
    y = 315
    li=(780-(textWidth+x))
    can.line(li,315,le,315)
    can.setFont('Times-Italic', 26)
    can.drawString(x, y, row.value)
    can.save()
# move to the beginning of the StringIO buffer
    packet.seek(0)
    new_pdf = PdfFileReader(packet)
# read your existing PDF
    existing_pdf = PdfFileReader(open("certificate.pdf", "rb"))
    output = PdfFileWriter()
# add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.getPage(0)
    page2 = new_pdf.getPage(0)
    page.mergePage(page2)
    output.addPage(page)
# finally, write "output" to a real file
    outputStream =open("page_{:0}.pdf".format(s), "wb")
    s+= 1
    output.write(outputStream)
outputStream.close()

