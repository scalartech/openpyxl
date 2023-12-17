from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.worksheet.worksheet import Worksheet

wb = Workbook()
ws: Worksheet = wb.active
ws['A1'] = 'You should see a scalar logo in the header'
# create an image and anchor it to the center header

# add to worksheet and anchor next to cells
ws.oddHeader.left.image = Image('image2.png')
ws.oddHeader.left.image.width = 98.4
ws.oddHeader.left.image.height = 28.5
ws.oddHeader.left.image._id = 1
ws.oddHeader.left.size = 10

ws.oddHeader.center.text = "#Portfolio_Company_Name"
ws.oddHeader.center.size = 12
ws.oddHeader.center.color = "0D0D0F"
ws.oddHeader.center.font = "Open Sans,Regular"

ws.oddHeader.right.text = """


&10&G"""
ws.oddHeader.right.image = Image('image1.png')
ws.oddHeader.right.image.width = 594
ws.oddHeader.right.image.height = 1.5
ws.oddHeader.right.size = 8
ws.oddHeader.right.image._id = 2

ws.HeaderFooter.scaleWithDoc = False
wb.save('logo.xlsx')