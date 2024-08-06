from openpyxl.reader.excel import load_workbook
from openpyxl.drawing.image import Image
import os


def test_header_footer_image_as_external_user(tmpdir):
    tmpdir.chdir()
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    dir_excel_file = os.path.join(BASE_DIR, f"tests/data/copy_test.xlsx".replace('/', os.sep))
    dir_image_file = Image(os.path.join(BASE_DIR, f"tests/data/logo-excel.png".replace('/', os.sep)))

    #load the workbook
    wb = load_workbook(dir_excel_file)

    #Update the worsheet with the image
    wb.worksheets[0].oddHeader.left.image = dir_image_file
    wb.worksheets[0].oddHeader.left.image.width = 98.4
    wb.worksheets[0].oddHeader.left.image.height = 28.5

    wb.worksheets[0].oddFooter.right.image = dir_image_file
    wb.worksheets[0].oddFooter.right.image.width = 98.4
    wb.worksheets[0].oddFooter.right.image.height = 28.5

    #save the excel file and now you can open the excel -> file -> print view to visualize the images
    wb.save(dir_excel_file)
