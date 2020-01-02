from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime

now=datetime.now()
filename=now.strftime("./out/%d-%m-%Y_%H-%M-%S.xlsx")

wb = Workbook()
wb.save(filename)
