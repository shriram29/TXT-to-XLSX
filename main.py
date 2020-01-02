from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime

inputFilePath="./in/2.txt"

now=datetime.now()
filename=now.strftime("./out/%d-%m-%Y_%H-%M-%S.xlsx")

#
# wb.save(filename)
# wb = load_workbook(filename)
#
# sheet1 =

def parse(x):
    out=[]
    for itm in x:
        if itm.isdigit() or (itm.startswith('-') and itm[1:].isdigit()):
            out.append(int(itm))
        else:
            out.append(itm.replace("_",":"))
    return out

try:
    with open(inputFilePath) as inputFile:
        wb = Workbook()
        wb.create_sheet('sheet1',0)
        active = wb['sheet1']
        line = inputFile.readline()
        while line:
           active.append(parse(line.split("\t|\t")))
           line = inputFile.readline()
        wb.save(filename)
except Exception as error:
    print("Error :" + repr(error)+" ")
