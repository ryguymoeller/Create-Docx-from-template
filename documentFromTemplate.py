# Program graps template and fills in variables using excel sheet

# Name template word doc "template.docx"
# Name excel file "rawdata"
# Look at template to name column headers in excel

from docxtpl import DocxTemplate
import jinja2
import shutil, os, re
from xlrd import open_workbook

# Function grabs data from excel and puts data into a dictionary
def excelToDict(fileloc, sheetNum):
    book = open_workbook(fileloc)
    sheet = book.sheet_by_index(sheetNum)

    # read header values into the list
    keys = [sheet.cell(0, col_index).value for col_index in range(sheet.ncols)]

    dict_list = []
    for row in range(1, sheet.nrows):
        d = {}
        for column in range(sheet.ncols):
            d.setdefault(keys[column], sheet.cell(row, column).value)
        dict_list.append(d)
    return dict_list


# Excel file with the values the program will fill the template
excelFile = r"rawdata.xlsx"

context = excelToDict(excelFile, 0)

filename = r"template.docx"

# Fills in template based on dictionary values
for i in range(len(context)):
    doc = DocxTemplate(filename)
    doc.render(context[i])
    os.makedirs('output', exist_ok=True)
    doc.save(f".\\output\\output{i+1}.docx")
