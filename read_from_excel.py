# # Reading an excel file using Python
# import openpyxl
#
#
from masteval import Trainee


def add_dates(trainee):
    # To open Workbook
    wb = openpyxl.load_workbook(loc)
    ws = wb.active
    pos = []
    for i in range(1, ws.max_row):
        if ws.cell(row=i, column=1).value.find(trainee.firstname) != -1:
            pos.append(i)
    if len(pos) == 0:
        return -1
    elif len(pos) > 1:
        if trainee.type <= 1:
            for i in pos:
                if ws.cell(row=i, column=4).value.find(trainee.name2) != -1:
                    ws.write()
    wb.save("loc")
    return 0