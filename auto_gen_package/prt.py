from openpyxl.worksheet.worksheet import Worksheet
#测试打印

def print_tp_sheet(tp_sheet:Worksheet):
    for row in tp_sheet.iter_rows(min_row=5):
        for cell in row:
            print(cell.value,end=" ")
    print("\n")


def print_sheet(sheet:Worksheet):
    for row in sheet.iter_rows():
        for cell in row:
            print(cell.value,end=" ")
    print("\n")
