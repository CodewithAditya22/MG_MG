import openpyxl
from dict import ts
filename = "2020-06-14 - STOCK UPDATE.xlsx"

wb = openpyxl.load_workbook(filename)
print(wb.sheetnames)
def pri_data_TSHIRT():
    """
    runs if user enter TSHIRT in sheet as tshirt have 2 value
    1. SKUID
    2. Size
    """
    SKUID_TSHIRT = input("Enter SKUID: ")
    SIZE_TSHIRT = input("Enter Size: ")
    data_TSHIRT = sh1[ts[SKUID_TSHIRT][SIZE_TSHIRT]].value
    print(data_TSHIRT)
#---------------------------------------------------------------
User_sheet_input = input("Enter Sheet name: ")
sh1 = wb[User_sheet_input]
if User_sheet_input == "T-SHIRT":
    pri_data_TSHIRT()


