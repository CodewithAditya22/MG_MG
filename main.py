import openpyxl
from dict import ts,dal_chadar
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
def pri_data_DAL_CHADAR():
    """
    FUNCTION TO PRINT VALUE IF USER
        ENTERS DAL CHADAR AS SKUID
    """
    SKUID_DAL_CHADAR = input("Enter SKUID")
    data_dalchadar = sh1[dal_chadar[SKUID_DAL_CHADAR]].value
    print(data_dalchadar)

#---------------------------------------------------------------
User_sheet_input = input("Enter Sheet name: ")
sh1 = wb[User_sheet_input]
if User_sheet_input == "T-SHIRT":
    pri_data_TSHIRT()
if User_sheet_input == "DAL CHADAR":
    pri_data_DAL_CHADAR()

