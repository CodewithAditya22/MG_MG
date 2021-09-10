import openpyxl
from dict import ts, dal_chadar

filename = "2020-06-14 - STOCK UPDATE.xlsx"

wb = openpyxl.load_workbook(filename)
print(wb.sheetnames)


def pri_data_TSHIRT():
    """
    runs if user enter TSHIRT in sheet as tshirt have 2 value
    1. SKUID
    2. Size
    """
    def substract_value():
        '''substract value of sold quantity from stock then saves'''
        subvalue_tshirt = int(input("Enter quantity of sold item"))
        CHV_TSHIRT = data_TSHIRT - subvalue_tshirt
        sh1[ts[SKUID_TSHIRT][SIZE_TSHIRT]].value = CHV_TSHIRT
        wb.save(filename)

    SKUID_TSHIRT = input("Enter SKUID: ")
    SIZE_TSHIRT = input("Enter Size: ")
    data_TSHIRT = sh1[ts[SKUID_TSHIRT][SIZE_TSHIRT]].value
    print("The current avaliable stock of", ts[SKUID_TSHIRT][SIZE_TSHIRT],
          "is: ", data_TSHIRT)
    what_to_do_tshirt = input(
        """Enter "Stockup" if adding stock of enter "Selling" if packing order"""
    )

    if what_to_do_tshirt == "Selling":
        substract_value()


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
