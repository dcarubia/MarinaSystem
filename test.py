from main import DataBase
import xlrd
import pandas as pd

    #loc = pd.read_excel('customer.xls', sheetNamme="sheet1")
    #loc.to_cvs('customer.csv', index=False)

customer = (r"C:\Users\fusar\PycharmProjects\MarinaSystem\customer.xls")
slip = (r"C:\Users\fusar\PycharmProjects\MarinaSystem\slip.xls")
db = DataBase()



# method for inserting excel file for testing
def insert_customer(file):
    # file is file name of workbook, sheet is sheet name
    workbook = xlrd.open_workbook(file, on_demand=True)
    # create excel obj
    sheet = workbook.sheet_by_index(0)
    # use add method from main class
    x = 1
    while sheet.cell(x, 0) != xlrd.empty_cell.value:
        firstname = sheet.cell(x, 0).value
        if sheet.cell(x, 1).value != xlrd.empty_cell.value:
            lastname = sheet.cell(x, 1).value
        if sheet.cell(x, 2).value != xlrd.empty_cell.value:
            phone = sheet.cell(x, 2).value
        if sheet.cell(x, 3).value != xlrd.empty_cell.value:
            street = sheet.cell(x, 3).value
        if sheet.cell(x, 4).value != xlrd.empty_cell.value:
            city = sheet.cell(x, 4).value
        if sheet.cell(x, 5).value != xlrd.empty_cell.value:
            state = sheet.cell(x, 5).value
        data = [firstname, lastname, phone, street, city, state]

        db.add_customer(data)
        x += 1
        #
    def insert_slip(file):
        # file is file name of workbook, sheet is sheet name
        workbook = xlrd.open_workbook(file, on_demand=True)
        # create excel obj
        sheet = workbook.sheet_by_index(0)
        # use add method from main class
        cust = 1
        while sheet.cell(cust, 0) != xlrd.empty_cell.value:
            current_lease = sheet.cell(cust, 0).value
            if sheet.cell(cust, 1).value != xlrd.empty_cell.value:
                max_length = sheet.cell(cust, 1).value
            if sheet.cell(cust, 2).value != xlrd.empty_cell.value:
                dock_ID = sheet.cell(cust, 2).value
            data = [current_lease, max_length, dock_ID, cust]

            db.add_slip(data)
            cust += 1

    insert_slip(slip)
    # insert_customer(customer)


