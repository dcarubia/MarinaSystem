from main import DataBase
import xlrd

db = DataBase()
result = db.search_customer(("", "", "b"))
print(result)


#method for inserting excel file for testing
  def insert_file(file, sheet)
      #file is file name of workbook, sheet is sheet name
      workbook = xlrd.open_workbook(fil, on_demand = Truee)
      #create excel obj
      worksheet = workbook.sheet_by_name(sheet)
      #use add method from main class
      x = 1
     while sheet.cell(x, 0)!= xlrd.empty_cell.value:
        firstname = cell(x, 0).value
       if sheet.cell(x, 1).value != xlrd.empty_cell.value:
            lastname = cell(x, 1).value
       if sheet.cell(x, 2).value  != xlrd.empty_cell.value:
          phone = cell(x, 2).value
       if sheet.cell(x, 3).value  != xlrd.empty_cell.value:
          street = cell(x, 3).value
        if sheet.cell(x, 4).value  != xlrd.empty_cell.value:
          city = cell(x, 4).value
        if sheet.cell(x, 5).value  != xlrd.empty_cell.value:
          state = cell(x, 5).value
       data = [firstname, lastname, phone, street, city, state]
      
      main.add_customer(main,data)
          
       if sheet.cell(x+1, 0) == xlrd.empty_cell.value:
        break 
      
      x += 1
        

    
      
