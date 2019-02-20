import openpyxl
path = "C:/Users/ryanwon7/Desktop/boeing-case-script/"
wb = openpyxl.load_workbook(path+'suppliers.xlsx')
type(wb)