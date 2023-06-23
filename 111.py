


import win32com.client as win32


fname = input("input:")
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fname, ReadOnly=True)

wb.SaveAs(fname + "x", FileFormat=51)  # FileFormat = 51 is for .xlsx extension
wb.Close(SaveChanges=1)  # FileFormat = 56 is for .xls extension
excel.Application.Quit()

