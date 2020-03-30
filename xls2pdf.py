from openpyxl import load_workbook
import win32com.client

o = win32com.client.Dispatch("Excel.Application")
o.Visible = False
wb_path = 'c:\\users\drsun\\OneDrive - LR\\MyPython\\PyExdCrv-20191215\\SCE_CoE.xlsx'
wb = o.Workbooks.Open(wb_path)

ws_index_list = [1] #say you want to print these sheets
path_to_pdf = 'c:\\users\drsun\\OneDrive - LR\\MyPython\\PyExdCrv-20191215\\SCE_CoE.pdf'


print_area = 'A1:G50'

for index in ws_index_list:
    #off-by-one so the user can start numbering the worksheets at 1
    ws = wb.Worksheets[index - 1]
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.PrintArea = print_area

wb.WorkSheets(ws_index_list).Select()
wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)