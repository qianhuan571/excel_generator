from openpyxl import *
from openpyxl.compat import range
from openpyxl.utils import get_column_letter

#def excel_gen(excelSour):
#getSour=load_workbook('Source.xlsx',read_only = True)

def dest_excel_init(excelSour):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "cycle cases"
    ws1.append([])
    ws1.append(['','Id','Prj','Chip','Board Type','Os','IDE','Target',\
                'Test Env','Module','Testcase','B-Res','Result','Comment',\
                'CRID','Testor'])
    ws2 = wb.create_sheet(title = "binary test")
    ws2.append(['Examples','Platform','Tester'])
    destName = excelSour[:-5]+'_detail.xlsx'
##    ws3 = wb.create_sheet(title="Data")
##    for row in range(10, 20):
##        for col in range(27, 54):
##            _ = ws3.cell(column=col, row=row, value="{0}".format(get_column_letter(col)))
##    print(ws3['AA10'].value)
    wb.save(filename = destName)
    return destName
