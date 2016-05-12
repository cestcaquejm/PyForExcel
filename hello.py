import sys
import datetime
import openpyxl

argument = sys.argv

print('using ' + argument[1])

wbname = argument[1]
# wbname = 'pythontest.xlsx'

wbnamesplited = wbname.split('.')

stBaseName = 'List'
stCreateSheetName = 'MasterSheet'

dtCurrentDatetime = datetime.datetime.today()

stExecuteTime = dtCurrentDatetime.strftime("%Y%m%d") 
stCreateFileName = wbnamesplited[0] + stExecuteTime + '.xlsx'

try:
    wb = openpyxl.load_workbook(wbname)
except:
    print('The file does not exist')

wb.create_sheet(stCreateSheetName)

wsBaseSheet = wb.get_sheet_by_name(stBaseName)
wsCreateSheet = wb.get_sheet_by_name(stCreateSheetName)

intBaseMaxRow = wsBaseSheet.max_row
intBaseMaxColumn = wsBaseSheet.max_column

tbBase = wsBaseSheet.rows
ltUseList = []

for c in tbBase:
    for r in c:
        if r.column == 'F':
            if r.value == '使用':
                ltUseList.append(c)

try:
    wb.save(stCreateFileName)
    print('The file was saved')
except:
    print('The file can be opend')

