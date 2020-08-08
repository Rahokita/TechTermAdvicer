import os
import openpyxl

excel_filename = 'example.xlsx'
file_passname = os.path.dirname(__file__)
excel_filename = file_passname + '/' + excel_filename

wb = openpyxl.load_workbook(excel_filename)

#シート名を出力する
sheetnames = wb.sheetnames
print(sheetnames)
#アクティブシートを取得する
print(wb.active)

#シートからセルを取得する
WS1 = wb.get_sheet_by_name('Sheet1')
print(WS1['B1'].value)
print(WS1.cell(row=1,column=2).value)
