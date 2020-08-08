import os
import openpyxl

excel_filename = 'example.xlsx'
file_passname = os.path.dirname(__file__)
excel_filename = file_passname + '/' + excel_filename

wb = openpyxl.load_workbook(excel_filename)

#シート名を出力する(12.3.2)
sheetnames = wb.sheetnames
print(sheetnames)

#アクティブシートを取得する
print(wb.active)

#シートからセルを取得する(12.3.3)
WS1 = wb['Sheet1']
print(WS1['B1'].value)
print(WS1.cell(row=1,column=2).value)

#シートのデータサイズ（行数、列数）を取得する
print(WS1.max_row)
print(WS1.max_column)
