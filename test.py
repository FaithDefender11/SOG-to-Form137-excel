import openpyxl


exceltest = openpyxl.load_workbook(r'C:\Users\user\Documents\To fix grading daehan\Folder\105366130092-FORM137-9.xlsx')
sheet = exceltest.active

cell = sheet.cell(28 , 1)
if 'Merged' in str(cell):
    print(cell, 'this is merged!')
else:
    print(cell, 'this is not merged!')

print(sheet['A31'].value)
print(sheet.merged_cells.ranges[0])