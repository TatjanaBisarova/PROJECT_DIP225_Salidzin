# 1.posms Bankas datu analīze S un L uzņēmumā
from openpyxl import load_workbook
wb = load_workbook('S_010123_310523.xlsx')
ws = wb['01-05.2023']
max_row = ws.max_row
summa = 0
for row in ws.iter_rows(min_row=2, max_col=3, max_row=max_row):
    for cell in row:
        if str(cell.value).startswith('B'):
             summa += float(cell.offset(column=3).value)
print(f"S_010123_310523 banka,EUR: {round(summa, 2)}")
wb.close()

wb_L = load_workbook('L_01_05.23.xlsx')
ws_L = wb_L['Banka']
max_row_L = ws_L.max_row
summa_L = 0
for row in ws_L.iter_rows(min_row=3, max_col=3, max_row=max_row_L):
  for cell in row:
    if cell.value is not None:
      if isinstance(cell.value, (int, float)):
        summa_L += cell.value
print(f"L_01_05.23 banka, EUR: {round(summa_L, 2)}")
wb_L.close()
