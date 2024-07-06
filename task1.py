import xlwings as xw

wb = xw.Book("TestTask1.xlsx")
sheet = wb.sheets['Sheet1']
data = sheet.range('A1').expand('table').value
headers = data[0]
status_col_index = headers.index("Status")
for row_idx, row in enumerate(data[1:], start=2):
    status = row[status_col_index]
    if status == "Done":
        sheet.range(f"A{row_idx}:C{row_idx}").color = (0, 255, 0)
    elif status == "In progress":
        sheet.range(f"A{row_idx}:C{row_idx}").color = (255, 0, 0)
wb.save()
wb.close()
