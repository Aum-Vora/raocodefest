from openpyxl import load_workbook

book = load_workbook('emails.xlsx')
sheet = book.active
rows = sheet.rows

# headers = next(rows)
headers = [cell.value for cell in next(rows)]
# print(headers)
all_rows = []

for row in rows:
    data = {}
    for title, cell in zip(headers, row):
        data[title] = cell.value

    all_rows.append(data)

print(all_rows)
