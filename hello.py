import xlsxwriter

workbook = xlsxwriter.Workbook('hello.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A1', '	Expenses')

expenses = (
	['Oyster card', 100],
	['Pizza', 50],
	['Macbook', 1000],
)

row = 1
col = 1

for item, cost in (expenses):
	worksheet.write(row, col, item)
	worksheet.write(row, col +1, cost)
	row += 1

worksheet.write(row, 1, 'Total')
worksheet.write(row, 2, '=SUM(C2:C4)')

workbook.close()