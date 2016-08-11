import xlsxwriter

workbook = xlsxwriter.Workbook('hello.xlsx')
worksheet1 = workbook.add_worksheet() #Defaults to Sheet1
worksheet2 = workbook.add_worksheet('Reemma')

bold = workbook.add_format({'bold': True})

money = workbook.add_format({'num_format': '$#,##0'})

worksheet1.write('A1', 'Item', bold)
worksheet1.write('B1', 'Cost', bold)

expenses = (
	['Oyster card', 100],
	['Pizza', 50],
	['Macbook', 1000],
)

row = 1
col = 1

for item, cost in (expenses):
	worksheet1.write(row, col, item)
	worksheet1.write(row, col +1, cost, money)
	row += 1

worksheet1.write(row, 1, 'Total', bold)
worksheet1.write(row, 2, '=SUM(C2:C4)', money)

worksheet2.write('A1', 'Reemma')

workbook.close()