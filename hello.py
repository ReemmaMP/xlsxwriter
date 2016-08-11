from datetime import datetime
import xlsxwriter

workbook = xlsxwriter.Workbook('hello.xlsx') #Creates the workbook
worksheet1 = workbook.add_worksheet() #Defaults to Sheet1
worksheet2 = workbook.add_worksheet('Reemma') #Sheet2 is now called Reemma

bold = workbook.add_format({'bold': True}) #Adds bold format
money_format = workbook.add_format({'num_format': '$#,##0'}) #adds number format
date_format = workbook.add_format({'num_format': 'mmmm d yyyy'}) #add the excel date format

worksheet1.set_column(1,2,15)

#Headers for the dummy data
worksheet1.write('B1', 'Item', bold)
worksheet1.write('C1', 'Date', bold)
worksheet1.write('D1', 'Cost', bold)

#Some random data
expenses = (
	['Oyster card', '2016-07-27', 100],
	['Pizza', '2016-07-17', 50],
	['Macbook', '2016-06-27', 1000],
)

row = 1
col = 1

#Iterates over the dummy data and writes it out each row at a time
for item, date_str, cost in (expenses):
	date = datetime.strptime(date_str, "%Y-%m-%d") #we need to convert the date string into a datetime object

	worksheet1.write(row, col, item)
	worksheet1.write(row, col +1, date, date_format)
	worksheet1.write(row, col +2, cost, money_format)
	row += 1

#Outputs a total
worksheet1.write(row, 1, 'Total', bold)
worksheet1.write(row, 3, '=SUM(C2:C4)', money_format)

worksheet2.write('A1', 'Reemma')

workbook.close()