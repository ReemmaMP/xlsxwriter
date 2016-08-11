from datetime import datetime
import xlsxwriter

workbook = xlsxwriter.Workbook('hello.xlsx') #Creates the workbook
worksheet1 = workbook.add_worksheet() #Defaults to Sheet1
worksheet2 = workbook.add_worksheet('Chart') #Sheet2 is now called Chart

bold = workbook.add_format({'bold': True}) #Adds bold format
money_format = workbook.add_format({'num_format': '$#,##0'}) #adds number format
date_format = workbook.add_format({'num_format': 'mmmm d yyyy'}) #add the excel date format
chart = workbook.add_chart({'type': 'column'}) #creating a chart object

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

worksheet2.write('A1', 'Simple chart')#title

#data to be plotted
somedata = [
	[1,3,5,7,9],
	[2,4,6,8,10],
	[5,7,9,3,2],
]

worksheet2.write_column('A2', somedata[0])
worksheet2.write_column('B2', somedata[1])
worksheet2.write_column('C2', somedata[2])

#adding series to configure the chart
chart.add_series({'values': '=Chart!$A$2:$A$6'})
chart.add_series({'values': '=Chart!$B$2:$B$6'})
chart.add_series({'values': '=Chart!$C$2:$C$6'})

worksheet2.insert_chart('A7', chart) #inserting the chart!!

workbook.close()
