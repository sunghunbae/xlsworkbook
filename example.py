from xlsworkbook import ExcelWorkBook

wb = ExcelWorkBook(filename="example.xlsx")
print "number of sheets=", wb.nsheets()
print wb.sheet_name()
wb.sheet_by_index(1)
print wb.sheet_name()
wb.sheet_by_name('Sheet3')
print wb.sheet_name()
wb.sheet_by_index(0)
print wb.sheet_name()

print "number of rows=", wb.nrows()
print "number of cols=", wb.ncols()
print

wb= wb.parse()

for row in wb :
	if row['Supplier'] == 'TCI' :
		print row['Name']
