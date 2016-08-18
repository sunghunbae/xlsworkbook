from xlsworkbook import ExcelWorkBook

nb = ExcelWorkBook(filename="example.xlsx").parse()
for row in nb :
	if row['Supplier'] == 'TCI' :
		print row['Name']
