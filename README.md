# xlsworkbook
reading Excel workbook (.xls, .xlsx) in python using openpyxl and xlrd

### Usage example
```
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
```

In the above example, each row is a dictionary object with keys from the header line.

```
row = {u'Supplier': u'Samchun', u'Package': u'500 g', 
       u'Storage': u'X001', u'Name': u'potassium bicarbonate', u'No': 1}
```

### Example input (converted to .csv)
<pre>
"No","Name","Storage","Package","Supplier"
1,"potassium bicarbonate","X001","500 g","Samchun"
2,"phosphorus pentachloride","X002","500 g","Lancaster"
3,"phthalic anhydride","X003","500 g","Aldrich"
4,"propionic acid","X004","500 mL","AJAX chemical"
5,"phthalic anhydride","X005","500 g","Aldrich"
6,"phenyl ether","X006","1 kg","Acros"
7,"polyphosphoric acid","X007","1 Kg","Aldrich"
8,"Phthalimide","X008","500 g","TCI"
9,"Phosphorus tribromide","X009","500 g","Aldrich"
--- omitted ---
</pre>

### Example output
<pre>
number of sheets= 3
Sheet3
Sheet2
Sheet3
reagents
number of rows= 120
number of cols= 5

Phthalimide
colchicine
5-Hydroxy-1-indanone
6-Hydroxy-1-indanone
p-Toluenethiol
3-Pyridinecarboxaldehyde
Ethyl 5-Bromovalerate
4-Cyanobenzaldehyde
Dibenzylamine
Methyl 4-(Aminomethyl)benzoate Hydrochloride
Benzenesulfonyl Chloride
1,4-Phenylenediamine Dihydrochloride
1-Methyl piperazine
2,6-dimetnylaniline
Pyrrolidine
4-(Methylamino)pyridine
2-Thiophenecarboxamide
4-Fluorobenzamide
Isovaleraldehyde
5-Formyl-2-thiophenecarboxylic Acid
5-Methylthiophene-2-carboxaldehyde
5-Amino-1-pentanol
Methyl Trifluoromethanesulfonate
</pre>
