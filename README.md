# xlsworkbook
reading Excel workbook (.xls, .xlsx) in python

### example
```
from xlsworkbook import ExcelWorkBook

nb = ExcelWorkBook(filename="example.xlsx").parse()
for row in nb :
    if row['Supplier'] == 'TCI' :
        print row['Name']
```
### example output
<pre>
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
