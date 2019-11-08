ruta = ''

import pyodbc
conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' %ruta)
cursor = conn.cursor()
tablas = [x for x in cursor.tables() if not x[2].startswith('MS')]
for tabla in tablas:
	tabla = tabla[2]
	query = 'SELECT * FROM [{}];'.format(tabla)
	data = cursor.execute(query).fetchall()
	with open(tabla+'.csv','w') as dump:
		columnas = [x[3] for x in cursor.columns() if x[2] == tabla]
		dump.write(';'.join((str(x) for x in columnas))+'\n')
		for row in data:
			dump.write(';'.join((str(x) for x in row))+'\n')