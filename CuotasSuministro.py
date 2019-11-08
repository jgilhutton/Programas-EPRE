import pyodbc
from os import walk
from os.path import isdir
from re import search

blacklist = []
data = []

def salir(mensaje=''):
	print(mensaje,'\nEnter para terminar...')
	input()
	exit()

def toCsv():
    csv = open('cuotas.csv','w', encoding = 'latin-1')
    csv.write('id_suministro; sumi_localidad; desc_cat_iva; tipo_inscripcion_ib; tipo_subinscripcion_ib; des_reparticion; descripcion_clase_ciiu; descrip_tarifa; id_documento; tipo_documento_devolucion; bimestre; cuota; importe_documento; fecha_emision; fecha_emision_cuota_anterior; id_devolucion; tipo_devolucion; importe_devolucion; nro_cuota; cuotas_pendientes; saldo_devolucion; importe_cuota_pura; importe_interes; importe_cuota_total; importe_devuelto; observacion\n')
    while True:
        data = (yield)
        for row in data:
            if not row: continue
            row = list(row)
            if len(row) == 26:
                row.append(row[-1])
                row[-2] = None
            row = '^'.join([str(x) for x in row]).replace('.',',')
            row = row.replace(';',',')
            row = row.replace('^',';')
            csv.write(row+'\n')

def getDir():
	dirs = []
	while True:
		dir = input('Carpeta:> ')
		if not dir and not dirs: salir()
		elif not dir: break
		elif not isdir(r'{}'.format(dir)):
			print('No existe el directorio ingresado.')
			continue
		else: dirs.append(dir)
	return dirs
	
def getFiles(folder):
	basesDeDatos = []
	for root,dirs,files in walk(folder):
		for file in files:
			if (file.endswith('.mdb') or file.endswith('.accdb')) and search('cuotas_devoluciones',file) and root != folder:
				basesDeDatos.append(root+'/'+file)
	return basesDeDatos
	
def getData(db,suministro):
	global blacklist
	tablasCuotas = []
	conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' %db)
	cursor = conn.cursor()
	tablas = tuple((x for x in cursor.tables()))
	
	for tabla in tablas: 
		tablaName = tabla[2]
		if search('_?cuotas_?devoluciones_?',tablaName.lower()) and tablaName not in blacklist:
			blacklist.append(tablaName)
			query = 'SELECT * FROM [{}] WHERE id_suministro = ?;'.format(tablaName)
			cursor.execute(query,(suministro,))
			dataSuministro = cursor.fetchall()
			yield dataSuministro
	

folders = getDir()
suministro = input('Suministro :> ')
volcarDatos = toCsv()
volcarDatos.send(None)
for folder in folders:
	files = getFiles(folder)
	for db in files:
		for data in getData(db,suministro):
			if data: volcarDatos.send(data)
salir()