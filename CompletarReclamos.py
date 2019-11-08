import pyodbc
from os import listdir
from os.path import isdir

tablaReclamos = 'reclamos_por_usuario'

def getDb(folder):
	c = 1
	dbs = []
	for file in listdir(folder):
		if file.endswith('.accdb') or file.endswith('.mdb'):
			print('{}) {}'.format(c,file))
			dbs.append(file)
			c+=1
	choice = int(input('Base de datos:> '))
	db = dbs[choice-1]
	return folder.replace('\\','/')+'/'+db
	
def getDir():
	global dir
	while True:
		dir = input('Carpeta:> ')
		if not isdir(r'{}'.format(dir)):
			print('No existe el directorio ingresado.')
			continue
		else: break
	if not dir: dir = getcwd()
	
def exportarTabla(listaReclamos):
	with open('{}_completa.csv'.format(tablaReclamos),'w') as db:
		db.write('NumerodeOrden;NumerodeReclamo;FechayHoradeReclamo;NumerodeSuministro;NombreUsuario;Domicilio;Departamento;VillaoBarrio;NumeroSeta;MotivoReclamo;CodigoTrabajo;TrabajoRealizado;FechaHoraInicioAtencion;FechaHoraFinAtencion;FechaHoraLlamado;Recepcionista;Descripción Falla;nombre_contratista1;Fecha Pasado a 1;nombre_contratista2;Fecha derivado a 2;OAC;Distribuidor;Interrupcion;Id_rpu\n')
		for reclamo in listaReclamos:
			string = ';'.join((str(x) for x in reclamo))+'\n'
			string = string.replace('None','')
			string = string.replace('.0;',';')
			db.write(string)

def getReclamos(db):
	dictReclamos = {}
	conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' %db)
	cursor = conn.cursor()
	query = 'SELECT * from {}'.format(tablaReclamos)
	cursor.execute(query)
	listaReclamos = cursor.fetchall()
	for reclamo in listaReclamos:
		if reclamo[1] not in dictReclamos:
			dictReclamos[reclamo[1]] = []
		dictReclamos[reclamo[1]].append(reclamo)
	return dictReclamos
	
if __name__ == '__main__':
	getDir()
	db = getDb(dir)
	dictReclamos = getReclamos(db)
	listaReclamos = []
	for idReclamo in dictReclamos:
		if len(dictReclamos[idReclamo]) > 1:
			if idReclamo == 679283.0:
				pass
			lRec = dictReclamos[idReclamo]
			primerReclamo = sorted(lRec,key=lambda x:len(tuple((y for y in x if y))),reverse=True)[0]
			listaReclamos.append(primerReclamo)
			lRec.remove(primerReclamo)
			for reclamo in lRec:
				indexDato = 3
				for dato in reclamo[3:]:
					if not dato and indexDato not in [14,24]:
						reclamo[indexDato] = primerReclamo[indexDato]
					indexDato+=1
						
				listaReclamos.append(reclamo)
		else: listaReclamos.append(dictReclamos[idReclamo][0])
	exportarTabla(listaReclamos)
	