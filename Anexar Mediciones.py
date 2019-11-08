import pyodbc
import openpyxl
from re import search
from os import walk
from rutas import dbMediciones

def getResultadosESJ():
	"""
	devuelve un diccionario con el resultado que manda la distribuidora
	"""
	tmp = []
	resultadosESJ = []
	
	for root, folder, file in walk(dir):
		tmp.append(file)
	files = [x for x in tmp[0]]
	
	for f in files:
		if search('(?i)resultados mediciones',f) and f.endswith('.xlsx'):
			resultadosESJ.append(f)
	if not resultadosESJ:
		print('No se encontro el informe de resultados de la distribuidora')
		input('Enter para terminar...')
		exit()
		
	if len(resultadosESJ) == 1:
		return resultadosESJ[0]
	elif len(resultadosESJ) > 1:
		for index,archivo in enumerate(resultadosESJ,start=1):
			print('{}) {}'.format(index,archivo))
		resultadosESJ = resultadosESJ[int(input('>: '))-1]
		
	return resultadosESJ
	
def getMediciones(resultados):
	mediciones = []
	index = 0
	try:resultadosMediciones = openpyxl.load_workbook('{}/{}'.format(dir,resultados))['Rdo Mediciones Usuario-Centros']
	except KeyError: resultadosMediciones = openpyxl.load_workbook('{}/{}'.format(dir,resultados))['']
	tmp = [[cell.value for cell in row] for row in resultadosMediciones]
	for m in tmp[8:]:
		mediciones.append(tuple([index]+m[:16]+[m[18]]+m[16:18]))
		index += 1
	return mediciones
		
if __name__ == '__main__':
	conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' %dbMediciones)
	cursor = conn.cursor()
	dir = input('Carpeta> ')

	resultados = getResultadosESJ()
	mediciones = getMediciones(resultados)

	cursor.execute('DELETE [MED_REM-2003-2014_RED].* FROM [MED_REM-2003-2014_RED]',())
	cursor.commit()

	query = 'INSERT INTO [MED_REM-2003-2014_RED] (Id,[PTO-MED],[FECHA-COL],[FECHA- RET],Resultado) VALUES (?,?,?,?,?);'
	cursor.executemany(query,[[x[0],str(x[1]),x[6],x[8],x[19]] for x in mediciones])
	cursor.commit()

	query = 'SELECT [MED_REM-2003-2014_RED].[PTO-MED], [MED_REM-2003-2014_RED].[FECHA-COL], [MED_REM-2003-2014_RED].[FECHA- RET], [MED_REM-2003-2014_RED].RESULTADO, [MED_REM-2003-2014_RED].[FLIKERS-ARMONICOS], suministros_sidac.SETA, distris.NOMBRE, centros.ET, suministros_sidac.LOCALIDAD, suministros_sidac.TARIFA FROM [MED_REM-2003-2014_RED] LEFT JOIN (((suministros_sidac LEFT JOIN setas ON suministros_sidac.SETA = setas.CODISE) LEFT JOIN distris ON setas.DITRISE = distris.DISTRI) LEFT JOIN centros ON distris.CENTRO = centros.CODISE) ON [MED_REM-2003-2014_RED].[PTO-MED] = suministros_sidac.CLIENTE;'
	cursor.execute(query,())
	dump = cursor.fetchall()
	cantidadOriginal = len(dump)

	query = 'SELECT * FROM [Base de Mediciones]'
	cursor.execute(query,())
	baseDeMediciones = cursor.fetchall()
	dump = list(filter(lambda x: x not in baseDeMediciones,dump))
	cantidadNueva = len(dump)
	if cantidadNueva == 0:
		print('No hay mediciones nuevas para anexar.')
		input('Enter para terminar...')
		exit()

	query = 'INSERT INTO [Base de Mediciones] ([PTO-MED],[FECHA-COL],[FECHA- RET],RESULTADO,[FLIKERS-ARMONICOS],SETA,NOMBRE,ET,LOCALIDAD,TARIFA) VALUES (?,?,?,?,?,?,?,?,?,?);'

	print('Cantidad de mediciones nuevas:',cantidadOriginal)
	print('Se anexan {} mediciones'.format(cantidadNueva))
	cursor.executemany(query,dump)
	cursor.commit()

	print('Listo...')
	input('Enter para terminar...')