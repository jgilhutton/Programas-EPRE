import pyodbc
import pickle
from os.path import isfile
from rutas import dbUsuarios

def getData():
	tabla = 'Usuarios'
	connUsuarios = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' %dbUsuarios)
	cursorUsuarios = connUsuarios.cursor()
	query = 'SELECT TARIFA,NUMERO_DE_SUMINISTRO,DISTRITO,[NOMBRE_ APELLIDO],DIRECCION,TIPO,ESTADO,NUMERO_DE_CENTRO,categoria,CODIGO_ALIMENTADOR_BT FROM {}'.format(tabla)
	cursorUsuarios.execute(query,())
	cursorUsuarios = cursorUsuarios.fetchall()
	return cursorUsuarios

def query(suministro):
	try: data = tablaUsuarios[suministro]
	except KeyError: 
		if __name__ == '__main__': print('No se encontro el suministro')
		return
	if __name__ != '__main__':
		return dict(zip(['Tarifa','Distrito','Nombre y Apellido','Direccion','Tipo','Estado','Seta','Categoria','Distribuidor'],data))
	for x,y in zip(['Tarifa','Distrito','Nombre y Apellido','Direccion','Tipo','Estado','Seta','Categoria','Distribuidor'],data):
		print('{}: {}'.format(x,y))
	
if not isfile('Recursos/U.pickle'):
	print('Generando pickle para tabla de usuarios...')
	usus = getData()
	tablaUsuarios = {x[1]:[x[0]]+list(x[2:]) for x in usus}
	with open('Recursos/U.pickle','wb') as d:
		pickle.dump(tablaUsuarios,d)
	print('OK\n')
else:
	with open('Recursos/U.pickle','rb') as d:
		tablaUsuarios = pickle.load(d)

if __name__ == '__main__':
	try:
		while True:
			print('\nSuministro: XXXXXXXXXXX\r',end='Suministro: ')
			suministro = input()
			query(suministro)
	except KeyboardInterrupt: exit()
