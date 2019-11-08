from os import walk
from re import search
from collections import defaultdict
import xlrd
import pickle
from os.path import isfile

def getFiles():
	excelsCompactos = []
	for root,dirs,files in walk('//Alfredo/Servidor/ACS/Producto Tecnico/REGISTROS/ESJ 2ª ETAPA - Campaña de Medición'):
		for file in files:
			if search('MT compactos MT',file) and file.endswith('.xls'):
				excelsCompactos.append(root.replace('\\','/')+'/'+file)
	return excelsCompactos

def getMonYear(archivo):
	regex = '(?P<month>(?<=MT 3\d_)\d\d(?=_\d\d))(?:_)(?P<year>(?<=\d\d_)\d\d(?=.xls))'
	busqueda = search(regex,archivo)
	return busqueda.groupdict()
	
def getData(archivo):
	wb = xlrd.open_workbook(archivo)
	try:ws = wb.sheet_by_name('compactos')
	except:
		try:ws = wb.sheet_by_name('MT compactos')
		except: return None,None
		
	compactos = defaultdict(tuple)
	mesanio = ''
	c = 1
	while True:
		try:
			row = ws.row(c)
			sum = row[0].value
			nom = row[1].value
			tv  = row[10].value
			ti = row[11].value
			dataTiempo = getMonYear(archivo)
			mes = dataTiempo['month']
			anio = dataTiempo['year']
			mesanio = '{}/{}'.format(mes,anio)
			compactos[sum+';'+str(nom)] = (tv,ti)
			c += 1
		except: break
	# print(mesanios)
	return compactos,mesanio

def dump(data,mesanios):
	with open('Compactos MT Historico.csv','w',encoding = 'utf-8') as d:
		d.write('Suministro;Nombre;'+';'.join(('="{}";="{}"'.format(x,x) for x in mesanios))+'\n')
		for sum in data.keys():
			sumi,nom = sum.split(';')
			linea = ['="{}";{}'.format(sumi,nom)]
			for ma in mesanios:
				if data[sum].__contains__(ma):
					linea += [data[sum][ma][0],data[sum][ma][1]]
				else:
					linea += ['','']
			linea += ['\n']
			d.write(';'.join((str(x) for x in linea)))

compactos = defaultdict(dict)
if isfile('Recursos/compactos.pickle'):
	with open('Recursos/compactos.pickle','rb') as c:
		archivosCompactos = pickle.load(c)
else:
	archivosCompactos = getFiles()
	with open('Recursos/compactos.pickle','wb') as c:
		pickle.dump(archivosCompactos,c)

mesanios = set()

for file in archivosCompactos:
	data,ma = getData(file)
	if not data:# or '/18' not in ma or '/19' not in ma:
		continue
	mesanios = mesanios.union({ma,})
	for key in data.keys():
		valores = data[key]
		compactos[key][ma] = (valores[0],valores[1])
mesanios = sorted(mesanios,key=lambda x: x.split('/')[1]+x.split('/')[0])
dump(compactos,mesanios)	















