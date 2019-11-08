from os import walk, mkdir, getcwd, rename
from re import search
import openpyxl
import xlrd
from xlutils.copy import copy
from xlwt import easyxf
from time import strptime,strftime
from rutas import informeResultadosDecsa,informeResultadosEsj

directorio = input('Carpeta:> ').replace('\\','/')
cwd = getcwd().replace('\\','/')
carpeta = cwd if not directorio else directorio
sheetInforme = 'Consulta1'
sheetDECSA = 'Hoja1'
sheetESJ = 'Rdo Mediciones Usuario-Centros'
reMeses = '(?i)enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre|semestre'
meses = {'enero':'01','febrero':'02','marzo':'03','abril':'04','mayo':'05','junio':'06','julio':'07','agosto':'08','septiembre':'09','octubre':'10','noviembre':'11','diciembre':'12'}

def getFiles(directory):
 files = []
 for root, dir, file in walk(directory):
  files.append(file)
 files = files[0]
 return files

def getExcel(archivo,informe):
	if archivo == 'decsa':
		excel = openpyxl.load_workbook('{}/{}'.format(carpeta,informe))
		sheet = excel[sheetDECSA]
	elif archivo == 'esj':
		excel = openpyxl.load_workbook('{}/{}'.format(carpeta,informe))
		sheet = excel[sheetESJ]
	elif archivo == 'informe':
		excel = xlrd.open_workbook(informe)
		excel = copy(excel)
		sheet = excel.get_sheet(0) 
	return excel,sheet
	
def getFechaESJ():
	dia = search('\d{1,2}(?= de )',carpeta).group().zfill(2)
	mes = search('(?i)(?<={} de )enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre'.format(dia),carpeta).group()
	return {'fecha':'{}/{}/20'.format(dia,meses[mes.lower()]),
			'dia':dia,
			'mes':mes}
	
def decsa(informe):
	excelDECSA,resultadosDECSA = getExcel('decsa',informe)
	excelInforme,informe = getExcel('informe',informeResultadosDecsa)
	mes = search(reMeses,carpeta)

	decremento = 0
	for i in range(4,20):
		j = i+1
		i = i - decremento
		archivo = resultadosDECSA['B%d'%j].value
		if archivo == None: break
		if search('(?i)fall',resultadosDECSA['H%d'%j].value):
			print('La medicion {} da fallida.'.format(archivo))
			decremento += 1
			continue
		informe.write(i,0,str(resultadosDECSA['C%d'%j].value).zfill(4))
		informe.write(i,1,str(resultadosDECSA['I%d'%j].value))
		informe.write(i,2,str(resultadosDECSA['J%d'%j].value))
		informe.write(i,3,0)
		informe.write(i,4,0)
		informe.write(i,5,0)
		informe.write(i,7,str(resultadosDECSA['B%d'%j].value[:8])+'.R32')
		informe.write(i,9,str(resultadosDECSA['B%d'%j].value[:8])+'.R32')
		if mes:
			informe.write(1,1,mes.group().upper())
	return excelInforme

def esj(informe):
	excelESJ,resultadosESJ = getExcel('esj',informe)
	excelInforme,informe = getExcel('informe',informeResultadosEsj)
	
	fecha = getFechaESJ()
	rowInforme = 4
	for row in range(8,300):
		if resultadosESJ['A%d'%row].value != None and resultadosESJ['C%d'%row].value == TIPO_MEDICION and resultadosESJ['Q%d'%row].value in r32s:
			if search('(?i)fall',resultadosESJ['R%d'%row].value):
				print('La medicion {} da fallida.'.format(resultadosESJ['Q%d'%row].value))
			try:
				fechaRes = resultadosESJ['D%d'%row].value.timetuple()
			except:
				fechaRes = strptime(resultadosESJ['D%d'%row].value.split()[0],'%d/%m/%Y')
			if fecha['fecha'] in strftime('%d/%m/%Y',fechaRes):
				informe.write(rowInforme,0,str(resultadosESJ['A%d'%row].value)) # SUMINISTRO
				informe.write(rowInforme,1,str(resultadosESJ['F%d'%row].value)) # FECHA INSTALACION
				informe.write(rowInforme,2,str(resultadosESJ['H%d'%row].value)) # FECHA RETIRO
				informe.write(rowInforme,3,str(str(resultadosESJ['K%d'%row].value)).replace('.',',')) # ENERGIA INICIO
				informe.write(rowInforme,4,str(str(resultadosESJ['N%d'%row].value)).replace('.',',')) # ENERGIA FIN
				informe.write(rowInforme,5,str(resultadosESJ['P%d'%row].value)) # FACTOR DE LECTURA
				informe.write(rowInforme,7,str(resultadosESJ['Q%d'%row].value)) # NOMBRE ARCHIVO
				informe.write(rowInforme,9,str(resultadosESJ['Q%d'%row].value)) # NOMBRE ARCHIVO
				rowInforme += 1
	
	informe.write(1,1,'{} de {}'.format(fecha['dia'],fecha['mes'].capitalize()))
	return excelInforme

files = getFiles(carpeta)

r32s = []
for f in files:
	if '.R32' in f:
		r32s.append(f)

DECSA = False
ESJ = False
try:
	informe = [x for x in files if search('TABLA RESULTADO (?:RE)?MEDICIONES',x)][0]
	DECSA = True
except:
	informe = [x for x in files if search('Resultados Mediciones del',x)][0]
	ESJ = True
if DECSA:
	excel = decsa(informe)
elif ESJ:
	tmp = input('Tipo de medicion:\n1)Medicion\tO\n2)Remedicion\tR\n:> ')
	if tmp == '1':
		TIPO_MEDICION = 'O'
	elif tmp == '2':
		TIPO_MEDICION = 'R'
	excel = esj(informe)

	
excel.save('%s/Informe Resultados1.xls'%carpeta)


