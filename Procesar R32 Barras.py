from os import walk, getcwd
from os.path import isdir
import ECAMEC
import openpyxl

def getDir():
	while True:
		dir = input('Carpeta:> ')
		if not isdir(r'{}'.format(dir)):
			print('No existe el directorio ingresado.')
			continue
		else: break
	if not dir: dir = getcwd()
	return dir
	
def getTree(dir):
	arbol = {}
	for root,dirs,files in walk(dir):
		if root.endswith('Archivos Enviados'):
			r32s = [x for x in files if x.endswith('.R32')]
			informeExcel = [x for x in files if x.endswith('.xlsx')]
			if informeExcel:
				informeExcel = informeExcel[0]
			arbol[root+'/'] = {'informe':root+'/'+informeExcel,'r32s':r32s}
	return arbol
	
def getR32Data(informe):
	dataR32 = {}
	excel = openpyxl.load_workbook(informe)
	sheet = excel['barras']
	for row in sheet:
		if row[8].value and row[8].value.endswith('.R32'):
			r32 = row[8].value
			dataR32[r32] = {'tv':row[6].value,'ti':row[7].value}
	return dataR32
		

def main():
	dir = getDir()
	arbol = getTree(dir)
	for ruta in arbol:
		print(ruta)
		data = arbol[ruta]
		informe = data['informe']
		r32s = data['r32s']
		
		dataR32 = getR32Data(informe)
		for file in r32s:
			if file not in dataR32: continue
			args = {'rutaProcesar':ruta+file,'outputDirectory':ruta,'TV':dataR32[file]['tv'],'TI':dataR32[file]['ti']}
			ecamec = ECAMEC.Ecamec(**args)
			try:
				ecamec.procesarR32(file)
			except Exception as e:
				print(file.center('-',50))
				print(e)
				continue

main()
		
		
		