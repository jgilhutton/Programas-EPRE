from os.path import isdir
import sys
import ProdTec
sys.path.append('./ECAMEC')
from ECAMEC import Ecamec

def getDir():
	while True:
		dir = input('Carpeta:> ')
		if not isdir(r'{}'.format(dir)):
			print('No existe el directorio ingresado.')
			continue
		else: break
	if not dir: dir = getcwd()
	return dir

dir = getDir()

ecamec = Ecamec.Ecamec(**{'rutaProcesar':dir,'outputDirectory':dir,'TV':1,'TI':1})
for archivo in ecamec.archivos:
	print(archivo)
	ecamec.procesarR32(archivo)
ProdTec.main(folder=dir,histo=True,imp=True)