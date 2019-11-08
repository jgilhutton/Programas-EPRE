from os import walk, mkdir, getcwd, rename
from re import search

def getFiles(directory):
	for root, dirs, files in walk(directory):
		archivos = files
		break
	return [x for x in archivos if x[-4:] in ['.dat','.err']]

if __name__ == '__main__':
	directorio = input('Carpeta:> ').replace('\\','/')
	cwd = getcwd().replace('\\','/')
	carpeta = cwd if not directorio else directorio
	files = getFiles(carpeta)
	for file in files:
		file2 = file[:8]+file[-4:]
		origen = '{}/{}'.format(carpeta,file)
		destino = '{}/{}'.format(carpeta,file2)
		rename(origen,destino)






