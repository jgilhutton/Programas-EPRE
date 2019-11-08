import win32api
import win32print
from time import sleep
from os import walk
from os.path import getmtime

if __name__ == '__main__':
	dirs = {1:'Notas 25m',2:'Notas Avance de Obra',3:'Notas D. Respuesta',4:'Notas Respuesta a CPP'}
	print("""
1) Notas 25m
2) Notas Avance de Obra
3) Notas Debida Respuesta
4) Notas CPP
5) Otro
	""")
	opt = int(input(':> '))
	if opt == 5: dir = input('Carpeta:> ')
	else: dir = dirs[opt]

	archivos = []
	for root, folder, files in walk(dir):
		archivos = files
		break
	archivos.sort(key = lambda x: getmtime(dir+'/'+x))

	nCopias = input('Numero de copias:> ')
	for file in archivos:
		print(file)
		for _ in range(int(nCopias)):
			win32api.ShellExecute(0,"print",dir+'/'+file,None,".",0)
			sleep(3)