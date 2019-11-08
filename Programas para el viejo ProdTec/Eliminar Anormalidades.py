from os import walk
from re import sub

def getFiles(directorio):
	files = []
	for root, dir, file in walk(directorio):
		files.append(file)
	return [x for x in files[0] if x[-4:] == '.dat']
	

directorio = input('Carpeta:> ')
archivos = getFiles(directorio)

for file in archivos:
	with open('\\'.join([directorio,file]),'r') as fread:
		f = fread.readlines()
		f = [sub('\tA\n$','\t\n',x) for x in f]
	with open('\\'.join([directorio,file]),'w') as d:
		for line in f:
			d.write(line)