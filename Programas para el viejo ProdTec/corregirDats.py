from os import walk

c = 2

def getFiles(directorio):
	files = []
	for root, dir, file in walk(directorio):
		files.append(file)
	return [x for x in files[0] if x[-4:] == '.dat']

carpeta = input('Carpeta:> ')
if not carpeta: carpeta = '.'
files = getFiles(carpeta)
for dat in files:
	with open('/'.join([carpeta,dat]),'r') as f:
		files = [x.strip('\n').split('\t') for x in f.readlines()]
	for registro in files[9:]:
		try:
			while c<15:
				valor = float(registro[c].replace(',','.'))
				if valor > 275.00:
					registro[c] = '160,00'
				c+=6
		except:
			pass
		c=2
	nuevoDat = []
	for linea in files:
		nuevoDat.append('\t'.join(linea))
	with open('/'.join([carpeta,dat]),'w') as f2:
		for linea in nuevoDat:
			f2.write(linea+'\n')