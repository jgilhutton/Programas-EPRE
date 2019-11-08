from os import walk
from re import search
from os.path import isfile

fallidasTXT = 'Recursos/fallidas.txt'
correctasTXT = 'Recursos/correctas.txt'
carpetasTXT = 'Recursos/carpetas.txt'

def logMala(dir):
	with open(fallidasTXT,'a+') as mala:
		mala.write(dir+'\n')

def logBuena(dir):
	with open(correctasTXT,'a+') as buena:
		buena.write(dir+'\n')

if __name__ == '__main__':
	if not isfile(carpetasTXT):
		directorio = input('Carpeta:> ')

		carpetas = []
		meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre']
		mesesRe = '|'.join(meses)
		regex = '(?i)(?P<mes>{})|(?P<dia>\d\d de (?:{}))|(?P<MT>.*mt.*|13[\.,]2)|(?P<REM>remed.*)'.format(mesesRe,mesesRe)
		for root,dirs,files in walk(directorio):
			root = root.replace('\\','/')
			if files == [] or len(list(filter(lambda x: x.endswith('.dat'),files))) == 0:
				if root in carpetas: carpetas.remove(root)
			else:
				carpetas.append(root)

		with open(carpetasTXT,'w') as c:
			c.write('\n'.join(carpetas))
		exit()
		
	with open(carpetasTXT,'r') as f:
		carpetas = [x.strip() for x in f.readlines()]
	if isfile(fallidasTXT):
		with open(fallidasTXT,'r') as malas:
			fallidas = [x.strip() for x in malas.readlines()]
		carpetas = list(filter(lambda x:x in fallidas,carpetas))
	with open(fallidasTXT,'w') as mala: pass

	import ProdTec	
	for dir in carpetas:
		print(dir)
		try: 
			op = ProdTec.main(folder=dir,histo=True,imp=False)
			if not op: logMala(dir)
			else: logBuena(dir)
		except Exception as e:
			print(e)
			logMala(dir)