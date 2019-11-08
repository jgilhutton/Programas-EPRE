import pip
from os import listdir

def instalar(package):
    if hasattr(pip, 'main'):
        pip.main(['install', package])
    else:
        pip._internal.main(['install', package])

if __name__ == '__main__':
	modulos = [x for x in listdir() if x.endswith('.py')]
	for modulo in modulos:
		print(modulo,end='... ')
		try: __import__(modulo.split('.')[0])
		except ImportError as error:
			paquete = error.name
			print('Intentando instalar {}'.format(paquete))
			try: instalar(paquete)
			except: print('No se pudo instalar')
		print('OK')
	input('Listo')