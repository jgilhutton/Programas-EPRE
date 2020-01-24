from shutil import copy
from shutil import SameFileError

# La lista de archivos fue eliminada por una cuestión de seguridad. Asi como los commits donde aparecen.
programas = # Lista de programas en el directorio raiz
multaServicio = # Lista de programas en el directorio "Calculo Multa servicio tecnico"
recursos = # Lista de assets en el directorio "Recursos"
			
total = sum(map(len,(recursos,multaServicio,programas)))			

def barraProgreso():
	barra = '[{}]'
	while True:
		progreso = yield
		parcial = '#'*int((progreso/total)*75)
		full = parcial.ljust(75,'-')
		print(barra.format(full)+'\r',end = '')

barra = barraProgreso()
barra.send(None)
progreso = 0
for script in programas:
	progreso += 1
	barra.send(progreso)
	try: copy(script,'.')
	except SameFileError: continue
	except FileNotFoundError as e:
		input('No se puede copiar {}. {}'.format(script,e))
for asset in recursos:
	progreso += 1
	barra.send(progreso)
	try: copy(asset,'./Recursos/')
	except SameFileError: continue
	except FileNotFoundError as e:
		input('No se puede copiar {}. {}'.format(script,e))
for script in multaServicio:
	progreso += 1
	barra.send(progreso)
	try: copy(script,'./Calculo de Multa Servicio Tecnico/')
	except SameFileError: continue
	except FileNotFoundError as e:
		input('No se puede copiar {}. {}'.format(script,e))
