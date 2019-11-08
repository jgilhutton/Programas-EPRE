from Interrupcion import Interrupcion
from Variables import *
from Usuario import Usuario
from multiprocessing import Pool, cpu_count
from time import time,mktime,strptime,strftime,sleep
from math import ceil
from os.path import isfile
from os import remove
from random import shuffle
import pyodbc
import pickle
import openpyxl

# Hacer las consultas a la base de datos y organizar la informacion lleva
# muchisimo tiempo y, si surge algún error en el prorgama, tendrìamos que volver a 
# hacer todo eso.
# Para evitar esto, una vez que se generan los diccionarios necesarios, el programa 
# los guarda en estos archivos para no tener que volver a generarlos y hacer las consultas a la db.
totalPickle = 'total.pickle'
IRMpickle = 'IRM.pickle'
IUMpickle = 'IUM.pickle'
Upickle = 'U.pickle'
Tpickle = 'T.pickle'

# Otras variables
choices = {'yes':['y','s','','si','yes'],'no':['n','no']}
cantidadNucleos = cpu_count()
fechaCambioUnitarios = mktime(strptime(fechaCambioUnitarios,'%d/%m/%Y'))

def salir(*args):
	print(' '.join([str(x) for x in args]))
	exit()
	
def blink(text):
	from time import sleep
	try:
		while True:
			print('{}\r'.format(text),end='')
			sleep(1)
			print(' '*len(text)+'\r',end='')
			sleep(1)
	except KeyboardInterrupt: exit()

def conectarConDB():
	"""
	Crea una conexion con la base de datos dbInterrupciones
	"""
	try:
		conexion = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' %dbInterrupciones)
		return conexion.cursor()
	except Exception as e:
		salir('Hubo un error al intentar conectarse con la base de datos "{}":\n{}'.format(dbInterrupciones,e))
		return None

def getInterrupciones(cursor):
	"""
	Trae la IRM de la db
	"""
	query = 'SELECT DISTINCT * FROM [{}];'.format(tablaIRM)
	try: cursor.execute(query)
	except:
		salir('ERROR\nHubo un error al intentar obtener las interrupciones\nRevisa si existe la tabla y si el nombre es "{}"'.format(tablaIRM))
	return cursor

def getCambioDeTarifas(cursor):
	"""
	Trae los cambios de tarifas
	"""
	query = 'SELECT DISTINCT * FROM [{}];'.format(tablaT)
	try: cursor.execute(query)
	except:
		print('ERROR\nHubo un error al intentar obtener los cambios de tarifas\nRevisa si existe la tabla y si el nombre es "{}"'.format(tablaT))
		return None
	return cursor
	
def getUsuariosAfectados(cursor):
	"""
	Devuelve todos los usuarios afectados por interrupciones.
	"""
	cursorUsuariosAfectados = cursor
	query = 'SELECT DISTINCT Suministro,Interrupcion,Orden_Reposicion FROM [{}];'.format(tablaIUM)
	try: cursorUsuariosAfectados.execute(query)
	except:
		salir('ERROR\nHubo un error al intentar obtener los usuarios afectados\nRevisa si existe la tabla y si el nombre es "{}"'.format(tablaIUM))
	return cursorUsuariosAfectados
		
def getDataSuministros(cursor):
	"""
	Devuelve la tabla de suministros con los datos de c/un
	"""
	cursorUsuarios = cursor
	query = 'SELECT DISTINCT Suministro,Nombre,Numero,Calle,Departamento,Tiposervicio,Tarifa,Consumo_anual,Seta FROM [{}];'.format(tablaUsuarios)
	try: cursorUsuarios.execute(query)
	except:
		salir('ERROR\nHubo un error al intentar obtener los datos de los usuarios\nRevisa si existe la tabla y si el nombre es "{}"'.format(tablaUsuarios))
	try:
		return cursorUsuarios
	except:
		print(suministro)
		exit()		

def printProgreso(total,contador,id):
	"""
	Imprime en pantalla el porcentaje completado del trabajo
	"""	
	porcentaje = round((contador/total)*100,2)
	print('Thread {}: {}% Completado.\r'.format(id,str(porcentaje)),end='')
	
def aCSV(fileName,thread,datos):
	"""
	Exporta todos los datos procesados a un archivo de extension csv
	los campos están separados por punto y coma: ";"
	"""
	with open('%s%s.csv'%(fileName,thread),'w') as d:
		for fila in datos:
			d.write(';'.join([str(x).replace('.',',') for x in fila])+'\n')

def procesar(suministrosAfectados,threadNo):
	datos = [] # Aquí va toda la burundanga
	datosENS = []
	interrupcionesNoEncontradas = []

	# Se fija si están los pickles y carga los archivos.
	with open(IRMpickle, 'rb') as irm:
		try:
			interrupcionesReposicionesMensual = pickle.load(irm)
		except EOFError:
			remove(IRMpickle)
	
	# Lo mismo que recien pero con la IUM
	with open(IUMpickle,'rb') as ium:
		try:
			intUsuarios = pickle.load(ium)
		except EOFError:
			remove(IUMpickle)

	# Lo mismo que recien pero con la tabla de usuarios
	with open(Upickle,'rb') as du:
		try:
			dataSuministros = pickle.load(du)
		except:
			remove(Upickle)

	with open(Tpickle,'rb') as t:
		try:
			cambioDeTarifas = pickle.load(t)
		except:
			remove(Tpickle)

	total = len(suministrosAfectados)	
	
	contador = 1
	for suministro in suministrosAfectados:
		datosLocal = []
		multa = 0.0
		indexCambio = 0	
		try:
			interrupcionesDelUsuario = intUsuarios[suministro]
			if procesarEstasInterrupciones:
				interrupcionesDelUsuario = list(filter(lambda x: x['Interrupcion'] in procesarEstasInterrupciones,interrupcionesDelUsuario))
				if not interrupcionesDelUsuario: continue
			# tomo los datos del suministro,
			dataSuministro = dataSuministros[suministro]
		except KeyError:
			continue
		# y con esos datos creo un OBJETO Usuario.
		# Para màs informacion ver el archivo Usuario.py
		usuario = Usuario(dataSuministro)
		# Ordeno los cortes segun el inicio
		interrupcionesDelUsuario.sort(key = lambda x: int(x['Interrupcion'][2:]))

		# para cada interrupcion que afecta al usuario
		for corte in interrupcionesDelUsuario:
			# Tomo el id y el orden de reposicion
			# Ej: id = MT|BT.... orden_reposicion = 2
			id,oRepo = corte['Interrupcion'],corte['Orden_Reposicion']
			# Con el id, oRepo y los datos del corte, armo un OBJETO Interrupcion
			# Pàra más información ver el archivo Interrupcion.py
			try: interrupcion = Interrupcion(interrupcionesReposicionesMensual[id][oRepo],id,oRepo,fechaCambioUnitarios)
			except KeyError:
				if id not in interrupcionesNoEncontradas:
					interrupcionesNoEncontradas.append(id)
					print('No se encontraron datos sobre la int. {}'.format(id))
				continue
			# Calculo los minutos por banda horaria del corte. Esta funcion esta en el archivo Interrupcion.py
			interrupcion.minutosPorBandaHoraria()

			# saco el factor fa. 0.33,0.66,1.00
			if interrupcion.tipoTension == 'B':
				if usuario.tipoServicio == 'M' and interrupcion.faseCorte == 'M': fa = 1.0/3.0
				elif usuario.tipoServicio == 'M' and interrupcion.faseCorte == 'B': fa = (1.0/3.0)*2.0
				else: fa = 1.0
			elif interrupcion.tipoTension == 'M':
				if usuario.tipoServicio == 'M' and interrupcion.faseCorte == 'M': fa = 1.0/3.0
				else: fa = 1.0

			# Calculo la energia no suministrada que va a ir a la tabla final
			ENS = usuario.calcularENS(interrupcion,fa)
			usuario.ENS += ENS

			datosENS.append((usuario.suministro,
							interrupcion.nombre,
							interrupcion.ordenReposicion,
							interrupcion.motivoEPRE,
							interrupcion.penalizable,
							interrupcion.duracion/60.0,
							ENS,
							))

			# Si no es penalizable paso a la siguiente. Solo proceso las que penalizan
			if not interrupcion.penalizable: continue

			# sumo la duracion al total de tiempo y fa al total de cortes
			usuario.totalT += interrupcion.duracion 		# total tiempo sin servicio
			usuario.totalTpen += interrupcion.duracion*fa 	# Este tiempo es el que se tiene en cuenta para la multa
			usuario.totalQ += fa 							# total cortes
			
			# Manejo los eventuales cambios de tarifa del suministro
			flag = True
			if flag:
				sumiStr = str(usuario.suministro)+'_'+str(indexCambio)
				if cambioDeTarifas.__contains__(sumiStr):
					if mktime(interrupcion.inicioEPRE) > cambioDeTarifas[sumiStr]['fecha']:
						try:
							ptaje = (usuario.totalT-interrupcion.duracion)/usuario.limT
							usuario.setTarifa(int(cambioDeTarifas[sumiStr]['tNueva']))
							usuario.totalTpen = ptaje*usuario.limT
							usuario.totalTpen += interrupcion.duracion
							cambioDeTarifas.pop(sumiStr)
						except ZeroDivisionError: pass
					else:
						usuario.setTarifa(int(cambioDeTarifas[sumiStr]['tVieja']))
			else: flag = False

			# si no penaliza previamente
			if not usuario.penaliza:
				# si supera la cantidad limite de cortes
				if usuario.totalQ > usuario.limQ:
					# Si pasa de 5.3 a 6.3, se debe tomar solo la diferencia entre el 6.3 y el limite de 6.
					# Lo mismo para los otros limites (4 o 3)
					diff = usuario.totalQ-usuario.limQ
					# sumo duracion al tiempo penalizado
					usuario.tPenalizado += interrupcion.duracion*diff
					# Declaro que el usuario penaliza asi falla la comparacion de la linea 186
					usuario.penaliza = True
					# calculo la multa
					multa = usuario.calcularMulta(interrupcion,fa)*fa*diff
				# si supera el limite de tiempo
				elif usuario.totalTpen > usuario.limT:
					# calculo pa porcion penalizable del corte restandole el limite de tiempo
					tiempo = usuario.totalTpen - usuario.limT
					# fin e inicio en segundos
					fin = mktime(interrupcion.finEPRE)
					inicio = fin-tiempo
					# recalculo los minutos por banda horaria del corte con el nuevo inicio
					minutosBHRecalculados = interrupcion.recalcularMinBH(inicio,fin)
					interrupcion.minutosBH = minutosBHRecalculados
					# sumo duracion al tiempo penalizado
					usuario.tPenalizado += tiempo
					# Declaro que el usuario penaliza asi falla la comparacion de la linea 186
					usuario.penaliza = True
					# calculo la multa y se la sumo a la variable local "multa"
					multa = usuario.calcularMulta(interrupcion,fa)
				# si no supera los limites, no hace nada
				else: pass			
			# si el usuario ya penaliza
			elif usuario.penaliza:
                # sumo duracion al tiempo penalizado
				usuario.tPenalizado += interrupcion.duracion
				# calculo la multa y se la sumo a la variable local "multa"
				multa = usuario.calcularMulta(interrupcion,fa)

			# le paso el valor parcial local de la multa al objeto Usuario
			usuario.multa += multa

			# Guardo todo en la lista de datos que van a ir a la tabla final
			datosLocal.append((
				str(usuario.suministro).zfill(11),
				usuario.tarifa,
				usuario.seta,
				interrupcion.nombre,
				float(interrupcion.ordenReposicion),
				strftime('%d/%m/%y %H:%M:%S',interrupcion.inicioEPRE),
				strftime('%d/%m/%y %H:%M:%S',interrupcion.finEPRE),
				float(interrupcion.duracion/60.0),
				float(ENS),
				float(usuario.valorizacion) if interrupcion.unitarios == 'Nuevos' else float(usuario.valorizacionPrevia),
				float(multa),
				float(usuario.totalQ),
				float(usuario.limQ),
				float(usuario.limT/60.0),
				float(usuario.totalT/60.0),
				float(usuario.totalTpen/60.0),
				float(usuario.tPenalizado/60.0),
				usuario.tipoServicio,
				float(usuario.consumoAnual),
								   ))
		
		# le paso el valor local FINAL de "multa" al objeto Usuario asi mantenemos todo en un mismo lugar 
		usuario.multa = multa
		
		# Guardo todos los datos que van a ir a la tabla final
		if usuario.penaliza:
			for i in datosLocal:
				datos.append(i)
		
		# imprimo el progreso en pantalla cada 1000 suministros procesados
		if contador%1000 == 0:
			printProgreso(total,contador,threadNo)
		
		contador += 1

	print('Thread',threadNo,': 100% Completado  ')
	if datos:
		aCSV('datos',threadNo,datos)
	if datosENS:
		aCSV('datosENS',threadNo,datosENS)

	return True
	

if __name__ == '__main__':
	print("""
		########################################################
		Cualquier problema con este programa ya saben qué hacer:
		\t-Juan Ignacio Gil-Hutton
		\t-Tel: 264 5067132
		\t-mail: jgilhutton@gmail.com
		########################################################

		""")
	input('1) Comprobar los unitarios')
	input('2) Verificar la variable "fechaCambioUnitarios"')
	input('3) Borrar los pickles si hiciste algun cambio en las tablas')
	input('4) Fijarse si la variable "procesarEstosSuministros" sea la deseada')
	input('5) Fijarse si la variable "procesarEstasInterrupciones" sea la deseada')
	input('6) De-comentar la variable "limitesTemporales" si queres usar limites=0')
	input('ENTER para empezar...')
	######################################
	# Armo las estructuras de datos aqui #
	# para que sea mucho mas raido el    #
	# programa. Manjeo pikles y traigo   #
	# toda la informacion de la DB       #
	# Se fija si están los pickles y carga los archivos.
	# Si no están, los crea. Esto va a tardar un poquito.
	if not isfile(IUMpickle):
		print('Generando pickle para IUM...\r',end='Generando pickle para IUM... ')
		ints = getUsuariosAfectados(conectarConDB())
		ints = ints.fetchall() # [[sum,int,orepo],...]
		intUsuarios = {}  # {'12345678912':[{'Interrupcion':int,'Orden_Reposicion':orepo},...], ...}
		for sum in ints:
			if not intUsuarios.__contains__(sum[0]): intUsuarios[sum[0]] = []
			d = dict(zip(['Interrupcion','Orden_Reposicion'],sum[1:])) # {'Interrupcion':int,'Orden_Reposicion':orepo}
			intUsuarios[sum[0]].append(d)
		with open(IUMpickle,'wb') as ium:
			pickle.dump(intUsuarios,ium)
		print('OK')
	else:
		print('Usando pickle previo para IUM')
		with open(IUMpickle,'rb') as ium:
			try:
				intUsuarios = pickle.load(ium)
			except EOFError:
				remove(IUMpickle)
				salir('Quedó un pickle viejo corrupto. Ya lo eliminé. Echá a andar el programa devuelta')

	if not isfile(Upickle):
		print('Generando pickle para Usuarios...\r',end='Generando pickle para Usuarios... ')
		dataS = getDataSuministros(conectarConDB())
		dataS = dataS.fetchall()
		dataSuministros = {}
		for dSum in dataS:
			dataSuministros[dSum[0]] = dict(zip(['Suministro','Nombre','Numero','Calle','Departamento','Tiposervicio','Tarifa','Consumo_anual','Seta'],dSum))
		with open(Upickle,'wb') as du:
			pickle.dump(dataSuministros,du)
		print('OK')
	else: print('Usando pickle previo para Usuarios')

	if not isfile(IRMpickle):
		print('Generando pickle para IRM...\r',end='Generando pickle para IRM... ')
		interrupcionesReposicionesMensual= getInterrupciones(conectarConDB())
		interrupcionesReposicionesMensual = interrupcionesReposicionesMensual.fetchall()
		ints = tuple(set([x[0] for x in interrupcionesReposicionesMensual]))
		interrupcionesReposicionesMensual = dict([(int,dict([(x[1],dict(zip(['Dispositivo_Operado_A','Dispositivo_Operado_C','Inicio_ESJ','Fin_ESJ','Inicio_EPRE','Fin_EPRE','Distribuidor','Motivo_ESJ','Motivo_EPRE','Descripcion_EPRE','Motivo_Interno','Nivel','Administracion','Tipo_Elemento','Usuarios_FS','Usuarios_Repuestos','Descripcion','Fase','Computable','Localidad','Id'],x[2:]))) for x in interrupcionesReposicionesMensual if x[0] == int])) for int in ints])
		with open(IRMpickle,'wb') as irm:
			pickle.dump(interrupcionesReposicionesMensual,irm)
		print('OK')
	else: print('Usando pickle previo para IRM')

	if not isfile(Tpickle):
		print('Generando pickle para Cambio de Tarifas...\r',end='Generando pickle para Cambio de Tarifas... ')
		cambioDeTarifas = getCambioDeTarifas(conectarConDB())
		tempDict = {}
		if cambioDeTarifas:
			cambioDeTarifas = cambioDeTarifas.fetchall()
			sums = tuple(set([x[0] for x in cambioDeTarifas]))
			for sumi in sums:
				indexCambio = 0
				for x in cambioDeTarifas:
					if x[0] == sumi:
						tempDict['_'.join([str(sumi),str(indexCambio)])] = dict(zip(['fecha','tVieja','tNueva'],[(x[1]-x[1].utcfromtimestamp(0)).total_seconds(),x[2],x[3]]))
						indexCambio += 1
		print('OK')
		with open(Tpickle,'wb') as t:
			pickle.dump(tempDict,t)
	else: print('Usando pickle previo para Cambio de Tarifas')
	######################################
	######################################

	# Saco una lista con todos los suministros que fueron afectados por interrupciones.
	# Los que no tuvieron cortes no están en esta lista.
	suministrosAfectados = list(intUsuarios.keys())

	if procesarEstosSuministros:
		suministrosAfectados = procesarEstosSuministros		

	total = len(suministrosAfectados)

	# Muy importante desordenar la lista de suministros
	shuffle(suministrosAfectados)
	# Como los suministros nuevos tienden a tener menos cortes que los viejos,
	# los ultimos threads, al tener todos estos suministros nuevos, van a terminar antes
	# y desperdiciarán tiempo. Al desordenar la lista, los suministros se distribuyen mas uniformemente
	# y los procesos terminan más o menos al mismo tiempo
	# Si no me creen, borren esta linea y vean lo que pasa. Asombroso, de verdad, muy muy curioso

	inicio = time()
	print("""
		Usando factor de inversion = {}
		Usando factor de estimulo = {}
		Total suministros afectados = {}
		""".format(factorDeInversion,factorDeEstimulo,total))
	
	# Divido la lista de suministros en partes del mismo largo
	# La cantidad de partes depende de la cantidad de procesadores que tenga la máquina.
	frameSize = ceil(total/cantidadNucleos)
	print('Tamaño de parte:',frameSize)
	print()
	print('Procesanding...')
	frames = [suministrosAfectados[i:i + frameSize] for i in range(0, total, frameSize)]

	# Armo los procesos y les paso la informacion a cada uno
	with Pool(cantidadNucleos) as p:
		args = tuple(zip([frame for frame in frames],[id for id in range(cantidadNucleos)]))
		res = p.starmap(procesar, args)

	# Aqui ya va a estar todo procesado. Los procesos exportaron sus datos a archivos csv
	# Abro cada archivo y le anexo todas las filas a la variable "datos"
	datos = ''
	datosENS = ''
	for i in range(cantidadNucleos):
		try:
			with open('datos%s.csv'%i, 'r') as d:
				datos += d.read()
			remove('datos%s.csv'%i)
		except FileNotFoundError: continue
	for i in range(cantidadNucleos):
		try:
			with open('datosENS%s.csv'%i, 'r') as d:
				datosENS += d.read()
			remove('datosENS%s.csv'%i)
		except FileNotFoundError: continue

	print()
	# Con toda la informacion en la variable "datos", abrio un archivo csv nuevo y le echo toda
	# la burundanga ahi. Después abrir esta burundanga con el Access e importarla ahi
	sliced = 1
	if len(datos.split('\n')) > 1000000:
		datos = datos.split('\n')[:1000000]
		datos1 = datos.split('\n')[1000000:]
		sliced = 2
	for i in range(sliced):
		with open('data{}.csv'.format(i),'w') as dump:
			print('Generando data.csv{}...\r'.format(i),end='Generando data{}.csv... '.format(i))
			dump.write('Suministro;Tarifa;Seta;Interrupcion;Orden de Reposicion;Inicio;Fin;Duracion;ENS usuario;Valorizacion;Multa;Cantidad de cortes;Lim Frecuencia;Lim Tiempo;Tiempo sin servicio;Tiempo*fa;Tiempo Penalizado;Tipo de usuario;Consumo anual\n')
			if i == 0:
				dump.write(datos.replace('.',','))
			elif i == 1:
				dump.write(datos1.replace('.',','))
			print('OK')
	sliced = 1
	if len(datosENS.split('\n')) > 1000000:
		datosENS0 = datosENS.split('\n')[:1000000]
		datosENS1 = datosENS.split('\n')[1000000:]
		sliced = 2
	else: datosENS0 = datosENS.split('\n')
	for i in range(sliced):
		with open('dataENS{}.csv'.format(i),'w') as dump:
			print('Generando dataENS{}.csv...\r'.format(i),end='Generando dataENS{}.csv... '.format(i))
			dump.write('Suministro;Interrupcion;Orden_Reposicion;Motivo_EPRE;Penalizable;Duracion;ENS\n')
			if i == 0:
				datosENS0 = '\n'.join(datosENS0)
				dump.write(datosENS0.replace('.',','))
			elif i == 1:
				datosENS1 = '\n'.join(datosENS1)
				dump.write(datosENS1.replace('.',','))
			print('OK')
		
	fin = time()-inicio
	mins = fin / 60
	print()
	print('Tiempo transcurrido:',round(mins,2),'minutos.')
	blink('CTRL+C para terminar')

# COLORÍN COLORADO, ESTO LLEVÓ MUCHO TIEMPO MÁS QUE EL ESPERADO