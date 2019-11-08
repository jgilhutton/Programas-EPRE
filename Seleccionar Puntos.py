import pyodbc
import DataSuministro
import openpyxl
from types import SimpleNamespace
from collections import defaultdict
from math import ceil
from re import search

redTotalDb = '//Alfredo/Servidor/ACS/Actualizacion Redes/2019/9-Setiembre 2019/Actualizacion Redes/2019-10-20 Redtotal.mdb'
radioPorcentualDesdeSeta = 30
minimaDiferenciaPorcentualEntreColasDeLinea = 50
radioDesdeSeta = 40
maximaDistanciaDeAlternativo = 100

# Porcentaje de la distancia entre el punto más alejado de la seta y la seta:
# SETA---------------------------_|_---------SUM
#                               |30%|
minimaDistanciaEntrePuntos = 0.3

def getDatos(suministro):
	puntoReclamo = None
	data = DataSuministro.query(suministro)
	if not data:
		salir('No se encontró el suministro.\nSi existe en la tabla usuarios, eliminar el archivo U.pickle de la carpeta Recursos y volver a intentar.')
	suministrosSeta = []
	connDb = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' %redTotalDb)
	cursor = connDb.cursor()
	
	query = 'SELECT CODISE,SETA_X,SETA_Y FROM setas WHERE CODISE = ?;'
	seta = cursor.execute(query,(data['Seta'],)).fetchall()[0]
	seta = SimpleNamespace(**dict(zip(('id','x','y'),(seta))))
	
	query = 'SELECT CLIENTE,CIRCUITO,CLIEN_X,CLIEN_Y FROM clientes WHERE CIRCUITO = ?;'
	puntos = cursor.execute(query,(seta.id,)).fetchall()
	for punto in puntos:
		punto = SimpleNamespace(**dict(zip(('suministro','seta','x','y'),punto)))
		if punto.suministro == suministro:
			puntoReclamo = punto
		dataPunto = DataSuministro.query(punto.suministro)
		if not dataPunto or dataPunto['Estado'] != '1' or search('AP',dataPunto['Tarifa']): continue
		punto.distancia = calcularDistancia(punto,seta)
		punto.__dict__.update(dataPunto)
		suministrosSeta.append(punto)
	
	if not puntoReclamo: salir('El suministro {} tiene una seta distinta en la tabla clientes'.format(suministro))

	return puntoReclamo,suministrosSeta,seta
	
def salir(mensaje=''):
	print(mensaje)
	input('Enter para terminar...')
	exit()
	
def calcularDistancia(punto1,punto2):
	x = punto2.x-punto1.x
	y = punto2.y-punto1.y
	distancia = (x**2+y**2)**0.5
	return distancia
	
def getAlternativo(puntos,pto):
	puntosMismoTipo = filter(lambda x: x.Tipo == pto.Tipo and x != pto,puntos)
	try: alternativo = min(puntosMismoTipo,key = lambda x: calcularDistancia(pto,x))
	except ValueError:
		alternativo = min(puntosMismoTipo,key = lambda x: calcularDistancia(pto,x))
	if calcularDistancia(alternativo,pto) > maximaDistanciaDeAlternativo:
		pto.alternativo = None
		return pto
	pto.alternativo = alternativo.suministro
	pto.alternativoTipo = alternativo.Tipo
	return pto
	
	
def checkDistancias(puntos,pto,maximaDistancia):
	if not puntos: return True
	for punto in puntos:
		dist = calcularDistancia(punto,pto)
		dRel = abs(dist / pto.distancia)*100
		minimaDistancia = maximaDistancia*minimaDistanciaEntrePuntos
		if dRel < 75 or dist < minimaDistancia: break
	else: return True
	return False
	
def getDistanciaRelativaSeta(punto):
	return ((punto.distancia-distanciaMin)/distanciaMax)*100
	
def getColasDeLinea(puntos,tipo,ptaje=50):
	puntos = [x for x in puntos if x.Tipo == tipo]
	if not puntos: return None
	ptoMasLejano = max(puntos,key = lambda x: x.distancia)
	puntos.remove(ptoMasLejano)
	colasDeLinea = []
	yield ptoMasLejano
	while True:
		ptosMasLejanos = []
		for pto in puntos:
			if calcularDistancia(ptoMasLejano,pto) > ptoMasLejano.distancia:
				ptosMasLejanos.append(pto)
		puntos = ptosMasLejanos
		if not puntos: break
		ptoLejano = max(puntos,key = lambda x: x.distancia)
		dRel = (ptoLejano.distancia / ptoMasLejano.distancia)*100
		if dRel < ptaje: break
		ptaje = ptaje/dRel*100
		ptoMasLejano = ptoLejano
		yield ptoMasLejano
		puntos.remove(ptoMasLejano)
		

def seleccionar(suministroReclamo,cantidadDePuntos,puntoTrif):
	global distanciaMax,distanciaMin
	puntoLejano,puntoCercano = None,None
	puntosSeleccionados = []
	index = 0

	puntoReclamo, puntos, seta = getDatos(suministroReclamo)
	puntoReclamo.alternativo = puntoReclamo.suministro
	puntoReclamo.alternativoTipo = puntoReclamo.Tipo
	puntosSeleccionados.append(puntoReclamo)
	if puntoReclamo.Tipo == '3': puntoTrif = True
	cantidadDePuntos -= 1
	distanciaMax = max((x.distancia for x in puntos))
	distanciaMin = min((x.distancia for x in puntos))
	colasDeLineaMono = getColasDeLinea(puntos,'1',minimaDiferenciaPorcentualEntreColasDeLinea)
	colasDeLineaTrif = getColasDeLinea(puntos,'3',minimaDiferenciaPorcentualEntreColasDeLinea)
	distRelPtoReclamo = getDistanciaRelativaSeta(puntoReclamo)

	if distRelPtoReclamo < radioPorcentualDesdeSeta or puntoReclamo.distancia < radioDesdeSeta:
		while cantidadDePuntos:
			try:
				if not puntoTrif:
					pto = colasDeLineaTrif.send(None)
					puntoTrif = True
				else: pto = colasDeLineaMono.send(None)
			except StopIteration: break
			if checkDistancias(puntosSeleccionados,pto,distanciaMax):
				pto = getAlternativo(puntos,pto)
				puntosSeleccionados.append(pto)
				cantidadDePuntos -= 1
	else:
		if cantidadDePuntos:
			puntosTipoDiferente = filter(lambda x: x.Tipo != puntoReclamo.Tipo ,puntos)
			try:
				puntoDebajoDeSeta = min(puntosTipoDiferente,key = lambda x: x.distancia)
				if not (getDistanciaRelativaSeta(puntoDebajoDeSeta) < radioPorcentualDesdeSeta and checkDistancias(puntosSeleccionados,puntoDebajoDeSeta,distanciaMax)): raise ValueError
				if puntoDebajoDeSeta.Tipo == '3' and puntoTrif: raise ValueError
			except ValueError:
				puntosMismoTipo = filter(lambda x: x.Tipo == puntoReclamo.Tipo ,puntos)
				puntoDebajoDeSeta = min(puntosMismoTipo,key = lambda x: x.distancia)

			puntoDebajoDeSeta = getAlternativo(puntos,puntoDebajoDeSeta)
			puntosSeleccionados.append(puntoDebajoDeSeta)
			if puntoDebajoDeSeta.Tipo == '3': puntoTrif = True
			cantidadDePuntos -= 1


		while cantidadDePuntos:
			try:
				if not puntoTrif:
					pto = colasDeLineaTrif.send(None)
					puntoTrif = True
				else: pto = colasDeLineaMono.send(None)
			except StopIteration: break
			if calcularDistancia(pto,puntoReclamo) < puntoReclamo.distancia or not checkDistancias(puntosSeleccionados,pto,distanciaMax): continue
			else:
				pto = getAlternativo(puntos,pto)
				puntosSeleccionados.append(pto)
				cantidadDePuntos -= 1
	
	return puntosSeleccionados
	
def mostrarEnPantalla(seleccion):
	seta = ''
	for punto in seleccion:
		if punto.Seta != seta:
			print('\nTIPO\tSUMINISTRO\tTIPO\tSETA')
		seta = punto.Seta
		print('\t'.join(('BASICO',punto.suministro,punto.Tipo,punto.Seta)))
		if punto.alternativo:
			print('\t'.join(('ALTERN',punto.alternativo,punto.alternativoTipo,punto.Seta)))

def main():
	seleccionDePuntos = []
	suministrosReclamos = []
	while True:
		sum = input('Suministro reclamo:> ')
		if not sum: break
		elif len(sum) != 11:
			print('Suministro invalido')
			continue
		else: suministrosReclamos.append(sum)
	cantidadReclamos = len(suministrosReclamos)
	print('Cantidad de puntos para la seleccion:> 16\r',end = 'Cantidad de puntos para la seleccion:> ')
	cantidadDePuntosSeleccion = input()
	if not cantidadDePuntosSeleccion: cantidadDePuntosSeleccion = 16
	else: cantidadDePuntosSeleccion = int(cantidadDePuntosSeleccion)
	print('Cantidad de puntos trifasicos:> 4\r',end = 'Cantidad de puntos trifasicos:> ')
	cantidadTrifasicos = input()
	if not cantidadTrifasicos: cantidadTrifasicos = 4
	else: cantidadTrifasicos = int(cantidadTrifasicos)
	
	maximoPuntosPorReclamo = ceil(cantidadDePuntosSeleccion/len(suministrosReclamos))
	
	trifs = [DataSuministro.query(x)['Tipo'] for x in suministrosReclamos].count('3')
	if trifs > cantidadTrifasicos:
		salir('Los suministros de reclamos suman {} puntos trifasicos.'.format(trifs))
		
	suministrosReclamos.sort(key=lambda x: int(DataSuministro.query(x)['Tipo']),reverse=True)
	
	for sum in suministrosReclamos:
		if DataSuministro.query(sum)['Tipo'] == '3' or cantidadTrifasicos <= 0: puntoTrif = True
		else: puntoTrif = False
		seleccionDePuntos += seleccionar(sum,maximoPuntosPorReclamo,puntoTrif)
		cantidadTrifasicos -= 1
	
	mostrarEnPantalla(seleccionDePuntos)
	salir('Listo.')
	

main()	