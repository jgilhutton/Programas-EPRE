from time import strptime,localtime,mktime
from Variables import MotivosEPRE as motivosEpre
from Variables import fallasCL

class Interrupcion:
	def __init__(self,int,id,oRepo,fechaCambiounitarios):
		self.nombre 		= id
		self.ordenReposicion = oRepo
		self.inicioEPRE 	= strptime(int['Inicio_EPRE'].strftime('%d/%m/%Y %H:%M:%S'),'%d/%m/%Y %H:%M:%S')	#Tupla
		self.finEPRE 		= strptime(int['Fin_EPRE'].strftime('%d/%m/%Y %H:%M:%S'),'%d/%m/%Y %H:%M:%S')	#Tupla
		if mktime(self.finEPRE) < fechaCambiounitarios:	self.unitarios = 'Viejos'
		else: self.unitarios = 'Nuevos'
		self.duracion 		= mktime(self.finEPRE)-mktime(self.inicioEPRE)	# en segundos
		self.motivoEPRE 	= int['Motivo_EPRE']
		self.cViento		= fallasCL[self.motivoEPRE] if fallasCL.__contains__(self.motivoEPRE) else 1
		self.tipoTension 	= int['Nivel']	# Baja o Media (B,M)
		self.faseCorte 		= int['Fase']	# Monofasico, Bifasico o Trifasico (M,B,T)
		self.minutosBH 		= dict(zip(range(24),[0 for _ in range(24)])) # {0:0,1:0...}
		self.penalizable 	= (self.duracion > 180.0 and motivosEpre[self.motivoEPRE])	# True|False
		
	def __str__(self):
		return self.nombre
		
	def minutosPorBandaHoraria(self):
		fin = self.finEPRE
		duracion = self.duracion
		horaFin = fin.tm_hour
		segundosEnUltimaHora = fin.tm_min*60+fin.tm_sec
		self.minutosBH[horaFin] += segundosEnUltimaHora/60.0
		fin = localtime(mktime(fin)-segundosEnUltimaHora)
		duracion -= segundosEnUltimaHora
		while duracion >= 3600:
			tuplaFin = localtime(mktime(fin)-3600)
			fin = tuplaFin
			self.minutosBH[tuplaFin.tm_hour] += 60.0
			duracion -= 3600
		tuplaFin = localtime(mktime(fin)-duracion)
		horaFin = tuplaFin.tm_hour
		self.minutosBH[horaFin] += duracion/60.0
		
	def recalcularMinBH(self,inicio,fin):
		"""
		Hace lo mismo que la anterior pero esta sirve para los cortes en los que...
		adivinen qué.... muy bien adivinaron!!! -> para los cortes en los que hay que
		recalcular los minutos de la banda horaria
		"""
		minutosBH = dict(zip(range(24),[0 for _ in range(24)]))
		duracion = fin-inicio
		horaFin = localtime(fin).tm_hour
		segundosEnUltimaHora = localtime(fin).tm_min*60+localtime(fin).tm_sec
		minutosBH[horaFin] += segundosEnUltimaHora/60.0
		fin = fin-segundosEnUltimaHora
		duracion -= segundosEnUltimaHora
		while duracion >= 3600:
			tuplaFin = localtime(fin-3600)
			fin = mktime(tuplaFin)
			minutosBH[tuplaFin.tm_hour] += 60.0
			duracion -= 3600
		tuplaFin = localtime(fin-duracion)
		horaFin = tuplaFin.tm_hour
		minutosBH[horaFin] += duracion/60.0
		return minutosBH

		
	
		
