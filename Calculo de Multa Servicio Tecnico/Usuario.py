from Variables import *
if 'limitesTemporales' in dir():
	limites = limitesTemporales

def salir(*args):
	print(' '.join([str(x) for x in args]))
	exit()

class Usuario:
	def __init__(self,dataSuministro):
		self.setTarifa(dataSuministro['Tarifa'])
		self.suministro 	= dataSuministro['Suministro']
		self.nombre 		= dataSuministro['Nombre']
		self.numero 		= dataSuministro['Numero']
		self.calle 			= dataSuministro['Calle']
		self.departamento 	= dataSuministro['Departamento']
		self.tipoServicio 	= dataSuministro['Tiposervicio']
		self.consumoAnual 	= dataSuministro['Consumo_anual']
		self.seta 			= str(dataSuministro['Seta'])
		self.totalQ 		= 0.0
		self.totalT 		= 0.0
		self.totalTpen		= 0.0
		self.tPenalizado	= 0.0
		self.ENS			= 0.0
		self.multa 			= 0.0
		self.penaliza		= False
		self.minutosBH		= dict(zip(range(24),[0 for _ in range(24)]))
		
	def setTarifa(self,T):
		self.idTarifa = int(T)
		try:self.tarifa 	= tarifas[self.idTarifa]['nombre']
		except KeyError: salir('No se encontró la tarifa con código %d en la lista'%self.idTarifa)
		self.tipoTarKi		= tarifas[self.idTarifa]['tipo']
		self.ki				= Ki[tarifas[self.idTarifa]['Ki']]
		self.valorizacion	= unitarios[tarifas[self.idTarifa]['Ki']]*factorDeInversion*factorDeEstimulo
		self.valorizacionPrevia	= unitariosPrevios[tarifas[self.idTarifa]['Ki']]*factorDeInversion*factorDeEstimulo
		self.limQ			= limites[self.tipoTarKi]['cantidad']
		self.limT			= limites[self.tipoTarKi]['duracion']
		self.tipoUsuario	= limites[self.tipoTarKi]['descripcion']

	def agregarInterrupcion(self,int):
		self.interrupciones.append(int)
		
	def sumarMinutosBH(self,minutosBH):
		for hora in minutosBH:
			self.minutosBH[hora] += minutosBH[hora]
			
	def calcularENS(self,interrupcion,fa):
		ENS = 0.0
		for hora in interrupcion.minutosBH:
			minutos = interrupcion.minutosBH[hora]
			ki = self.ki[hora]
			ENS += self.consumoAnual/525600*minutos*fa*ki
		return ENS
			
	def calcularMulta(self,interrupcion,fa):
		multa = 0.0
		#print(interrupcion)
		for hora in interrupcion.minutosBH:
			minutos = interrupcion.minutosBH[hora]*fa
			if not minutos: continue
			ENS = self.consumoAnual/525600*minutos
			ki = self.ki[hora]
			if interrupcion.unitarios == 'Nuevos': unitario = self.valorizacion
			elif interrupcion.unitarios == 'Viejos': unitario = self.valorizacionPrevia
			cv = interrupcion.cViento
			# Calculo final para esta banda horaria
			multa += ENS*ki*unitario*cv
		# Calculo final para toda la interrupcion
		return multa

		
	
		
