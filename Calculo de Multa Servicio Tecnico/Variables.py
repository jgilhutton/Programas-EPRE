procesarEstosSuministros = []
procesarEstasInterrupciones = []
# limitesTemporales = {'A':{'descripcion':'Usuarios en AT/ST','cantidad':0,'duracion':0},
		# 'B':{'descripcion':'Usuarios en MT	4','cantidad':0,'duracion':0},
		# 'C':{'descripcion':'Usuarios en BT (pequenas y medianas demandas)','cantidad':0,'duracion':0},
		# 'D':{'descripcion':'Usuarios en BT(grandes demandas)','cantidad':0,'duracion':0},}
		
# El "tipo" es una letra [A,B,C,D] que corresponde a su valor en la variable "limites"
# El "Ki" es una tarifa que corresponde a su valor en la variable "Ki"
tarifas = {	8012:{'nombreExtendido':'T1R1-CONS (Consorcios) Espacios Verdes-Bombas de Edificios','tipo':'C','Ki':'T1-R','nombre':'T1R1-CONS'},
			8013:{'nombreExtendido':'T1R2-CONS (Consorcios) Espacios Verdes-Bombas de Edificios','tipo':'C','Ki':'T1-R','nombre':'T1R2-CONS'},
			8014:{'nombreExtendido':'T1R3-CONS (Consorcios) Espacios Verdes-Bombas de Edificios','tipo':'C','Ki':'T1-R','nombre':'T1R3-CONS'},
			8100:{'nombreExtendido':'Consumo bimestral inferior o igual a 220 k','tipo':'C','Ki':'T1-R','nombre':'T1-R1'},
			8101:{'nombreExtendido':'Consumo bimestral mayor a 220 kWh y hasta 580 kWh','tipo':'C','Ki':'T1-R','nombre':'T1-R2'},
			8102:{'nombreExtendido':'Consumo bimestral inferior o igual a 240 k','tipo':'C','Ki':'T1-G','nombre':'T1-G1'},
			8103:{'nombreExtendido':'Consumo bimestral mayor a 240 kWh y hasta 580 kWh','tipo':'C','Ki':'T1-G','nombre':'T1-G2'},
			8104:{'nombreExtendido':'Alumbrado Publico','tipo':'C','Ki':'T1-AP','nombre':'T1-AP'},
			8106:{'nombreExtendido':'T3 Baja Tension','tipo':'D','Ki':'T3-BT','nombre':'T3-BT'},
			8114:{'nombreExtendido':'Consumo bimestral mayor a 580 kWh','tipo':'C','Ki':'T1-R','nombre':'T1-R3'},
			8115:{'nombreExtendido':'Uso residencial - Consumo bimestral mayor a 580 kWh','tipo':'C','Ki':'T1-G','nombre':'T1-G3'},
			8116:{'nombreExtendido':'T2 Sin Medicion de Potencia (10 kWh a 20 kWh)','tipo':'C','Ki':'T1-G','nombre':'T2-SMP'},
			8117:{'nombreExtendido':'T2 Con Medicion de Potencia (20 kWh a 50 kWh)','tipo':'C','Ki':'T2','nombre':'T2-CMP'},
			8118:{'nombreExtendido':'T3 Media Tension con uso de red','tipo':'B','Ki':'T3-MT','nombre':'T3-MT-13.2-R'},
			8119:{'nombreExtendido':'T3 Media Tension sin uso de red','tipo':'B','Ki':'T3-MT','nombre':'T3-MT-13.2-B'},
			8120:{'nombreExtendido':'T3 Media Tension (no interesa el tipo de conexion)','tipo':'B','Ki':'T3-MT','nombre':'T3-MT-33'},
			8121:{'nombreExtendido':'Riego Agricola Sin Diferimiento','tipo':'D','Ki':'T3-BT','nombre':'TRA-RSD'},
			8122:{'nombreExtendido':'Riego Agricola Con Diferimiento','tipo':'D','Ki':'T3-BT','nombre':'TRA-RCD'},
			8124:{'nombreExtendido':'T4 Baja Tension se aplica para peaje','tipo':'D','Ki':'T3-BT','nombre':'T4-BT'},
			8125:{'nombreExtendido':'T4-MT-13.2-R','tipo':'B','Ki':'T4-MT','nombre':'T4-MT-13.2-R'},
			8127:{'nombreExtendido':'T4-MT-33','tipo':'B','Ki':'T4-MT','nombre':'T4-MT-33'},
			8128:{'nombreExtendido':'T4-MT-132','tipo':'A','Ki':'T4-MT','nombre':'T4-AT'},
			8129:{'nombreExtendido':'Tarifa edificio propio - tipo G','tipo':'C','Ki':'T1-G','nombre':'TEDP-G'},
			8130:{'nombreExtendido':'Tarifa edificio propio - T2 CMP','tipo':'C','Ki':'T2','nombre':'TEDP2-CMP'},
			8131:{'nombreExtendido':'Tarifa edificio propio - T3 BT','tipo':'D','Ki':'T3-BT','nombre':'TEDP3-BT'},
			8136:{'nombreExtendido':'Alumbrado Publico','tipo':'C','Ki':'T1-AP','nombre':'T1-AP'},
			8138:{'nombreExtendido':'T4-MT-132','tipo':'A','Ki':'T4-AT','nombre':'T4-AT'},
			}

# True = PENALIZA || False = NO PENALIZA
MotivosEPRE = {	'CAE_MT':False,
				'CSC_BT': False,
				'CSC_MT': False,
				'FBT': True,
				'FCL': False,
				'FCL_BT': False,
				'FCL_MT': False,
				'FEX': True,
				'FIU_BT': False,
				'FIU_MT': False,
				'FMT': True,
				'FOD': False,
				'FOT_BT': False,
				'FOT_EX': False,
				'FOT_MT': False,
				'FRA_BT': False,
				'FRA_MT': False,
				'FSI_BT': False,
				'FSI_MT': False,
				'MAN': True,
				'MAN_BT': True,
				'MAN_MT': True,
				'MFA_BT': True,
				'MFA_MT': True,
				'MNP_BT': True,
				'MNP_MT': True,
				'MRC_BT': True,
				'MRC_MT': True,
				'OCI_BT': True,
				'OCI_MT': True,
				'SUS_BT': False,
				'SUS_MT':False,
				'82_CV1':True,
				'82_CV2':True,
				'74_CV1':True,
				'74_CV2':True,}
		
limites = {'A':{'descripcion':'Usuarios en AT/ST','cantidad':3.0,'duracion':120*60.0},
		'B':{'descripcion':'Usuarios en MT	4','cantidad':4.0,'duracion':180*60.0},
		'C':{'descripcion':'Usuarios en BT (pequenas y medianas demandas)','cantidad':6.0,'duracion':600*60.0},
		'D':{'descripcion':'Usuarios en BT(grandes demandas)','cantidad':6.0,'duracion':360*60.0},}

Ki = {
	'T1-AP':[2.40,2.40,2.40,2.40,2.40,1.25,0,0,0,0,0,0,0,0,0,0,0,0,0,1.20,2.40,2.40,2.40,2.40],
	'T1-G':[.48,.48,.44,.44,.52,.81,.97,1.16,1.37,1.46,1.53,1.50,1.37,1.37,1.37,1.33,1.34,1.12,1.03,.96,.79,.79,.70,.63],
	'T1-R':[.85,.66,.50,.50,.50,.50,.59,.71,1.01,1.27,1.30,1.18,1.18,1.18,1.05,1.05,1.05,1.11,1.23,1.69,1.93,1.23,.99,.78],
	'T2':[.82,.82,.82,.82,.82,.82,.82,1.02,1.14,1.14,1.11,1.11,1.34,1.34,1.34,1.34,1.34,1.17,.73,.87,.87,.82,.82,.82],
	'T3-AT':[.65,.65,.63,.63,.67,.81,.89,1.09,1.25,1.30,1.32,1.30,1.36,1.36,1.36,1.33,1.34,1.15,.88,.92,.83,.80,.76,.73],
	'T3-BT':[.82,.82,.82,.82,.82,.82,.82,1.02,1.14,1.14,1.11,1.11,1.34,1.34,1.34,1.34,1.34,1.15,.73,.87,.87,.82,.82,.82],
	'T3-MT':[.65,.65,.63,.63,.67,.81,.89,1.09,1.25,1.30,1.32,1.30,1.36,1.36,1.36,1.33,1.34,1.15,.88,.92,.83,.80,.76,.73],
	'T4-AT':[.65,.65,.63,.63,.67,.81,.89,1.09,1.25,1.30,1.32,1.30,1.36,1.36,1.36,1.33,1.34,1.15,.88,.92,.83,.80,.76,.73],
	'T4-BT':[.82,.82,.82,.82,.82,.82,.82,1.02,1.14,1.14,1.11,1.11,1.34,1.34,1.34,1.34,1.34,1.15,.73,.87,.87,.82,.82,.82],
	'T4-MT':[.65,.65,.63,.63,.67,.81,.89,1.09,1.25,1.30,1.32,1.30,1.36,1.36,1.36,1.33,1.34,1.15,.88,.92,.83,.80,.76,.73],
	'T4-AT':[.65,.65,.63,.63,.67,.81,.89,1.09,1.25,1.30,1.32,1.30,1.36,1.36,1.36,1.33,1.34,1.15,.88,.92,.83,.80,.76,.73],
	'TRA':[.82,.82,.82,.82,.82,.82,.82,1.02,1.14,1.14,1.11,1.11,1.34,1.34,1.34,1.34,1.34,1.15,.73,.87,.87,.82,.82,.82],}

factorDeInversion = 0.7
factorDeEstimulo = 1
factorViento1 = 0.4
factorViento2 = 0.65

fallasCL = {'82_CV1':factorViento1,
			'82_CV2':factorViento2,
			'74_CV1':factorViento1,
			'74_CV2':factorViento2,}

fechaCambioUnitarios = '23/01/2019'
dbInterrupciones = ''
tablaIRM = 'interrupciones_reposiciones_mensual' 
tablaIUM = 'interrupciones_por_usuario_mensual'
tablaUsuarios = 'datos comerciales del usuario'
tablaT = 'cambio_de_tarifa_por_usuario'
tablaUnitarios = 'Unitarios.xlsx'

dbInterrupciones = dbInterrupciones.replace('\\','/')

def cargarUnitarios():
	import openpyxl
	from time import mktime,strptime
	global unitarios,unitariosPrevios
	tarifas = ['T1-R','T1-G','T1-AP','T2','T3-BT','T3-MT','T3-AT','T4-MT','T4-AT'] # Mismo orden en el que aparecen en el archivo Unitarios.xlsx
	unitariosPrevios = {}
	unitarios = {}
	unitariosWB = openpyxl.load_workbook(tablaUnitarios)
	unitariosWS = unitariosWB['Energia NO Suministrada']

	col = 1
	while True:
		try:
			fecha = unitariosWS.cell(1,col).value.strftime('%d/%m/%Y')
			if fecha == fechaCambioUnitarios:
				col -= 1
				break
		except AttributeError: pass
		col += 1
		
	indexTarifa = 0
	for row in unitariosWS:
		if type(row[col].value) == float:
			unitarios[tarifas[indexTarifa]] = row[col].value
			unitariosPrevios[tarifas[indexTarifa]] = row[col-1].value
			indexTarifa += 1
cargarUnitarios()