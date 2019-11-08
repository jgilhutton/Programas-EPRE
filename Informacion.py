cortesEspeciales = ['MT1906160041',]
#################################################################
tipoTarifas = {
				'T1-AP' 		: 'BT',
				'T1-R1' 		: 'BT',
				'T1-R2' 		: 'BT',
				'T1-R3' 		: 'BT',
				'T1-G1' 		: 'BT',
				'T1-G2' 		: 'BT',
				'T1-G3' 		: 'BT',
				'T2-CMP' 		: 'BT',
				'T2-SMP' 		: 'BT',
				'T3-BT' 		: 'BT',
				'T3-BT' 		: 'BT',
				'TRA-RCD' 		: 'BT',
				'TRA-RSD' 		: 'BT',
				'T3-MT-13,2-R' 	: 'MT13,2',
				'T3-MT-13.2-R' 	: 'MT13,2',
				'T4-MT-13,2-R' 	: 'MT13,2',
				'T4-MT-13.2-R' 	: 'MT13,2',
				'T3-MT-33' 		: 'MT33',
				'T4-MT-33' 		: 'MT33',
			}
#################################################################
Ki = {
	'1R'	:[0.85,0.66,0.50,0.50,0.50,0.50,0.59,0.71,1.01,1.27,1.30,1.18,1.18,1.18,1.05,1.05,1.05,1.11,1.23,1.69,1.93,1.23,0.99,0.78],
	'1G'	:[0.48,0.48,0.44,0.44,0.52,0.81,0.97,1.16,1.37,1.46,1.53,1.50,1.37,1.37,1.37,1.33,1.34,1.12,1.03,0.96,0.79,0.79,0.70,0.63],
	'1AP'	:[2.40,2.40,2.40,2.40,2.40,1.25,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,1.20,2.40,2.40,2.40,2.40],
	'2'		:[0.82,0.82,0.82,0.82,0.82,0.82,0.82,1.02,1.14,1.14,1.11,1.11,1.34,1.34,1.34,1.34,1.34,1.17,0.73,0.87,0.87,0.82,0.82,0.82],
	'3BT'	:[0.82,0.82,0.82,0.82,0.82,0.82,0.82,1.02,1.14,1.14,1.11,1.11,1.34,1.34,1.34,1.34,1.34,1.15,0.73,0.87,0.87,0.82,0.82,0.82],
	'3MT'	:[0.65,0.65,0.63,0.63,0.67,0.81,0.89,1.09,1.25,1.30,1.32,1.30,1.36,1.36,1.36,1.33,1.34,1.15,0.88,0.92,0.83,0.80,0.76,0.73],
	'3AT'	:[0.65,0.65,0.63,0.63,0.67,0.81,0.89,1.09,1.25,1.30,1.32,1.30,1.36,1.36,1.36,1.33,1.34,1.15,0.88,0.92,0.83,0.80,0.76,0.73],
	}
#################################################################
relacionTarifas = {
					'T1-AP' 		: '1AP',
					'T1-R1'			: '1R',
					'T1-R2' 		: '1R',
					'T1-R3' 		: '1R',
					'T1-G1' 		: '1G',
					'T1-G2' 		: '1G',
					'T1-G3' 		: '1G',
					'T2-CMP' 		: '2',
					'T2-SMP' 		: '2',
					'T3-BT' 		: '3BT',
					'T3-BT' 		: '3BT',
					'TRA-RCD' 		: '3BT',
					'TRA-RSD' 		: '3BT',
					'T3-MT-13,2-R' 	: '3MT',
					'T3-MT-13.2-R' 	: '3MT',
					'T3-MT-33' 		: '3MT',
					'T4-MT-13,2-R' 	: '3MT', 
					'T4-MT-13.2-R' 	: '3MT',
					'T4-MT-33' 		: '3MT',
				}
#################################################################
valoresLimite = {	
				'Maximo porcentaje de registros fuera de rango':0.03,	# 3% del total de registros validos
				### Limites minimos y maximos para considerar registros validos ###
				'BT':{'baja':0.75,'alta':1.25},	# 165V y 275V (25%)
				'MT13,2':{'baja':0.5,'alta':1.5},	# (50%)
				'MT33':{'baja':0.5,'alta':1.5},	# (50%)
				'qBarras':{'baja':0.5,'alta':1.5},
				###################################################################
				### limites para flicker y thd ###
				'flicker':1.0,
				'thd':{'MT33':3.0,'BT':8.0,'MT13,2':8.0},
				'thdBarras':8.0,
				##################################
				### limites para determinar si los registros son penalizables ###
				'AE':{'BT':0.07,'MT13,2':0.07,'MT33':0.07},	# 7% , 7% , 7%
				'RU':{'BT':0.07,'MT13,2':0.07,'MT33':0.07},	# 7% , 5% , 5%
				'SU':{'BT':0.05,'MT13,2':0.05,'MT33':0.05},	# 5% , 5% , 5%
				'penBarras':0.07,
				#################################################################
				}
#################################################################
fechaFinRegistrosCortes = '30/09/19'
#################################################################
factoresDeInversion = {
						'ESJ'	: {'01/01/13':0.7,
								   '23/07/19':0.5},
						'DECSA'	: {'01/01/13':1},
					}
#################################################################
tablasUsuarios = {
					'ESJ'	: 'Usuarios',
					'DECSA' : 'UsuariosDECSA',
				}
#################################################################
tensionesNominales = {
						'BT'	:220.00,
						'MT13,2':7621.02,
						'MT33'	:19052.55,
					}
#################################################################			
tipoInstalacion = {
					1:'Monofasica',
					3:'Trifasica',
				}
#################################################################					
posicionesGraficos = {
						1:{
							1:'H4',
							2:'H19',
							3:'H30'},
						3:{
							1:'H4',
							2:'H18',
							3:'H25'},
					
					}
#################################################################
categorias = {
				'RU':'Rural',
				'AE':'Aerea',
				'SU':'Subterranea',
			}
#################################################################
sheets = {
			'ESJ'	:'Rdo Mediciones Usuario-Centros',
			'DECSA'	:'Hoja1',
			'BARRAS':'barras',
		}
#################################################################
choices = {
			'yes':['y','s','','si','seeeee metele que son pasteles'],
			'no' :['n','no','no gracias, ya tengo'],
		}
#################################################################
distribuidoras = {
					'ESJ'	:'ENERGIA SAN JUAN',
					'DECSA'	:'DECSA',
				}
#################################################################
cantidad = {'1':'un',
			'2':'dos',
			'3':'tres',
			'4':'cuatro',
			'5':'cinco',
			'6':'seis',
			'7':'siete',
			'8':'ocho',
			'9':'nueve',
			'10':'diez',
			'11':'once',
			'12':'doce',
			'13':'trece',
			'14':'catorce',
			'15':'quince',
			'16':'dieciseis',
		}
plazos = cantidad
#################################################################
meses = {
			'enero'			:'01',
			'febrero'		:'02',
			'marzo'			:'03',
			'abril'			:'04',
			'mayo'			:'05',
			'junio'			:'06',
			'julio'			:'07',
			'agosto'		:'08',
			'septiembre'	:'09',
			'octubre'		:'10',
			'noviembre'		:'11',
			'diciembre'		:'12',
			'diciembreplus'	:'00',
			'eneroplus'		:'13', #No le den bola a este.
		}
#################################################################
mesesRev = {
			'06':'junio',
			'12':'diciembre',
			'02':'febrero',
			'00':'diciembre',
			'03':'marzo',
			'08':'agosto',
			'11':'noviembre',
			'04':'abril',
			'01':'enero',
			'09':'septiembre',
			'13':'enero',
			'10':'octubre',
			'05':'mayo',
			'07':'julio',
		}
#################################################################
mesesR32 = {
			'1':'enero',
			'2':'febrero',
			'3':'marzo',
			'4':'abril',
			'5':'mayo',
			'6':'junio',
			'7':'julio',
			'8':'agosto',
			'9':'septiembre',
			'O':'octubre',
			'N':'noviembre',
			'D':'diciembre',
		}
#################################################################
penalizaTipo = {
				'Servicio':False,
				'Producto':False,
			}
#################################################################
tiposUsuario = {
				1:'el Usuario',
				2:'la Usuaria',
				3:'los Usuarios',
				4:'las Usuarias',
			}
#################################################################
tiposSuministro = {
					1:'el Suministro',
					2:'los Suministros',
				}
#################################################################
reResultadoNoPenalizado = '(?i)correct(?:a|o)+|no penaliza(?:da)*|NP'
#################################################################
reResultadoFallido = '(?i)fa(?:ll)?(?:ida)?'
#################################################################
resultadosMediciones = {
						'NP':False,
						'P':False,
						'F':False,
					}
#################################################################
posicionesGraficosBarras = {
							1:'H4',
							2:'H18',
							3:'H25',
						}
#################################################################
celdasPlantillaMonofasica = {'mes':'B5','año':'B6','tipo':'B7','categoria':'B8','seta':'B9','tarifa':'B10',
					'departamento':'B11','usuario':'B12','fechaInicio':'F6','horaInicio':'F7',
					'fechaFin':'F8','horaFin':'F9','archivo':'F10','suministro':'F11','direccion':'E12',
					'resultado':'D13','totalRegistros':'C16','totalRegistrosF1':'D16','totalRegistrosF2':'E16',
					'totalRegistrosF3':'F16','totalRegistrosSobretension':'C17','totalRegistrosSobretensionF1':'D17',
					'totalRegistrosSobretensionF2':'E17','totalRegistrosSobretensionF3':'F17','totalRegistrosSubtension':'C18',
					'totalRegistrosSubtensionF1':'D18','totalRegistrosSubtensionF2':'E18','totalRegistrosSubtensionF3':'F18',
					'totalRegistrosPenalizados':'C19','totalRegistrosPenalizadosF1':'D19','totalRegistrosPenalizadosF2':'E19',
					'totalRegistrosPenalizadosF3':'F19','energiaTotal':'C20','energiaTotalF1':'D20','energiaTotalF2':'E20',
					'energiaTotalF3':'F20','energiaSobretension':'C21','energiaSubtension':'C22','energiaPenalizada':'C23',
					'multaFueraDeRango':'C24','thdTotal':'C26','thdF1':'D26','thdF2':'E26','thdF3':'F26','thdFueraDeRango':'C27',
					'thdFueraDeRangoF1':'D27','thdFueraDeRangoF2':'E27','thdFueraDeRangoF3':'F27','thdPenalizable':'C28',
					'thdPenalizableF1':'D28','thdPenalizableF2':'E28','thdPenalizableF3':'F28','flicker':'C30','flickerF1':'D30',
					'flickerF2':'E30','flickerF3':'F30','flickerFueraDeRango':'C31','flickerFueraDeRangoF1':'D31',
					'flickerFueraDeRangoF2':'E31','flickerFueraDeRangoF3':'F31','flickerPenalizable':'C32','flickerPenalizableF1':'D32',
					'flickerPenalizableF2':'E32','flickerPenalizableF3':'F32','promedioTension':'C34','promedioTensionF1':'D34',
					'promedioTensionF2':'E34','promedioTensionF3':'F34','tensionMaxima':'C35','tensionMaximaF1':'D35','tensionMaximaF2':'E35',
					'tensionMaximaF3':'F35','tensionMinima':'C36','tensionMinimaF1':'D36','tensionMinimaF2':'E36','tensionMinimaF3':'F36',
					'apartamientoMaximo':'C37','apartamientoMaximoF1':'D37','apartamientoMaximoF2':'E37','apartamientoMaximoF3':'F37',
					'apartamientoPromedio':'C38','apartamientoPromedioF1':'D38','apartamientoPromedioF2':'E38','apartamientoPromedioF3':'F38','distribuidora':'C3',
					}
#################################################################
celdasPlantillaTrifasica = {'distribuidora':'C3','mes':'B5','año':'B6','tipo':'B7','categoria':'B8','seta':'B9','tarifa':'B10',
					'departamento':'B11','usuario':'B12','fechaInicio':'F6','horaInicio':'F7',
					'fechaFin':'F8','horaFin':'F9','archivo':'F10','suministro':'F11','direccion':'E12',
					'resultado':'D13','totalRegistros':'C16','totalRegistrosF1':'D16','totalRegistrosF2':'E16',
					'totalRegistrosF3':'F16','totalRegistrosSobretension':'C17','totalRegistrosSobretensionF1':'D17',
					'totalRegistrosSobretensionF2':'E17','totalRegistrosSobretensionF3':'F17','totalRegistrosSubtension':'C18',
					'totalRegistrosSubtensionF1':'D18','totalRegistrosSubtensionF2':'E18','totalRegistrosSubtensionF3':'F18',
					'totalRegistrosPenalizados':'C19','totalRegistrosPenalizadosF1':'D19','totalRegistrosPenalizadosF2':'E19',
					'totalRegistrosPenalizadosF3':'F19','energiaTotal':'C20','energiaTotalF1':'D20','energiaTotalF2':'E20',
					'energiaTotalF3':'F20','energiaSobretension':'C21','energiaSubtension':'C22','energiaPenalizada':'C23',
					'multaFueraDeRango':'C24','thdTotal':'C25','flicker':'C26','promedioTension':'C27','promedioTensionF1':'D27',
					'promedioTensionF2':'E27','promedioTensionF3':'F27','tensionMaxima':'C28','tensionMaximaF1':'D28','tensionMaximaF2':'E28',
					'tensionMaximaF3':'F28','tensionMinima':'C29','tensionMinimaF1':'D29','tensionMinimaF2':'E29','tensionMinimaF3':'F29',
					'apartamientoPromedio':'C30','apartamientoPromedioF1':'D30','apartamientoPromedioF2':'E30','apartamientoPromedioF3':'F30',
					}
#################################################################
def manejarErrores(mainFunction):
	def wrapper(*args,**kwargs):
		try:
			mainFunction(*args,**kwargs)
			input('\nEnter para terminar...')
		except (SystemExit,KeyboardInterrupt): pass
		except:
			from sys import exc_info
			from traceback import extract_tb,format_list
			print('ERROR')
			error = exc_info()
			string = ''.join(format_list(extract_tb(error[2])))
			string+= '\n{} {}'.format(error[0],error[1])
			print(string)
			print('\nIntentando mandar mail a jgilhutton@gmail.com con información del error... ',end='')
			try:
				import win32com.client
				outlook = win32com.client.Dispatch("Outlook.Application")
				newMail = outlook.CreateItem(0)
				newMail.Subject = 'ERROR en script'
				newMail.Body = string
				newMail.To = 'jgilhutton@gmail.com'
				newMail.Send()
				print('OK')
				input('Enter para terminar...')
			except:
				print('ERROR')
				print('Si querés notificar del error abrí el Outlook y volvé a ejecutar el programa, reproduciendo las condiciones exactas para generar el error')
				input('Enter para terminar...')
		exit()
	return wrapper