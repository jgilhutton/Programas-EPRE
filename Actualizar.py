from shutil import copy
from shutil import SameFileError

programas = ['//Fabricio/Programas Juani/Actualizar.py',
			'//Fabricio/Programas Juani/Analisis Mediciones Enviadas.py',
			'//Fabricio/Programas Juani/Anexar Mediciones.py',
			'//Fabricio/Programas Juani/ArreglaDatsDECSA.py',
			'//Fabricio/Programas Juani/Barras.py',
			'//Fabricio/Programas Juani/CompletarReclamos.py',
			'//Fabricio/Programas Juani/Cortes.py',
			'//Fabricio/Programas Juani/CortesTresAños.py',
			'//Fabricio/Programas Juani/Data para Rotura.py',
			'//Fabricio/Programas Juani/DataSuministro.py',
			'//Fabricio/Programas Juani/EliminarCeros.py',
			'//Fabricio/Programas Juani/Imprimir Directorio.py',
			'//Fabricio/Programas Juani/Indices CCSPU y TCSPU.py',
			'//Fabricio/Programas Juani/Informacion.py',
			'//Fabricio/Programas Juani/Mediciones.py',
			'//Fabricio/Programas Juani/ordenarESJ.py',
			'//Fabricio/Programas Juani/Pases.py',
			'//Fabricio/Programas Juani/Procesar Semestre.py',
			'//Fabricio/Programas Juani/ProdTec.py',
			'//Fabricio/Programas Juani/Seleccion De Puntos.py',
			'//Fabricio/Programas Juani/Seleccionar Puntos.py',
			'//Fabricio/Programas Juani/Trabajo Tedioso.py',
			'//Fabricio/Programas Juani/CuotasSuministro.py',
			'//Fabricio/Programas Juani/Expedientes Demorados.py',
			'//Fabricio/Programas Juani/rutas.py',
			'//Fabricio/Programas Juani/Procesar R32 Barras.py',
			'//Fabricio/Programas Juani/Procesar Mediciones.py']
multaServicio = ['//Fabricio/Programas Juani/Calculo de Multa Servicio Tecnico/Calcular Multa.py',
				'//Fabricio/Programas Juani/Calculo de Multa Servicio Tecnico/Variables.py',
				'//Fabricio/Programas Juani/Calculo de Multa Servicio Tecnico/Interrupcion.py',
				'//Fabricio/Programas Juani/Calculo de Multa Servicio Tecnico/Usuario.py',
				'//Fabricio/Programas Juani/Calculo de Multa Servicio Tecnico/Unitarios.xlsx']
recursos = ['//Pasante/Programas Juani Pasante/Recursos/avanceDeObra.docx',
			'//Pasante/Programas Juani Pasante/Recursos/blanco.docx',
			'//Pasante/Programas Juani Pasante/Recursos/CompactosMT.xlsx',
			'//Pasante/Programas Juani Pasante/Recursos/comparacionCortes.xlsx',
			'//Pasante/Programas Juani Pasante/Recursos/comparacionCortesBackup.xlsx',
			'//Pasante/Programas Juani Pasante/Recursos/Cortes 3 años.docx',
			'//Pasante/Programas Juani Pasante/Recursos/Cortes Semestrales.xlsx',
			'//Pasante/Programas Juani Pasante/Recursos/Cortes.docx',
			'//Pasante/Programas Juani Pasante/Recursos/Cprocesado 3 años.docx',
			'//Pasante/Programas Juani Pasante/Recursos/Cprocesado.docx',
			'//Pasante/Programas Juani Pasante/Recursos/DebidaRespuesta.docx',
			'//Pasante/Programas Juani Pasante/Recursos/IGG 1NP.docx',
			'//Pasante/Programas Juani Pasante/Recursos/IGG 2P.docx',
			'//Pasante/Programas Juani Pasante/Recursos/IGG RES.docx',
			'//Pasante/Programas Juani Pasante/Recursos/IGG Sumario Demoras.docx',
			'//Pasante/Programas Juani Pasante/Recursos/Indices.xlsx',
			'//Pasante/Programas Juani Pasante/Recursos/Informe Resultados DECSA.xls',
			'//Pasante/Programas Juani Pasante/Recursos/Informe Resultados ESJ.xls',
			'//Pasante/Programas Juani Pasante/Recursos/Mediciones.docx',
			'//Pasante/Programas Juani Pasante/Recursos/modelo25m.docx',
			'//Pasante/Programas Juani Pasante/Recursos/Mprocesado.docx',
			'//Pasante/Programas Juani Pasante/Recursos/ND 1NP.docx',
			'//Pasante/Programas Juani Pasante/Recursos/ND 2P.docx',
			'//Pasante/Programas Juani Pasante/Recursos/NotaCPP.docx',
			'//Pasante/Programas Juani Pasante/Recursos/NU.docx',
			'//Pasante/Programas Juani Pasante/Recursos/Pase Archivo Resolucion.docx',
			'//Pasante/Programas Juani Pasante/Recursos/Pase Archivo.docx',
			'//Pasante/Programas Juani Pasante/Recursos/paseArchivo.docx',
			'//Pasante/Programas Juani Pasante/Recursos/paseArchivo2.docx',
			'//Pasante/Programas Juani Pasante/Recursos/PaseGerenciaGeneral.docx',
			'//Pasante/Programas Juani Pasante/Recursos/PaseGerenciaGeneral2.docx',
			'//Pasante/Programas Juani Pasante/Recursos/Planilla Cortes.xlsx',
			'//Pasante/Programas Juani Pasante/Recursos/Plantilla Barras.xlsx',
			'//Pasante/Programas Juani Pasante/Recursos/Plantilla CMP Multas Etapa II.xlsx',
			'//Pasante/Programas Juani Pasante/Recursos/Plantilla Rotura.xlsx',
			'//Pasante/Programas Juani Pasante/Recursos/Plantilla SMP Multas Etapa II.xlsx',
			'//Pasante/Programas Juani Pasante/Recursos/PlantillaSeleccion.docx',
			'//Pasante/Programas Juani Pasante/Recursos/Resolucion.docx',
			'//Pasante/Programas Juani Pasante/Recursos/Resumen Mediciones Barras.xlsx',
			'//Pasante/Programas Juani Pasante/Recursos/Sumario Demoras.docx',
			'//Pasante/Programas Juani Pasante/Recursos/Unitarios-23-1-2018.docx',
			'//Pasante/Programas Juani Pasante/Recursos/Unitarios-23-1-2019.docx',
			'//Pasante/Programas Juani Pasante/Recursos/Unitarios-23-7-2017.docx',
			'//Pasante/Programas Juani Pasante/Recursos/Unitarios-23-7-2018.docx',
			'//Pasante/Programas Juani Pasante/Recursos/Unitarios-23-7-2019.docx',
			'//Pasante/Programas Juani Pasante/Recursos/DemoradosPlanilla.xlsx',]
			
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