from os import walk, mkdir, getcwd, rename

encabezado1 = 'Fecha\tHora\tVR\tVmxR\tVmiR\tIa R\tCos FI R\tEa R\tVS\tVmxS\tVmiS\tIa S\tCos FI S\tEa S\tVT\tVmxT\tVmiT\tIa T\tCos FI T\tEa T\tArmónica\tFlicker\tPa Tot.\tEa Tot.\tAnormalidad\n'
encabezado2 = '\t\tV\tV\tV\tA\tp.u.\tKWh\tV\tV\tV\tA\tp.u.\tKWh\tV\tV\tV\tA\tp.u.\tKWh\t%\t%\tKW\tKWh\t\n'

def arreglar(carpeta):
	files = []
	for root, dir, file in walk(carpeta):
		files = [dat for dat in file if dat.endswith('.dat')]
		break

	for dat in files:
		with open('{}/{}'.format(carpeta,dat),'r',encoding='latin1') as f:
			registros = [x.strip('\n').split('\t') for x in f.readlines()]
		
		start = 9
		if 'Demanda Horaria' in registros[7][0]:
			registros[1] = registros[1]+['\t','\t']+registros[7]
			registros[2] = registros[2]+['\t']+registros[8]
			registros[3] = registros[3]+['\t']+registros[9]
			start = 13
		else: continue

		
		nuevoDat = []
		for reg in registros[start:]:
			E = abs(float(reg[11].lstrip().strip().replace(',','.')))/3.0
			E3 = abs(float(reg[11].lstrip().strip().replace(',','.')))
			energia = '{0:.8f}'.format(E).replace('.',',')
			energia3 = '{0:.8f}'.format(E3).replace('.',',')
			
			linea = [reg[0],
			reg[1],
			reg[2],
			0,
			0,
			reg[3],
			0,
			energia,
			reg[4],
			0,
			0,
			reg[5],
			0,
			energia,
			reg[6],
			0,
			0,
			reg[7],
			0,
			energia,
			0,
			0,
			reg[8],
			energia3,
			reg[-1]]
			
			lineaStr = [str(x) for x in linea]

			nuevoDat.append('\t'.join(lineaStr))
		
		with open('{}/{}'.format(carpeta,dat),'w',encoding='latin1') as f2:
			for i in registros[:start-6]:
				f2.write('\t'.join(i)+'\n')
			f2.write(encabezado1)
			f2.write(encabezado2)
			for i in nuevoDat:
				f2.write(i+'\n')
			
if __name__ == '__main__':
	directorio = input('Carpeta:> ')
	cwd = getcwd()
	carpeta = cwd if not directorio else directorio
	carpeta = carpeta.replace('\\','/')
	arreglar(carpeta)