import win32api
import win32print
from time import localtime,sleep
from docx import Document
from docx.shared import RGBColor
from rutas import paseArchivo,paseArchivoTemp,paseGG,paseGGtemp
from re import search
from os import getcwd

if __name__ == '__main__':
	plantillas = [paseGG,paseArchivo]
	plantillasTemp = [paseGGtemp,paseArchivoTemp]
	print('Pase a:\n1) Gerencia General\n2) Archivo')
	choice = int(input('>: ')) - 1 
	print()
	plantilla = Document(plantillas[choice])
	fecha = '/'.join([str(x).zfill(2) for x in localtime()[:3][::-1]])

	for parrafo in plantilla.paragraphs:
		for run in parrafo.runs:
			try:
				if str(run.element.rPr.color.val) == '0000FF':
					run.text = run.text.replace('{','')
					run.text = run.text.replace('}','')
					variable = search('(?<={)?fecha(?=})?',run.text).group()
					if variable:
						run.text = run.text.replace(variable,fecha)
						run.element.rPr.color.val = RGBColor(0x00,0x00,0x00)
						break
					else:
						pass
			except:
				pass

	pl = plantillasTemp[choice]
	cwd = getcwd().replace('\\','/') + '/'
	pl = cwd+pl
	plantilla.save(pl)

	for _ in range(int(input('Numero de copias:> '))):
		win32api.ShellExecute(0,"print",pl,None,".",0)
		sleep(4)
