def alarma(tiempo=0):
	from time import sleep
	sleep(tiempo)
	while True:
		print('\a'*1)
		sleep(1)

if __name__ == '__main__':
	print('Tiempo (min):> ')
	tiempo = abs(int(input())*60)
	alarma(tiempo)