import os
# Se define el nombre de la carpeta o directorio a crear

for i in range(11,18):
	directorio = "/home/egiptian/Documentos/Universidad/8vo Semestre/Comunicaciones 2🤩/CiscoCourse/Cap"+str(i)
	try:
    		os.mkdir(directorio)
	except OSError:
    		print("La creación del directorio %s falló" % directorio)
	else:
    		print("Se ha creado el directorio: %s " % directorio)
