Attribute VB_Name = "History"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''History''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'[Wizard 04/09/05]
'Bueno, termine con la lista de bugs paso a escribir
'los cambios en el cliente:
'--------------------------------------
'*Ahora cuando quieras poner modo combate estando
'trabajando no podras, el error era la posicion del
'paquete TAB
'*Si tenes casteado un hechizo, y desactivas el modo
'combate se descasteara, para evitar que tiren hechi
'sos fuera de modo combate.
'*Cuando te paralizan, no deberias salir dis-
'parado al removerte manteniedno las teclas de mov
'ya que puse un exit sub en el CheckKeys, es algo
'que hay que probar.
'*Usando los botones de PMSG ya no se enviaran mensa-
'jes vacios, es decir "", y cunado se envie " "(un
'espacio) lo tomara como el hablar comun, y sacara
'cartel.
'*Al escribir /MEDITAR tambien ahora estara medido
'por el intervalo
'*Ahora el mapa 106 se llama Mar.
'*Agregue 6 mensajes a la lista
'*Ya no se cuelga el socket, si pones volver al
'crear el personaje
'*El panel gm ahora ordenara los nicks en forma
'alfavetica
'*Arregle el error de que a los reales nunca les cam-
'biaba el rango.
'*ARregle lo de que cuando morias no se te iba la estu-
'pidez sin enviar paqutes por suerte.
'************************08/09/05 Wizard*************
'*La pantalla no titilara si se actualiza la consola
'aunque esten en INVENTARIO
'*Arreglados errores ortograficos en los mensajes
'(Eso lo hizo el word digo Nacho)
'*Puse que los nombres de los del consejo se vean dif
'*Negue la entrada con doble ao:P
'*Agregado que con la W, manda el comando /Trabajando
'*Agregue el color del consejo de las sombras, pienso
'que quedaria mas lindo si se llamara Consilio:P
'255,50,0.
'*************************Marche**************************
'Agregado el cbay
'Si se toca una tecla se deja de trabajar
'Casper traspasables
'Arregle la carga dinamica. No tendria que colgar cuando se juega mucho.
'Antish.
'
