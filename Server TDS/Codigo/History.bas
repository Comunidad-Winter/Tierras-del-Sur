Attribute VB_Name = "History"
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''History''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'[Wizard 04/09/05]
'*Arreglado el bug, de que con el F8 podias trabajar
'sin stamina.
'*Se actualizan los nicks, al entrar o salir de un
'clan y al hacerte crimi o ciuda.
'*Los consejeros no comercian via comercio seguro
'*Los gms dejan log, al comerciar con usuarios.
'*Con el comando /LEERPARTY, veras los PMSG que
'envian los miembros, dejas de leer con /CALLARPARTY
'*Arreglado el bug que al comenzar las mujeres altas
'se veian con el body de las ropas de hombre.
'*Agregue la vestimenta comun de gnomas y enanas
'*Arregle el bug de q no empezaban con la animacion
'de la daga.
'*Cambie el Slot donde recivis las pociones rojas
'asi no se sobreponen a las azules que te daban
'si eras clase magica, ahora ya da rojas y azules.
'*Ahora el onlinegm deja logs.
'*Al crear un portal ya no crea una pocion en la pos
'target.
'*En los chars guarda 2 IPS, que se actualizan solo
'si son diferentes.
'*Los logs de los gms, ahora se encuentran en la car
'peta GMS, y la de consejeros en su intenrior
'*Ahora si llueve no recuperras stamina aunque estes
'desnudo o con hambre y sed en 0, como pasaba.
'*La estupidez y el veneno, se van al morir.
'*El pescador tiene la navegacion sin restriccion de
'nivel
'*Al meditar, si recivis daño, dejas de meditar.
'*Al meditar, recivis hechizos que no son de daño
'no ves su animacion, ya que en teoria deberian
'ser tapadas por el aura.
'*Ya no se puede remover paralisis a criaturas Para
'lizadas.
'*Los newbies si salen del newbie dungeon se les cae
'el oro al morir.
'*Los elementales de fuego y tierra, no afectan a
'las criaturas.
'*Arreglado el bug que al golpear a un usuario, el
'atacante suba tacticas de combate, en vez de hacer
'lo la victima
'*Arreglados los horrores de redaccion de los subs
'de enlistar a faccion que producian errores de en-
'trega
'*El /ira, si estas invisible va mas lejos para no
'chocar
'*Si una criatura esta inmovilizada o paralizada
'lo dice al cliquearla.
'
'*************************07/09/05 Wizard*************
'*Arreglado el bug que cuando deslogueas, muerto y
'navegando, cuando logueas te ves vivo.-
'Para esto movi toda entrega de Bodys al conectar
'hacia el LoadUserInit, ya que se le daba el body 2
'veces.
'*Ahora al morir se actualiza el inventario.
'*Cambie algunos mensajes en el /retirar
'*Posible arreglo en el spawn de las criaturas mari
'nas, le di un Optional al sub ClosestLegalPos, para
'que este le diera al "LegalPos" si el agua era vali
'da, esto podria arreglar el No-Spawn, y arregle el
'Spawn en tierra de los npc marinos, agregando que
'se fije que haya agua si el aguavalida esta OK
'en un function cuyo nombre no recuerdo:P
'*Ahora cuando domas una criatura, guarda el PrevMap
'osea el mapa donde la domaste, y eso se manda al
'Crear el npc como optional, y si es <> a 0, entonce
'respawndea en el Mapa donde lo domaron y no donde
'murio.
'************************08/09/05 Wizard**************
'*Ahora los GMS, pueden atraversar cualquier portal
'no importa que nivel sean.
'*El pescador, pesca 1 de cada 10 peces raro, y si
'navega son 1 de cada 5. La cantidad se guarda en un
'flag llamado pececitos.
'************************13/09/05 Wizard**********************
'*Arregle el bug que cuando te reviven en el agua,
'caminas sobre la misma.
'*Arregle que podias tirate provocar hambre a ti
'mismo.
'*Ahora para golpear un Npc no hostil deberas tener
'el seguro desactivado, y al matarlo dara puntos de
'bandido y no de asesino.
'*Al dañar con elementales a una criatura, dara exp
'por golpe.
'*Agregue la variable FRIO a los mapas, esta designa
'si en el mapa al estar desnudo te restara vida.
'*Los clanes ahora son REALES, CAOS , o NEUTROS:
'El cliente al fundar clan, da para elegir 3 chk
'esto setea la Alineacion, puede ser:
'1-Neutro
'2-Real
'3-Caos
'0-No eligio y salta error en el cliente:D:D:D
'................
'*El robar ahora, nececita stamina y el intervalo es
'el del atacar.
'*La variable SKILLM nos dice cauntos skills nececita
'para usar tal objeto.(SKills en Magia)
'*Para atacar con arco es necesario estar en modo
'combate.
'*El MakeUserchar, ahora envia 4 si el user pertenece
'a algun consejo.- El cliente sabra si es bander o
'caos fijandose si es crimi o ciuda;)
'Tambien envia q se vean diferente al clik y ahora
'dice BANDERBILL y no BANDERBILLE ja:P
'
'
'
'
'


