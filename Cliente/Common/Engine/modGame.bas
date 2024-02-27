Attribute VB_Name = "CLI_Game"

Option Explicit

Public SuperWater As Boolean

Public Timer_Caminar As New clsPerformanceTimer

Public Consola_Clan             As vWControlChat

Public UserPasos As Byte
Private ultimo As Long
Private direcciones(1 To 4) As Long

Private Declare Function GetActiveWindow Lib "user32" () As Long


Sub CheckKeys()
    
    If Not UserMoving = 0 Then Exit Sub
    
    If GetActiveWindow > 0 Then
    
        If IScombate = False Then
            If frmMain.SendTxt.Visible = True Then Exit Sub
        End If
        
        If GetKeyState(vbKeyNorte) < 0 Then
            Moverme NORTH
            LastKeyPress = NORTH
            LastKeyPressTime = GetTimer
            Exit Sub
        End If
    
        'Move right
        If GetKeyState(vbKeyEste) < 0 Then
            Moverme EAST
            LastKeyPress = EAST
            LastKeyPressTime = GetTimer
            Exit Sub
        End If
    
        'Move down
        If GetKeyState(vbKeySur) < 0 Then
            Moverme SOUTH
            LastKeyPress = SOUTH
            LastKeyPressTime = GetTimer
            Exit Sub
        End If
    
        'Move left
        If GetKeyState(vbKeyOeste) < 0 Then
            Moverme WEST
            LastKeyPress = WEST
            LastKeyPressTime = GetTimer
            Exit Sub
        End If
    End If
    
    If Not MovimientoDefault = E_Heading.None Then
        Moverme MovimientoDefault
    End If


End Sub

Public Sub seMueveElPersonaje(direccion As E_Heading)
    Call Char_Move_by_Head(UserCharIndex, direccion)
             
    Call Engine_MoveScreen(direccion)
             
    Call BorrarB(direccion)
             
    UserPasos = UserPasos + 1
             
    Call actualizarMapaNombre
End Sub

Public Sub actualizarMapaNombre()
    frmMain.Coord2.Caption = "Mapa " & UserMap & " [" & UserPos.X & "," & UserPos.Y & "]"
    frmMain.MinimapUser.top = frmMain.Minimapa.top + UserPos.Y * frmMain.Minimapa.height / ALTO_MAPA
    frmMain.MinimapUser.left = frmMain.Minimapa.left + UserPos.X * frmMain.Minimapa.width / ANCHO_MAPA
End Sub

Sub Moverme(ByVal direccion As E_Heading)
     
    Dim nuevoX%, nuevoY%, PaqueteMirar%, PaqueteCaminar%, PaqueteMoverCasper%, paquete_grabadora%
    Dim BloqueadorCharIndex As Integer
    
    'Intervalo del caminar
    ' tiempo_que_tarda_en_hacer_32px = velocidad / engineBaseSpeed / 32
    If Timer_Caminar.Time(True) * CharList(UserCharIndex).Velocidad.X < 32 / engineBaseSpeed Then
        Exit Sub
    End If
    
    'Si esta haciendo determinadas acciones no se puede mover
    If Comerciando = True Or UserMeditar = True Or Bovedeando = True Or TiempoReto > 0 Then Exit Sub
    If UserDescansar = True Or UserMeditar = True Then Exit Sub
    
    'Si el personaje esta trabajando y no esta en centinela dejar de trabajar
    If Istrabajando = True And Not UserStats(SlotStats).UserCentinela Then Call modMiPersonaje.DejarDeTrabajar
    
    'No se
    If Cartel Then Cartel = False
       
    Select Case direccion
        Case E_Heading.NORTH
            nuevoX = UserPos.X
            nuevoY = UserPos.Y - 1
            
            PaqueteMirar = Paquetes.MirarNorte
            PaqueteMoverCasper = Paquetes.MNorteM
            PaqueteCaminar = Paquetes.MNorth
            paquete_grabadora = sPaquetes.MoverNorth
        Case E_Heading.EAST
            nuevoX = UserPos.X + 1
            nuevoY = UserPos.Y
            
            PaqueteMirar = Paquetes.MirarEste
            PaqueteMoverCasper = Paquetes.MEsteM
            PaqueteCaminar = Paquetes.MEast
            paquete_grabadora = sPaquetes.MoverEast
        Case E_Heading.WEST
            nuevoX = UserPos.X - 1
            nuevoY = UserPos.Y
        
            PaqueteMirar = Paquetes.MirarOeste
            PaqueteMoverCasper = Paquetes.MOesteM
            PaqueteCaminar = Paquetes.MWest
            paquete_grabadora = sPaquetes.MoverWest
        Case E_Heading.SOUTH
            nuevoX = UserPos.X
            nuevoY = UserPos.Y + 1
            
            PaqueteMirar = Paquetes.MirarSur
            PaqueteMoverCasper = Paquetes.MSurM
            PaqueteCaminar = Paquetes.MSouth
            paquete_grabadora = sPaquetes.MoverSouth
    End Select
        
     'Puede caminar hacia este tile?
     If UserStats(SlotStats).UserParalizado = False And modTriggers.PuedoCaminar(UserPos.X, UserPos.Y, direccion, UserNavegando) Then
             
         BloqueadorCharIndex = CharMap(nuevoX, nuevoY)
    
         '¿Hay un personaje donde quiero pasar?
         If BloqueadorCharIndex = 0 Then
            Timer_Caminar.Time
            
            Call EnviarPaquete(PaqueteCaminar)

            Call seMueveElPersonaje(direccion)
            
            Call CrearAccion(Chr(paquete_grabadora))
            Exit Sub
         Else
             'Esta muerto?
             If CharList(BloqueadorCharIndex).muerto Then
                 'Intento traspasarlo
                 Call EnviarPaquete(PaqueteMoverCasper)
                  
                 Timer_Caminar.Time
                 Exit Sub
             End If
         End If
     End If
     
     'No puedo caminar hacia ese tile o hay un usuario vivo
     If CharList(UserCharIndex).heading <> direccion Then
         EnviarPaquete PaqueteMirar
         'Timer_Caminar.Time
         CharList(UserCharIndex).heading = direccion
         CharList(UserCharIndex).invheading = direccion
     End If
End Sub

Public Sub Iniciar_Constantes_De_Juego()
 
ReDim NpcsMensajes(0 To 110)

NpcsMensajes(1) = "¡Hola forastero! Bienvenido a nuestra humilde aldea."
NpcsMensajes(5) = "Hola hijo mío, soy el Sacerdote de este pueblo. Si quieres ser resucitado escribe /RESUCITAR."
NpcsMensajes(6) = "Daré cárcel a todos los rivales."
NpcsMensajes(7) = "Bienvenido, rata de alcantarilla, tengo algunas armas que pueden ayudarte en tus viajes."
NpcsMensajes(8) = "Hola, jovencito, tengo algunas de las frutas más frescas y saludables de todo Argentum."
NpcsMensajes(9) = "¡Hola, tengo las mejores manzanas de la zona al mejor precio!"
NpcsMensajes(10) = "¡¿Hola, os gustaria oir una hermosa melodía?!"
NpcsMensajes(11) = "¡Hola forastero! Bienvenido a nuestra humilde aldea."
NpcsMensajes(12) = "¡Hola forastero! confeccionamos las mejores ropas de la zona."
NpcsMensajes(13) = "Hola, soy el gobernador de Ullathorpe ¡Bienvenido a nuestra pequeña aldea!"
NpcsMensajes(14) = "Hola, tengo las mejores pocimas de la zona."
NpcsMensajes(15) = "¡Bienvenido a mi taberna, viajero!"
NpcsMensajes(24) = "Hola, bienvenido a la cadena de finanzas Goliath. Somos la más grande cadena de finanzas en Argentum, contamos con bancos en las ciudades más importantes."
NpcsMensajes(33) = "¡Hola muchacho!, si deseas vender o comprar madera escribe /COMERCIAR."
NpcsMensajes(34) = "Compramos y vendemos las mejores propiedades de la zona."
NpcsMensajes(35) = "¡Hola, forastero! vendemos las mejores casas de Nix."
NpcsMensajes(36) = "Hola, bienvenido a la cadena de finanzas Goliath. Somos la más grande cadena de finanzas en Argentum, contamos con bancos en las ciudades más importantes."
NpcsMensajes(37) = "Bienvenido, rata de alcantarilla, tengo algunas armas que pueden ayudarte en tus viajes."
NpcsMensajes(38) = "Bienvenido, rata de alcantarilla."
NpcsMensajes(39) = "¡Hola forastero! confeccionamos las mejores ropas de la zona."
NpcsMensajes(40) = "Tengo algunas de las frutas más frescas y saludables de todo Argentum. Si deseas comerciar escribe /COMERCIAR."
NpcsMensajes(41) = "¡Bienvenido al gremio de pescadores de Nix!"
NpcsMensajes(42) = "¡Hola aventurero! Si estais hambriento habeis venido al lugar apropiador, tengo todo tipo de frutas."
NpcsMensajes(43) = "¿¡Estais sediento!?"
NpcsMensajes(44) = "¡Hola muchacho!, bienvenido a mi humilde Carpintería."
NpcsMensajes(45) = "¡Huye de aquí mientras puedas!"
NpcsMensajes(46) = "Bienvenido a mi pequeño negocio... Tengo algunos hechizos que podrían interesarte."
NpcsMensajes(47) = "Bienvenido a mi pequeño negocio... Tengo algunos hechizos que podrían interesarte."
NpcsMensajes(48) = "Tengo los mejores minerales de todo Argentum."
NpcsMensajes(49) = "DARÉ CÁRCEL A TODOS LOS ENEMIGOS."
NpcsMensajes(50) = "¡Hola! Bienvenido al 'Mesón Hostigado', ¿Qué deseas beber?"
NpcsMensajes(51) = "Bienvenido a mi pequeño negocio... Tengo algunos hechizos que podrían interesarte.."
NpcsMensajes(52) = "¡Hola, forastero! vendemos las mejores propiedades de Banderbill."
NpcsMensajes(53) = "Dispongo de todo tipo de provisiones."
NpcsMensajes(54) = "¡Hola forastero! Confeccionamos las mejores ropas de la zona."
NpcsMensajes(55) = "Bienvenido, rata de alcantarilla, tengo algunas armas que pueden ayudarte en tus viajes. Si estás perdido escribe /AYUDA antes de que te expulse de mi Herrería.~" & RGB(255, 0, 255)
NpcsMensajes(56) = "Bienvenido, rata de alcantarilla."
NpcsMensajes(57) = "Hola, bienvenido a la cadena de finanzas Goliath. Somos la más grande cadena de finanzas en Argentum, contamos con bancos en las ciudades más importantes.~" & RGB(255, 0, 255)
NpcsMensajes(58) = "¡Bienvenido al gremio de pescadores!"
NpcsMensajes(59) = "¡Hola muchacho!, si deseas vender o comprar madera escribe /COMERCIAR."
NpcsMensajes(60) = "¡Hola amigo! Soy el Maestro de armas de la Milicia de Banderbill, si deseas que traiga una criatura sólo debes pedirmelo. (Escribe /ENTRENAR)"
NpcsMensajes(61) = "Bienvenido a mi pequeño negocio... Tengo algunos hechizos que podrían interesarte."
NpcsMensajes(62) = "¡Hola, forastero! Vendemos las mejores propiedades."
NpcsMensajes(63) = "Hola hijo mío, soy el Sacerdote de este pueblo. Si quieres ser resucitado escribe /RESUCITAR."
NpcsMensajes(64) = "¡Hola forastero, bienvenido a nuestro pueblo!"
NpcsMensajes(68) = "¡Hola forastero, bienvenido a Lindos!"
NpcsMensajes(69) = "¡Hola buen amigo! cuento con las mejores frutas de estas comarcas, por algunas monedas de oro podrá obtener muy buenas provisiones."
NpcsMensajes(70) = "¡Hola forastero! confeccionamos las mejores ropas de la zona."
NpcsMensajes(71) = "Bienvenido, rata de alcantarilla, tengo algunas armas que pueden ayudarte en tus viajes. Si estás perdido escribe /AYUDA antes de que te expulse de mi Herrería."
NpcsMensajes(72) = "Hola buen amigo, yo soy el Rey de Banderbill. Mi función es garantizar que todo marche bien en estas tierras y combatir toda manifestación del mal. Si deseas enlistarte en mis tropas escribe /ENLISTAR."
NpcsMensajes(73) = "Daré cárcel a todos los enemigos."
NpcsMensajes(80) = "¡¡Hola buen amigo, compro y vendo todo tipo de mercancías!!"
NpcsMensajes(81) = "Ohhh hijo mío!! que la llama de la virtud arda en tu corazón!!!"
NpcsMensajes(82) = "Ohhh hijo mío!! que la llama de la virtud arda en tu corazón!!!"
NpcsMensajes(83) = "Ohhh hijo mío!! que la llama de la virtud arda en tu corazón!!!"
NpcsMensajes(84) = "Ohhh hijo mío!! que la llama de la virtud arda en tu corazón!!!"
NpcsMensajes(85) = "Ohhh hijo mío!! que la llama de la virtud arda en tu corazón!!!"
NpcsMensajes(86) = "Ohhh hijo mío!! que la llama de la virtud arda en tu corazón!!!"
NpcsMensajes(87) = "Ohhh hijo mío!! que la llama de la virtud arda en tu corazón!!!"
NpcsMensajes(98) = "Si deseas unirte a la Legión Oscura escribe /ENLISTAR."
NpcsMensajes(99) = "¡Hola forastero! confeccionamos las mejores ropas de la zona."
NpcsMensajes(100) = "¡Haz obtenido el rango máximo!."
'Look For Empty places to Send Messages stored in server.
NpcsMensajes(2) = "No puedo traer mas criaturas, mata las existentes."
NpcsMensajes(3) = "No tengo ningun interes en comerciar."
NpcsMensajes(4) = "Tienes #1 monedas de oro en tu cuenta."
NpcsMensajes(16) = "¡¡No perteneces a las tropas reales!!"
NpcsMensajes(17) = "Tu deber es combatir a los integrantes del ejército escarlata, mientras mas asesines mejor sera tu recompensa.~" & RGB(0, 100, 255)
NpcsMensajes(18) = "¡¡No perteneces a las tropas del caos!!"
NpcsMensajes(19) = "Tu deber es sembrar el terror en tus enemigos, mientras mas índigos mates mejor sera tu recompensa.~" & RGB(255, 100, 0)
NpcsMensajes(20) = "Serás bienvenido si deseas regresar."
NpcsMensajes(21) = "¡Sal de aqui bufon!"
NpcsMensajes(22) = "Ya volveras arrastrandote."
NpcsMensajes(23) = "¡¡Sal de aqui maldito escarlata!!."
NpcsMensajes(25) = "No perteneces a ninguna fuerza."
NpcsMensajes(26) = "Tienes #1 monedas de oro en tu cuenta."
NpcsMensajes(27) = "No tienes esa cantidad."
NpcsMensajes(28) = "¡¡Ya perteneces a las tropas reales, ve a combatir escarlatas!!"
NpcsMensajes(29) = "¡¡¡Maldito insolente!!!¡¡¡vete de aqui seguidor de las sombras!!!"
NpcsMensajes(30) = "Para unirte a nuestras fuerzas debes matar #1 integrantes el Ejército Escarlata, solo has matado #2."
NpcsMensajes(32) = "Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!"
NpcsMensajes(74) = "¡¡¡Bienvenido a al Ejercito Imperial!!!, aqui tienes tu armadura. Por cada centena de Escarlatas que acabes te dare un recompensa, buena suerte soldado."
NpcsMensajes(75) = "Ya perteneces a las tropas del caos."
NpcsMensajes(76) = "¡¡¡Las sombras reinaran en Argentum, largate de aqui estupido índigo!!!"
NpcsMensajes(77) = "No permitiré que ningún insecto real ingrese ¡Traidor del Rey!"
NpcsMensajes(78) = "Para unirte a nuestras tropas debes matar al menos #1 integrantes del Ejército Índigo, tu solo has matado #2."
NpcsMensajes(79) = "Para unirte a nuestras tropas debes ser al menos nivel #1."
NpcsMensajes(88) = "¡¡¡Bienvenido a al lado oscuro!!!, aqui tienes tu armadura. Si matas muchos índigos te recompensare, buena suerte soldado!."
NpcsMensajes(89) = "!!Felicitaciones has ganado!! ahora apurate a recojer los items, tienes 20 segundos antes de que la arena se cierre..."
NpcsMensajes(90) = "¡¡Ve con el banquero del pueblo!!"
NpcsMensajes(91) = "¡¡Ve con el sacerdote del pueblo!!"
NpcsMensajes(92) = "Gracias por el alago viajero, pero yo solo reino en mi hogar."
'facciones
NpcsMensajes(101) = "Para subir de rango debes matar #1, tu solo has matado #2."
NpcsMensajes(102) = "Para subir de rango debes ser al menos nivel #1."
NpcsMensajes(103) = "Para subir de rango debes pagar #1 monedas de oro. Tú no tienes esa cantidad."

ReDim MensajesCompuestos(0 To 100)

MensajesCompuestos(1) = "El Personaje #1 desea ingresar a la party.~255~200~200~1~0"
MensajesCompuestos(2) = "Debes tipear el comando /CENTINELA #1."
MensajesCompuestos(3) = "Has ganado #1 puntos de experiencia.~255~0~0~1~0~"
MensajesCompuestos(4) = "Te ha apuñalado #1 por #2.~255~0~0~1~0~"
MensajesCompuestos(5) = "Has apuñalado a #1 por #2.~255~0~0~1~0~"
MensajesCompuestos(6) = "Has apuñalado la criatura por #1." & vbCrLf & _
"Tu golpe total es de #2.~255~0~0~1~0~"
MensajesCompuestos(7) = "Su golpe total ha sido #1.~255~0~0~1~0~"
MensajesCompuestos(8) = "Tu golpe total es de #1.~255~0~0~1~0~"
MensajesCompuestos(9) = "#1 te ha quitado #2 puntos de vida.~255~0~0~1~0~"
MensajesCompuestos(10) = "Le has causado #1 puntos de daño a la criatura!.~255~0~0~1~0~"
MensajesCompuestos(11) = "Le has restaurado #1 puntos de vida a #2.~255~0~0~1~0~"
MensajesCompuestos(12) = "#1 te ha restaurado #2 puntos de vida.~255~0~0~1~0~"
MensajesCompuestos(13) = "Te has restaurado #1 puntos de vida.~255~0~0~1~0"
MensajesCompuestos(14) = "#1"
MensajesCompuestos(15) = "Has ganado #1 puntos de vida.~65~190~156~0~0"
MensajesCompuestos(16) = "Has ganado #1 puntos de mana.~65~190~156~0~0"
MensajesCompuestos(17) = "Tu golpe maximo aumento en #1 puntos.~65~190~156~0~0"
MensajesCompuestos(18) = "Tu golpe minimo aumento en #1 puntos.~65~190~156~0~0"
MensajesCompuestos(19) = "¡Has recuperado #1 puntos de mana!.~65~190~156~0~0"
MensajesCompuestos(20) = "¡Has obtenido #1 lingotes!.~65~190~156~0~0"
MensajesCompuestos(21) = "¡Has fundado el clan numero #1 de Tierras del Sur!.~65~190~156~0~0"
MensajesCompuestos(22) = "%%%%%%%%%% INFORMACION DEL HECHIZO %%%%%%%%%%" & vbCrLf & _
"Nombre: #1" & vbCrLf & _
"Descripcion: #2" & vbCrLf & _
"Skill requerido: #3" & vbCrLf & _
"Mana necesario: #4" & vbCrLf & _
"Stamina necesaria: #5" & vbCrLf & _
"%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%~65~190~156~0~0"
MensajesCompuestos(23) = "¡Cerrando...Se cerrará el juego en #1 segundos!.~65~190~156~0~0"
MensajesCompuestos(24) = "Para usar este objeto necesitas #1 skills en magia.~65~190~156~0~0"
MensajesCompuestos(25) = "Record de usuarios conectados en simultáneo." & "Hay #1 usuarios.~65~190~156~0~0"
MensajesCompuestos(26) = "Numero de usuarios: #1.~65~190~156~0~0"
MensajesCompuestos(27) = "Has sido encarcelado, deberas permanecer en la carcel #1 minutos.~65~190~156~0~0"
MensajesCompuestos(28) = "#1 te ha encarcelado, deberas permanecer en la carcel #2 minutos.~65~190~156~0~0"
MensajesCompuestos(29) = "Le has robado #1 monedas de oro a #2.~65~190~156~0~0"
MensajesCompuestos(30) = " Te han robado #1 monedas de oro.~65~190~156~0~0"
MensajesCompuestos(31) = "Has robado #1 #2.~65~190~156~0~0"
MensajesCompuestos(32) = "¡Has mejorado tu skill #1 en un punto!. Ahora tienes #2 puntos.~65~190~156~0~0"
MensajesCompuestos(33) = "#1 no tiene oro.~65~190~156~0~0"
MensajesCompuestos(34) = "Has ganado #1 skillpoints.~65~190~156~0~0"
MensajesCompuestos(35) = "Para usar este barco necesitas #1 puntos en navegacion.~65~190~156~0~0"
MensajesCompuestos(36) = "Ahora debes esperar que #1 acepte el reto.~65~190~156~0~0"
MensajesCompuestos(37) = "Abandonas la party liderada por #1.~255~200~200~1~0"
MensajesCompuestos(38) = "Durante la misma has conseguido #1 puntos de experiencia!.~255~200~200~1~0"
MensajesCompuestos(39) = "#1~255~200~200~1~0"
MensajesCompuestos(40) = "#1 ha dejado de comerciar con vos.~255~255~255~0~0"
MensajesCompuestos(41) = "Para usar este objeto necesitas #1 skills en mineria.~65~190~156~0~0"

ReDim Msgboxes(0 To 21) 'hasta en esto ahorramos!

Msgboxes(1) = "El nombre o la clave son invalidos. Por favor reingreselos."
Msgboxes(2) = "El servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde."
Msgboxes(3) = "No es posible usar mas de un personaje de la cuenta al mismo tiempo."
Msgboxes(4) = "El usuario está saliendo."
Msgboxes(5) = "Perdon, un usuario con el mismo nombre se há logueado."
Msgboxes(6) = "Error en el personaje."
Msgboxes(7) = "Se te ha prohibido la entrada a Tierras Del Sur.#."
Msgboxes(8) = "Se te ha prohibido la entrada a Tierras Del Sur."
Msgboxes(9) = "Ya existe el personaje."
Msgboxes(10) = "No se puede acceder al personaje debido a que se encuentra en Modo Candado."
Msgboxes(11) = "Servidor restringido a Administradores. Por favor intente en unos momentos."
Msgboxes(12) = "El PIN contiene caracteres invalidos."
Msgboxes(13) = "El personaje se encuentra bloqueado. Debes desbloquearlo entrando a tu cuenta premium."
Msgboxes(14) = "#" 'Mensaje en donde se puede poner cualquier cosa
' Mensajes para servidor premium
Msgboxes(15) = "Debes habilitar la creacion de este pj desde la cuenta premium."
Msgboxes(16) = "Debes ingresar el PIN y MAIL de la cuenta premium."
Msgboxes(17) = "No puedes crear personajes si tu premium está vencida."
Msgboxes(18) = "Tu cuenta no es premium y se te acabo el tiempo gratuito. Debes cargar tiempo premium o esperar al próximo mes."
Msgboxes(19) = "Máximo de (2) personajes conectados en simultaneo por cuenta alcanzado."

Msgboxes(20) = "El personaje se encuentra bloqueado. Debes desbloquearlo desde la cuenta premium."
Msgboxes(21) = "No es posible usar mas de dos personajes de la cuenta al mismo tiempo."

UserRazaDesc(1) = "Son la raza mas común y equilibrada en Argentum. Debido a sus habilidades son recomendables para casi cualquier profesión."
UserRazaDesc(2) = "Seres de poca altura, contextura rechoncha, fuertes y muy temperamentales. Sus principales características son el combate cuerpo a cuerpo."
UserRazaDesc(3) = "Son seres tranquilos de gran sabiduría y belleza, que habitan los rincones más apartados de los bosques de Argentum en una gran comunidad."
UserRazaDesc(4) = "Comprenden el más siniestro y malvado segmento de la población elfica. Debido a su gran maldad, se han desviado del sendero de la sabiduría, por lo cual han ido perdiendo su magnífica inteligencia."
UserRazaDesc(5) = "Son una raza pequeña y amistosa. Poseen una notable agilidad e inteligencia, son los mas indicadados para el uso de armas mágicas y de sortilegios de menor clase "

LogDebug "  Constantes de texto iniciadas."
End Sub
