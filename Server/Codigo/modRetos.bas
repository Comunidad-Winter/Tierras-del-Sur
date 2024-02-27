Attribute VB_Name = "modRetos"
Option Explicit

#If TDSFacil Then
Private Const TIEMPO_MAXIMO = 30 'Cantidad de minutos como maximo que puede durar el reto
#Else
Private Const TIEMPO_MAXIMO = 60 'Cantidad de minutos como maximo que puede durar el reto
#End If

Private cantidadRetosActivos As Byte 'Cantidad de retos que se estan jugando en estos momentos
Public permitirOro As Boolean
Public permitirItems As Boolean
Public permitirPlantado As Boolean
Public permitir3vs3 As Boolean
Public permitirResu As Boolean
Public cantidadMaximaRetos As Byte

Public Const MIN_NIVEL = 18 'Nivel minimo
Public Const MIN_APUESTA = 5000 'Apuesta minima
Public Const MAX_APUESTA = 10000000 'Apuesta maxima que pueden realizar en oro
Public Const VALOR_RETO = 20000 'Lo que le sale a cada usuario jugar un reto

Public ACT_RETO As Boolean 'Los retos estan activados?

Public Sub iniciar()

    ACT_RETO = True
    
    ' Configuración del Sistema de Retos
    modRetos.permitirItems = True
    modRetos.permitirOro = True
    modRetos.permitir3vs3 = True
    modRetos.permitirPlantado = True
    modRetos.permitirResu = True
    
    modRetos.cantidadMaximaRetos = 18
End Sub
Private Function puedeJugarReto(personaje As User, ByRef error As String) As Boolean

    puedeJugarReto = False
                          
    'El personaje esta en otro evento el cual no esta desarrollandose?
    If Not personaje.evento Is Nothing Then
        If personaje.evento.getEstadoEvento = eEstadoEvento.Desarrollandose Then
            error = "El personaje " & personaje.Name & " se encuentra ocupado en otro evento."
            Exit Function
        End If
    End If
    
    If personaje.Stats.ELV < MIN_NIVEL Then
        error = "El personaje " & personaje.Name & " no tiene el nivel suficiente para retar."
        Exit Function
    End If

    If MapInfo(personaje.pos.map).Pk = True Then
        error = "El personaje " & personaje.Name & " no está en zona segura."
        Exit Function
    End If
    
    puedeJugarReto = True
End Function

Public Sub crear(Lider As User, infoCreacion As String)
       
    ' Parametros
    Dim modo As eModoApuesta
    Dim cantidadOroApostado As Long
    Dim cantidadParticipantes As Byte     ' El total incluyendo a todos los equipos
    Dim cantidadEquipos As Byte           ' Siempre va a ser 2
    Dim cantidadIntegrantesEquipo As Byte ' Esto me marca de cuanto es cada equipo: 1vs1, 2vs2, 3vs3.
    Dim condicionMaxItems As CondicionEventoLimiteItem  ' Para limitar las rojas
    Dim condicionSinCascoNiEscudo As CondicionEventoSinCascoEscudo
    Dim limitarRojas As Boolean
    Dim cantidadRojas As Integer
    
    ' Auxiliares
    Dim equiposInfo() As String
    Dim integrantesInfo() As String
    Dim posicionInfoEquipos As Byte

    Dim estaOnline As Boolean
    Dim error As String
    Dim presentacionString As String
    Dim equipoIntegranteID() As Long       ' Necesario para pasarselo al sistema de solicitudes del reto
    Dim UserIndex As Integer
    Dim hayError As Boolean
    Dim plantado As Boolean
    Dim valeResu As Boolean
    Dim limitarCascoYEscudo As Boolean
    
    ' Loops
    Dim loopEquipo As Byte
    Dim loopIntegrante As Byte
    Dim loopParticipante As Byte
    Dim loopIntegranteLista As Byte
    
    Dim reto As cReto
    
    ' ¿Tiene el argumento?
    If infoCreacion = "" Then Exit Sub

    ' - Parametros
    cantidadEquipos = 2
    cantidadIntegrantesEquipo = StringToByte(infoCreacion, 1)
    modo = StringToByte(infoCreacion, 2)
    plantado = StringToByte(infoCreacion, 3) = 1
    cantidadOroApostado = StringToLong(infoCreacion, 4)
    valeResu = (StringToByte(infoCreacion, 8) = 1)
    limitarRojas = (StringToByte(infoCreacion, 9) = 1)
   
    cantidadParticipantes = 2 * cantidadIntegrantesEquipo
    If limitarRojas Then
        cantidadRojas = STI(infoCreacion, 10)
    Else
        cantidadRojas = 0
    End If
    
    limitarCascoYEscudo = (StringToByte(infoCreacion, 12) = 1)
     
    ' Esto no deberia suceder, pero por las dudas
    If plantado And Not cantidadIntegrantesEquipo = 1 Then plantado = False
    If limitarCascoYEscudo And Not cantidadIntegrantesEquipo = 1 Then limitarCascoYEscudo = False
    
    posicionInfoEquipos = 13
    
    ' 1vs1, 2vs2, 3vs3
    ' - Sistema de bloqueo de retos
    If modRetos.ACT_RETO = False Then
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(59), Lider.UserIndex, ToIndex
        Exit Sub
    End If
    
    If plantado And modRetos.permitirPlantado = False Then
        EnviarPaquete Paquetes.mensajeinfo, "Temporalmente los retos plantados estan deshabilitados.", Lider.UserIndex, ToIndex
        Exit Sub
    End If
    
    If valeResu And modRetos.permitirResu = False Then
        EnviarPaquete Paquetes.mensajeinfo, "Temporalmente los retos con el hechizo resucitar estan deshabilitados.", Lider.UserIndex, ToIndex
        Exit Sub
    End If
    
    If cantidadIntegrantesEquipo = 3 And modRetos.permitir3vs3 = False Then
        EnviarPaquete Paquetes.mensajeinfo, "Temporalmente los retos 3vs3 estan deshabilitados.", Lider.UserIndex, ToIndex
        Exit Sub
    End If
    
    If modo = eModoApuesta.oro And modRetos.permitirOro = False Then
        EnviarPaquete Paquetes.mensajeinfo, "Temporalmente los retos por sólo oro estan deshabilitados. Debes jugar por oro e items.", Lider.UserIndex, ToIndex
        Exit Sub
    End If

    If (modo = oroitems Or modo = Items) And modRetos.permitirItems = False Then
        EnviarPaquete Paquetes.mensajeinfo, "Temporalmentelos retos por items se encuentran desactivados. Debes jugar retos sólo por oro.", Lider.UserIndex, ToIndex
        Exit Sub
    End If
    
    'Se fija si se pueden seguir creando retos. Si no se alcanzo el maximo de retos en simultaneo
    If cantidadRetosActivos >= cantidadMaximaRetos Then
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(62), Lider.UserIndex, ToIndex
        Exit Sub
    End If
    
    'Chequeo el tema de la apuesta
    If modo = eModoApuesta.oro Or modo = eModoApuesta.oroitems Then
        If cantidadOroApostado < MIN_APUESTA Then ' apuesta minima
            EnviarPaquete Paquetes.MensajeSimple2, Chr$(61), Lider.UserIndex, ToIndex
            Exit Sub
        ElseIf cantidadOroApostado > MAX_APUESTA Then ' apuesta maxima
            EnviarPaquete Paquetes.mensajeinfo, "La apuesta máxima son " & MAX_APUESTA & " monedas de oro.", Lider.UserIndex, ToIndex
            Exit Sub
        End If
    End If
    
    ' - Validaciones del Creador
    ' Esta en otro evento?
    If Not Lider.evento Is Nothing Then
        ' El evento esta desarrolandose?. Puede tener un reto "Preparandose"
        If Lider.evento.getEstadoEvento = eEstadoEvento.Desarrollandose Then
            EnviarPaquete Paquetes.mensajeinfo, "No puedes crear un reto si estas participando en uno.", Lider.UserIndex, ToIndex
            Exit Sub
        End If
    End If

    'Si el usuario esta muerto no puede crear un evento
    If Lider.flags.Muerto = 1 Then  'Esta muerto?
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(63), Lider.UserIndex, ToIndex
        Exit Sub
    End If

    'Tiene el nivel minimo para crear un reto ?
    If Lider.Stats.ELV < MIN_NIVEL Then  'Minimo nivel
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(60), Lider.UserIndex, ToIndex
        Exit Sub
    End If

    If MapInfo(Lider.pos.map).Pk Then
        EnviarPaquete Paquetes.mensajeinfo, "Para crear un reto debes estar en Zona segura.", Lider.UserIndex, ToIndex
        Exit Sub
    End If


    'Redimensiono el vector donde voy a guardar los userindex
    'de los usuarios a medida que los voy procesando
    'Esto es para no tener que volver a buscar todos los userindex con la funcion
    'No creo el reto en el primer ciclo, por si alguno de los posibles participantes
    'no cumple con alguna condición entonces tengo que eliminarlo. y prefiero no estar
    'creando y eliminando objetos
    ReDim integrantesIndex(1 To cantidadParticipantes) As Integer
    
    ' Obtengo a los equipos
    equiposInfo = Split(mid$(infoCreacion, posicionInfoEquipos + 1), "|")
    
    loopIntegranteLista = 1
    
    For loopEquipo = 1 To cantidadEquipos

            integrantesInfo = Split(equiposInfo(loopEquipo - 1), "-")
            
            For loopIntegrante = 0 To UBound(integrantesInfo)
        
                UserIndex = NameIndex(integrantesInfo(loopIntegrante))
                
                estaOnline = True
                If UserIndex = 0 Then
                    estaOnline = False
                Else
                    #If testeo = 0 Then
                    If UserList(UserIndex).flags.Privilegios > 0 Then estaOnline = False
                    #End If
                    
                    If MapInfo(UserList(UserIndex).pos.map).Pk = True Then estaOnline = False
                End If
                
                'Esta online?
                If estaOnline Then
                
                    If puedeJugarReto(UserList(UserIndex), error) Then
                        
                        'Chequeo si esta repetido
                        For loopParticipante = 1 To loopIntegranteLista
                            If integrantesIndex(loopParticipante) = UserIndex Then
                                'En este caso termina porque hubo una falla en el Cliente
                                Exit Sub
                            End If
                        Next loopParticipante

                        'Al pedo hacer esto si ya detectamos algún error.
                        If Not hayError Then
                            'Todo ok
                            integrantesIndex(loopIntegranteLista) = UserIndex

                            If loopIntegrante = UBound(integrantesInfo) - 1 Then
                                presentacionString = presentacionString & UserList(UserIndex).Name & "(" & UserList(UserIndex).Stats.ELV & ") y "
                            ElseIf loopIntegrante < UBound(integrantesInfo) Then
                                presentacionString = presentacionString & UserList(UserIndex).Name & "(" & UserList(UserIndex).Stats.ELV & "), "
                            Else
                                presentacionString = presentacionString & UserList(UserIndex).Name & "(" & UserList(UserIndex).Stats.ELV & ")"
                            End If
                            
                            loopIntegranteLista = loopIntegranteLista + 1
                        End If
                        
                    Else
                        hayError = True
                        EnviarPaquete Paquetes.mensajeinfo, error, Lider.UserIndex, ToIndex
                    End If
                Else
                    'El personaje no esta online
                    hayError = True
                    EnviarPaquete Paquetes.mensajeinfo, "El personaje " & integrantesInfo(loopIntegrante) & " no está en un mapa seguro.", Lider.UserIndex, ToIndex
                End If

            Next loopIntegrante

            ' Agregamos el Vs en el Medio
            If loopEquipo < cantidadEquipos Then presentacionString = presentacionString & " Vs "

    Next loopEquipo

    ' Si surgio algun error salgo
    If hayError Then Exit Sub

    ' Termino de Generar el Stirng de Presentacion
    If modo = eModoApuesta.oro Then
        presentacionString = presentacionString & ". Apuesta " & cantidadOroApostado & " monedas de oro."
    ElseIf modo = eModoApuesta.Items Then
        presentacionString = presentacionString & ". Apuesta los items."
    Else
        presentacionString = presentacionString & ". Apuesta " & cantidadOroApostado & " monedas de oro y los items."
    End If

    ' ¿Estan limitadas las rojas?
    If limitarRojas Then
        presentacionString = presentacionString & " Limite de Pociones Rojas por personaje: " & cantidadRojas & "."
    End If
    
    If valeResu Then
        presentacionString = presentacionString & " VALE RESUCITAR."
    End If
    
    If limitarCascoYEscudo Then
        presentacionString = presentacionString & " No vale utilizar Cascos ni Escudos."
    End If
    
    If plantado = True Then
        presentacionString = Lider.Name & " te desafía A PLANTAR en el reto: " & presentacionString & " Para aceptar el plante escribe /RETAR " & Lider.Name & " o /RECHAZAR " & Lider.Name & " para negarselo."
    Else
        presentacionString = Lider.Name & " te invita a participar del reto " & presentacionString & " Para aceptar escribe /RETAR " & Lider.Name & " o /RECHAZAR " & Lider.Name & " para negarselo."
    End If
    
    Debug.Print presentacionString
    
    Dim indexLiderVector(1 To 1) As Integer

    indexLiderVector(1) = Lider.UserIndex
        
    ' ¿Limite de rojas?
    If limitarRojas Then
        Dim EventoObjetoRestringido(1 To 1) As tEventoObjetoRestringido
        
        EventoObjetoRestringido(1).cantidad = cantidadRojas
        EventoObjetoRestringido(1).tipo = eRangoLimite.maximo
        EventoObjetoRestringido(1).id = 38
                
        Set condicionMaxItems = New CondicionEventoLimiteItem
        
        Call condicionMaxItems.setParametros(EventoObjetoRestringido, False, False)
        
        If Not condicionMaxItems.iCondicionEvento_puedeIngresarEquipo(indexLiderVector) Then
            EnviarPaquete Paquetes.mensajeinfo, "No cumplis con las condición de máxima cantidad de rojas.", Lider.UserIndex, ToIndex
            Exit Sub
        End If
    End If
    
    ' Limite de Casco e Items
    If limitarCascoYEscudo Then
        Set condicionSinCascoNiEscudo = New CondicionEventoSinCascoEscudo
               
        If Not condicionSinCascoNiEscudo.iCondicionEvento_puedeIngresarEquipo(indexLiderVector) Then
            EnviarPaquete Paquetes.mensajeinfo, "No cumplis con las condición de NO uso de Cascos y Escudos.", Lider.UserIndex, ToIndex
            Exit Sub
        End If
    End If
    

    ' Creo el Reto
    Set reto = New cReto

    ' Cantidad de equipos
    Call reto.setCantidadEquipos(2)

    'Selecciono el modo del evento y por cuanto es
    Call reto.setModoApuesta(modo, cantidadOroApostado)

    'Establezco el tiempo maximo que puede tardar el reto en terminarse.
    Call reto.setTiempoMaximo(TIEMPO_MAXIMO)
    
    ' Si es plantado necesitamos un ring especial
    If plantado Then Call reto.iEvento_setTiporing(eRingTipo.ringPlantado + eRingTipo.ringTorneo)
    
    ' Vale resu  o no
    Call reto.setValeResu(valeResu)
    
    ' ¿Condicion de items?
    If Not condicionMaxItems Is Nothing Then Call reto.iEvento_agregarCondicionIngreso(condicionMaxItems)
    ' Condicion de NO usar Cascos y Escudos?
    If Not condicionSinCascoNiEscudo Is Nothing Then Call reto.iEvento_agregarCondicionIngreso(condicionSinCascoNiEscudo)
        
    ' Agrego los equipos al evento y mientras tanto les aviso
    ' Como así se recorrio al principio, en la lista de integrantes lo tenemos en el mismo orden
    ReDim equipoIntegranteID(1 To cantidadIntegrantesEquipo) As Long

    loopIntegranteLista = 1
    
    For loopEquipo = 1 To cantidadEquipos

        For loopIntegrante = 1 To cantidadIntegrantesEquipo

            equipoIntegranteID(loopIntegrante) = UserList(integrantesIndex(loopIntegranteLista)).id

            If Lider.UserIndex <> integrantesIndex(loopIntegranteLista) Then
                'Envio el mensaje de invitacion
                EnviarPaquete Paquetes.MensajeTalk, presentacionString, integrantesIndex(loopIntegranteLista), ToIndex
            End If

            loopIntegranteLista = loopIntegranteLista + 1

        Next loopIntegrante

        Call reto.agregarEquipo(equipoIntegranteID, loopEquipo)

    Next loopEquipo


    'El que lo creo, obviamente, acepta la solicitud,
    Call reto.aceptarSolicitud(Lider.id, Lider.UserIndex)

    'El tipo tenia un evento creado por el que nunca se inicio y ahora quiere crear otor
    'en su remplazo para ver si tiene más exito?
    If Not Lider.evento Is Nothing Then
        If Lider.evento.getEstadoEvento = eEstadoEvento.Preparacion Then
            Dim evento As iEvento
            Set evento = Lider.evento
            'Lo quito de la lista ya
            Call modEventos.quitarEvento(evento)
            'Cancelo ese evento
            Call Lider.evento.cancelar
            Set evento = Nothing
        End If
    End If

    'Le asigno el evento al que lo creo, al lider
    Set Lider.evento = reto
    
    'Le aviso que se creo correctamente
    If cantidadIntegrantesEquipo = 1 Then
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(36) & UserList(integrantesIndex(2)).Name, Lider.UserIndex, ToIndex
    Else
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(127), Lider.UserIndex, ToIndex
    End If
        
    'Agrego el evento a la lista de eventos
    Call modEventos.agregarEvento(reto)
End Sub

Public Sub terminoReto()
    cantidadRetosActivos = cantidadRetosActivos - 1
End Sub

Public Sub comenzoReto()
    'Aumento la cantida de retos que se estan jugando
    cantidadRetosActivos = cantidadRetosActivos + 1
End Sub

Public Function getCantidadRetosActivos() As Byte
    getCantidadRetosActivos = cantidadRetosActivos
End Function

Public Sub rechazarSolicitud(Usuario As User, Solicitante As String)
    Dim index As Integer
    Dim reto As cReto
    
    index = NameIndex(Solicitante)
    
    'esta online el tipo que supuestamente creo el reto ?
    If index = 0 Then
        ' Pongo un mensaje generico para que no puedan descubrir si se esta Online o no
        EnviarPaquete Paquetes.MensajeTalk, Solicitante & " no te está invitando a un reto.", Usuario.UserIndex, ToIndex
        Exit Sub
    End If

    'El tipo hizo un reto?
    If UserList(index).evento Is Nothing Then
        'No tiene un reto
        EnviarPaquete Paquetes.MensajeTalk, Solicitante & " no te está invitando a un reto.", Usuario.UserIndex, ToIndex
        Exit Sub
    End If
    
    'Ok esta en un evento. Este evento esta en preparacion?
    If Not UserList(index).evento.getEstadoEvento = eEstadoEvento.Preparacion Then
        EnviarPaquete Paquetes.MensajeTalk, Solicitante & " no te está invitando a un reto.", Usuario.UserIndex, ToIndex
        Exit Sub
    End If
    
    ' Esta en un evento, en preparacion, ¿es un reto?
    If Not UserList(index).evento.getTipoEvento = eTipoEvento.reto Then
        'Esta en otro evento que nada que ver
        EnviarPaquete Paquetes.MensajeTalk, Solicitante & " no te está invitando a un reto.", Usuario.UserIndex, ToIndex
        Exit Sub
    End If
    
    Set reto = UserList(index).evento
    
    ' Cancelamos el resto
    If reto.rechazarSolicitud(Usuario.id, Usuario.UserIndex) Then
        EnviarPaquete Paquetes.MensajeTalk, "Has rechazado la solicitud que te envió " & Solicitante & ".", Usuario.UserIndex, ToIndex
    Else
        'Esta en otro evento que nada que ver
        EnviarPaquete Paquetes.MensajeTalk, Solicitante & " no te está invitando a un reto.", Usuario.UserIndex, ToIndex
    End If
End Sub

Public Sub aceptarSolicitud(UserIndex As Integer, Solicitante As String)

    Dim index  As Integer
    Dim resultadoAceptacion As Byte
    Dim reto As cReto
    
    'El usuario que quiere ingresar a este reto esta en otro?
    If Not UserList(UserIndex).evento Is Nothing Then
        If UserList(UserIndex).evento.getEstadoEvento = Desarrollandose Then
            EnviarPaquete Paquetes.MensajeTalk, "No puedes participar de un reto si estas participando en otro evento.", UserIndex, ToIndex
            Exit Sub
        End If
    End If
   
    'Se fija si se pueden seguir creando retos. Si no se alcanzo
    'el maximo de retos en simultaneo
    If cantidadRetosActivos >= modRetos.cantidadMaximaRetos Then
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(62), UserIndex, ToIndex
        Exit Sub
    End If
    
    index = NameIndex(Solicitante)
    
    'esta online el tipo que supuestamente creo el reto ?
    If index = 0 Then
        'No pongo que esta online asi no buscan personas online (incluidos gms) poniendo /RETAR NICKBUSCADO
        EnviarPaquete Paquetes.MensajeTalk, Solicitante & " no te invitó a ningún reto.", UserIndex, ToIndex
        Exit Sub
    End If
    
    'El tipo hizo un reto?
    If UserList(index).evento Is Nothing Then
        'No tiene un reto
        EnviarPaquete Paquetes.MensajeTalk, Solicitante & " no te invitó a ningún reto.", UserIndex, ToIndex
        Exit Sub
    End If

    'Ok esta en un evento. Este evento esta en preparacion?
    If Not UserList(index).evento.getEstadoEvento = eEstadoEvento.Preparacion Then
        'No te invito..
        EnviarPaquete Paquetes.MensajeTalk, Solicitante & " no te invitó a ningún reto.", UserIndex, ToIndex
        Exit Sub
    End If
    
    ' Esta en un evento, en preparacion, ¿es un reto?
    If Not UserList(index).evento.getTipoEvento = eTipoEvento.reto Then
        'Esta en otro evento que nada que ver
        EnviarPaquete Paquetes.MensajeTalk, UserList(index).Name & " no te invitó a ningún reto.", UserIndex, ToIndex
        Exit Sub
    End If
                       
    Set reto = UserList(index).evento
                
    If Not reto.puedeIngresar(UserIndex) Then
        EnviarPaquete Paquetes.MensajeTalk, "No cumplís con las condiciones para ingresar al Reto.", UserIndex, ToIndex
        Exit Sub
    End If
                
    resultadoAceptacion = reto.aceptarSolicitud(UserList(UserIndex).id, UserIndex)
            
    If resultadoAceptacion = 0 Then 'Solicitud aceptada
        If reto.listoParaEmpezar() Then
            'Comienzo el reto
            Call reto.comenzar
        Else   'Si es el ultimo que acepta no tiene valor que le mande al lider que este acept?
            'Le aviso al lider y al participante que acepto
            EnviarPaquete Paquetes.MensajeTalk, "Aceptaste el reto. Ahora debes esperar que los demás acepten. Si te arrepientes escribe /RECHAZAR " & Solicitante & ".", UserIndex, ToIndex
            EnviarPaquete Paquetes.MensajeTalk, UserList(UserIndex).Name & " aceptó tu invitación para jugar el reto.", index, ToIndex
        End If
    ElseIf resultadoAceptacion = 1 Then 'Ya habia aceptado
        EnviarPaquete Paquetes.MensajeTalk, "Tu ya has aceptado el reto. Ahora debes esperar que los demás acepten.", UserIndex, ToIndex
    ElseIf resultadoAceptacion = 2 Then 'No fue invitado
        EnviarPaquete Paquetes.MensajeTalk, UserList(index).Name & " no te invitó a ningún reto.", UserIndex, ToIndex
    End If

End Sub

Public Sub registrar(Ganador As User, Perdedor As User, oro As Long)
    Dim sql As String
    
    sql = "INSERT DELAYED " & DB_NAME_PRINCIPAL & ".juego_logs_retos_1(IDCUENTA1, IDPJ1, IDCUENTA2, IDPJ2, ORO)" & _
            " values(" & Ganador.IDCuenta & "," & Ganador.id & "," & Perdedor.IDCuenta & "," & Perdedor.id & "," & oro & ")"
    
    conn.Execute sql, , adExecuteNoRecords
End Sub
