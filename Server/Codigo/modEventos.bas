Attribute VB_Name = "modEventos"
Option Explicit

Public Enum eTipoEvento
    reto = 1
    torneo = 2
End Enum

Public Enum eResultadoInscripcion
    correcta = 1
    noTieneOro = 2
    inscripcionCerrada = 3
    desconocido = 4
End Enum

' Formas de identificar a un equipo
Public Enum eIdentificacionEquipos
    identificaPersonajes = 1
    identificaClan = 2
    identificaFaccion = 3
End Enum

Private Eventos() As iEvento
Private cantidadEventos As Byte
Private cantidadEventosNoRetos As Byte

'Personas que participan de un evento y cerraron
Private usuariosOffline As Collection

#If TDSFacil = 1 Then
    Private Const CANTIDAD_CAPACIDAD_INICIAL = 40
#Else
    Private Const CANTIDAD_CAPACIDAD_INICIAL = 20
#End If

Private Const MULTIPLICADOR_CANTIDAD_CRECIMIENTO = 2
Private Const MULTIPLICADOR_CANTIDAD_REDUCCION = 2

Public Sub cancelarSolicitudesPendientesParaElEvento(evento As iEvento)
    Dim loopUser As Integer
    
    For loopUser = 1 To LastUser
        ' Personaje jugando?
        If UserList(loopUser).flags.UserLogged Then
            ' ¿Tiene una solicitud?
            If Not UserList(loopUser).solicitudEvento Is Nothing Then
                ' ¿La solicitud es de este evento?
                If UserList(loopUser).solicitudEvento.getEvento Is evento Then
                    ' Lo cancelamos
                    Set UserList(loopUser).solicitudEvento = Nothing
                    EnviarPaquete Paquetes.mensajeinfo, "Las inscripciones del evento han finalizado. Tus compañeros no aceptaron todos la solicitud o lo hicieron demasiado tarde.", loopUser, ToIndex
                End If
            End If
        End If
    Next
End Sub

Public Sub iniciarEstructuraEventos()

    Dim loopC As Byte
    
    'Creo el espacio inicial para guardar los eventos que se vayan generando
    ReDim Eventos(1 To CANTIDAD_CAPACIDAD_INICIAL) As iEvento
    
    For loopC = 1 To UBound(Eventos)
        Set Eventos(loopC) = Nothing
    Next

    'Creo la coleccion donde voy a guardar los usuarios que salen
    Set usuariosOffline = New Collection
    
End Sub

Private Function getIndexLibre() As Byte
    
    Dim loopC As Byte
    
    For loopC = 1 To UBound(Eventos)
        If Eventos(loopC) Is Nothing Then
            getIndexLibre = loopC
            Exit Function
        End If
    Next loopC
    
    getIndexLibre = 0

End Function

Public Sub agregarEvento(ByRef evento As iEvento)

    If cantidadEventos = UBound(Eventos) Then
        ReDim Preserve Eventos(1 To Int(cantidadEventos * MULTIPLICADOR_CANTIDAD_CRECIMIENTO)) As iEvento
    End If
    
    Set Eventos(getIndexLibre) = evento
    
    cantidadEventos = cantidadEventos + 1
End Sub

Public Function publicarEvento(nombreEvento As String) As Boolean
    Dim loopC As Integer
    
    publicarEvento = False
    
    'Obtengo el evento
    For loopC = 1 To UBound(Eventos)
    
        If Not Eventos(loopC) Is Nothing Then
                        
            If Eventos(loopC).getNombre = nombreEvento Then
                
                ' Chequeo que sea el estado valido para hacer esta operacion
                If Not Eventos(loopC).getEstadoEvento = esperandoConfirmacionInicio Then Exit For
                
                ' Lo publico
                Call Eventos(loopC).publicar
                    
                ' Los cuento cuando se publica
                If Eventos(loopC).getTipoEvento() <> eTipoEvento.reto Then cantidadEventosNoRetos = cantidadEventosNoRetos + 1
    
                ' Todo ok
                publicarEvento = True
                Exit Function
            End If
            
        End If
    Next

    publicarEvento = False
End Function

Public Function cancelarEvento(nombreEvento As String) As Boolean
    Dim loopC As Integer
    Dim evento As iEvento
    
    cancelarEvento = False
    
    'Obtengo el evento
    For loopC = 1 To UBound(Eventos)
    
        If Not Eventos(loopC) Is Nothing Then
                        
            If Eventos(loopC).getNombre = nombreEvento Then
                Set evento = Eventos(loopC)
                Call quitarEvento(evento)
                Call evento.cancelar
                cancelarEvento = True
                Exit For
            End If
            
        End If
    Next

End Function

Public Sub quitarEvento(ByRef evento As iEvento)
    Dim loopC As Byte
    Dim loopUser As Integer
    
    Call LogEventos("Se quita el evento" & evento.getNombre)
    
    For loopC = 1 To UBound(Eventos)
        If Eventos(loopC) Is evento Then
            
            If Eventos(loopC).getTipoEvento <> eTipoEvento.reto And Not Eventos(loopC).getEstadoEvento = esperandoConfirmacionInicio Then
                cantidadEventosNoRetos = cantidadEventosNoRetos - 1
            End If
            
            Set Eventos(loopC) = Nothing
            
            cantidadEventos = cantidadEventos - 1
            
            Exit For
        End If
    Next loopC
End Sub

Public Sub procesarTimeOutMinuto()
    Dim loopC As Byte
    
    For loopC = 1 To UBound(Eventos)
    
        If Not Eventos(loopC) Is Nothing Then
            
            If Eventos(loopC).getEstadoEvento = eEstadoEvento.Terminado Then
                Call LogEventos("Se quita el evento" & Eventos(loopC).getNombre)
                
                If Eventos(loopC).getTipoEvento <> eTipoEvento.reto Then
                    cantidadEventosNoRetos = cantidadEventosNoRetos - 1
                End If
                
                cantidadEventos = cantidadEventos - 1
                Set Eventos(loopC) = Nothing
                
            Else
                Call Eventos(loopC).timeOutMinuto
            End If
            
        End If
    Next
    
    'Si la capacidad del array es mas de X veces superior a la cantidad de elementos que hay
    'Osea que tiene espacio al pedo, redimensiono el array
    If UBound(Eventos) >= MULTIPLICADOR_CANTIDAD_CRECIMIENTO * cantidadEventos And UBound(Eventos) > CANTIDAD_CAPACIDAD_INICIAL Then
        Call compactarYRedimensionarEventos
    End If
   
End Sub

Private Sub loguearEstadoEventos()

Dim loopEvento As Integer

    Call LogEventos("-----------------------------------")
    
    For loopEvento = 1 To UBound(Eventos)
        If Eventos(loopEvento) Is Nothing Then
            Call LogEventos(loopEvento & "> Vacio")
        Else
            Call LogEventos(loopEvento & "> " & Eventos(loopEvento).getNombre)
        End If
    Next loopEvento
    
    Call LogEventos("-----------------------------------")

End Sub
'Compacta la lista de eventos cuando esta se fragmenta. Ejemplo la lista
'mide 100 elementos pero solo hay 2 usados. El 1 y el 100

Private Sub compactarYRedimensionarEventos()

Dim cantidadOriginal As Integer
Dim nuevaCapacidad As Integer

Dim loopInferior As Integer
Dim loopSuperior As Integer

Call loguearEstadoEventos

cantidadOriginal = UBound(Eventos)
nuevaCapacidad = Int(cantidadOriginal / MULTIPLICADOR_CANTIDAD_REDUCCION)

loopInferior = 1
loopSuperior = cantidadOriginal

Do While Not loopSuperior < loopInferior
        
    Do While Not Eventos(loopInferior) Is Nothing And loopInferior < UBound(Eventos)
        loopInferior = loopInferior + 1
    Loop
        
    Do While Eventos(loopSuperior) Is Nothing And loopSuperior > 0
        loopSuperior = loopSuperior - 1
    Loop
        
    If loopSuperior > loopInferior Then
        Set Eventos(loopInferior) = Eventos(loopSuperior)
        Set Eventos(loopSuperior) = Nothing
        
        loopSuperior = loopSuperior - 1
        loopInferior = loopInferior + 1
    End If
Loop

'Redimensiono
ReDim Preserve Eventos(1 To nuevaCapacidad) As iEvento

Call LogEventos("Se compacta la lista de eventos de " & cantidadOriginal & " a " & nuevaCapacidad)
End Sub

Public Function getCantidadEnventos() As Byte
    getCantidadEnventos = cantidadEventos
End Function

Public Function getCantidadEnventosNoRetos() As Byte
    getCantidadEnventosNoRetos = cantidadEventosNoRetos
End Function

Public Function getMayorIndex() As Byte
    getMayorIndex = UBound(Eventos)
End Function

Public Function obtenerEstadoTorneos() As String
    Dim Estado As String
    Dim loopC As Integer
    Dim cantidadEventos As Integer

    cantidadEventos = 0
        
    For loopC = 1 To UBound(Eventos)
    
        If Not Eventos(loopC) Is Nothing Then
            
            If Not Eventos(loopC).getTipoEvento = eTipoEvento.reto Then
                Estado = Estado & Eventos(loopC).getNombre & "-"
                
                If Eventos(loopC).getEstadoEvento = eEstadoEvento.Terminado Then
                    Estado = Estado & " Terminado"
                ElseIf Eventos(loopC).getEstadoEvento = eEstadoEvento.Preparacion Then
                    Estado = Estado & " Preparando"
                ElseIf Eventos(loopC).getEstadoEvento = eEstadoEvento.Desarrollandose Then
                    Estado = Estado & " Desarrollandose"
                ElseIf Eventos(loopC).getEstadoEvento = eEstadoEvento.esperandoConfirmacionInicio Then
                    Estado = Estado & " Esperando confirmacion"
                End If
                
                Estado = Estado & " Time " & Eventos(loopC).getTimeTranscurrido & "||"
            End If
        End If
    Next
    
    obtenerEstadoTorneos = Estado
End Function


Public Sub verEstadoEventos(lista As VB.ListBox)
    Dim Estado As String
    Dim loopC As Integer
    Dim cantidadEventos As Integer
    
    cantidadEventos = 0
        
    For loopC = 1 To UBound(Eventos)
    
        If Not Eventos(loopC) Is Nothing Then
                    
            Estado = loopC & "-" & Eventos(loopC).getNombre & "-"
            
            If Eventos(loopC).getTipoEvento = eTipoEvento.reto Then
                Estado = Estado & " (Reto)"
            Else
                Estado = Estado & " (Torneo:" & Eventos(loopC).obtenerIDPersistencia() & ")"
            End If
                    
            If Eventos(loopC).getEstadoEvento = eEstadoEvento.Terminado Then
                Estado = Estado & " Terminado"
            ElseIf Eventos(loopC).getEstadoEvento = eEstadoEvento.Preparacion Then
                Estado = Estado & " Preparando"
            ElseIf Eventos(loopC).getEstadoEvento = eEstadoEvento.Desarrollandose Then
                Estado = Estado & " Desarrollandose"
            ElseIf Eventos(loopC).getEstadoEvento = eEstadoEvento.esperandoConfirmacionInicio Then
                Estado = Estado & " Esperando confirmacion"
            End If
            
            Estado = Estado & " Gan " & Eventos(loopC).getIDGanador & " Time " & Eventos(loopC).getTimeTranscurrido
            
            lista.AddItem Estado
            
            cantidadEventos = cantidadEventos + 1
        End If
    Next
    lista.AddItem "Cantidad eventos: " & cantidadEventos
End Sub

Public Function getEventoByNombre(ByVal nombre As String) As iEvento
    
    Dim loopC As Integer
    
    nombre = UCase(nombre)
    
    For loopC = 1 To UBound(Eventos)
    
        If Not Eventos(loopC) Is Nothing Then
            If UCase(Eventos(loopC).getNombre) = nombre Then
                Set getEventoByNombre = Eventos(loopC)
                Exit Function
            End If
        End If
    Next
    
    Set getEventoByNombre = Nothing
End Function

'Agrega el usuario a la lista de usuarios offline
Public Sub agregarUsuarioOffline(UserID As Long, ByRef evento As iEvento)
    
    'Creo el objeto donde guardo los datos
    Dim usuarioInfo As cUsuarioOfflineEvento
    Set usuarioInfo = New cUsuarioOfflineEvento
    
    usuarioInfo.UserID = UserID
    Set usuarioInfo.evento = evento
    
    'Lo agrego a la lista
    Call usuariosOffline.Add(usuarioInfo)
    
End Sub

'Busca si el usuario estaba en un evento y cerro
Public Sub reEstablecerEventoUsuario(ByRef Usuario As User, UserIndex As Integer)
    
    'Recorro la lista de eventos buscando el userid. Si esta, tengo el evento. Sino es nothing
    Dim loopUsuario As Integer
    
    For loopUsuario = 1 To usuariosOffline.Count
            If usuariosOffline(loopUsuario).UserID = Usuario.id Then
                'Le asigno el evento
                Set Usuario.evento = usuariosOffline(loopUsuario).evento
                'Le aviso al evento que este volvio a loguear
                Call usuariosOffline(loopUsuario).evento.usuarioIngreso(UserIndex, Usuario.id)
                'Lo remuevo de la lista de usuaros offline
                Call usuariosOffline.Remove(loopUsuario)
                Exit Sub
            End If
    Next loopUsuario
    
    'No tiene ningun evento
    Set Usuario.evento = Nothing
End Sub

'Quita la relacion que tiene un usuario con un evento. Cuando este usuario loguee, no se le va a asignar este
'evento
Public Sub quitarReferenciaUsuarioEvento(idUsuario As Long)
    'Recorro la lista de eventos buscando el userid. Si esta, tengo el evento. Sino es nothing
    Dim loopUsuario As Integer
    
    For loopUsuario = 1 To usuariosOffline.Count
            If usuariosOffline(loopUsuario).UserID = idUsuario Then
                'Lo remuevo de la lista de usuaros offline
                Call usuariosOffline.Remove(loopUsuario)
                Exit Sub
            End If
    Next loopUsuario

End Sub

Private Function obtenerListaEventosNoRetos() As String
    Dim loopC As Byte
    Dim cantidadTorneos As Byte
    Dim temp As String
    Dim encontrados As Byte
    
    cantidadTorneos = modEventos.getCantidadEnventosNoRetos()
    encontrados = 0
    temp = ""
    For loopC = 1 To UBound(Eventos)
        If Not Eventos(loopC) Is Nothing Then
            If Not Eventos(loopC).getTipoEvento = eTipoEvento.reto Then
                If Not Eventos(loopC).getEstadoEvento = eEstadoEvento.esperandoConfirmacionInicio Then
                    
                    encontrados = encontrados + 1
                    
                    If encontrados = cantidadTorneos - 1 Then ' Ante ultimo
                        temp = temp & Eventos(loopC).getNombre & " y "
                    ElseIf encontrados < cantidadTorneos Then
                        temp = temp & Eventos(loopC).getNombre & ", "
                    Else ' Ultimo
                        temp = temp & Eventos(loopC).getNombre
                    End If
                    
                End If
            End If
        End If
    Next loopC
    
    obtenerListaEventosNoRetos = temp
End Function

Private Function obtenerInfoPrimerEventoNoReto() As String
    Dim loopC As Byte
    Dim cantidadTorneos As Byte
    Dim temp As String
    Dim encontrados As Byte
    
    cantidadTorneos = modEventos.getCantidadEnventosNoRetos()

    temp = ""
    
    For loopC = 1 To UBound(Eventos)
        ' ¿Evento?
        If Not Eventos(loopC) Is Nothing Then
            ' ¿No es reto?
            If Not Eventos(loopC).getTipoEvento = eTipoEvento.reto Then
                '¿Publicado?
                If Not Eventos(loopC).getEstadoEvento = eEstadoEvento.esperandoConfirmacionInicio Then
                    temp = Eventos(loopC).getNombre & "> " & Eventos(loopC).getDescripcion & vbCrLf & Eventos(loopC).obtenerInfoEstado
                End If
            End If
        End If
    Next loopC
    
    obtenerInfoPrimerEventoNoReto = temp
End Function

'Envia al usuariro la ifnormacion de un evento en base al nombre
Public Sub enviarListaTorneos(UserIndex As Integer)
    
    Dim cantidadTorneos As Byte
    Dim listaTorneos As String
    
    cantidadTorneos = modEventos.getCantidadEnventosNoRetos()
        
    If cantidadTorneos = 0 Then
        listaTorneos = "En este momento no se esta realizando algún evento."
        EnviarPaquete Paquetes.mensajeinfo, listaTorneos, UserIndex, ToIndex
    ElseIf cantidadTorneos = 1 Then
        listaTorneos = obtenerInfoPrimerEventoNoReto()
        EnviarPaquete Paquetes.MensajeTalk, listaTorneos, UserIndex, ToIndex
    Else
        listaTorneos = "Actualmente se están desarrollando los siguientes eventos: " & obtenerListaEventosNoRetos() & ". Escribe /EVENTO nombre del evento para obtener más informacion."
        EnviarPaquete Paquetes.mensajeinfo, listaTorneos, UserIndex, ToIndex
    End If
        
    
    
End Sub

'Envia al usuariro la ifnormacion de un evento en base al nombre
Public Sub enviarInformacionEvento(UserIndex As Integer, nombreEvento As String)
    Dim evento As iEvento
    Set evento = getEventoByNombre(nombreEvento)
    
    If evento Is Nothing Then
        EnviarPaquete Paquetes.mensajeinfo, "No hay ningún evento con el nombre " & nombreEvento & ".", UserIndex, ToIndex
    Else
        Dim infoEstado As String
        If evento.getTipoEvento <> eTipoEvento.reto Then  'Los retos no califican
            EnviarPaquete Paquetes.MensajeTalk, evento.getNombre & "> " & evento.getDescripcion & vbCrLf & evento.obtenerInfoEstado, UserIndex, ToIndex
        End If
    End If
    
End Sub

Public Function obtenerIDParaPersistirEvento() As Long
    'Tengo que obtener un id para la tabla
    Dim sql As String
    Dim infoIDEvento As ADODB.Recordset
    
    sql = "UPDATE " & DB_NAME_PRINCIPAL & ".juego_torneos_eventos SET ULTIMOIDEVENTO = LAST_INSERT_ID(ULTIMOIDEVENTO + 1) WHERE ULTIMOIDEVENTO>=0;"
    
    conn.Execute sql, , adExecuteNoRecords Or adCmdText
            
    sql = "SELECT LAST_INSERT_ID() AS IDEVENTO;"
    
    Set infoIDEvento = conn.Execute(sql, , adCmdText)
    
    obtenerIDParaPersistirEvento = CInt(infoIDEvento!IDEVENTO)
    
    'Libero
    infoIDEvento.Close
    Set infoIDEvento = Nothing
End Function

Public Sub guardarEstadoEventos()
    Dim Estado As String
    Dim loopC As Integer
    Dim cantidadEventos As Integer
    
    cantidadEventos = 0
        
    For loopC = 1 To UBound(Eventos)
        If Not Eventos(loopC) Is Nothing Then
            If Not Eventos(loopC).getTipoEvento = eTipoEvento.reto Then
                Call Eventos(loopC).guardar
            End If
        End If
    Next
End Sub

Public Function establecerGanadorEvento(nombreEvento As String, NombreEquipo As String) As Boolean
    Dim evento As iEvento
    
    Set evento = modEventos.getEventoByNombre(nombreEvento)
    
    If Not evento Is Nothing Then
        establecerGanadorEvento = evento.establecerGanadorManualmente(NombreEquipo)
    Else
        establecerGanadorEvento = False
    End If
    
End Function


