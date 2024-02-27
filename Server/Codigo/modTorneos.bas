Attribute VB_Name = "modTorneos"
Option Explicit

Public Type tIntegrantesEquipoTorneo
    id As Long
    UserIndex As Integer
    nick As String
    posOriginal As WorldPos
    cantidadAdvertencias As Byte
    Estado As eEstadoIntegranteEquipo
    cantidadOroPagadoInscripcion As Long
End Type

'Cada tabla de resultados del torneo esta compuesto por un registro de estos
Public Type tEquipoTablaTorneo
    idTablaPersistencia As Long
    idEquipo As Byte
    NombreEquipo As String
    integrantes() As tIntegrantesEquipoTorneo
    cantidadIntegrantes As Byte
    cantidadIntegrantesDescalificados As Byte
    
    Estado As eEstadoEquipoTorneo
    
    'informacion estadisticas sobre el equipo en el torneo
    cantidadCombatesGanados As Byte
    cantidadCombatesJugados As Byte
    cantidadCombatesEmpatados As Byte
    
    cantidadRoundsGanados As Byte
    cantidadRoundsJugados As Byte
    
    'Aux
    tickInscripcion As Integer
End Type

Public Enum eFormatoDisplayEquipo
    NombreEquipo = 1
    nombresJugadores = 2
    completo = 3
End Enum

Public Enum eEstadoIntegranteEquipo
    Jugando = 1
    Descalificando = 2
End Enum

Public Enum eEstadoEquipoTorneo
    participando = 0 'actualmente se encuentra participando en el evento
    termino = 2 'El personaje termino su participacion en el evento por algo que NO fue ser descalifcado
    descalificado = 3 'El personaje terino su partiicpacion en el evento porque fue descalificado
End Enum

'Genera el mensaje de
'Para participar escribe /PARTICIPAR <NOMBRE DEL EVENTO>-<NOMBRE COMPA 1>-<NOMBRE COMPA 2>-<NOMBRE COMPA N>
Public Function generarMensajeParticipar(nombreEvento As String, cantidadParticipantes As Byte) As String

Dim loopC As Byte
    
    If cantidadParticipantes = 1 Then
        generarMensajeParticipar = " Para participar escribe /PARTICIPAR " & nombreEvento
    Else
    
        generarMensajeParticipar = " Para participar escribe /PARTICIPAR " & nombreEvento
        
        'El ultimo se lo pongo manualmente
        For loopC = 1 To cantidadParticipantes - 2
            generarMensajeParticipar = generarMensajeParticipar & "-Nombre Compa " & loopC
        Next loopC
        
        generarMensajeParticipar = generarMensajeParticipar & "-Nombre Compa " & loopC
    End If
    
   
End Function

' ******************************************************************************************
' Funciones para pasar de equipos a lista que los representan

'Obtiene el nombre de los equipos separados por <SEPARADOR> terminando con un punto.
'Solo tiene en cuenta a los equipos que estan participando
'
' separador. String que separa a los equipos.
' Si soloJugando = true. Muestra solo los participantes que no fueron descalificados.
Public Function obtenerStringPrensetacion(tablaEquipos() As tEquipoTablaTorneo, separador As String, soloJugando As Boolean, formato As eFormatoDisplayEquipo) As String

    Dim loopEquipo As Byte
    Dim cantidadEquipos As Byte
    
    obtenerStringPrensetacion = ""
    
    cantidadEquipos = UBound(tablaEquipos)
    
    For loopEquipo = 1 To cantidadEquipos
    
        If tablaEquipos(loopEquipo).Estado = eEstadoEquipoTorneo.participando Then
            obtenerStringPrensetacion = obtenerStringPrensetacion & obtenerStringEquipo(tablaEquipos(loopEquipo), soloJugando, formato)
            
            If loopEquipo < cantidadEquipos Then
                obtenerStringPrensetacion = obtenerStringPrensetacion & separador
            Else
                obtenerStringPrensetacion = obtenerStringPrensetacion & "."
            End If
        End If
    Next loopEquipo

   Exit Function

End Function

' Devuelve un string del equipo, separado por "," e "y" indicando el estado de cada uno.
Public Function obtenerStringEquipoConEstado(equipo As tEquipoTablaTorneo, formato As eFormatoDisplayEquipo) As String

    Dim loopIntegrante As Byte
    Dim cantidadIntegrantesProcesados As Byte
    Dim cantidadIntegrantes As Byte
       
    cantidadIntegrantes = equipo.cantidadIntegrantes
    
    cantidadIntegrantesProcesados = 1
    
    obtenerStringEquipoConEstado = ""
    
    If formato And eFormatoDisplayEquipo.NombreEquipo Then
        If Not equipo.NombreEquipo = vbNullString Then
            obtenerStringEquipoConEstado = equipo.NombreEquipo & ": "
        End If
    End If
    
    For loopIntegrante = 1 To equipo.cantidadIntegrantes
        
        If equipo.integrantes(loopIntegrante).Estado = eEstadoIntegranteEquipo.Descalificando Then
            obtenerStringEquipoConEstado = obtenerStringEquipoConEstado & equipo.integrantes(loopIntegrante).nick & " (De)"
        ElseIf equipo.integrantes(loopIntegrante).UserIndex = 0 Then
            obtenerStringEquipoConEstado = obtenerStringEquipoConEstado & equipo.integrantes(loopIntegrante).nick & " (Off " & equipo.integrantes(loopIntegrante).cantidadAdvertencias & ")"
        Else
            obtenerStringEquipoConEstado = obtenerStringEquipoConEstado & equipo.integrantes(loopIntegrante).nick
        End If
        
        If loopIntegrante = cantidadIntegrantes - 1 Then
            obtenerStringEquipoConEstado = obtenerStringEquipoConEstado & " y "
        ElseIf loopIntegrante < cantidadIntegrantes Then
            obtenerStringEquipoConEstado = obtenerStringEquipoConEstado & ", "
        End If
            
        cantidadIntegrantesProcesados = cantidadIntegrantesProcesados + 1
    Next

End Function

'Obtiene los integrantes del equipo separados por  "," e "y"
'Solo tiene en cuenta a los participantes que estan jugando
'TO-DO Re veer lo del separador
Public Function obtenerStringEquipo(equipo As tEquipoTablaTorneo, soloJugando As Boolean, formato As eFormatoDisplayEquipo) As String

    Dim loopIntegrante As Byte
    Dim cantidadIntegrantesProcesados As Byte
    Dim cantidadIntegrantesEnString As Byte
    
    ' ¿Tiene alguna identificacion especial?
    If Not equipo.NombreEquipo = vbNullString Then
        If formato = eFormatoDisplayEquipo.NombreEquipo Then
            obtenerStringEquipo = equipo.NombreEquipo
            Exit Function
        ElseIf (formato And eFormatoDisplayEquipo.NombreEquipo) Then
            obtenerStringEquipo = equipo.NombreEquipo & ": "
        End If
    End If
    
    ' Sino vamos con los nombres de los personajes
    If soloJugando Then
        cantidadIntegrantesEnString = equipo.cantidadIntegrantes - equipo.cantidadIntegrantesDescalificados
    Else
        cantidadIntegrantesEnString = equipo.cantidadIntegrantes
    End If
    
    cantidadIntegrantesProcesados = 1
    
    For loopIntegrante = 1 To equipo.cantidadIntegrantes
        
        If Not equipo.integrantes(loopIntegrante).Estado = eEstadoIntegranteEquipo.Descalificando Or soloJugando = False Then

            obtenerStringEquipo = obtenerStringEquipo & equipo.integrantes(loopIntegrante).nick
            
            If loopIntegrante = cantidadIntegrantesEnString - 1 Then
                obtenerStringEquipo = obtenerStringEquipo & " y "
            ElseIf loopIntegrante < cantidadIntegrantesEnString Then
                obtenerStringEquipo = obtenerStringEquipo & ", "
            End If
            
             cantidadIntegrantesProcesados = cantidadIntegrantesProcesados + 1
        End If
    Next

End Function

Public Function estaPersonajeEnEquipo(equipo As tEquipoTablaTorneo, idPersonaje As Long) As Boolean
    Dim loopIntegrante As Integer
    
    For loopIntegrante = 1 To equipo.cantidadIntegrantes
        If equipo.integrantes(loopIntegrante).id = idPersonaje Then
            estaPersonajeEnEquipo = True
            Exit Function
        End If
    Next

    estaPersonajeEnEquipo = False
End Function
Public Function sonCompaneros(tablaEquipos() As tEquipoTablaTorneo, idPersonaje1 As Long, IDPersonaje2 As Long) As Boolean

    Dim posEquipoTabla As Byte
    
    ' Obtenemos el equipo
    posEquipoTabla = obtenerPosicionEnTablaPersonaje(tablaEquipos, idPersonaje1)
    
    ' MM lo encontre? Esto no deberia pasar
    If posEquipoTabla = 0 Then
        sonCompaneros = False
        Exit Function
    End If
    
    If Not estaPersonajeEnEquipo(tablaEquipos(posEquipoTabla), IDPersonaje2) Then
        sonCompaneros = False
        Exit Function
    End If

    sonCompaneros = True

End Function



'Obtener la posicion dentro de la tabla del equipo
'@param        tablaequipos() Required. tEquipoTablaTorneo object.
'@param        idEquipo Required. Byte.
'@return       Byte. La posición del equipo en la tabla
'@rem
Public Function obtenerPosTablaIDEquipo(tablaEquipos() As tEquipoTablaTorneo, idEquipo As Byte) As Byte
    
    
    Dim loopEquipo As Byte
    Dim cantidadEquipos As Byte
    
    cantidadEquipos = UBound(tablaEquipos)
    
    For loopEquipo = 1 To cantidadEquipos
    
            If tablaEquipos(loopEquipo).idEquipo = idEquipo Then
                obtenerPosTablaIDEquipo = loopEquipo
                Exit Function
            End If
    Next loopEquipo
    obtenerPosTablaIDEquipo = 0
    Exit Function
End Function

Public Function obtenerPosicionEnTablaPersonaje(tablaEquipos() As tEquipoTablaTorneo, idPersonaje As Long) As Byte

    Dim loopIntegrante As Byte
    Dim loopEquipo As Byte
    Dim cantidadEquipos As Byte
        
    cantidadEquipos = UBound(tablaEquipos)
    
    For loopEquipo = 1 To cantidadEquipos
        
        With tablaEquipos(loopEquipo)
                
            If .Estado = participando Then
            
                For loopIntegrante = 1 To .cantidadIntegrantes
                                
                    If .integrantes(loopIntegrante).id = idPersonaje Then
                        obtenerPosicionEnTablaPersonaje = loopEquipo
                        Exit Function
                    End If
                Next
                
            End If
                    
        End With
    Next

obtenerPosicionEnTablaPersonaje = 0
End Function
'Esta funcion devuelve el id del equipo al cual pertenece un personaje.
' O 0 en caso de que el personaje no pertenezca a ningun equipo
'El algoritmo se fija en los equipos que estan participando
Public Function obtenerIDEquipoPersonaje(tablaEquipos() As tEquipoTablaTorneo, idPersonaje As Long) As Byte

    Dim loopIntegrante As Byte
    Dim loopEquipo As Byte
    Dim cantidadEquipos As Byte
        
    cantidadEquipos = UBound(tablaEquipos)
    
    For loopEquipo = 1 To cantidadEquipos
        
        With tablaEquipos(loopEquipo)
                
            If .Estado = participando Then
            
                For loopIntegrante = 1 To .cantidadIntegrantes
                                
                    If .integrantes(loopIntegrante).id = idPersonaje Then
                        obtenerIDEquipoPersonaje = tablaEquipos(loopEquipo).idEquipo
                        Exit Function
                    End If
                Next
                
            End If
                    
        End With
    Next

obtenerIDEquipoPersonaje = 0
End Function


'Actualiza el userindex que garda la tabla de equipos
'Devuelve false en caso de que el personaje no este en ninguna equipo
'que esta participando del evento
Public Function actualizarUserIndexPersonaje(tablaEquipos() As tEquipoTablaTorneo, idPersonaje As Long, NuevoUserIndex As Integer)


    Dim loopIntegrante As Byte
    Dim loopEquipo As Byte
    Dim cantidadEquipos As Byte
        
    cantidadEquipos = UBound(tablaEquipos)
    
    For loopEquipo = 1 To cantidadEquipos
        
        With tablaEquipos(loopEquipo)
                
            If .Estado = participando Then
            
                For loopIntegrante = 1 To .cantidadIntegrantes
                                
                    If .integrantes(loopIntegrante).id = idPersonaje Then
                        .integrantes(loopIntegrante).UserIndex = NuevoUserIndex
                        actualizarUserIndexPersonaje = True
                        Exit Function
                    End If
                Next
                
            End If
                    
        End With
    Next

    actualizarUserIndexPersonaje = False

End Function

Public Sub entregarOro(id As Long, oro As Long, mensaje As String)

    Dim UserIndex As Integer
    
    UserIndex = IDIndex(id)
    
    If UserIndex > 0 Then 'Esta online
        'Le doy el oro.
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + oro
        'Lo actualizo
        EnviarPaquete Paquetes.EnviarOro, Codify(UserList(UserIndex).Stats.GLD), UserIndex, ToIndex
    
        If mensaje <> "" Then
            EnviarPaquete Paquetes.MensajeTalk, mensaje, UserIndex, ToIndex
        End If
        
        Call LogTorneos("Se le entrega " & oro & " a " & UserList(UserIndex).Name & "." & mensaje)
    Else  'Esta offline
        If oro > 0 Then
            Call modUsuarios.entregarOroOffline(id, oro)
        End If
    End If
End Sub

Public Sub desecharIntegranteEquipo(Integrante As tIntegrantesEquipoTorneo, premioOro As Long, Optional ByVal oroInscripcion = False)
    Dim oroADar As Long
    
    If oroInscripcion Then
        oroADar = Integrante.cantidadOroPagadoInscripcion
    Else
        oroADar = premioOro
    End If
    
    If Integrante.UserIndex > 0 Then 'Esta online
    
        'Si corresponde le doy oro
        If oroADar > 0 Then
            'Le doy el oro.
            UserList(Integrante.UserIndex).Stats.GLD = UserList(Integrante.UserIndex).Stats.GLD + oroADar
            'Lo actualizo
            EnviarPaquete Paquetes.EnviarOro, Codify(UserList(Integrante.UserIndex).Stats.GLD), Integrante.UserIndex, ToIndex
        End If
        
        'La transporto. 'TO-DO poner un transportar posta...
        Call modUsuarios.transportarUsuarioOnline(Integrante.UserIndex, Integrante.posOriginal.map, Integrante.posOriginal.x, Integrante.posOriginal.y, False, False)
        
        'Le quito la referencia al evento
        Set UserList(Integrante.UserIndex).evento = Nothing
        
    Else 'Esta offline
        
        'Le doy el oro
        If oroADar > 0 Then
            Call modUsuarios.entregarOroOffline(Integrante.id, oroADar)
        End If
        
        'Lo tranporto
        Call modUsuarios.transportarUsuarioOffline(Integrante.id, Integrante.posOriginal.map, CByte(Integrante.posOriginal.x), CByte(Integrante.posOriginal.y))
        
        'Le quito la referencia
        Call modEventos.quitarReferenciaUsuarioEvento(Integrante.id)
    End If
        
End Sub
Public Sub desecharEquipo(equipo As tEquipoTablaTorneo, premioOro As Long, Optional ByVal EsOroInscripcion = False)
    
    Dim Integrante As tIntegrantesEquipoTorneo
    Dim loopIntegrante As Byte
            
    For loopIntegrante = 1 To equipo.cantidadIntegrantes
    
        If equipo.integrantes(loopIntegrante).Estado = eEstadoIntegranteEquipo.Jugando Then
            Call desecharIntegranteEquipo(equipo.integrantes(loopIntegrante), premioOro, EsOroInscripcion)
        End If

    Next loopIntegrante

End Sub

Public Sub cargarReglasbasicasHechizos(ByRef reglas() As Boolean)

    ReDim reglas(1 To 41) As Boolean  'TODO. Cambiar el 41 por una variable de cantidad max de hechizos
    
    Dim loopHechizo As Byte
    
    For loopHechizo = 1 To 41
        reglas(loopHechizo) = True
    Next
    
    reglas(eHechizos.Ayuda_espiritu_indomable) = False
    reglas(eHechizos.Debilidad) = False
    reglas(eHechizos.Implorar_ayuda) = False
    reglas(eHechizos.Invisibilidad) = False
    reglas(eHechizos.Invocar_elemetanl_fuego) = False
    reglas(eHechizos.Invocar_Mascotas) = False
    reglas(eHechizos.Invocar_Zombies) = False
    reglas(eHechizos.Invocoar_elemental_agua) = False
    reglas(eHechizos.Invocoar_elemental_tierra) = False
    reglas(eHechizos.Llamado_naturaleza) = False
    reglas(eHechizos.Provocar_Hambre) = False
    reglas(eHechizos.Resucitar) = False
    reglas(eHechizos.Terrible_Hambre) = False
    reglas(eHechizos.Torpeza) = False
    reglas(eHechizos.Mimetismo) = False

    
End Sub


'Enviar mensajaes a los equipos
Public Sub enviarMensajeEquipo(ByRef equipo As tEquipoTablaTorneo, ByVal mensaje As String, Optional nombreEvento As String = vbNullString)
   
    Dim loopIntegrante As Byte
    Dim UserIndex As Integer
       
    If nombreEvento <> vbNullString Then
        mensaje = nombreEvento & "-> " & mensaje
    End If
    
    For loopIntegrante = 1 To equipo.cantidadIntegrantes
    
        If equipo.integrantes(loopIntegrante).Estado = eEstadoIntegranteEquipo.Jugando Then
            UserIndex = equipo.integrantes(loopIntegrante).UserIndex
            'Si es mayor a 0 quiere decir que esta onlne
            If UserIndex > 0 Then
                EnviarPaquete Paquetes.MensajeTalk, mensaje, UserIndex, ToIndex
            End If
        End If
    Next
    
End Sub

Public Sub enviarMensajeEquipos(ByRef equipos() As tEquipoTablaTorneo, ByVal mensaje As String, Optional nombreEvento As String = vbNullString)

    Dim loopEquipo As Byte
        
    If nombreEvento <> vbNullString Then
        mensaje = nombreEvento & "-> " & mensaje
    End If
    
    For loopEquipo = 1 To UBound(equipos)
    
        If equipos(loopEquipo).Estado = eEstadoEquipoTorneo.participando Then
            Call enviarMensajeEquipo(equipos(loopEquipo), mensaje)
        End If
        
    Next loopEquipo
    
    Call LogTorneos(mensaje)
    EnviarPaquete Paquetes.MensajeTalk, "RT " & mensaje, 0, ToAdmins
End Sub

Public Sub enviarMensajeGlobal(mensaje As String, nombreEvento As String)
    EnviarPaquete Paquetes.MensajeTalk, nombreEvento & "-> " & mensaje, 0, ToAll

    Call LogTorneos(nombreEvento & "-> " & mensaje)
End Sub

Public Sub loguearTabla(tablaEquipos() As tEquipoTablaTorneo)
    Dim tabla As String
    Dim stringEquipo As String
    Dim loopC As Byte
    
    tabla = "Esta es la tabla final de posiciones:"
    
    For loopC = 1 To UBound(tablaEquipos)
    
        stringEquipo = stringEquipo = mid$(modTorneos.obtenerStringEquipo(tablaEquipos(loopC), False, eFormatoDisplayEquipo.completo), 1, 25)

        stringEquipo = tablaEquipos(loopC).idEquipo & "-" & stringEquipo & String$(26 - Len(stringEquipo), " ") & vbTab & _
                        "Combates Gan: " & Format(tablaEquipos(loopC).cantidadCombatesGanados, "#0") & _
                        ". Rounds Gan: " & Format(tablaEquipos(loopC).cantidadRoundsGanados, "#0") & _
                        ". Rounds Per: " & Format(tablaEquipos(loopC).cantidadRoundsJugados - tablaEquipos(loopC).cantidadRoundsGanados, "#0")

        If tablaEquipos(loopC).Estado = eEstadoEquipoTorneo.descalificado Then
            tabla = tabla & vbCrLf & loopC & ") " & stringEquipo & " (descalificado)"
        Else
            tabla = tabla & vbCrLf & loopC & ") " & stringEquipo
        End If
    Next loopC

    Call LogTorneos(tabla)
End Sub

Public Sub enviarTabla(tablaEquipos() As tEquipoTablaTorneo, nombreEvento As String, aQuien As Byte)
    Dim tabla As String
    Dim stringEquipo As String
    Dim loopC As Byte
    
    tabla = nombreEvento & "-> Esta es la tabla final de posiciones:"
    
    For loopC = 1 To UBound(tablaEquipos)
    
        stringEquipo = mid$(modTorneos.obtenerStringEquipo(tablaEquipos(loopC), False, NombreEquipo), 1, 25)

        stringEquipo = stringEquipo & String$(26 - Len(stringEquipo), " ") & vbTab & _
                        "Combates Gan: " & Format(tablaEquipos(loopC).cantidadCombatesGanados, "#0") & _
                        ". Rounds Gan: " & Format(tablaEquipos(loopC).cantidadRoundsGanados, "#0") & _
                        ". Rounds Per: " & Format(tablaEquipos(loopC).cantidadRoundsJugados - tablaEquipos(loopC).cantidadRoundsGanados, "#0")

        If tablaEquipos(loopC).Estado = eEstadoEquipoTorneo.descalificado Then
            tabla = tabla & vbCrLf & loopC & ") " & stringEquipo & " (descalificado)"
        Else
            tabla = tabla & vbCrLf & loopC & ") " & stringEquipo
        End If
    Next loopC
    
    EnviarPaquete Paquetes.MensajeTalk, tabla, 0, ToAll
    
    Call LogTorneos(tabla)
End Sub

Public Function obtenerCantidadEquiposJugando(tablaEquipos() As tEquipoTablaTorneo) As Byte
    Dim loopEquipo As Byte
    
    obtenerCantidadEquiposJugando = 0
    
    For loopEquipo = 1 To UBound(tablaEquipos)
        If tablaEquipos(loopEquipo).Estado = eEstadoEquipoTorneo.participando Then
            obtenerCantidadEquiposJugando = obtenerCantidadEquiposJugando + 1
        End If
    Next
    
End Function

'**
' Devuelve una tabla ordenada por:
' 1º Cantidad de combates ganados
' 2º Mayor diferencia cantidad de rounds ganados - cantidad de rounds perdidos
' 3º Menor Cantidad de rounds perdidos.
' 4º Estado (primero no descalificado)
' 5º ID del equipo (que debe coincidir con el orden en el cual fue inscripto)
'@param        tablaEquipos() Required. tEquipoTablaTorneo object.
'@rem
Public Sub ordenarTabla(tablaEquipos() As tEquipoTablaTorneo)

    Dim loopEquipo As Byte
    Dim primero As Byte
    Dim loopBusqueda As Byte
    
    Dim diferencia1 As Integer
    Dim diferencia2 As Integer
    
    Dim auxTablaEquipo As tEquipoTablaTorneo
    
    For loopEquipo = 1 To UBound(tablaEquipos)
    
        primero = loopEquipo
    
        For loopBusqueda = loopEquipo + 1 To UBound(tablaEquipos)
    
            '1º Cantidad de comabtes ganados
            If tablaEquipos(primero).cantidadCombatesGanados < tablaEquipos(loopBusqueda).cantidadCombatesGanados Then
            
                primero = loopBusqueda
                
            ElseIf tablaEquipos(primero).cantidadCombatesGanados = tablaEquipos(loopBusqueda).cantidadCombatesGanados Then
                
                diferencia1 = CInt(tablaEquipos(primero).cantidadRoundsGanados) - CInt(tablaEquipos(primero).cantidadRoundsJugados - tablaEquipos(primero).cantidadRoundsGanados)
                diferencia2 = CInt(tablaEquipos(loopBusqueda).cantidadRoundsGanados) - CInt(tablaEquipos(loopBusqueda).cantidadRoundsJugados - tablaEquipos(loopBusqueda).cantidadRoundsGanados)
                
                '2º Cantidad de rounds ganados
                If diferencia1 < diferencia2 Then
                
                    primero = loopBusqueda
                    
                ElseIf diferencia1 = diferencia2 Then
                    
                    '3º Cantidad de rounds jugados (lo que es lo mismo que decir menor cantidad de rounds perdidos)
                    If tablaEquipos(primero).cantidadRoundsJugados > tablaEquipos(loopBusqueda).cantidadRoundsJugados Then
                        
                        primero = loopBusqueda
                        
                    ElseIf tablaEquipos(primero).cantidadRoundsJugados = tablaEquipos(loopBusqueda).cantidadRoundsJugados Then
                    
                        '4º Estado del equipo
                        If tablaEquipos(primero).Estado = eEstadoEquipoTorneo.descalificado And Not tablaEquipos(loopBusqueda).Estado = eEstadoEquipoTorneo.descalificado Then
                            primero = loopBusqueda
                        Else
                            '5º Por quien envio primero el /PARTICIPAR
                            If tablaEquipos(primero).tickInscripcion > tablaEquipos(loopBusqueda).tickInscripcion Then
                                primero = loopBusqueda
                            End If
                        End If
                    End If
                End If
            End If
        
        Next loopBusqueda
        
        auxTablaEquipo = tablaEquipos(loopEquipo)
        tablaEquipos(loopEquipo) = tablaEquipos(primero)
        tablaEquipos(primero) = auxTablaEquipo
        
    Next loopEquipo
End Sub


Public Sub cargarTabla(idTabla As Long, tabla() As tEquipoTablaTorneo)
    Dim sql As String
    Dim infoTabla As Recordset
    Dim cantidadEquipos As Byte
    Dim loopEquipo As Byte
    
    Dim infoEquipo As Recordset
    Dim loopIntegrante As Byte
    
   
    sql = "SELECT * FROM " & DB_NAME_PRINCIPAL & ".juego_torneos_tablaequipos WHERE IDTABLA=" & idTabla

    Set infoTabla = conn.Execute(sql, cantidadEquipos)

    ReDim tabla(1 To cantidadEquipos) As tEquipoTablaTorneo
    
    For loopEquipo = 1 To cantidadEquipos
        With tabla(loopEquipo)
                .idEquipo = infoTabla!idEquipo
                
                'Info de los combates
                .cantidadCombatesGanados = infoTabla!COMBATESGANADOS
                .cantidadCombatesEmpatados = infoTabla!COMBATESEMPATADOS
                .cantidadCombatesJugados = infoTabla!COMBATESJUGADOS
            
                'Info de los rounds
                .cantidadRoundsGanados = infoTabla!ROUNDSGANADOS
                .cantidadRoundsJugados = infoTabla!ROUNDSJUGADOS

                'Estado del equipo
                If infoTabla!Estado = "PARTICIPANDO" Then
                    .Estado = eEstadoEquipoTorneo.participando
                ElseIf infoTabla!Estado = "TERMINO" Then
                    .Estado = eEstadoEquipoTorneo.termino
                Else
                    .Estado = eEstadoEquipoTorneo.descalificado
                End If
                
                'Cargamos el equipo Seteamos la cantidad de integrantes. Inicialmente ninguno esta descalificado
                sql = "SELECT * FROM " & DB_NAME_PRINCIPAL & ".juego_torneos_tablaequipos_integrantes WHERE IDTABLAEQUIPO=" & idTabla & " AND IDEquipo=" & .idEquipo
                Set infoEquipo = conn.Execute(sql, .cantidadIntegrantes)
                
                .cantidadIntegrantesDescalificados = 0
                
                ReDim .integrantes(1 To .cantidadIntegrantes) As tIntegrantesEquipoTorneo
                
                For loopIntegrante = 1 To .cantidadIntegrantes
                    .integrantes(loopIntegrante).id = infoEquipo!idjugador
                    .integrantes(loopIntegrante).posOriginal.map = infoEquipo!MAPAORIGINAL
                    .integrantes(loopIntegrante).posOriginal.x = infoEquipo!MAPA_X
                    .integrantes(loopIntegrante).posOriginal.y = infoEquipo!MAPA_Y
                    .integrantes(loopIntegrante).nick = infoEquipo!nick
                    .integrantes(loopIntegrante).cantidadAdvertencias = 0
                    .integrantes(loopIntegrante).UserIndex = IDIndex(infoEquipo!idjugador)
                    
                    If infoEquipo!Estado = "JUGANDO" Then
                        .integrantes(loopIntegrante).Estado = eEstadoIntegranteEquipo.Jugando
                    ElseIf infoEquipo!Estado = "DESCALIFICADO" Then
                        .integrantes(loopIntegrante).Estado = eEstadoIntegranteEquipo.Descalificando
                        .cantidadIntegrantesDescalificados = .cantidadIntegrantesDescalificados + 1
                    End If
                    
                    'Siguiente integrante
                    infoEquipo.MoveNext
                Next loopIntegrante
        End With
        'Siguiente equipo
        infoTabla.MoveNext
    Next loopEquipo
End Sub
Private Sub nuevaTabla(tablaEquipos() As tEquipoTablaTorneo)

    Dim sql As String
    Dim sqlIntegrantes As String
    Dim infoTabla As ADODB.Recordset
    Dim loopE As Byte
    Dim idTabla As Long
    Dim loopJugador As Byte

    'Tengo que obtener un id para la tabla
    sql = "UPDATE " & DB_NAME_PRINCIPAL & ".juego_torneos_tablaequipos_id SET ULTIMOIDTABLA = LAST_INSERT_ID(ULTIMOIDTABLA + 1) WHERE ULTIMOIDTABLA>=0;"
    conn.Execute sql, , adExecuteNoRecords
    
    sql = "SELECT LAST_INSERT_ID() AS IDTABLA;"
    Set infoTabla = conn.Execute(sql, , adCmdText)
    
    idTabla = CLng(infoTabla!idTabla)
    
    infoTabla.Close
    Set infoTabla = Nothing
    
    'Asigno el ID de la tabla a cada equipo
    For loopE = 1 To UBound(tablaEquipos)
        tablaEquipos(loopE).idTablaPersistencia = idTabla
    Next
    
    'Agrego los equipos
    sql = "INSERT INTO " & DB_NAME_PRINCIPAL & ".juego_torneos_tablaequipos(IDTABLA , IDEQUIPO) VALUES"
     'Integrantes del equipo
    sqlIntegrantes = "INSERT INTO " & DB_NAME_PRINCIPAL & ".juego_torneos_tablaequipos_integrantes(IDTABLAEQUIPO, IDEQUIPO, IDJUGADOR, NICK, MAPAORIGINAL, MAPA_X, MAPA_Y) VALUES"
            
    For loopE = 1 To UBound(tablaEquipos)
        With tablaEquipos(loopE)
            ' Genero el sql para el equipo
            If loopE = UBound(tablaEquipos) Then
                sql = sql & "(" & idTabla & "," & .idEquipo & ");"
            Else
                sql = sql & "(" & idTabla & "," & .idEquipo & "), "
            End If
            
            For loopJugador = 1 To .cantidadIntegrantes
                If loopE = UBound(tablaEquipos) And loopJugador = .cantidadIntegrantes Then
                    sqlIntegrantes = sqlIntegrantes & " (" & idTabla & "," & .idEquipo & "," & .integrantes(loopJugador).id & _
                    ", '" & .integrantes(loopJugador).nick & "', " & _
                    .integrantes(loopJugador).posOriginal.map & ", " & _
                    .integrantes(loopJugador).posOriginal.x & ", " & .integrantes(loopJugador).posOriginal.y & ");"
                Else
                     sqlIntegrantes = sqlIntegrantes & " (" & idTabla & "," & .idEquipo & "," & .integrantes(loopJugador).id & _
                    ", '" & .integrantes(loopJugador).nick & "', " & _
                     .integrantes(loopJugador).posOriginal.map & ", " & _
                    .integrantes(loopJugador).posOriginal.x & ", " & .integrantes(loopJugador).posOriginal.y & "),"
                End If
            Next loopJugador
        End With
    Next loopE
    
    
    'Guardo los equipos
    conn.Execute (sql), , adExecuteNoRecords
    'Guardo los integrantes de los equipos
    conn.Execute (sqlIntegrantes), , adExecuteNoRecords
End Sub
Private Sub actualizarTabla(tablaEquipos() As tEquipoTablaTorneo)
    Dim sql As String
    Dim sqlIntegrantes As String
    Dim loopE As Byte
    Dim idTabla As Long
    Dim loopJugador As Byte

    Dim sEstado As String
    
    idTabla = tablaEquipos(1).idTablaPersistencia
    'Agrego los equipos
   
     'Integrantes del equipo
    sqlIntegrantes = "INSERT INTO " & DB_NAME_PRINCIPAL & ".juego_torneos_tablaequipos_integrantes(IDTABLAEQUIPO, IDEQUIPO, IDJUGADOR, MAPAORIGINAL, MAPA_X, MAPA_Y) VALUES"
            
    For loopE = 1 To UBound(tablaEquipos)

        With tablaEquipos(loopE)
        
            If .Estado = eEstadoEquipoTorneo.participando Then
                sEstado = "PARTICIPANDO"
            ElseIf .Estado = eEstadoEquipoTorneo.descalificado Then
                sEstado = "DESCALIFICADO"
            Else
                sEstado = "TERMINO"
            End If
            
            sql = "UPDATE " & DB_NAME_PRINCIPAL & ".juego_torneos_tablaequipos SET Estado='" & sEstado & "'" & _
                ", COMBATESGANADOS=" & .cantidadCombatesGanados & _
                ", COMBATESEMPATADOS=" & .cantidadCombatesEmpatados & _
                ", COMBATESJUGADOS=" & .cantidadCombatesJugados & _
                ", ROUNDSGANADOS=" & .cantidadRoundsGanados & _
                ", ROUNDSJUGADOS=" & .cantidadRoundsJugados & _
                " WHERE IDTABLA=" & idTabla & " AND IDEQUIPO=" & .idEquipo

            conn.Execute sql, , adExecuteNoRecords
            
            For loopJugador = 1 To .cantidadIntegrantes
                If .integrantes(loopJugador).Estado = eEstadoIntegranteEquipo.Jugando Then
                    sEstado = "JUGANDO"
                Else
                    sEstado = "DESCALIFICADO"
                End If
                
                sqlIntegrantes = "UPDATE " & DB_NAME_PRINCIPAL & ".juego_torneos_tablaequipos_integrantes SET Estado='" & sEstado & "'" & _
                    " WHERE IDTABLAEQUIPO=" & idTabla & " AND IDEQUIPO=" & .idEquipo & " AND IDJUGADOR=" & .integrantes(loopJugador).id

                conn.Execute (sqlIntegrantes), , adExecuteNoRecords
           Next loopJugador
        End With
        'Guardo los equipos
    Next loopE
End Sub
Public Sub guardarTabla(tablaEquipos() As tEquipoTablaTorneo)

'Actualizo o guardo'
If tablaEquipos(1).idTablaPersistencia = 0 Then
    Call nuevaTabla(tablaEquipos)
    Call actualizarTabla(tablaEquipos)
Else
    Call actualizarTabla(tablaEquipos)
End If

End Sub
Public Sub testOrdenarTabla()

    Dim tabla(1 To 6) As tEquipoTablaTorneo

    tabla(1).idEquipo = 1
    tabla(1).cantidadCombatesGanados = 4
    tabla(1).cantidadRoundsGanados = 2
    tabla(1).cantidadRoundsJugados = 17
    tabla(1).Estado = eEstadoEquipoTorneo.termino
    ReDim tabla(1).integrantes(1 To 2)
    tabla(1).integrantes(1).nick = "Goku II"
    tabla(1).integrantes(2).nick = "Fucking Gobling"
    tabla(1).integrantes(1).id = 1
    tabla(1).integrantes(2).id = 2
    tabla(1).cantidadIntegrantes = 2

    tabla(2).idEquipo = 2
    tabla(2).cantidadCombatesGanados = 3
    tabla(2).cantidadRoundsGanados = 6
    tabla(2).cantidadRoundsJugados = 6
    tabla(2).Estado = eEstadoEquipoTorneo.termino
    ReDim tabla(2).integrantes(1 To 2)
    tabla(2).integrantes(1).nick = "Picton"
    tabla(2).integrantes(2).nick = "Pom"
    tabla(2).integrantes(1).id = 10
    tabla(2).integrantes(2).id = 20
    tabla(2).cantidadIntegrantes = 2

    tabla(3).idEquipo = 3
    tabla(3).cantidadCombatesGanados = 1
    tabla(3).cantidadRoundsGanados = 3
    tabla(3).cantidadRoundsJugados = 11
    tabla(3).Estado = eEstadoEquipoTorneo.termino
    ReDim tabla(3).integrantes(1 To 2)
    tabla(3).integrantes(1).nick = "Althyr"
    tabla(3).integrantes(2).nick = "Shinichi"
    tabla(3).cantidadIntegrantes = 2

    tabla(4).idEquipo = 4
    tabla(4).cantidadCombatesGanados = 1
    tabla(4).cantidadRoundsGanados = 4
    tabla(4).cantidadRoundsJugados = 10
    tabla(4).Estado = eEstadoEquipoTorneo.termino
    ReDim tabla(4).integrantes(1 To 2)
    tabla(4).integrantes(1).nick = "Pepe IV"
    tabla(4).integrantes(1).nick = "Jose"
    tabla(4).cantidadIntegrantes = 2

    tabla(5).idEquipo = 5
    tabla(5).cantidadCombatesGanados = 1
    tabla(5).cantidadRoundsGanados = 4
    tabla(5).cantidadRoundsJugados = 10
    tabla(5).Estado = eEstadoEquipoTorneo.termino
    ReDim tabla(5).integrantes(1 To 2)
    tabla(5).integrantes(1).nick = "Pepe IV"
    tabla(5).integrantes(1).nick = "Jose"
    tabla(5).cantidadIntegrantes = 2

    tabla(6).idEquipo = 6
    tabla(6).cantidadCombatesGanados = 1
    tabla(6).cantidadRoundsGanados = 4
    tabla(6).cantidadRoundsJugados = 10
    tabla(6).Estado = eEstadoEquipoTorneo.termino
    ReDim tabla(6).integrantes(1 To 2)
    tabla(6).integrantes(1).nick = "Pepe IV"
    tabla(6).integrantes(1).nick = "Jose"
    tabla(6).cantidadIntegrantes = 2

   ' Call modTorneos.guardarTabla(tabla)

    Dim tabla2() As tEquipoTablaTorneo
    
    Call modTorneos.cargarTabla(28, tabla2)

End Sub

Public Function obtenerNombreIdentificacionPersonajes(integrantesIndexs() As Integer, comoIdentificarEquipo As eIdentificacionEquipos) As String
    Dim UserIndex As Integer
    
    ' Con un solo usuario nos alcanza
    UserIndex = integrantesIndexs(1)
    
    Select Case comoIdentificarEquipo
    
        Case eIdentificacionEquipos.identificaFaccion
        
            If UserList(UserIndex).faccion.ArmadaReal = 1 Then
                obtenerNombreIdentificacionPersonajes = "Armada Real"
            ElseIf UserList(UserIndex).faccion.FuerzasCaos = 1 Then
                obtenerNombreIdentificacionPersonajes = "Legión Oscura"
            ElseIf UserList(UserIndex).faccion.alineacion = eAlineaciones.caos Then
                obtenerNombreIdentificacionPersonajes = "Ejército Escarlata"
            ElseIf UserList(UserIndex).faccion.alineacion = eAlineaciones.Real Then
                obtenerNombreIdentificacionPersonajes = "Ejército Índigo"
            ElseIf UserList(UserIndex).faccion.alineacion = eAlineaciones.Neutro Then
                obtenerNombreIdentificacionPersonajes = "Rebeldes"
            End If
        Case eIdentificacionEquipos.identificaClan
        
            If Not UserList(UserIndex).ClanRef Is Nothing Then
                obtenerNombreIdentificacionPersonajes = "<" & UserList(UserIndex).ClanRef.getNombre & ">"
            Else
                obtenerNombreIdentificacionPersonajes = vbNullString
            End If
            
        Case Else ' Nombre de los personajes
        
            obtenerNombreIdentificacionPersonajes = vbNullString
            
    End Select
    

End Function
