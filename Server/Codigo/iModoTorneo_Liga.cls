VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iModoTorneo_Liga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iModoTorneo

Private Type tHojaRuta
    posTablaContrincante As Byte
    nombreRonda As String
End Type

Private idUnicoEvento As Long
Private nombreEvento As String 'Nombre del "sub" evento
Private necesidadRoundGanadosGanador As Byte

Private tablaEquipos() As tEquipoTablaTorneo 'Tabla donde esta la informacion de los equipos
Private cantidadEquipos As Byte

Private cantidadRondas As Byte
Private crucesPendientes As Integer

Private deathMatchs() As iModoTorneo_DeathMach
Private deathMatchsLibres() As Boolean
'Estoy arrays van de 1 a al cantidad de equipos que haya
'El indice es la posicion del equipo en la tabla
Private eventosEquipo() As iModoTorneo_DeathMach
Private rondaActualEquipo() As Byte
Private hojaRuta() As tHojaRuta

Private Estado As eEstadoEvento

Private rings() As tRing 'rings que tiene a disposicion pra utilizar el evento.

Private idaYvuelta As Byte

Private maxima_cantidad_advertencias As Byte ' Cantidad de llamados a un usuario que debe participar antes de ser descalificado


Private Const TIEMPO_SEGUNDO_COMBATE = 2
Private Const CANTIDAD_DE_COMBATES_GRUPO_INICIAL = 4

Private Sub Class_Initialize()
    Estado = eEstadoEvento.Preparacion
    idaYvuelta = 0
End Sub

Private Sub Class_Terminate()
    Call LogTorneos(nombreEvento & "-> Se elimina de memoria.")
End Sub

Public Sub iModoTorneo_cargar(idliga As Long)

    Dim sql As String
    Dim idTabla As Long
    Dim infoLiga As Recordset
    Dim auxCrucesPendientes As Integer
    Dim loopEquipo As Byte
    Dim info As Recordset
    ' Obtengo los datos globales
    idUnicoEvento = idliga
    sql = "SELECT * FROM " & DB_NAME_PRINCIPAL & ".juego_torneos_liga WHERE IDEVENTO=" & idliga
    
    Set infoLiga = conn.Execute(sql)
    
    idTabla = infoLiga!IDTABLAEQUIPOS
    crucesPendientes = infoLiga!crucesPendientes
    
    Call modTorneos.cargarTabla(idTabla, tablaEquipos)
    
    ' Cargo la situacion de cada equipo
     For loopEquipo = 1 To UBound(tablaEquipos)
        sql = "SELECT NUMEROETAPA FROM " & DB_NAME_PRINCIPAL & ".juego_torneos_equipoetapa WHERE IDEVENTO=" & idUnicoEvento & " AND IDEQUIPO=" & tablaEquipos(loopEquipo).idEquipo

        Set info = conn.Execute(sql)
        rondaActualEquipo(loopEquipo) = info!NUMEROETAPA
    Next
End Sub

Public Sub crearPersistencia()
    Dim sql As String
    Dim loopEquipo As Byte
    
    idUnicoEvento = obtenerIDParaPersistirEvento()
    'Persisto la tabla
    
    Call modTorneos.guardarTabla(tablaEquipos)
    
    'Persisto info general (relacion evento tabla)
    sql = "INSERT INTO " & DB_NAME_PRINCIPAL & ".juego_torneos_liga(IDEVENTO, IDTABLAEQUIPOS, CRUCESPENDIENTES) " & _
            "VALUES(" & idUnicoEvento & "," & tablaEquipos(1).idTablaPersistencia & "," & crucesPendientes & ");"

    Call conn.Execute(sql)
    
    'Persisto cada equipo y en que sub evento esta
    sql = "INSERT INTO " & DB_NAME_PRINCIPAL & ".juego_torneos_equipoetapa(IDEVENTO, IDEQUIPO) VALUES"
    
    For loopEquipo = 1 To UBound(tablaEquipos)
        If loopEquipo = UBound(tablaEquipos) Then
            sql = sql & "(" & idUnicoEvento & "," & tablaEquipos(loopEquipo).idEquipo & ");"
        Else
            sql = sql & "(" & idUnicoEvento & "," & tablaEquipos(loopEquipo).idEquipo & "), "
        End If
    Next
    
    Call conn.Execute(sql)
    
    Call modTorneos.guardarTabla(tablaEquipos)
End Sub
Public Sub iModoTorneo_guardar()
    Dim sql As String
    Dim loopEquipo As Byte
    'Persisto la tabla de equipos
    Call modTorneos.guardarTabla(tablaEquipos)
    
    'Persisto al cantidad de combates que faltan
    sql = "UPDATE " & DB_NAME_PRINCIPAL & ".juego_torneos_liga SET CRUCESPENDIENTES = " & crucesPendientes & " WHERE IDEVENTO=" & idUnicoEvento
    Call conn.Execute(sql)
    
    'Persisto cada equipo y en que sub evento esta
    For loopEquipo = 1 To UBound(tablaEquipos)
        sql = "UPDATE " & DB_NAME_PRINCIPAL & ".juego_torneos_equipoetapa SET NUMEROETAPA = " & rondaActualEquipo(loopEquipo) & " WHERE IDEVENTO=" & idUnicoEvento & " AND IDEQUIPO=" & tablaEquipos(loopEquipo).idEquipo

        Call conn.Execute(sql)
    Next
End Sub

Friend Function iModoTorneo_obtenerEquipo(idEquipo As Byte) As tEquipoTablaTorneo
    Dim posEquipoTabla As Byte
    
    posEquipoTabla = modTorneos.obtenerPosTablaIDEquipo(tablaEquipos, idEquipo)
    
    If Not eventosEquipo(posEquipoTabla) Is Nothing Then
        iModoTorneo_obtenerEquipo = eventosEquipo(posEquipoTabla).iModoTorneo_obtenerEquipo(idEquipo)
    Else
        iModoTorneo_obtenerEquipo = tablaEquipos(posEquipoTabla)
    End If
End Function

Public Sub iModoTorneo_cancelar()

    Dim loopSubEvento As Byte
    Dim i As Integer
    
    Dim tablaAux() As tEquipoTablaTorneo
    
    For loopSubEvento = 1 To UBound(deathMatchs)
    
        Call deathMatchs(loopSubEvento).iModoTorneo_cancelar
        
        tablaAux() = deathMatchs(loopSubEvento).iModoTorneo_obtenerTabla
        
        For i = 1 To UBound(tablaAux)
            tablaEquipos(modTorneos.obtenerPosTablaIDEquipo(tablaAux, tablaAux(i).idEquipo)) = tablaAux(i)
        Next i
        
    Next
    
    Call terminar
    
End Sub

Public Sub iModoTorneo_comenzar(delay As Byte)
    'Voy a genrar y configurar los deathmatchs que necesito
    Dim cantidadDeathMatchs As Byte
    Dim loopDeathMatch As Byte

    'Guardo la tabla con la que voy a procesar
    '*************************************************************************.
    If idUnicoEvento = 0 Then
        Me.crearPersistencia
    End If
    '**************************************************************************
    
    'Cantidad de deatmatchs que necesitamos para los combates
    cantidadDeathMatchs = cantidadEquipos \ 2 'Divison entera
    
    ReDim deathMatchs(1 To cantidadDeathMatchs) As iModoTorneo_DeathMach
    ReDim deathMatchsLibres(1 To cantidadDeathMatchs) As Boolean
    
    For loopDeathMatch = 1 To cantidadDeathMatchs
    
        deathMatchsLibres(loopDeathMatch) = True
        
        Set deathMatchs(loopDeathMatch) = New iModoTorneo_DeathMach
        
        'Configuro
        Call deathMatchs(loopDeathMatch).iModoTorneo_setRing(rings(loopDeathMatch))
        Call deathMatchs(loopDeathMatch).iModoTorneo_setNecesidadRoundGanadosGanador(necesidadRoundGanadosGanador)
        Call deathMatchs(loopDeathMatch).iModoTorneo_setMaximaCantidadAdvertencias(maxima_cantidad_advertencias)
        
    Next loopDeathMatch
    
    ' Generamos el fixture
    Call generarFixture
    
    ' Que comiencen los ruces
    Call generarCruces
End Sub

Public Function iModoTorneo_isTerminado() As Boolean
    iModoTorneo_isTerminado = (Estado = eEstadoEvento.Terminado)
End Function

Private Function iModoTorneo_obtenerIDPersistencia() As Integer
    iModoTorneo_obtenerIDPersistencia = idUnicoEvento
End Function

Public Sub iModoTorneo_pasarMinuto()
    Dim loopSubEvento As Byte
    
    'Les aviso a mis sub eventos que paso un minuto
    For loopSubEvento = 1 To UBound(deathMatchs)
    
        If Not deathMatchs(loopSubEvento) Is Nothing Then
            If Not deathMatchs(loopSubEvento).iModoTorneo_isTerminado Then
                deathMatchs(loopSubEvento).iModoTorneo_pasarMinuto
                If deathMatchs(loopSubEvento).iModoTorneo_isTerminado Then Call procesarGanador(deathMatchs(loopSubEvento))
            End If
        End If
    
    Next loopSubEvento
End Sub

Public Sub iModoTorneo_setMaximaCantidadAdvertencias(ByVal cantidad As Byte)
    maxima_cantidad_advertencias = cantidad
End Sub

Public Sub iModoTorneo_setNombreEvento(nombre_ As String)
    nombreEvento = nombre_
End Sub


Private Sub procesarGanador(ByVal deathmatch As iModoTorneo_DeathMach)

    Dim tabla() As tEquipoTablaTorneo
    Dim posGanador As Byte
    Dim posPerdedor As Byte
    Dim auxByte As Byte
    
    
    'Obtengo la tabla del evento
    tabla = deathmatch.iModoTorneo_obtenerTabla()
    
    'En el primer lugar esta el ganador..
    'Estos equipo en mi tabala general en que posicion estan?
    posGanador = obtenerPosTablaIDEquipo(tablaEquipos, tabla(1).idEquipo)
    posPerdedor = obtenerPosTablaIDEquipo(tablaEquipos, tabla(2).idEquipo)
    
    'Actualizo la tabla
    tablaEquipos(posGanador) = tabla(1)
    tablaEquipos(posPerdedor) = tabla(2)
    
    'La posGanador seimpre va a ser la menor de la tabla (o la mayorr mirando de abajo hacia arriba)
    'Les desrelaciono el deathtmatch con el equipo
    Set eventosEquipo(posGanador) = Nothing
    Set eventosEquipo(posPerdedor) = Nothing
    
    crucesPendientes = crucesPendientes - 1
    
    Call LogTorneos(nombreEvento & " Liga -> Cruces pendientes " & crucesPendientes)
     
    'Libero el deathmatch
    For auxByte = 1 To UBound(deathMatchs)
        If deathMatchs(auxByte) Is deathmatch Then
            deathMatchsLibres(auxByte) = True
        End If
    Next
    
    'Es indiferente que me fije en cualqueira de los dos equipos
    If rondaActualEquipo(posPerdedor) = cantidadRondas Then
    
        If crucesPendientes = 0 Then
            'Ahora si ordenamos la tabla
            Call modTorneos.loguearTabla(tablaEquipos)
            Call modTorneos.ordenarTabla(tablaEquipos)
            
            Call LogTorneos(nombreEvento & "-> " & " la liga termino.")
            Call terminar
        Else
            Call modTorneos.enviarMensajeEquipo(tablaEquipos(posPerdedor), "Tu participación terminó pero el torneo aún no. Gracias por participar, en breve se informará el resultado final del torneo y, si corresponde, se te entregará el premio.", nombreEvento)
            Call modTorneos.enviarMensajeEquipo(tablaEquipos(posGanador), "Tu participación terminó pero el torneo aún no. Gracias por participar, en breve se informará el resultado final del torneo y, si corresponde, se te entregará el premio.", nombreEvento)
            
            Call LogTorneos(nombreEvento & "-> " & " Se terminó la participación de los equipos de " & modTorneos.obtenerStringEquipo(tablaEquipos(posPerdedor), False, eFormatoDisplayEquipo.completo) & " y " & modTorneos.obtenerStringEquipo(tablaEquipos(posGanador), False, eFormatoDisplayEquipo.completo))
            
            If tablaEquipos(posPerdedor).Estado = eEstadoEquipoTorneo.participando Then
                tablaEquipos(posPerdedor).Estado = eEstadoEquipoTorneo.termino
                Call modTorneos.desecharEquipo(tablaEquipos(posPerdedor), 0)
            End If
            
            If tablaEquipos(posGanador).Estado = eEstadoEquipoTorneo.participando Then
                tablaEquipos(posGanador).Estado = eEstadoEquipoTorneo.termino
                Call modTorneos.desecharEquipo(tablaEquipos(posGanador), 0)
            End If
        End If
    Else
        'Pasan a la siguiente ronda
        rondaActualEquipo(posPerdedor) = rondaActualEquipo(posPerdedor) + 1
        rondaActualEquipo(posGanador) = rondaActualEquipo(posGanador) + 1
    
        Call continuarNuevoCruce(posPerdedor, True)
        Call continuarNuevoCruce(posGanador, True)
    End If
End Sub

Public Sub terminar()
    Dim loopC  As Byte

    'Libero los deahmatchs
    For loopC = 1 To UBound(deathMatchs)
        If Not deathMatchs(loopC) Is Nothing Then
            Set deathMatchs(loopC) = Nothing
            LogTorneos (nombreEvento & "> Libero el deathmatch " & loopC)
        End If
    Next
    
    'Reseteo la memoria
    For loopC = 1 To cantidadEquipos
        Set eventosEquipo(tablaEquipos(loopC).idEquipo) = Nothing
    Next loopC
    
    Estado = eEstadoEvento.Terminado
End Sub

Public Function continuarNuevoCruce(posEquipo As Byte, esGanador As Boolean, Optional ByVal tiempoComienzo As Byte = TIEMPO_SEGUNDO_COMBATE) As Integer
    Dim posEquipoContrario As Byte
    Dim rondaActual As Byte
    
    continuarNuevoCruce = 0 'Retornamos la posicion en la tabla del equipo con el cual va a jugar
    
    rondaActual = rondaActualEquipo(posEquipo)
    posEquipoContrario = hojaRuta(posEquipo, rondaActual).posTablaContrincante
    
    If posEquipoContrario > 0 Then
        If rondaActualEquipo(hojaRuta(posEquipo, rondaActual).posTablaContrincante) = rondaActual Then
            Call armarCruce(posEquipo, posEquipoContrario, hojaRuta(posEquipo, rondaActual).nombreRonda, tiempoComienzo)
            continuarNuevoCruce = posEquipoContrario
        Else
            'El equipo con el que tiene que jugar aún no esta definido.
            Call modTorneos.enviarMensajeEquipo(tablaEquipos(posEquipo), "Tu próximo adversario (" & modTorneos.obtenerStringEquipo(tablaEquipos(posEquipoContrario), True, eFormatoDisplayEquipo.completo) & ") todavía se encuentra combatiendo. Tomate un descanso hasta que termine su combate.", nombreEvento)
        End If
    Else
        Call modTorneos.enviarMensajeEquipo(tablaEquipos(posEquipo), "En la " & hojaRuta(posEquipo, rondaActual).nombreRonda & " no participás ya que es tu fecha libre. Debido a que hay una cantidad impar de equipos, para cada equipo hay una fecha en la que no juega.", nombreEvento)
        
        If Not rondaActualEquipo(posEquipo) = cantidadRondas Then
            rondaActualEquipo(posEquipo) = rondaActualEquipo(posEquipo) + 1
            continuarNuevoCruce = continuarNuevoCruce(posEquipo, True)
        Else
            Call modTorneos.enviarMensajeEquipo(tablaEquipos(posEquipo), "Tu participación terminó pero el torneo aún no. Gracias por participar, en breve se informará el resultado final del torneo y, si corresponde, se te entregará el premio.", nombreEvento)
            tablaEquipos(posEquipo).Estado = eEstadoEquipoTorneo.termino
            Call modTorneos.desecharEquipo(tablaEquipos(posEquipo), 0)
            
            Call LogTorneos(nombreEvento & "-> " & " Se terminó la participación de los equipos de " & modTorneos.obtenerStringEquipo(tablaEquipos(posEquipo), False, eFormatoDisplayEquipo.completo))
        End If
    End If
            
End Function


Private Sub iModoTorneo_setParametros(parametros As String)

    idaYvuelta = CByte(val(parametros))
    
End Sub

Friend Sub iModoTorneo_setRing(ring As tRing)
End Sub

'Tabla de equipos
Friend Sub iModoTorneo_setTablaEquipos(tabla() As tEquipoTablaTorneo)

    Dim loopEquipo As Byte

    tablaEquipos = tabla
   
    cantidadEquipos = UBound(tablaEquipos)

    If idaYvuelta = 1 Then
        cantidadRondas = (cantidadEquipos - 1) * 2 + cantidadEquipos Mod 2
    Else
        cantidadRondas = cantidadEquipos - 1 + cantidadEquipos Mod 2
    End If
    
    crucesPendientes = (cantidadEquipos \ 2) * cantidadRondas
    
    ReDim rondaActualEquipo(1 To cantidadEquipos)
    ReDim relacionIDPosicion(1 To cantidadEquipos)
    ReDim eventosEquipo(1 To cantidadEquipos)
    ReDim hojaRuta(1 To cantidadEquipos, 1 To cantidadRondas)
    
    For loopEquipo = 1 To cantidadEquipos
        rondaActualEquipo(loopEquipo) = 1
    Next loopEquipo
    
    Call LogTorneos(nombreEvento & " Liga -> Cruces pendientes " & crucesPendientes)
End Sub

Private Sub armarCruce(posEnTablaEq1 As Byte, posEnTablaEq2 As Byte, nombreCruce As String, tiempoDelay As Byte)
    Dim tablaAux(1 To 2) As tEquipoTablaTorneo
        
    tablaAux(1) = tablaEquipos(posEnTablaEq1)
    tablaAux(2) = tablaEquipos(posEnTablaEq2)
    
    'Tengo que buscar un deathmatch....
    Dim numeroDeathLibre As Byte
    
    For numeroDeathLibre = 1 To UBound(deathMatchs)
        
        If deathMatchsLibres(numeroDeathLibre) Then
        
                deathMatchsLibres(numeroDeathLibre) = False 'Marco el death como usado
                
                Set eventosEquipo(posEnTablaEq1) = deathMatchs(numeroDeathLibre)
                Set eventosEquipo(posEnTablaEq2) = deathMatchs(numeroDeathLibre)
    
                Call deathMatchs(numeroDeathLibre).iModoTorneo_setNombreEvento(nombreEvento & " " & nombreCruce)
                Call deathMatchs(numeroDeathLibre).iModoTorneo_setTablaEquipos(tablaAux)

                Call deathMatchs(numeroDeathLibre).iModoTorneo_comenzar(tiempoDelay)
    
                If deathMatchs(numeroDeathLibre).iModoTorneo_isTerminado Then Call procesarGanador(deathMatchs(numeroDeathLibre))
            
                Exit Sub
        End If
    Next
    
    Call LogTorneos(nombreEvento & "-> ¡ERROR! No se encontro deathmatch libre!")
    
End Sub
Private Sub generarCruces()
Dim loopEquipo As Byte
Dim equipos() As Byte
Dim combatesArmados As Byte
Dim equipoConElQueVaAluchar As Byte

ReDim equipos(1 To cantidadEquipos) As Byte
' Proceso cada equipo

combatesArmados = 0

For loopEquipo = 1 To cantidadEquipos
    '¿Ya lo procese a este?
    If equipos(loopEquipo) = 0 And Not tablaEquipos(loopEquipo).Estado = termino Then
    
        equipoConElQueVaAluchar = continuarNuevoCruce(loopEquipo, False, (combatesArmados / CANTIDAD_DE_COMBATES_GRUPO_INICIAL) + 1)
        combatesArmados = combatesArmados + 1
        'Lo marco como procesado
        equipos(loopEquipo) = 1
        'Marco que procese tambien al contricante para no procesarlo dos veces
        If (equipoConElQueVaAluchar > 0) Then equipos(equipoConElQueVaAluchar) = 1
    End If
Next
End Sub

Private Sub generarFixture()
Dim equipos() As Byte
Dim i As Integer
Dim loopRonda As Byte
Dim pivote As Byte
Dim auxint As Integer

'Si la cantidad de equipos es impar debo
ReDim equipos(1 To cantidadEquipos + cantidadEquipos Mod 2) As Byte

For i = 1 To cantidadEquipos
    equipos(i) = i
Next i

If cantidadEquipos Mod 2 = 1 Then
    equipos(cantidadEquipos + 1) = 0
End If
    
'NO pongo la cantidad de rondas porque si es ida y vuelta voy a tener el doble. Y me alcanza con la primera parte
'Del array en este caso ya que la otra es la misma pero espejada
For loopRonda = 1 To cantidadEquipos + cantidadEquipos Mod 2 - 1
    
    If cantidadEquipos > 2 Then
        If loopRonda Mod 2 = 0 Then
            'Ronda par, vengo de un impar
            'El pivote pasa de la prmera posicion a la ultima. Y todos se corren uno desde la ultima posicion para atras
            pivote = equipos(2)
            
            For i = 1 To UBound(equipos) - 1 Step 2
                
                If i - 2 > 0 Then
                    equipos(i - 2) = equipos(i + 1)
                End If
                
                equipos(i + 1) = equipos(i)
            Next i
            
            equipos(UBound(equipos) - 1) = pivote
        Else
            pivote = equipos(UBound(equipos) - 1)
            
            For auxint = UBound(equipos) - 1 To 2 Step -1
                equipos(auxint) = equipos(auxint - 1)
            Next auxint
            
            equipos(2) = pivote
        End If
    End If
    
   
    For i = 1 To UBound(equipos) Step 2
    
        If equipos(i) > 0 Then
            hojaRuta(equipos(i), loopRonda).nombreRonda = loopRonda & "º Fecha"
            hojaRuta(equipos(i), loopRonda).posTablaContrincante = equipos(i + 1)
        End If
        
        If equipos(i + 1) > 0 Then
            hojaRuta(equipos(i + 1), loopRonda).nombreRonda = loopRonda & "º Fecha"
            hojaRuta(equipos(i + 1), loopRonda).posTablaContrincante = equipos(i)
        End If
             
    Next i
    
Next loopRonda

End Sub

Friend Sub iModoTorneo_setRings(ring() As tRing)
    rings() = ring()
End Sub

Public Sub iModoTorneo_setNecesidadRoundGanadosGanador(cantidad As Byte)
    necesidadRoundGanadosGanador = cantidad
End Sub

' Devuelve la tabla de equipos.
Friend Function iModoTorneo_obtenerTabla() As tEquipoTablaTorneo()
    iModoTorneo_obtenerTabla = tablaEquipos
End Function

Public Sub iModoTorneo_usuarioAbandono(UserIndex As Integer)
    Call LogTorneos(nombreEvento & "Liga -> Abandono " & UserList(UserIndex).Name)
    
    Dim posEquipoTabla As Byte
    
    'Los usuarios son dueños de otro evento?
    'De cual?
    posEquipoTabla = modTorneos.obtenerPosicionEnTablaPersonaje(tablaEquipos, UserList(UserIndex).id)
        
    'El jugador no esta en ningun sub evento, es prioridad mia
    If eventosEquipo(posEquipoTabla) Is Nothing Then
        
        'Obtengo el integrante
        'Lo marco como descalificado
        'Me fijo si tengo que descalificar a todo el equipo
        'Si descalifico al equipo me fijo si queda más de un equipo
        '   Si queda un solo equipo, declaro al ganador
        Dim loopIntegrante As Byte
        
        With tablaEquipos(posEquipoTabla)
        
            For loopIntegrante = 1 To .cantidadIntegrantes
                    
                    If .integrantes(loopIntegrante).id = UserList(UserIndex).id Then
                                
                        .cantidadIntegrantesDescalificados = .cantidadIntegrantesDescalificados + 1
                                
                        .integrantes(loopIntegrante).Estado = eEstadoIntegranteEquipo.Descalificando
                                
                        Call modTorneos.desecharIntegranteEquipo(.integrantes(loopIntegrante), 0)
                            
                        'Le aviso al mundo
                        Call modTorneos.enviarMensajeGlobal(UserList(UserIndex).Name & " abandono el evento.", nombreEvento)
                                
                        'Le aviso al usuario que es un pancho por abandonar
                        EnviarPaquete Paquetes.MensajeTalk, "Abandonaste el evento. Esto se considera una falta de respeto grave para los usuarios participantes y los no participantes. Perdiste la inscripción y podrás ser penado.", UserIndex
                    End If
                            
                
            Next
        
            If Not tieneCantidadMinimaDeParticipantes(tablaEquipos(posEquipoTabla)) Then
               
               Call modTorneos.enviarMensajeEquipos(tablaEquipos, "Equipo de " & modTorneos.obtenerStringEquipo(tablaEquipos(posEquipoTabla), False, eFormatoDisplayEquipo.NombreEquipo) & " descalificado por no contar con la cantidad minima de integrantes.", nombreEvento)
               
               .Estado = eEstadoEquipoTorneo.descalificado
               
               Call modTorneos.desecharEquipo(tablaEquipos(posEquipoTabla), 0)
               
            End If
        
        End With
    Else 'No es de mi propiedad el jugador.
        'Le indico al evento que el jugador cerro
        Call eventosEquipo(posEquipoTabla).iModoTorneo_usuarioAbandono(UserIndex)
        
        '¿Este evento causo que termine el sub evento?
        If eventosEquipo(posEquipoTabla).iModoTorneo_isTerminado Then Call procesarGanador(eventosEquipo(posEquipoTabla))
    End If

End Sub

Public Sub iModoTorneo_usuarioCerro(UserIndex As Integer)

    Dim posEquipoTabla As Byte
    
    Call LogTorneos(nombreEvento & "Liga -> Cerro " & UserList(UserIndex).Name)
    
    'Los usuarios son dueños de otro evento?
    'De cual?
    posEquipoTabla = modTorneos.obtenerPosicionEnTablaPersonaje(tablaEquipos, UserList(UserIndex).id)
        
    'El jugador no esta en ningun sub evento, es prioridad mia
    If eventosEquipo(posEquipoTabla) Is Nothing Then
        'Lo marco como offline
        Call modTorneos.actualizarUserIndexPersonaje(tablaEquipos, UserList(UserIndex).id, 0)
        'Lo agrego a la lista de users offline
        Call modEventos.agregarUsuarioOffline(UserList(UserIndex).id, UserList(UserIndex).evento)
    Else 'No es de mi propiedad el jugador.
        'Le indico al evento que el jugador cerro
        Call eventosEquipo(posEquipoTabla).iModoTorneo_usuarioCerro(UserIndex)
        '¿Este evento causo que termine el sub evento?
        If eventosEquipo(posEquipoTabla).iModoTorneo_isTerminado Then Call procesarGanador(eventosEquipo(posEquipoTabla))
    End If

End Sub

Public Sub iModoTorneo_usuarioIngreso(UserIndex As Integer, UserID As Long)

    Dim posEquipoTabla As Byte

    Call LogTorneos(nombreEvento & "Liga" & "-> Ingreso " & UserList(UserIndex).Name)
    'Los usuarios son dueños de otro evento?
    'De cual?
    posEquipoTabla = modTorneos.obtenerPosicionEnTablaPersonaje(tablaEquipos, UserList(UserIndex).id)
         
    'El jugador no esta en ningun sub evento, es prioridad mia
    If eventosEquipo(posEquipoTabla) Is Nothing Then
        Call modTorneos.actualizarUserIndexPersonaje(tablaEquipos, UserID, UserIndex)
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call modUtilitarios.Preparando(UserIndex)
        End If
    Else
        Call eventosEquipo(posEquipoTabla).iModoTorneo_usuarioIngreso(UserIndex, UserID)
    End If

    Exit Sub
        
End Sub

Public Sub iModoTorneo_usuarioMuere(UserIndex As Integer)
    Dim posEquipoTabla As Byte
    
    posEquipoTabla = modTorneos.obtenerPosicionEnTablaPersonaje(tablaEquipos, UserList(UserIndex).id)
    
    If Not eventosEquipo(posEquipoTabla) Is Nothing Then
    
        Call eventosEquipo(posEquipoTabla).iModoTorneo_usuarioMuere(UserIndex)
        
        If eventosEquipo(posEquipoTabla).iModoTorneo_isTerminado Then
            Call procesarGanador(eventosEquipo(posEquipoTabla))
        End If
          Debug.Print UserList(UserIndex).Name
    End If
End Sub

Private Function tieneCantidadMinimaDeParticipantes(equipo As tEquipoTablaTorneo) As Boolean
    tieneCantidadMinimaDeParticipantes = (equipo.cantidadIntegrantesDescalificados < equipo.cantidadIntegrantes)
End Function

Private Sub iModoTorneo_usuarioRevive(UserIndexRevivido As Integer, UserIndexRevividor As Integer)
    Dim posEquipoTabla As Byte
        
    posEquipoTabla = modTorneos.obtenerPosicionEnTablaPersonaje(tablaEquipos, UserList(UserIndexRevivido).id)
        
    If Not eventosEquipo(posEquipoTabla) Is Nothing Then
        Call eventosEquipo(posEquipoTabla).iModoTorneo_usuarioRevive(UserIndexRevivido, UserIndexRevividor)
    End If
End Sub
