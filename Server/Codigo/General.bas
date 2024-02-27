Attribute VB_Name = "General"
Global LeerNPCs As New clsLeerInis
Global LeerNPCsHostiles As New clsLeerInis

Public clanes As cClanes

Option Explicit

Sub Bloquear(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, map As Integer, ByVal x As Integer, ByVal y As Integer, b As Byte)
    'b=1 bloquea el tile en (x,y)
    'b=0 desbloquea el tile indicado
    EnviarPaquete Paquetes.BloquearTile, ITS(x) & ITS(y) & b, sndIndex, sndRoute, sndMap
End Sub

Sub EnviarSpawnList(ByVal UserIndex As Integer)
Dim k As Integer, SD As String
SD = UBound(SpawnList) & ","

For k = 1 To UBound(SpawnList)
    SD = SD & SpawnList(k).NpcName & ","
Next k

EnviarPaquete Paquetes.pEnviarSpawnList, SD, UserIndex, ToIndex, 0
End Sub

Sub MostrarNumUsers()
    frmMain.CantUsuarios.Caption = "Numero de usuarios jugando: " & NumUsers
End Sub

Sub Restart()
'Se asegura de que los sockets estan cerrados e ignora cualquier err
If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

Dim loopC As Integer

'Cierra el socket de escucha
If SockListen >= 0 Then Call apiclosesocket(SockListen)
'Inicia el socket de escucha
SockListen = ListenForConnect(Puerto, hWndMsg, "")

For loopC = 1 To MaxUsers
    Call CloseSocket(loopC)
Next

ReDim UserList(1 To MaxUsers)

For loopC = 1 To MaxUsers
    UserList(loopC).ConnID = INVALID_SOCKET
    UserList(loopC).InicioConexion = 0
    UserList(loopC).ConfirmacionConexion = 0
Next loopC

LastUser = 0
NumUsers = 0
NumUsersPremium = 0

ReDim NpcList(1 To MAXNPCS) As npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer

Call LoadSini
Call LoadOBJData
Call CargarHechizos

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
'Log it

Call Logs.LogMain("Servidor reiniciado.")

'Ocultar
If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TiempoInvocacion
' DateTime  : 13/02/2007 19:48
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub TiempoInvocacion(ByVal UserIndex As Integer, tiempoTranscurrido As Long)
Dim i As Integer

For i = 1 To MAXMASCOTAS

    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        If NpcList(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
           NpcList(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = _
           NpcList(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia - tiempoTranscurrido
            
            If NpcList(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia <= 0 Then
                Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
            End If
            
        End If
    End If
Next i

End Sub

Public Sub EfectoMimetismo(personaje As User, tiempoTranscurrido As Long)

If personaje.Counters.Mimetismo < IntervaloMimetizado Then
    personaje.Counters.Mimetismo = personaje.Counters.Mimetismo + tiempoTranscurrido
Else
    ' Se me termino el hechizo!!!
    Call modMimetismo.finalizarEfecto(personaje)
End If
            
End Sub

'---------------------------------------------------------------------------------------
' Procedure : EfectoFrio
' DateTime  : 13/02/2007 19:48
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub EfectoFrio(ByVal UserIndex As Integer, tiempo As Long, ByRef EnviarStats As Boolean)

If UserList(UserIndex).flags.Privilegios > 0 Then Exit Sub

If UserList(UserIndex).Counters.Frio < IntervaloFrio Then
  UserList(UserIndex).Counters.Frio = UserList(UserIndex).Counters.Frio + tiempo
Else
   If MapInfo(UserList(UserIndex).pos.map).Terreno = "NIEVE" Then
        If UserList(UserIndex).Invent.ArmourEqpObjIndex <> 665 And UserList(UserIndex).Invent.ArmourEqpObjIndex <> 666 And UserList(UserIndex).Invent.ArmourEqpObjIndex <> 667 Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(10), UserIndex
            If Nevando Then
            UserList(UserIndex).Stats.minHP = UserList(UserIndex).Stats.minHP - 45
            Else
            UserList(UserIndex).Stats.minHP = UserList(UserIndex).Stats.minHP - 30
            End If
        End If
   Else
        If UserList(UserIndex).Invent.ArmourEqpObjIndex = 0 Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(10), UserIndex
        UserList(UserIndex).Stats.minHP = UserList(UserIndex).Stats.minHP - 10
        Else
        UserList(UserIndex).Counters.Frio = 0
        Exit Sub
        End If
   End If
    UserList(UserIndex).Counters.Frio = 0
    
    If UserList(UserIndex).Stats.minHP < 1 Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(22), UserIndex
        UserList(UserIndex).Stats.minHP = 0
        Call UserDie(UserIndex, False)
    Else
        EnviarStats = True
    End If
End If

End Sub


Public Sub EfectoCalor(ByRef personaje As User, tiempo As Long, ByRef EnviarStats As Boolean)

If personaje.flags.Privilegios > 0 Then Exit Sub

If personaje.Counters.Calor < IntervaloCalor Then
    personaje.Counters.Calor = personaje.Counters.Calor + tiempo
    Exit Sub
End If
    
If personaje.Invent.CollarObjIndex = Objetos_Constantes.COLLAR Then
    personaje.Counters.Calor = 0
    Exit Sub
End If
    
Dim ticks As Integer
    
' Calculo cuantas veces tiene que aplicarle
ticks = personaje.Counters.Calor \ IntervaloCalor

' Ajusto el tempororizador
personaje.Counters.Calor = personaje.Counters.Calor - ticks * IntervaloCalor
personaje.Stats.minHP = personaje.Stats.minHP - (30 * ticks)

EnviarPaquete Paquetes.mensajeinfo, "Hace demasiado calor, afecta tu salud.", personaje.UserIndex, ToIndex
   
If personaje.Stats.minHP < 1 Then
    personaje.Stats.minHP = 0
    Call UserDie(personaje.UserIndex, False)
Else
    EnviarStats = True
End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : EfectoInvisibilidad
' DateTime  : 13/02/2007 19:46
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub EfectoInvisibilidad(ByRef personaje As User, TiempoTrancurrido)

If personaje.Counters.Invisibilidad < IntervaloInvisible Then
    personaje.Counters.Invisibilidad = personaje.Counters.Invisibilidad + TiempoTrancurrido
Else
    Call quitarInvisibilidad(personaje)
    EnviarPaquete Paquetes.MensajeSimple, Chr$(23), personaje.UserIndex
End If

End Sub

Public Sub quitarInvisibilidad(personaje As User)
  personaje.Counters.Invisibilidad = 0
  personaje.flags.Invisible = 0
  EnviarPaquete Paquetes.Visible, ITS(personaje.Char.charIndex), personaje.UserIndex, ToMap, personaje.pos.map
End Sub

'---------------------------------------------------------------------------------------
' Procedure : EfectoParalisisNpc
' DateTime  : 13/02/2007 19:46
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub EfectoParalisisNpc(ByVal npcIndex As Integer, tiempo As Long)

'Debug.Print Tiempo
If NpcList(npcIndex).Contadores.Paralisis > 0 Then
    NpcList(npcIndex).Contadores.Paralisis = NpcList(npcIndex).Contadores.Paralisis - tiempo
Else
    NpcList(npcIndex).flags.Paralizado = 0
    NpcList(npcIndex).flags.Inmovilizado = 0
End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : EfectoParalisisUser
' DateTime  : 13/02/2007 19:40
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub EfectoParalisisUser(ByRef personaje As User, ByVal tiempoTranscurrido As Long)

If personaje.Counters.Paralisis > 0 Then
    personaje.Counters.Paralisis = personaje.Counters.Paralisis - tiempoTranscurrido
    
    If personaje.Counters.Paralisis > 1000 And personaje.flags.paralizadoPor > 0 And (personaje.clase = eClases.Guerrero Or personaje.clase = eClases.Cazador) Then
        If UserList(personaje.flags.paralizadoPor).flags.UserLogged Then
            If Not estaEnArea(personaje, UserList(personaje.flags.paralizadoPor)) Then
                personaje.Counters.Paralisis = 1000
            End If
        Else
            personaje.Counters.Paralisis = 1000
        End If
    End If
    
Else
    Call quitarParalisis(personaje)
End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : RecStamina
' DateTime  : 13/02/2007 19:49
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub RecStamina(personaje As User, ByVal TipoRecuperacion As Byte, ByRef EnviarStats As Boolean)

Dim massta As Integer


If TipoRecuperacion < 5 Then
    ' Aumenta la Energia
    
    ' ¿Ya está a full?
    If personaje.Stats.MinSta = personaje.Stats.MaxSta Then Exit Sub
    
    ' Si tiene hambre y sed no puede recuperar energia
    If personaje.Stats.minham = 0 Or personaje.Stats.minAgu = 0 Then Exit Sub
    
    Select Case TipoRecuperacion
        Case 1 'El personaje se encuentra descansando
        massta = CInt(RandomNumber(1, Porcentaje(personaje.Stats.MaxSta, 5))) * 3
        Case 2 'Se encuentra vestido pero sin el comando /Descansar
        massta = CInt(RandomNumber(1, Porcentaje(personaje.Stats.MaxSta, 5))) * 1.1
        Case 3 'Se encuentra lloviendo pero descansado
        massta = CInt(RandomNumber(Porcentaje(personaje.Stats.MaxSta, 2), Porcentaje(personaje.Stats.MaxSta, 5)))
        Case 4 ' Se encuentra vestido pero esta lloviendo y sin descansar
        massta = CInt(Porcentaje(personaje.Stats.MaxSta, 2))
    End Select
    
    ' Aumento
    personaje.Stats.MinSta = personaje.Stats.MinSta + massta
    
    ' ¿Supere el máximo?
    If personaje.Stats.MinSta > personaje.Stats.MaxSta Then personaje.Stats.MinSta = personaje.Stats.MaxSta
    
Else
    ' Resta energia
    
    ' Si no tiene, cancelamos
    If personaje.Stats.MinSta = 0 Then Exit Sub
    
    ' Calculamos cuanto
    massta = CInt(RandomNumber(Porcentaje(personaje.Stats.MaxSta, 1), Porcentaje(personaje.Stats.MaxSta, 1) + 1))
    
    ' Restamos
    If personaje.Stats.MinSta - massta > 0 Then
        personaje.Stats.MinSta = personaje.Stats.MinSta - massta
    Else
        personaje.Stats.MinSta = 0
    End If

End If

EnviarStats = True

End Sub


'---------------------------------------------------------------------------------------
' Procedure : EfectoVeneno
' DateTime  : 13/02/2007 19:44
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub EfectoVeneno(ByRef personaje As User, EnviarStats As Boolean, tiempoTranscurrido)

Dim intervaloVeneno As Long

' Ticket #59
If personaje.Stats.ELV <= 20 Then
    intervaloVeneno = 60000
Else
    intervaloVeneno = 10000
End If

If personaje.Counters.Veneno < intervaloVeneno Then
    personaje.Counters.Veneno = personaje.Counters.Veneno + tiempoTranscurrido
Else
    personaje.Counters.Veneno = 0
        
    personaje.Stats.minHP = personaje.Stats.minHP - (personaje.Stats.MaxHP * 0.1)
  
    If personaje.Stats.minHP < 1 Then
        Call UserDie(personaje.UserIndex, False)
        EnviarPaquete Paquetes.mensajeinfo, "Has muerto por envenenamiento.", personaje.UserIndex, ToIndex
    End If
  
  EnviarStats = True
End If

End Sub

Public Sub DuracionPociones(UserIndex As Integer, tiempoTranscurrido As Long)
'Controla la duracion de las pociones
If UserList(UserIndex).flags.DuracionEfecto > 0 Then

    UserList(UserIndex).flags.DuracionEfecto = UserList(UserIndex).flags.DuracionEfecto - tiempoTranscurrido
    
    If UserList(UserIndex).flags.DuracionEfecto <= 5000 And UserList(UserIndex).flags.ShowDopa = False Then
        EnviarPaquete Paquetes.EnviarFA, LongToString(UserList(UserIndex).flags.DuracionEfecto) & ITS(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad)) & ITS(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza)), UserIndex, ToIndex
        UserList(UserIndex).flags.ShowDopa = True
    End If
    
    If UserList(UserIndex).flags.DuracionEfecto <= 0 Then
        'volvemos los atributos al estado normal
        Dim loopX As Integer
        For loopX = 1 To NUMATRIBUTOS
              UserList(UserIndex).Stats.UserAtributos(loopX) = UserList(UserIndex).Stats.UserAtributosBackUP(loopX)
        Next
        UserList(UserIndex).flags.ShowDopa = False
   End If
End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HambreYSed
' DateTime  : 13/02/2007 19:49
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub HambreYSed(UserIndex As Integer, fenviarAyS As Boolean, tiempo As Long)
If UserList(UserIndex).flags.Privilegios >= 2 Then Exit Sub

If UserList(UserIndex).Stats.minAgu > 0 Then
    If UserList(UserIndex).Counters.AGUACounter < IntervaloSed Then
          UserList(UserIndex).Counters.AGUACounter = UserList(UserIndex).Counters.AGUACounter + tiempo
    Else
          UserList(UserIndex).Counters.AGUACounter = 0
          UserList(UserIndex).Stats.minAgu = UserList(UserIndex).Stats.minAgu - 10
                            
          If UserList(UserIndex).Stats.minAgu <= 0 Then
               UserList(UserIndex).Stats.minAgu = 0
               UserList(UserIndex).flags.Sed = 1
          End If
          fenviarAyS = True
                            
    End If
End If
'hambre
If UserList(UserIndex).Stats.minham > 0 Then
   If UserList(UserIndex).Counters.COMCounter < IntervaloHambre Then
        UserList(UserIndex).Counters.COMCounter = UserList(UserIndex).Counters.COMCounter + tiempo
   Else
        UserList(UserIndex).Counters.COMCounter = 0
        UserList(UserIndex).Stats.minham = UserList(UserIndex).Stats.minham - 10
        If UserList(UserIndex).Stats.minham < 0 Then
               UserList(UserIndex).Stats.minham = 0
               UserList(UserIndex).flags.Hambre = 1
        End If
        fenviarAyS = True
    End If
End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Sanar
' DateTime  : 13/02/2007 19:49
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Sanar(UserIndex As Integer, EnviarStats As Boolean, intervalo As Long)

If MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).Trigger = 1 And _
   MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).Trigger = 2 And _
   MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).Trigger = 4 Then Exit Sub

Dim mashit As Integer
'con el paso del tiempo va sanando....pero muy lentamente ;-)
If UserList(UserIndex).Stats.minHP < UserList(UserIndex).Stats.MaxHP Then
   If UserList(UserIndex).Counters.HPCounter < intervalo Then
      UserList(UserIndex).Counters.HPCounter = UserList(UserIndex).Counters.HPCounter + 1
   Else
      mashit = CInt(RandomNumber(2, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5)))
                           
      UserList(UserIndex).Counters.HPCounter = 0
      UserList(UserIndex).Stats.minHP = UserList(UserIndex).Stats.minHP + mashit
      If UserList(UserIndex).Stats.minHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.minHP = UserList(UserIndex).Stats.MaxHP
         EnviarPaquete Paquetes.MensajeSimple, Chr$(40), UserIndex
          EnviarStats = True
      End If
End If

End Sub

Public Sub CargaNpcsDat()

Dim npcfile As String

npcfile = DatPath & "NPCs.dat"
LeerNPCs.Abrir npcfile
npcfile = DatPath & "NPCs-HOSTILES.dat"
LeerNPCsHostiles.Abrir npcfile

End Sub

'---------------------------------------------------------------------------------------
' Procedure : PasarSegundo
' DateTime  : 18/02/2007 21:26
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub PasarSegundo()

Dim UserIndex As Integer
Dim tiempo As Long
   
tiempo = GetTickCount
    
For UserIndex = 1 To LastUser
    
    ' Usuarios Logueados
    If UserList(UserIndex).flags.UserLogged Then
        ' Cuenta Regresiva de Retos
        If UserList(UserIndex).Counters.combateRegresiva > 0 Then
            UserList(UserIndex).Counters.combateRegresiva = UserList(UserIndex).Counters.combateRegresiva - 1
            EnviarPaquete Paquetes.TiempoReto, ByteToString(UserList(UserIndex).Counters.combateRegresiva), UserIndex, ToIndex
        End If
            
        'Cerrar usuario
        If Not UserList(UserIndex).flags.Saliendo = eTipoSalida.NoSaliendo Then
            
            If UserList(UserIndex).Counters.Salir <= tiempo Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(41), UserIndex
                If Not CloseSocket(UserIndex) Then LogError ("Pasar Segundo Cerrar User")
            End If
                
        End If
    End If
        
Next UserIndex

Call modResucitar.procesarResucitacionesPendientes(tiempo)

End Sub

Public Sub GuardarTodosLosUsuarios(Optional GuardarOnline As Boolean = False)
    Dim i As Integer

    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            Call SaveUser(i, GuardarOnline)
        End If
    Next i
End Sub

'Guardar online establece si el saveuser setea el Online en 0 (offline) o en 1 (online)
Sub GuardarUsuarios(Optional GuardarOnline As Boolean = False)
    
    haciendoBK = True
    EnviarPaquete Paquetes.Pausa, "", 0, ToAll
    EnviarPaquete Paquetes.MensajeSimple, Chr$(43), 0, ToAll
    
    Call GuardarTodosLosUsuarios(GuardarOnline)
    
    EnviarPaquete Paquetes.MensajeSimple, Chr$(44), 0, ToAll
    EnviarPaquete Paquetes.Pausa, "", 0, ToAll
    haciendoBK = False
End Sub
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
    'WorldSave
    Call DoBackUp
    ' experiencias
    Call mdParty.ActualizaExperiencias
    'Guardar Pjs
    Call GuardarUsuarios
    'Guilds
    'Call SaveGuildsDB
    'Chauuu
    Unload frmMain
End Sub

Public Sub cargarAtributosPersonajeOffline(nombrePersonaje As String, ByRef infoPersonaje As ADODB.Recordset, atributos As String, paraActualizar As Boolean)

If paraActualizar Then
    'Necesito obtener el ID para poder actualizarlo luego utilizando la clave
    sql = "SELECT ID," & atributos & " FROM " & DB_NAME_PRINCIPAL & ".usuarios WHERE NickB='" & nombrePersonaje & "'"
    
    'Tengo que abrirlo utilizando un nuevo recordset
    Set infoPersonaje = New ADODB.Recordset
    infoPersonaje.CursorLocation = adUseClient
    infoPersonaje.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
Else
    sql = "SELECT " & atributos & " FROM " & DB_NAME_PRINCIPAL & ".usuarios WHERE NickB='" & nombrePersonaje & "'"
    Set infoPersonaje = conn.Execute(sql, , adCmdText)
End If

End Sub

Public Sub iniciarEstructuras()
    Set GmsGroup = New EstructurasLib.ColaConBloques
    Set TrabajadoresGroup = New EstructurasLib.ColaConBloques
    Set Ayuda = New EstructurasLib.ColaConBloques

    ReDim NpcList(1 To MAXNPCS) As npc 'NPCS
    ReDim CharList(1 To MAXCHARS) As Integer
    
    
    ReDim Centinelas(0) 'EL YIND :)
    
    Call mdParty.iniciar
End Sub
