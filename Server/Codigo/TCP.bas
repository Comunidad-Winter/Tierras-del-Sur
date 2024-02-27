Attribute VB_Name = "TCP"
Option Explicit

'RUTAS DE ENVIO DE DATOS
Public Const ToIndex = 0 'Envia a un solo User
Public Const ToAll = 1 'A todos los Users
Public Const ToMap = 2 'Todos los Usuarios en el mapa
Public Const ToPCArea = 3 'Todos los Users en el area de un user determinado
Public Const ToNone = 4 'Ninguno
Public Const ToAllButIndex = 5 'Todos menos el index
Public Const ToMapButIndex = 6 'Todos en el mapa menos el indice
'Public Const ToGM = 7 DESUSO
Public Const ToNPCArea = 8 'Todos los Users en el area de un user determinado
Public Const ToGuildMembers = 9
Public Const ToAdmins = 10
Public Const ToPCAreaButIndex = 11
Public Const ToAdminsArea = 12
'Public Const ToDiosesYclan = 13 DESUSO
Public Const ToConsejo = 14
Public Const ToConsejoCaos = 15
Public Const ToDeadArea = 16
'Public Const ToDeadAreaButIndex = 17 DESUSO
Public Const ToAreaButIndex = 18 'Sistema de areas moviendo de personas
Public Const ToAreaNPC = 19 'Sistema de areas moviendo de npcs
Public Const ToArea = 20
'Autoconteo
Public Conteo As Byte



Function isNombreValido(nombre As String) As Boolean

If AsciiValidos(nombre) = False Then
    isNombreValido = False
    Exit Function
End If

If DobleEspacios(nombre) = True Then
    isNombreValido = False
    Exit Function
End If

isNombreValido = True

End Function
Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer
'Solo son validos letras y espacios
cad = LCase$(cad)
For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    If (car < 97 Or car > 122) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
Next i
AsciiValidos = True
End Function

Function ValidateSkills(ByVal UserIndex As Integer) As Boolean
Dim loopC As Integer
For loopC = 1 To NUMSKILLS
    If UserList(UserIndex).Stats.UserSkills(loopC) < 0 Then
        Exit Function
        If UserList(UserIndex).Stats.UserSkills(loopC) > 100 Then UserList(UserIndex).Stats.UserSkills(loopC) = 100
    End If
Next loopC
ValidateSkills = True
End Function

'CSEH: Nada
Public Function CloseSocket(ByVal UserIndex As Integer) As Boolean
    On Error GoTo errhandler
    Dim resultado As Boolean
    Dim rastreandoFecha As Boolean
    
    
    rastreandoFecha = False
    
    ' ¿Hay un personaje cargado en este UserIndex?
    If UserList(UserIndex).flags.UserLogged Then
        ' Actualizamos la cantidad de Onlines comunes y premium
        If NumUsers > 0 Then NumUsers = NumUsers - 1
        If UserList(UserIndex).Premium Then NumUsersPremium = NumUsersPremium - 1
        ' Cerramos el Personaje
        Call CloseUser(UserIndex)
    End If
    
    'Cierro el socket
    Call CloseSocketSL(UserList(UserIndex))
        
    'Limpio el slots y el personaje
    Call ResetUserSlot(UserIndex)
    
    ' Liberamos este Slot
    resultado = liberarUserIndex(UserIndex)
    
    If rastreandoFecha = True Then
        CloseSocket = False
    Else
        CloseSocket = resultado
    End If
Exit Function
errhandler:
    ' Logueo ya
    LogError ("Error en Close Socket" & Err.Description)
    
    ' Nos seguramos de resetear si o si el slot
    Call ResetUserSlot(UserIndex)

    ' Cerramos el Socket?
    If Not UserList(UserIndex).ConnID = INVALID_SOCKET Then Call CloseSocketSL(UserList(UserIndex))
    
    CloseSocket = False
End Function

Public Sub CloseSocketSL(ByRef personaje As User)

Debug.Print "CloseSocketSL"

' Nos aseguramos de estar cerrando un socket valido
If Not personaje.ConnID = INVALID_SOCKET Then

    ' Remuevo la relacion entre el SocketID y el UserIndex
    Call BorraSlotSock(personaje.ConnID)
        
    ' Cierro el Socket
    Call WSApiCloseSocket(personaje.ConnID)

    ' Marco como nulo al Socket
    personaje.ConnID = INVALID_SOCKET
End If
End Sub


Public Function EnviarDatosASlot(ByVal UserIndex As Integer, ByVal datos As String) As Long

Dim Ret As Long

' Validaciones
If UserIndex = 0 Then Exit Function ' Por las dudas
If UserList(UserIndex).ConnID = INVALID_SOCKET Then Exit Function ' ¿Esta relacionado a un PJ?

' Estadisticas
TCPESStats.BitesEnviadosMinuto = TCPESStats.BitesEnviadosMinuto + LenB(datos) * 8
TCPESStats.PaquetesEnviadosMinuto = TCPESStats.PaquetesEnviadosMinuto + 1

' Encriptamos
datos = CryptStr(datos, UserList(UserIndex).CryptOffset)

' Enviamos
Ret = WsApiEnviar(UserIndex, datos)

' Verificamos el exito. Si fallo es porque se perdio la conexion y no nos dimos cuenta hasta ahora
' POrque sino no hubiesemos llegado hasta acá ya que .ConnID seria invalido
If Ret <> 0 Then Call CierreForzadoPorDesconexion(UserIndex)

EnviarDatosASlot = Ret

End Function
Private Function generarHeader(ByRef longData As Integer) As String

'255 -> chr$(254)
'256 -> chr$(255) & its (0)
'257 -> chr$(255) & its (1)
'No puede haber paquetes con longitud 0
If longData > 255 Then
    generarHeader = Chr$(255) & ITS(longData - 256)
Else
    generarHeader = Chr$(longData - 1)
End If

End Function
'---------------------------------------------------------------------------------------
' Procedure : Senddata
' DateTime  : 19/02/2007 19:19
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub Senddata(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal sndData As String)
  
Dim loopC As Integer
Dim map As Integer
Dim UserIndex As Integer
Dim listaUsuarios As EstructurasLib.ColaConBloques

sndData = generarHeader(Len(sndData)) & sndData

Select Case sndRoute
    Case ToNone
        Exit Sub
    Case ToAdmins
        GmsGroup.itIniciar
        
        Do While GmsGroup.ithasNext
            UserIndex = GmsGroup.itnext
            Call EnviarDatosASlot(UserIndex, sndData)
        Loop
        
        Exit Sub
    Case ToAll
        For loopC = 1 To LastUser
            If UserList(loopC).flags.UserLogged Then 'Esta logeado como usuario?
                Call EnviarDatosASlot(loopC, sndData)
            End If
        Next loopC
        Exit Sub
    Case ToAllButIndex
        LogError ("Llamada a ToAllButIndex")
        Exit Sub
    Case ToMap
        
        If sndMap = 0 And sndIndex > 0 Then
            map = UserList(sndIndex).pos.map
        Else
            map = sndMap
        End If

        Set listaUsuarios = MapInfo(map).usuarios
        listaUsuarios.itIniciar
        
        Do While (listaUsuarios.ithasNext)
            Call EnviarDatosASlot(listaUsuarios.itnext, sndData)
        Loop

        Exit Sub
        
    Case ToMapButIndex
        If sndMap = 0 And sndIndex > 0 Then
            map = UserList(sndIndex).pos.map
        Else
            map = sndMap
        End If

        Set listaUsuarios = MapInfo(map).usuarios
        listaUsuarios.itIniciar
        
        Do While (listaUsuarios.ithasNext)
            UserIndex = listaUsuarios.itnext
            If UserIndex <> sndIndex Then
                Call EnviarDatosASlot(UserIndex, sndData)
            End If
        Loop
        
        Exit Sub
    Case ToGuildMembers
    
        Set listaUsuarios = UserList(sndIndex).ClanRef.getIntegrantesOnline
        listaUsuarios.itIniciar
        
        Do While (listaUsuarios.ithasNext)
            Call EnviarDatosASlot(listaUsuarios.itnext, sndData)
        Loop
        
        Exit Sub
    Case ToPCArea
    
        If sndMap = 0 And sndIndex > 0 Then
            map = UserList(sndIndex).pos.map
        Else
            map = sndMap
        End If

        Set listaUsuarios = MapInfo(map).usuarios
        listaUsuarios.itIniciar
        
        Do While listaUsuarios.ithasNext
            UserIndex = listaUsuarios.itnext
            If Distance(UserList(sndIndex).pos.x, UserList(sndIndex).pos.y, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y) <= Max_Distance Then
                Call EnviarDatosASlot(UserIndex, sndData)
            End If
        Loop
        
        Exit Sub
    Case ToDeadArea
    
        If sndMap = 0 And sndIndex > 0 Then
            map = UserList(sndIndex).pos.map
        Else
            map = sndMap
        End If

        Set listaUsuarios = MapInfo(map).usuarios
        listaUsuarios.itIniciar
        
        Do While listaUsuarios.ithasNext
            UserIndex = listaUsuarios.itnext
            If Distance(UserList(sndIndex).pos.x, UserList(sndIndex).pos.y, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y) <= Max_Distance And UserList(UserIndex).flags.Muerto Then
                Call EnviarDatosASlot(UserIndex, sndData)
            End If
        Loop
        Exit Sub
        
    Case ToPCAreaButIndex
        LogError ("Se llama a ToPCAreaButIndex")
        Exit Sub
    Case ToAdminsArea
        GmsGroup.itIniciar
        
        If sndMap = 0 And sndIndex > 0 Then
            map = UserList(sndIndex).pos.map
        Else
            map = sndMap
        End If
        
        Do While GmsGroup.ithasNext
            UserIndex = GmsGroup.itnext
            If UserList(sndIndex).pos.map = UserList(UserIndex).pos.map Then
                If Distance(UserList(sndIndex).pos.x, UserList(sndIndex).pos.y, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y) <= Max_Distance And UserList(UserIndex).flags.Privilegios > 0 Then
                    Call EnviarDatosASlot(UserIndex, sndData)
                End If
            End If
        Loop
        'Otra opcion recorriendo los usuarios que hay en el mapa.
        'Se toma la primera porque supongo que hay menos gms que personajes en los mapas
        
    Case ToNPCArea
    
        map = NpcList(sndIndex).pos.map
        
        Set listaUsuarios = MapInfo(map).usuarios
                
        listaUsuarios.itIniciar
        
        Do While listaUsuarios.ithasNext
            UserIndex = listaUsuarios.itnext
            If Distance(NpcList(sndIndex).pos.x, NpcList(sndIndex).pos.y, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y) <= Max_Distance Then
                Call EnviarDatosASlot(UserIndex, sndData)
            End If
        Loop
        
        Exit Sub
        
    Case ToIndex
    
        Call EnviarDatosASlot(sndIndex, sndData)
        
        Exit Sub
        
    Case ToConsejo
        For loopC = 1 To LastUser
            If UserList(loopC).flags.PertAlCons > 0 Then
                Call EnviarDatosASlot(loopC, sndData)
            End If
        Next loopC
        Exit Sub
    Case ToConsejoCaos
        For loopC = 1 To LastUser
            If UserList(loopC).flags.PertAlConsCaos > 0 Then
                Call EnviarDatosASlot(loopC, sndData)
            End If
        Next loopC
        Exit Sub
    Case ToAreaButIndex 'Para el nuevo sistema de areas
        If sndMap = 0 And sndIndex > 0 Then
            map = UserList(sndIndex).pos.map
        Else
            map = sndMap
        End If
        
        Set listaUsuarios = MapInfo(map).usuarios
        listaUsuarios.itIniciar
        
        Do While listaUsuarios.ithasNext
            UserIndex = listaUsuarios.itnext
            If UserIndex <> sndIndex Then
                If Abs(UserList(sndIndex).pos.x - UserList(UserIndex).pos.x) <= 12 And Abs(UserList(sndIndex).pos.y - UserList(UserIndex).pos.y) <= 12 Then
                    Call EnviarDatosASlot(UserIndex, sndData)
                End If
            End If
        Loop
        Exit Sub
    Case ToAreaNPC
        map = NpcList(sndIndex).pos.map
        
        Set listaUsuarios = MapInfo(map).usuarios
        listaUsuarios.itIniciar

        Do While listaUsuarios.ithasNext
            UserIndex = listaUsuarios.itnext
            If Abs(NpcList(sndIndex).pos.x - UserList(UserIndex).pos.x) <= 12 And Abs(NpcList(sndIndex).pos.y - UserList(UserIndex).pos.y) <= 12 Then
                Call EnviarDatosASlot(UserIndex, sndData)
            End If
        Loop
        
        Exit Sub
    Case ToArea
        If sndMap = 0 And sndIndex > 0 Then
            map = UserList(sndIndex).pos.map
        Else
            map = sndMap
        End If
        
        Set listaUsuarios = MapInfo(map).usuarios
        listaUsuarios.itIniciar
        
        Do While listaUsuarios.ithasNext
            UserIndex = listaUsuarios.itnext
            If Abs(UserList(sndIndex).pos.x - UserList(UserIndex).pos.x) <= 12 And Abs(UserList(sndIndex).pos.y - UserList(UserIndex).pos.y) <= 12 Then
                Call EnviarDatosASlot(UserIndex, sndData)
            End If
        Loop
        
        Exit Sub
   End Select

End Sub

Function EstaPCarea(index As Integer, Index2 As Integer) As Boolean
Dim x As Integer, y As Integer

For y = UserList(index).pos.y - BORDE_TILES_INUTILIZABLE + 1 To UserList(index).pos.y + BORDE_TILES_INUTILIZABLE - 1
        For x = UserList(index).pos.x - BORDE_TILES_INUTILIZABLE + 1 To UserList(index).pos.x + BORDE_TILES_INUTILIZABLE - 1
            If MapData(UserList(index).pos.map, x, y).UserIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        Next x
Next y
EstaPCarea = False
End Function

Function HayPCarea(pos As WorldPos) As Boolean
Dim x As Integer, y As Integer
For y = pos.y - RangoY + 1 To pos.y + RangoY - 1
        For x = pos.x - RangoX + 1 To pos.x + RangoX - 1
            If x >= SV_Constantes.X_MINIMO_JUGABLE And y >= SV_Constantes.Y_MAXIMO_JUGABLE And x <= SV_Constantes.X_MAXIMO_JUGABLE And y <= SV_Constantes.X_MAXIMO_JUGABLE Then
                If MapData(pos.map, x, y).UserIndex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next x
Next y
HayPCarea = False
End Function

Function HayOBJarea(pos As WorldPos, ObjIndex As Integer) As Boolean
Dim x As Integer, y As Integer
For y = pos.y - RangoY + 1 To pos.y + RangoY - 1
        For x = pos.x - RangoX + 1 To pos.x + RangoX - 1
            If MapData(pos.map, x, y).OBJInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function
            End If
        
        Next x
Next y
HayOBJarea = False
End Function

Function ValidateChr(ByVal UserIndex As Integer) As Boolean
If UserList(UserIndex).Char.Body = 0 Then
    Call DarCuerpoDesnudo(UserList(UserIndex))
End If
ValidateChr = UserList(UserIndex).Char.Body <> 0 And _
(UserList(UserIndex).Char.Head <> 0 Or (UserList(UserIndex).Char.Head = 0 And UserList(UserIndex).flags.Navegando = 1)) And ValidateSkills(UserIndex)
End Function

Sub ConnectUser(ByVal UserIndex As Integer, idPersonaje As Long, Password As String)
Dim tempIndex As Integer

' ¿Hay lugar para este usuario?
If NumUsers >= MaxUsers Then
    EnviarPaquete mbox, Chr$(2), UserIndex
    If Not CloseSocket(UserIndex) Then LogError ("Connect user 1")
    Exit Sub
End If

'¿Ya esta conectado el personaje?
tempIndex = personajeYaEstaLogueado(idPersonaje)

If tempIndex > 0 Then
    If Not UserList(tempIndex).flags.Saliendo = eTipoSalida.NoSaliendo Then
        ' Le avisamos que el personaje está cerrando
        EnviarPaquete mbox, Chr$(4), UserIndex
    Else
        ' El Personaje está Online
        EnviarPaquete mbox, Chr$(5), UserIndex
    End If
    
    If Not CloseSocket(UserIndex) Then LogError ("Connect user 2")
    Exit Sub
End If

#If testeo = 0 Then
    'El usuario ya esta con otro personaje?
    tempIndex = usuarioYaConectado(UserList(UserIndex).MacAddress)
    
    If tempIndex > 0 Then
    
        If UserList(tempIndex).flags.Saliendo = NoSaliendo Then
            EnviarPaquete mbox, Chr$(3), UserIndex
        Else
            EnviarPaquete mbox, Chr$(14) & "Un personaje tuyo se encuentra saliendo del juego. Por favor, aguarda.", UserIndex
        End If
        
        If Not CloseSocket(UserIndex) Then Call LogError("Connect User 3")
        Exit Sub
    End If
#End If

'Solicitamos la informacion a la base de datos el personaje
If modMySql.solicitarInfoPersonaje(UserIndex, idPersonaje, Password) = False Then
    EnviarPaquete mbox, Chr$(14) & "En estos momentos no es posible ingresar al juego. Por favor intenta en 1 minuto.", UserIndex, ToIndex
    If Not CloseSocket(UserIndex) Then Call LogError("connect user 4")
    Exit Sub
End If

End Sub

Public Function conectarPersonaje(infoPersonaje As Recordset, UserIndex As Integer) As Boolean

Dim desban As Boolean
Dim tempdate As Date
Dim esPremium As Boolean
Dim N As Integer
'Inicialmente es falso.
desban = False

conectarPersonaje = False

'¿El personaje esta online?
If infoPersonaje!online = 1 Then
    'Liberamos la memoria
    infoPersonaje.Close
    Set infoPersonaje = Nothing
    '\
    EnviarPaquete mbox, Chr$(5), UserIndex
    Exit Function
End If

'¿El personaje tiene el mismo nombre e ID con el que se autentifico el usuario en el login?
If StrComp(infoPersonaje!nickb, UserList(UserIndex).Name, vbTextCompare) Then
    'TODO Manejar error
    'Liberamos la memoria
    infoPersonaje.Close
    Set infoPersonaje = Nothing
    '\
    Exit Function
End If

'¿El perosnaje esta baneado?
If infoPersonaje!banb = 1 Then
'Veamos si le corresponde unban
    If infoPersonaje!Unban = "NUNCA" Or infoPersonaje!Unban = "" Then 'Cagamos no se lo desbaneamos
        EnviarPaquete mbox, Chr$(7) & infoPersonaje!banrazb, UserIndex
        'Liberamos la memoria
        infoPersonaje.Close
        Set infoPersonaje = Nothing
        '\
        Exit Function
    Else
        tempdate = infoPersonaje!Unban
        If Date >= tempdate Then 'Groso pj unbaneado!
            'para que no quede en el confesionario
            desban = True
        Else
            EnviarPaquete mbox, Chr$(7) & infoPersonaje!banrazb, UserIndex
            'Liberamos la memoria
            infoPersonaje.Close
            Set infoPersonaje = Nothing
            '\
            Exit Function
        End If
    End If
End If

'¿Tiene una cuenta?
If infoPersonaje!IDCuenta > 0 Then

    'Esta bloqueada?
    If infoPersonaje!bloqueada = "SI" Then
        EnviarPaquete mbox, Chr$(14) & "Tu cuenta se encuentra BLOQUEADA. Para más información ingresá a tu Cuenta.", UserIndex, ToIndex
        'Liberamos la memoria
        infoPersonaje.Close
        Set infoPersonaje = Nothing
        '\
        Exit Function
    End If

    'Esta activada?
    If Not infoPersonaje!Estado = "ACTIVADA" Then
        EnviarPaquete mbox, Chr$(14) & "Antes de ingresar al juego debés activar tu Cuenta.", UserIndex, ToIndex
        'Liberamos la memoria
        infoPersonaje.Close
        Set infoPersonaje = Nothing
        '\
        Exit Function
    End If
    
Else
    EnviarPaquete mbox, Chr$(14) & "Hay un problema on tu personaje. Por favor, envia soporte.", UserIndex, ToIndex
    'Liberamos la memoria
    infoPersonaje.Close
    Set infoPersonaje = Nothing
    '\
    Exit Function
End If

If infoPersonaje!esPremium = "SI" Then
    esPremium = True
Else
    esPremium = False
End If

#If TDSFacil And testeo = 0 Then

    'If Not esPremium Then
    '    If infoPersonaje!SEGUNDOS_TDSF <= 0 Then
            'Liberamos la memoria
    '        infoPersonaje.Close
    '        Set infoPersonaje = Nothing
            '\
    '        EnviarPaquete mbox, Chr$(18), UserIndex
    '        Exit Function
    '    End If
    'End If
    
#End If

' El personaje esta en modo candado? MercadoAo
If infoPersonaje!ModoCandado = 1 Then
    'Liberamos la memoria
    infoPersonaje.Close
    Set infoPersonaje = Nothing
    '\
    EnviarPaquete mbox, Chr$(10), UserIndex
    Exit Function
End If

' El personaje esta bloqueado? Beneficio cuenta premium
If infoPersonaje!bloqueado = 1 Then
    EnviarPaquete mbox, Chr$(20), UserIndex
    'Liberamos la memoria
    infoPersonaje.Close
    Set infoPersonaje = Nothing
    '\
    Exit Function
End If

#If testeo = 0 And TDSFacil = 1 Then 'Máximo de personajes por cuenta

    Dim tempbyte As Byte
    tempbyte = 0
    
    For N = 1 To LastUser
        If UserList(N).flags.UserLogged = True And UserIndex <> N Then
            If UserList(N).IDCuenta = infoPersonaje!IDCuenta Then
                If tempbyte = 0 And esPremium Then
                    tempbyte = 1
                Else
                     'Liberamos la memoria
                    infoPersonaje.Close
                    Set infoPersonaje = Nothing
                    
                    '\
                    If tempbyte = 0 Then
                        EnviarPaquete mbox, Chr$(3), UserIndex
                    Else
                        EnviarPaquete mbox, Chr$(21), UserIndex
                    End If
                    Exit Function
                End If
            End If
        End If
    Next N

#End If 'Testeo

With UserList(UserIndex)
    .id = infoPersonaje!id
    .Password = infoPersonaje!passwordb
    
    .Premium = esPremium
    .IDCuenta = infoPersonaje!IDCuenta
    
    '#If TDSFacil = 1 Then
    '    .segundosPremium = infoPersonaje!SEGUNDOS_TDSF
    '#End If
    
    'Reseteamos los FLAGS
    .flags.TargetNPC = 0
    .flags.TargetNpcTipo = 0
    .flags.TargetObj = 0
    .flags.TargetUser = 0
    .Char.FX = 0
    
    'Cargamos los datos del personaje
    If LoadUserInit(UserIndex, infoPersonaje) = False Then
        'Liberamos la memoria
        infoPersonaje.Close
        Set infoPersonaje = Nothing
        '\
        EnviarPaquete mBox2, "Error al cargar el personaje. Consulte a un administrador.", UserIndex
        Exit Function
    End If
    
    If LoadUserStats(UserIndex, infoPersonaje) = False Then
        'Liberamos la memoria
        infoPersonaje.Close
        Set infoPersonaje = Nothing
        '\
        EnviarPaquete mBox2, "Error al cargar el personaje. Consulte a un administrador.", UserIndex
        Exit Function
    End If
    
    If LoadUserReputacion(UserIndex, infoPersonaje) = False Then
        'Liberamos la memoria
        infoPersonaje.Close
        Set infoPersonaje = Nothing
        '\
        EnviarPaquete mBox2, "Error al cargar el personaje. Consulte a un administrador.", UserIndex
        Exit Function
    End If
'**************************************************************************************************
'                  TODOS LOS DATOS DEL PERSONAJE YA ESTAN CARGADOS EN MEMORIA
'**************************************************************************************************
    'Liberamos la memoria
    infoPersonaje.Close
    Set infoPersonaje = Nothing
    '\
    
    'El persnaje esta ok?
    If Not ValidateChr(UserIndex) Then
        EnviarPaquete mbox, Chr$(6), UserIndex
        Exit Function
    End If

    'Posicion de comienzo
    If .pos.map = 0 Then
        .pos = GetCiudad(UserList(UserIndex))
    End If

    If Not SV_PosicionesValidas.existeMapa(.pos.map) Then
        EnviarPaquete mbox, Chr$(14) & "EL PJ se encuenta en un mapa invalido.", UserIndex
        Exit Function
    End If
   
    If haciendoBK Then
        EnviarPaquete Pausa, "", UserIndex
        EnviarPaquete MensajeSimple, Chr$(52), UserIndex
        Exit Function
    End If
    
    If EnPausa Then
        EnviarPaquete Pausa, "", UserIndex
        EnviarPaquete MensajeSimple, Chr(53), UserIndex
        Exit Function
    End If
    
    If EnTesting And .Stats.ELV < 40 Then
        EnviarPaquete mbox, Chr$(14) & "Servidor restringido. Sólo pueden ingresar personajes nivel 40 o más.", UserIndex
        Exit Function
    End If
    
    'Controlamos si el server esta registringido a gms
    If ServerSoloGMs = 1 And .flags.Privilegios = 0 Then
        EnviarPaquete mbox, Chr$(11), UserIndex
        Exit Function
    End If
    
    'Si ya hay alguien en esa posicicion no puede ingresar
    If MapData(.pos.map, .pos.x, .pos.y).UserIndex <> 0 Then
        If Not ObtenerPosicionMasCercana(UserList(UserIndex)) Then
            EnviarPaquete mbox, Chr$(14) & "Ya hay un personaje en tu posición. Intenta ingresar nuevamente en unos instantes.", UserIndex
            Exit Function
        End If
    End If
    
'**************************************************************************************************
'                   ESTA TODO CORRECTO, SE LE ENVIA AL USUARIO SUS DATOS Y SE ACTIVA COMO LOGUEADO
'**************************************************************************************************
    'TELEFRAG
    If .Invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
    If .Invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
    If .Invent.WeaponEqpSlot = 0 Then .Char.WeaponAnim = NingunArma

    If .flags.Paralizado Then
        Call enviarParalizado(UserList(.UserIndex))
    End If

    If .flags.Envenenado = 1 Then EnviarPaquete Paquetes.EstaEnvenenado, "", UserIndex, ToIndex
        

    EnviarPaquete ChangeMap, ITS(.pos.map) & ITS(MapInfo(.pos.map).climaActual) & MapInfo(.pos.map).Terreno & "," & MapInfo(.pos.map).zona & "," & MapInfo(.pos.map).Name, UserIndex
    EnviarPaquete ChangeMusic, ByteToString(val(mid(MapInfo(.pos.map).Music, 1, 1))), UserIndex
    
    Call enviarPosicion(UserList(.UserIndex))
    
    Call UpdateUserInv(True, UserIndex, 0)
    Call UpdateUserHechizos(True, UserIndex, 0)
    Call UpdateUserMap(UserIndex)
    Call EnviarHambreYsed(UserIndex)
    Call SendMOTD(UserIndex)
    
    .flags.UserLogged = True 'A partir de ahora el CloseSocket hace guardado del personaje (CloseUser)
    
     'Agrego la referencia del personaje al mapa
    MapInfo(.pos.map).usuarios.agregar (UserIndex)

    If .Stats.SkillPts > 0 Then
        Call EnviarSkills(UserIndex)
        Call EnviarSubirNivel(UserIndex, .Stats.SkillPts)
    End If
      
    If .flags.Privilegios > 0 Then
        Call LogGM(.id, HelperIP.longToIP(.ip) & "-" & .MacAddress, "LOGIN")
        GmsGroup.agregar UserIndex
    End If
        
   .Counters.IdleCount = 0

    If .flags.Navegando = 1 Then EnviarPaquete Paquetes.Navega, "", UserIndex, ToIndex

    .flags.PermitirDragAndDrop = True

     'Crea  el personaje del usuario
     ' MOFICIADO
    Call modPersonaje_TCP.MakeUserChar(UserList(UserIndex), 0, ToMap)
    EnviarPaquete IndiceChar, ITS(.Char.charIndex), UserIndex
    
     'Intervalos
    If .clase = eClases.Guerrero Then
        EnviarPaquete loguea, ByteToString(.flags.Privilegios) & IntervaloTotalG & "-" & IntervalosAntiLanzarAutomatico, UserIndex
    ElseIf .clase = eClases.Cazador Then
        EnviarPaquete loguea, ByteToString(.flags.Privilegios) & IntervaloTotalC & "-" & IntervalosAntiLanzarAutomatico, UserIndex
    Else
        EnviarPaquete loguea, ByteToString(.flags.Privilegios) & IntervaloTotal & "-" & IntervalosAntiLanzarAutomatico, UserIndex
    End If
    
    Call establecerIntervalos(UserList(UserIndex))
      '\
    
    EnviarPaquete Paquetes.HechizoFX, ITS(.Char.charIndex) & ByteToString(FXWARP) & ITS(0) & Chr$(SND_WARP), UserIndex, ToPCArea, .pos.map

      'Mascotas
    If .NroMacotas > 0 Then
        For N = 1 To MAXMASCOTAS
            If .MascotasType(N) > 0 Then
                .MascotasIndex(N) = SpawnNpc(.MascotasType(N), .pos, True, True)
            
                If Not .MascotasIndex(N) > MAXNPCS Then
                    NpcList(.MascotasIndex(N)).MaestroUser = UserIndex
                    Call FollowAmo(.MascotasIndex(N))
                 Else
                    .MascotasIndex(N) = 0
                End If
            End If
        Next N
    End If

      'Clan
    If .GuildInfo.id > 0 Then
          'Si tiene clan le aviso a los integrantes que se conecto
        EnviarPaquete Paquetes.MensajeClan1, .Name & " se ha conectado.", UserIndex, ToGuildMembers
        
          'Agrego al usuario a la lista de Onlines del clan
        .ClanRef.setOnline UserIndex
        
         ' Le aviso al usuario que ahora  pertenece a un clan
        EnviarPaquete Paquetes.infoClan, "1", UserIndex, ToIndex
        
          'Le envio las novedades.
         Call mdClanes.EnviarNovedadesClan(UserIndex)
    Else
         ' Le aviso al usuario que ahora  pertenece a un clan
        EnviarPaquete Paquetes.infoClan, "0", UserIndex, ToIndex
    End If
      '\
    
      'Eventos.
      'Estaba en un evento y cerro?
    Call modEventos.reEstablecerEventoUsuario(UserList(UserIndex), UserIndex)
    
     'Avisamos si hay algun evento haciendose
    If .evento Is Nothing And modEventos.getCantidadEnventosNoRetos > 0 Then
        If modEventos.getCantidadEnventosNoRetos = 1 Then
            EnviarPaquete Paquetes.MensajeTalk, "En estos momentos estamos realizando un evento automático. Escribe /EVENTO para obtener más info.", UserIndex, ToIndex
        Else
            EnviarPaquete Paquetes.MensajeTalk, "En estos momentos estamos realizando eventos automáticos. Escribe /EVENTO para obtener más info.", UserIndex, ToIndex
        End If
    End If
    
    Call SendUserStatsBox(UserIndex)

    If desban Then
        Call WarpUserChar(UserIndex, 1, 50, 50, True)
    End If
    
    .FechaIngreso = Now
    
     #If TDSFacil = 1 Then
        'If Not .Premium Then
        '    EnviarPaquete Paquetes.MensajeFight, "Tu cuenta no es premium. El tiempo restante que podrás jugar sin ser premium es de " & segundosAHoras(.segundosPremium) & ".", UserIndex, ToIndex
        ' End If
      #End If
   '***********************************************************************************
   'Pongo el personaje como online
    sql = "UPDATE " & DB_NAME_PRINCIPAL & ".usuarios set Online=1 WHERE ID = " & .id

    frmMysqlAuxiliar.cargadorPersonajes.Execute sql, , adExecuteNoRecords Or adCmdText
   '************************************************************************************
   '                      ACTUALIZO EL NUMERO DE USUARIOS                              *
    NumUsers = NumUsers + 1
    
    If esPremium Then NumUsersPremium = NumUsersPremium + 1
    
    If NumUsers > DayStats.MaxUsuarios Then DayStats.MaxUsuarios = NumUsers
    If NumUsers < DayStats.MinUsuarios Then DayStats.MinUsuarios = NumUsers
    
    If NumUsers > RecordUsuarios Then
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(25) & NumUsers, 0, ToAll
        RecordUsuarios = NumUsers
        Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(RecordUsuarios))
     End If

    Call MostrarNumUsers
   '************************************************************************************
    If NumUsers > 700 Then
        If frmMain.GameTimer.Interval <> 450 Then frmMain.GameTimer.Interval = 450
    ElseIf NumUsers > 600 Then
        If frmMain.GameTimer.Interval <> 400 Then frmMain.GameTimer.Interval = 400
      Else
        If frmMain.GameTimer.Interval <> 500 Then frmMain.GameTimer.Interval = 500
      End If
    '************************************************************************************
  End With

conectarPersonaje = True

End Function


Sub SendMOTD(ByVal UserIndex As Integer)
Dim j As Integer
'Mensaje del dia
EnviarPaquete MensajeSimple, Chr(54), UserIndex
For j = 1 To MaxLines
    EnviarPaquete Paquetes.MensajeCompuesto, Chr$(14) & MOTD(j).texto, UserIndex
Next j
End Sub

Sub ResetFacciones(ByVal UserIndex As Integer)
    With UserList(UserIndex).faccion
        .ArmadaReal = 0
        .FuerzasCaos = 0
        .CiudadanosMatados = 0
        .CriminalesMatados = 0
        .RecibioArmaduraCaos = 0
        .RecibioArmaduraReal = 0
        .RecibioExpInicialCaos = 0
        .RecibioExpInicialReal = 0
        .RecompensasCaos = 0
        .RecompensasReal = 0
    End With
End Sub

Sub ResetContadores(ByVal UserIndex As Integer)
    With UserList(UserIndex).Counters
        .AGUACounter = 0
        .COMCounter = 0
        .Frio = 0
        .HPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .Paralisis = 0
        .Pena = 0
        .STACounter = 0
        .Veneno = 0
        .TimerLanzarSpell = 0
        .TimerPuedeAtacar = 0
        .TimerUsarClic = 0
        .TimerUsarU = 0
        .FotoDenuncia = 0
    End With
End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)
    With UserList(UserIndex).Char
        .Body = 0
        .CascoAnim = 0
        .charIndex = 0
        .FX = 0
        .Head = 0
        .loops = 0
        .heading = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
    End With
End Sub

Sub ResetBasicUserInfo(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        .Name = ""
        .Password = ""
        .desc = ""
        .pos.map = 0
        .pos.x = 0
        .pos.y = 0
        .ip = 0
        .RDBuffer = ""
        .IDCuenta = 0
        .Premium = False
        .TokSolicitudDePersonaje = 0
        
        .clase = eClases.indefinido
        .Raza = eRazas.indefinido
        .Genero = eGeneros.indefinido
        
        .Email = ""
        .Hogar = ""
        .CentinelaID = 0  'EL YIND
        .PacketNumber = 1
        .MinPacketNumber = 1
        .Stats.Banco = 0
        .Stats.ELV = 0
        .Stats.ELU = 0
        .Stats.Exp = 0
        .Stats.Def = 0
        .Stats.NPCsMuertos = 0
        .Stats.UsuariosMatados = 0
        .Stats.SkillPts = 0
        
        .controlCheat.VecesAtack = 0
        .controlCheat.rompeIntervalo = 0
        .controlCheat.vecesCheatEngine = 0
    End With
End Sub

Sub ResetReputacion(ByVal UserIndex As Integer)
    With UserList(UserIndex).Reputacion
        .AsesinoRep = 0
        .BandidoRep = 0
        .BurguesRep = 0
        .LadronesRep = 0
        .NobleRep = 0
        .PlebeRep = 0
        .NobleRep = 0
        .promedio = 0
    End With
End Sub

Sub ResetGuildInfo(ByVal UserIndex As Integer)
    With UserList(UserIndex).GuildInfo
        .id = 0
        .ClanFundadoID = 0
        .echadas = 0
        .EsGuildLeader = 0
        .FundoClan = 0
        .GuildName = ""
        .Solicitudes = 0
        .SolicitudesRechazadas = 0
        .VecesFueGuildLeader = 0
        .ClanesParticipo = 0
        .GuildPoints = 0
    End With
End Sub

Sub ResetUserFlags(ByVal UserIndex As Integer)
    With UserList(UserIndex).flags
        .Comerciando = False
        .Ban = 0
        .DuracionEfecto = 0
        .TargetNPC = 0
        .TargetNpcTipo = 0
        .TargetObj = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        .TargetUser = 0
        .Hambre = 0
        .Sed = 0
        .Descansar = False
        .modoCombate = False
        .Navegando = 0
        .Oculto = 0
        .Envenenado = 0
        .Invisible = 0
        .Paralizado = 0
        .paralizadoPor = 0
        .Meditando = 0
        .Privilegios = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .hechizo = 0
        .Trabajando = False
        .PertAlCons = 0
        .PertAlConsCaos = 0
    End With
End Sub

Sub ResetUserSpells(ByVal UserIndex As Integer)
Dim loopC As Integer
For loopC = 1 To MAXUSERHECHIZOS
    UserList(UserIndex).Stats.UserHechizos(loopC) = 0
Next
End Sub

Sub ResetUserPets(ByVal UserIndex As Integer)
Dim loopC As Integer

UserList(UserIndex).NroMacotas = 0
For loopC = 1 To MAXMASCOTAS
    UserList(UserIndex).MascotasIndex(loopC) = 0
    UserList(UserIndex).MascotasType(loopC) = 0
Next loopC
End Sub


Public Sub ResetUserMascotas(UserIndex As Integer)
Dim tempbyte As Byte

With UserList(UserIndex)
    .NroMacotas = 0
    .NroMascotasGuardadas = 0
    For tempbyte = 1 To MAXMASCOTAS
        .MascotasIndex(tempbyte) = 0
        .MascotasGuardadas(tempbyte) = 0
        .MascotasType(tempbyte) = 0
    Next
End With
End Sub

Public Sub ResetUserComercio(UserIndex As Integer)

Dim i As Byte

   With UserList(UserIndex)
        .ComUsu.Acepto = False
        .ComUsu.DestUsu = 0
        .ComUsu.DestNick = ""
        .flags.Comerciando = False
    
        For i = 0 To MAX_OBJETOS_COMERCIABLES
            .ComUsu.cant(i) = 0
            .ComUsu.objeto(i) = 0
            .ComUsu.ObjetoIndex(i) = 0
        Next
   End With
End Sub




Sub ResetUserBanco(ByVal UserIndex As Integer)
Dim loopC As Integer
For loopC = 1 To MAX_BANCOINVENTORY_SLOTS
      UserList(UserIndex).BancoInvent.Object(loopC).Amount = 0
      UserList(UserIndex).BancoInvent.Object(loopC).Equipped = 0
      UserList(UserIndex).BancoInvent.Object(loopC).ObjIndex = 0
Next
UserList(UserIndex).BancoInvent.NroItems = 0
End Sub

Public Sub LimpiarComercioSeguro(ByVal UserIndex As Integer)
With UserList(UserIndex).ComUsu
    If .DestUsu > 0 Then
        Call FinComerciarUsu(.DestUsu)
        Call FinComerciarUsu(UserIndex)
    End If
End With
End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)

Dim userTmp As User
Set UserList(UserIndex).ClanRef = Nothing

'Esto es para los eventos en general
Set UserList(UserIndex).solicitudEvento = Nothing
'**************************************
UserList(UserIndex) = userTmp
UserList(UserIndex).UserIndex = UserIndex
UserList(UserIndex).ConnID = INVALID_SOCKET
UserList(UserIndex).ConfirmacionConexion = 0
UserList(UserIndex).InicioConexion = 0

UserList(UserIndex).MacAddress = ""
UserList(UserIndex).CryptOffset = 0

UserList(UserIndex).eventoOcultar.fecha = 0
UserList(UserIndex).eventoOcultar.Posicion.x = 0
UserList(UserIndex).eventoOcultar.Posicion.y = 0

Call LimpiarComercioSeguro(UserIndex)
Call ResetFacciones(UserIndex)
Call ResetContadores(UserIndex)
Call ResetCharInfo(UserIndex)
Call ResetBasicUserInfo(UserIndex)
Call ResetReputacion(UserIndex)
Call ResetGuildInfo(UserIndex)
Call ResetUserFlags(UserIndex)
Call LimpiarInventario(UserIndex)
Call ResetUserSpells(UserIndex)
Call ResetUserPets(UserIndex)
Call ResetUserBanco(UserIndex)
Call ResetUserComercio(UserIndex)

End Sub

Sub CloseUser(ByVal UserIndex As Integer)

Dim x As Integer
Dim y As Integer
Dim map As Integer
Dim TempInt As Integer


With UserList(UserIndex)

    map = .pos.map
    x = .pos.x
    y = .pos.y
    .Char.FX = 0
    .Char.loops = 0
    
    ' Reseteo
    EnviarPaquete Paquetes.HechizoFX, ITS(.Char.charIndex) & ByteToString(0) & ITS(0), UserIndex, ToPCArea, .pos.map

    ' Marco como delsogueado
    .flags.UserLogged = False
    .flags.Saliendo = eTipoSalida.NoSaliendo

    'Le devolvemos el body y head originales
    If .flags.AdminInvisible = 1 Then Call DoAdminInvisible(UserIndex)
    
    'Si esta en party le devolvemos la experiencia
    If .PartyIndex > 0 Then Call mdParty.SalirDeParty(UserIndex)

    'Si esta trabajando dejar de estarlo
    If .flags.Trabajando Then Call DejarDeTrabajar(UserList(UserIndex))
   
    ' Si esta mimetizado, se desmimetiza
    If .flags.Mimetizado Then Call modMimetismo.finalizarEfecto(UserList(UserIndex))
    
    'Le aviso al evento donde el estaba que el usuario cerro
    If Not .evento Is Nothing Then
        If .evento.getNombre = "Reto" Then
            'El evento esta desarrollandose justo ahora?
            If .evento.getEstadoEvento = eEstadoEvento.Desarrollandose Then
                Call .evento.usuarioCerro(UserIndex)
            End If
        End If
    End If
    
        Debug.Print UserList(UserIndex).NroMacotas
        
    If Not .resucitacionPendiente Is Nothing Then
        Call modResucitar.cancelarResucitacion(.resucitacionPendiente)
    End If
    
    Call Anticheat_MemCheck.hook_cierraPersonaje(UserList(UserIndex))
    
        Debug.Print UserList(UserIndex).NroMacotas
        
    'Aviso al clan que el usuario cerro
    If .GuildInfo.id > 0 Then
        Call .ClanRef.setOffline(UserIndex)
        EnviarPaquete Paquetes.MensajeClan1, .Name & " se ha desconectado.", UserIndex, ToGuildMembers
    End If

    'Si es gm..
    If .flags.Privilegios > 0 Then
        'Es un gm lo quito de la lista
        Call GmsGroup.eliminar(UserIndex)
    
        TempInt = DateDiff("n", .FechaIngreso, Now) 'Calculo cuantos minutos estuvo
        
        Call WarpUserChar(UserIndex, MAPA_DESCANSO_GMS, RandomNumber(45, 55), RandomNumber(45, 55), True)
    
        Call LogGM(.id, str$(TempInt), "LOGOUT")
                
        conn.Execute "UPDATE " & DB_NAME_PRINCIPAL & ".juego_gms SET MinutosLogin=MinutosLogin + " & TempInt & " WHERE IDUsuario=" & .id, , adExecuteNoRecords
    End If

    conn.Execute "UPDATE " & DB_NAME_PRINCIPAL & ".juego_gms SET MinutosLogin=MinutosLogin + " & TempInt & " WHERE IDUsuario=" & .id, , adExecuteNoRecords
    
    ' Segundos de juego gratuitos
    #If TDSFacil = 1 Then
    
        ' Si el usuario no es premium
        If Not .Premium Then
            Call modCuentas.actualizarDatosCuenta(UserList(UserIndex))
        End If
    
    #End If
            
    'Quitar el dialogo
    map = .pos.map
    
    If MapInfo(map).usuarios.getCantidadElementos > 0 Then
        EnviarPaquete Paquetes.QDL, ITS(.Char.charIndex), UserIndex, ToMap
    End If
    
    ' Anti robo de npcs. Libera al npc
    If .LuchandoNPC > 0 Then
        Call AntiRoboNpc.resetearLuchador(NpcList(.LuchandoNPC))
    End If
    
    'Borro al personaje del mapa
    If .Char.charIndex > 0 Then
        Call EraseUserChar(ToMapButIndex, UserIndex, map, UserIndex)
    End If
    
    'Borrar mascotas
    Call GuardarMascotas(UserList(UserIndex))

    'Quito la referencia del personaje en el mapa
    Call MapInfo(map).usuarios.eliminar(UserIndex)

    ' Si el usuario habia dejado un msg en la gm's queue lo borramos
    Call Ayuda.eliminar(UserIndex)  'Si no esta no tira ningun error. Tarda lo mismo que un buscar

    'Guardo el personaje
    Call SaveUser(UserIndex)
    
    'Le aviso al evento donde el estaba que el usuario cerro
    If Not .evento Is Nothing Then
        If .evento.getNombre <> "Reto" Then
            'El evento esta desarrollandose justo ahora?
            If .evento.getEstadoEvento = eEstadoEvento.Desarrollandose Then
                Call .evento.usuarioCerro(UserIndex)
            End If
        ElseIf .evento.getNombre = "Reto" Then
            If .evento.getEstadoEvento = eEstadoEvento.Preparacion Then
                Call modEventos.quitarEvento(.evento)
                Call .evento.cancelar
            End If
        End If
        
        Set .evento = Nothing
    End If
End With

'Actualizo el numero de los onlines
Call MostrarNumUsers

End Sub

Sub ReloadSokcet()
    If NumUsers <= 0 Then
        Call WSApiReiniciarSockets
    End If
End Sub

Public Sub EcharPjsNoPrivilegiados()
Dim loopC As Long

' Se echa del juego a todos aquellos usuarios que estan logueados y no son Game Masters
For loopC = 1 To LastUser
    If UserList(loopC).flags.UserLogged Then
        If UserList(loopC).flags.Privilegios < 1 Then
            Call CloseSocket(loopC)
        End If
    End If
Next loopC

End Sub

Function ConvTextCapital(texto As String) As String
    ' Convierte el texto enviado en letra capital, la primera en mayúscula y el resto en minúculas
    ConvTextCapital = UCase(Left$(texto, 1)) & LCase(mid(texto, 2))
End Function
