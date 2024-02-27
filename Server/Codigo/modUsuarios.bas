Attribute VB_Name = "modUsuarios"

Option Explicit

'Retorna verdadero si el personaje existe
Public Function existePersonaje_Nombre(nombre As String) As Boolean
Dim infoPersonaje As ADODB.Recordset

sql = "SELECT ID FROM " & DB_NAME_PRINCIPAL & ".usuarios where NickB = '" & nombre & "'"

Set infoPersonaje = conn.Execute(sql, , adCmdText)

existePersonaje_Nombre = Not (infoPersonaje.EOF)

infoPersonaje.Close 'Cierro

End Function
' Crea un personaje en blanco retornado el ID asignado
Public Function crearPersonaje(nombre As String) As Long

Dim infoPJ As ADODB.Recordset

'Lo agrego
conn.Execute "INSERT INTO  " & DB_NAME_PRINCIPAL & ".usuarios(NickB) values('" & nombre & "')", , adExecuteNoRecords

'Obtengo el ID
sql = "SELECT last_insert_id() AS ID FROM " & DB_NAME_PRINCIPAL & ".usuarios"
Set infoPJ = conn.Execute(sql)  'Obtengo el ID del ultimo personaje agregado

crearPersonaje = infoPJ!id

'Cerramos
infoPJ.Close

End Function

Public Sub establecerClanFundado(ByRef Usuario As User, ByRef clan As cClan)
    
    Usuario.GuildInfo.FundoClan = 1
    Usuario.GuildInfo.EsGuildLeader = 1
    Usuario.GuildInfo.ClanFundadoID = clan.id
    Set Usuario.ClanRef = clan
    
    Usuario.GuildInfo.id = clan.id
    Usuario.GuildInfo.GuildName = clan.getNombre
    
End Sub


Public Sub establecerClanUsuarioOffline(idUsuario As Long, ClanID As Long)

    sql = "UPDATE " & DB_NAME_PRINCIPAL & ".usuarios SET IDClan=" & ClanID & ", ClanesParticipoB=ClanesParticipoB+1, GuildPointsB=GuildPointsB+25 WHERE ID = " & idUsuario
    conn.Execute (sql), , adExecuteNoRecords
    
End Sub


'Esta funcion asigna un clan a un usuario
'Factorice esto en una funcion porque lo uso en dos partes
'y ni da andar copiando y pegando codigo
Public Sub establecerClanAUsuarioOnline(UserIndex As Integer, clan As cClan)

    UserList(UserIndex).GuildInfo.GuildName = clan.getNombre()
    UserList(UserIndex).GuildInfo.id = clan.id
    
    Set UserList(UserIndex).ClanRef = clan
    
    'Actualizo el nombre del personaje con el clan en la pantalla de los usuarios
    Call modPersonaje_TCP.actualizarNick(UserList(UserIndex))
    
    EnviarPaquete Paquetes.infoClan, "1", UserIndex, ToIndex
    
    'Si el usuario esta online lo pongo como online en el clan
    Call clan.setOnline(UserIndex)
    
    'Guardo el usuario por als dudas
    Call SaveUser(UserIndex, 1)
    
End Sub

Public Sub quitarClanUsuario(idPersonaje As Long)
Dim UserIndex As Integer

UserIndex = IDIndex(idPersonaje)

If UserIndex > 0 Then
    Call modUsuarios.quitarClanUsuarioOnline(UserList(UserIndex), UserIndex, False)
Else
    Call modUsuarios.quitarClanUsuarioOffline(idPersonaje, False)
End If

End Sub
Private Sub quitarClanUsuarioOnline(ByRef Usuario As User, UserIndex As Integer, echado As Boolean)

    Set Usuario.ClanRef = Nothing
    Usuario.GuildInfo.GuildPoints = 0
    Usuario.GuildInfo.GuildName = ""
    Usuario.GuildInfo.id = 0
    Usuario.GuildInfo.EsGuildLeader = 0
    
    'Si fue echado aumento sus estadisticas
    If echado Then Usuario.GuildInfo.echadas = Usuario.GuildInfo.echadas + 1
    
    'Para tener un sistema consistente, por las dudas de que caiga el server, guardo al char
    Call SaveUser(UserIndex, 1)
    
    ' Le aviso al usuario que ahora NO pertenece a un clan
    EnviarPaquete Paquetes.infoClan, "0", UserIndex, ToIndex
    
    ' Actualizo el estado visual del personaje
    Call modPersonaje_TCP.actualizarNick(UserList(UserIndex))
End Sub

Private Sub quitarClanUsuarioOffline(idUsuario As Long, echado As Boolean)
    Dim echadoVal As Byte

    If echado Then
        echadoVal = 1
    Else
        echadoVal = 0
    End If
    
    'Actualizo el char
    conn.Execute ("UPDATE " & DB_NAME_PRINCIPAL & ".usuarios SET IDClan=0, EsGuildLeaderB=0, guildPtsB = 0, Echadasb = Echadasb+" & echadoVal & " WHERE ID= " & idUsuario), , adExecuteNoRecords
End Sub

'True: el personaje existe
'False: el personaje no se encuentra en la base de datos
Private Function existePersonaje(idPersonaje As Long) As Boolean

    Dim infoPersonaje As ADODB.Recordset

    sql = "SELECT ID FROM " & DB_NAME_PRINCIPAL & ".usuarios where ID = " & idPersonaje
    
    Set infoPersonaje = conn.Execute(sql)
    
    If Not infoPersonaje.EOF Then
        existePersonaje = True
    Else
        existePersonaje = False
    End If
    
    infoPersonaje.Close
    
End Function

Public Sub establecerLiderazgo(idPersonaje As Long, EsLider As Boolean)

Dim UnoOCero As Byte
Dim UserIndex As Integer

UserIndex = IDIndex(idPersonaje) 'Obtengo el userindex en caso de que el personaje este online

If EsLider Then UnoOCero = 1 Else UnoOCero = 0

If UserIndex > 0 Then
    'El personaje esta online
    UserList(UserIndex).GuildInfo.EsGuildLeader = UnoOCero
    'Guardo el personaje (como online) para que si se cae el server no quede colgado
    Call SaveUser(UserIndex, 1)
Else
    sql = "UPDATE " & DB_NAME_PRINCIPAL & ".usuarios SET EsGuildLeaderB=" & UnoOCero & " WHERE ID = " & idPersonaje
    conn.Execute sql, , adExecuteNoRecords
End If

End Sub

'0: El usuario no tiene clan
'> 0: El ID del clan del usuario
'-1: El personaje no existe
Public Function getIDClanUsuarioOffline(idPersonaje As Long) As Long
    Dim infoPersonaje As ADODB.Recordset
    
    sql = "SELECT IDClan FROM " & DB_NAME_PRINCIPAL & ".usuarios where ID = " & idPersonaje
    Set infoPersonaje = conn.Execute(sql)

    If Not infoPersonaje.EOF Then ' ¿Existe el personaje?
        'Devuelvo el id del clna
        getIDClanUsuarioOffline = CLng(infoPersonaje!IDClan)
    Else 'El personaje no existe
        getIDClanUsuarioOffline = -1
    End If
    
    infoPersonaje.Close
    Set infoPersonaje = Nothing
End Function

'Agrega experiencia al personaje, luego se fija si le alcanza para pasar de nivel
'No suma experiencia si el usuariro alcanzo el nivel maximp
Public Sub agregarExperiencia(UserIndex As Integer, ByVal Exp As Long)

    If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
        
        UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + Exp
        
        Call CheckUserLevel(UserList(UserIndex))
    
    End If
    
End Sub

'Agrega experiencia al personaje, luego se fija si le alcanza para pasar de nivel
'No suma experiencia si el usuariro alcanzo el nivel maximp
Public Sub agregarOro(personaje As User, ByVal cantidad As Long)
    
    personaje.Stats.GLD = personaje.Stats.GLD + cantidad
    
    EnviarPaquete Paquetes.EnviarOro, Codify(personaje.Stats.GLD), personaje.UserIndex, ToIndex
        
End Sub

Public Sub transportarUsuarioOffline(idUsuario As Long, ByVal mapa As Integer, ByVal x As Integer, ByVal y As Integer)
    
    Dim sql As String
    
    sql = "UPDATE " & DB_NAME_PRINCIPAL & ".usuarios SET Mapb =" & mapa & ", XB=" & x & ", YB=" & y & " WHERE ID=" & idUsuario
    
    conn.Execute sql, , adExecuteNoRecords
End Sub

Public Sub entregarOroOffline(idUsuario As Long, oro As Long)
    Dim sql As String
    
    sql = "UPDATE " & DB_NAME_PRINCIPAL & ".usuarios SET GLDB = GLDB + " & oro & " WHERE ID=" & idUsuario
    
    conn.Execute sql, , adExecuteNoRecords
End Sub

Public Sub matarUsuarioOffline(idUsuario As Long)
    Dim sql As String
    
    sql = "UPDATE " & DB_NAME_PRINCIPAL & ".usuarios SET MuertoB = 1 WHERE ID=" & idUsuario
    
    conn.Execute sql, , adExecuteNoRecords
End Sub


Public Sub transportarUsuario(idUsuario As Long, UserIndex As Integer, mapa As Integer, x As Integer, y As Integer, exacto As Boolean, PuedeAgua As Boolean)
    
    If UserIndex > 0 Then
        Call transportarUsuarioOnline(UserIndex, mapa, x, y, exacto, PuedeAgua)
    Else
        Call transportarUsuarioOffline(idUsuario, mapa, x, y)
    End If

End Sub

Public Sub transportarUsuarioOnline(UserIndex As Integer, mapa As Integer, x As Integer, y As Integer, exacto As Boolean, PuedeAgua As Boolean)
    
    Dim auxPos As WorldPos
    Dim nPos As WorldPos
    
    If exacto Then
        nPos.x = x
        nPos.map = mapa
        nPos.y = y
    Else
        auxPos.map = mapa
        auxPos.x = x
        auxPos.y = y
        
        Call ClosestLegalPos(auxPos, nPos, UserList(UserIndex))
    End If
    
    If nPos.x <> 0 And nPos.y <> 0 Then
        Call WarpUserChar(UserIndex, nPos.map, nPos.x, nPos.y, False)
    Else
        Call LogError("Fallo el transportar online con el personaje " & UserList(UserIndex).Name & " en mapa " & mapa & " X: " & x & "Y:" & y)
    End If

End Sub
