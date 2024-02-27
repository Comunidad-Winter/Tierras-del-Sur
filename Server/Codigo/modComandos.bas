Attribute VB_Name = "modComandos"
Option Explicit

Private Const MAX_CRIATURAS_MAPA = 200

Public Sub EliminarPortal(GameMaster As User, mapa As Integer, x As Byte, y As Byte)

If Not SV_PosicionesValidas.existePosicionMundo(mapa, x, y) Then
    EnviarPaquete Paquetes.mensajeinfo, "Tenes que hacer clic en donde está el portal que queres eliminar.", GameMaster.UserIndex, ToIndex
    Exit Sub
End If

'Puede hacer esto?
If GameMaster.flags.Privilegios = PRIV_GAMEMASTER Then
    If Not (modGameMasterEventos.esMapaDeEvento(mapa)) Then
        EnviarPaquete Paquetes.mensajeinfo, "Los GameMasters solo pueden eliminar portales en mapa de eventos. Segui ayudando a Tierras del Sur y algún día serás Dios.", GameMaster.UserIndex, ToIndex
        Exit Sub
    End If
End If
        
 'Hay un portal?
If Not ObjData(MapData(mapa, x, y).OBJInfo.ObjIndex).ObjType = OBJTYPE_TELEPORT Then
    EnviarPaquete Paquetes.mensajeinfo, "No hay un portal en la posición marcada.", GameMaster.UserIndex, ToIndex
    Exit Sub
End If

' Eliminamos el objeto
Call EraseObj(ToMap, 0, mapa, MapData(mapa, x, y).OBJInfo.Amount, mapa, x, y)
            
' Eliminamos la posición
Set MapData(mapa, x, y).accion = Nothing
   
LogGM GameMaster.id, mapa & "." & x & "." & y, "DT"
             
End Sub
Public Sub CrearPortal(GameMaster As User, ByVal mapa As Integer, ByVal x As Byte, ByVal y As Byte)

Dim posicionYCreacion As Integer
Dim MiObj As obj

If SV_PosicionesValidas.existeMapa(mapa) = False Then
    EnviarPaquete Paquetes.mensajeinfo, "El mapa " + mapa + " no existe.", GameMaster.UserIndex, ToIndex
    Exit Sub
End If

If SV_PosicionesValidas.existePosicionMundo(mapa, x, y) = False Then
    EnviarPaquete Paquetes.mensajeinfo, "La posición de destino no es válida", GameMaster.UserIndex, ToIndex
    Exit Sub
End If

posicionYCreacion = GameMaster.pos.y - 1

MiObj.ObjIndex = 378
MiObj.Amount = 1
                
'Creo el portal
MakeObj ToMap, 0, GameMaster.pos.map, MiObj, GameMaster.pos.map, GameMaster.pos.x, posicionYCreacion

' Creo la accion
Dim accion As cAccionExit
Set accion = New cAccionExit

Call accion.crear(mapa, x, y, 4, True)

Set MapData(GameMaster.pos.map, GameMaster.pos.x, posicionYCreacion).accion = accion
        
' Log
LogGM GameMaster.id, "O: " & GameMaster.pos.map & " D: " & mapa & "." & x & "." & y, "CT"

End Sub

Public Sub BanearUsuario(nombreGM As String, nombreUsuario As String, razon As String, Dias As Byte, GmOnline As Boolean)

Dim TempInt As Integer

TempInt = NameIndex(nombreUsuario)
If TempInt > 0 Then 'Esta on
    UserList(TempInt).flags.Ban = 1

    If LenB(UserList(TempInt).flags.Banrazon) > 0 Then
        UserList(TempInt).flags.Banrazon = UserList(TempInt).flags.Banrazon & vbCrLf & "Baneado por " & nombreGM & ". Razon: " & razon & Date
    Else
        UserList(TempInt).flags.Banrazon = "Baneado por " & nombreGM & ". Razon: " & razon & " " & Date
    End If
        
    If Dias > 0 Then UserList(TempInt).flags.Unban = Date + Dias Else UserList(TempInt).flags.Unban = "NUNCA"
    
    If Not CloseSocket(TempInt) Then Call LogError("Banear Usuario")
        
    EnviarPaquete Paquetes.MensajeServer, nombreGM & " ha baneado a " & nombreUsuario & ".", 0, ToAdmins
Else  ' Esta off
    Dim infoPersonaje As ADODB.Recordset

    Call cargarAtributosPersonajeOffline(nombreUsuario, infoPersonaje, "BANB, UNBAN, BANRAZB", True)
    
    If Not infoPersonaje.EOF Then
        If Dias > 0 Then 'Ban temporal
            infoPersonaje!Unban = Date + Dias + 1 ' Le sumo uno para que el dia no sea INCLUSIVE. Esto genera cada tanto soportes
        Else
            infoPersonaje!Unban = "NUNCA"
        End If
        
        infoPersonaje!banb = 1
        infoPersonaje!banrazb = infoPersonaje!banrazb & vbCrLf & "Baneado por " & nombreGM & ". Razon: " & razon & " " & Date
        infoPersonaje.Update
        EnviarPaquete Paquetes.MensajeServer, nombreGM & " ha baneado a " & nombreUsuario & ".", 0, ToAdmins
    Else
         If GmOnline Then EnviarPaquete Paquetes.MensajeSimple, Chr$(73), NameIndex(nombreGM), ToIndex
    End If
    
    'Liberamos
    infoPersonaje.Close
    Set infoPersonaje = Nothing
End If
    
End Sub

Public Sub invocarCriatura(IDCriatura As Integer, Renace As Boolean, GameMaster As User)

Dim npcIndex As Integer

If MapInfo(GameMaster.pos.map).NPCs.getCantidadElementos + 1 > MAX_CRIATURAS_MAPA Then
    EnviarPaquete Paquetes.mensajeinfo, "Se alcanzó la máxima cantidad de criaturas por mapa (" & MAX_CRIATURAS_MAPA & ").", GameMaster.UserIndex, ToIndex
    Exit Sub
End If

npcIndex = SpawnNpc(IDCriatura, GameMaster.pos, True, Renace)

If Renace Then
    LogGM GameMaster.id, NpcList(npcIndex).Name, "RACC"
Else
    LogGM GameMaster.id, NpcList(npcIndex).Name, "ACC"
End If

EnviarPaquete Paquetes.mensajeinfo, "Has invocado " & NpcList(npcIndex).Name, GameMaster.UserIndex, ToIndex

End Sub

#If TDSFacil = 1 Then
    Public Sub enviarTiempoGratisTDSF(personaje As User)
        Dim segundos As Long
        Dim jugadosAhora As Long
        
        ' Obtengo lo que jugo ahora
        jugadosAhora = DateDiff("s", personaje.FechaIngreso, Now)
        
        ' Obtengo los restantes
        If personaje.Premium Then
            segundos = personaje.segundosPremium
        ElseIf jugadosAhora > personaje.segundosPremium Then
            segundos = 0
        Else
            segundos = personaje.segundosPremium - jugadosAhora
        End If
        
        If personaje.Premium Then
            ' Chequeo que no sea negativo.
            If segundos > 0 Then
                EnviarPaquete Paquetes.MensajeFight, "Tu cuenta es Premium podés jugar Tierras del Sur Fácil sin limites. Si tu cuenta se vence antes de fin de mes, tendrás para jugar " & HelperTiempo.segundosAHoras(segundos) & ". Gracias por ayudar a Tierras del Sur.", personaje.UserIndex, ToIndex
            Else
                EnviarPaquete Paquetes.MensajeFight, "Tu cuenta es Premium podés jugar Tierras del Sur Fácil sin limites. Si tu cuenta se vence antes de fin de mes, ya no te quedan minutos gratuitos para jugar.", personaje.UserIndex, ToIndex
            End If
        Else
            If segundos < 60 Then
                EnviarPaquete Paquetes.MensajeFight, "El tiempo gratuito para jugar Tierras del Sur finalizó. Deberás esperar hasta el comienzo del próximo mes. Cargá tiempo premium para jugar sin limites y ayudar a Tierras del Sur a continuar y mejorar.", personaje.UserIndex, ToIndex
            Else
                EnviarPaquete Paquetes.MensajeFight, "Podés jugar Tierras del Sur Fácil sin ser premium " & HelperTiempo.segundosAHoras(segundos) & " más. Cargá tiempo premium para jugar sin limites y ayudar a Tierras del Sur a continuar y mejorar.", personaje.UserIndex, ToIndex
            End If
        End If
    
    End Sub
#End If

Public Sub enviarPenas(nombrePersonaje As String, GameMaster As User)

Dim TempInt As Integer

TempInt = NameIndex(nombrePersonaje)

If TempInt > 0 Then 'Esta on
    If LenB(UserList(TempInt).flags.Penasas) = 0 Then
        EnviarPaquete Paquetes.mensajeinfo, "El personaje no tiene penas", GameMaster.UserIndex, ToIndex
    Else
        EnviarPaquete Paquetes.mensajeinfo, "Penas del personaje " & UserList(TempInt).Name & ":", GameMaster.UserIndex, ToIndex
        EnviarPaquete Paquetes.mensajeinfo, UserList(TempInt).flags.Penasas, GameMaster.UserIndex, ToIndex
    End If
Else
    
    Dim infoPersonaje As ADODB.Recordset

    Call cargarAtributosPersonajeOffline(nombrePersonaje, infoPersonaje, "PENASASB, BANRAZB", False)
    
    If Not infoPersonaje.EOF Then
        If Len(infoPersonaje!penasasb) = 0 And Len(infoPersonaje!banrazb) = 0 Then
            EnviarPaquete Paquetes.mensajeinfo, "El personaje " & nombrePersonaje & " no tiene penas.", GameMaster.UserIndex, ToIndex
        Else
            EnviarPaquete Paquetes.mensajeinfo, "Penas del personaje " & nombrePersonaje & ":", GameMaster.UserIndex, ToIndex
            EnviarPaquete Paquetes.mensajeinfo, infoPersonaje!penasasb & infoPersonaje!banrazb, GameMaster.UserIndex, ToIndex
        End If
    Else
        EnviarPaquete Paquetes.mensajeinfo, "El personaje " & nombrePersonaje & " no existe.", GameMaster.UserIndex, ToIndex
    End If
    
    infoPersonaje.Close
    Set infoPersonaje = Nothing
End If
            
End Sub
