Attribute VB_Name = "Extra"
Option Explicit


Sub ClosestLegalPos(pos As WorldPos, ByRef nPos As WorldPos, ByRef personaje As User)
'*****************************************************************
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************
Dim Notfound As Boolean
Dim loopC As Integer
Dim tX As Integer
Dim tY As Integer
Dim minX As Integer
Dim maxX As Integer
Dim minY As Integer
Dim maxY As Integer


Notfound = True
nPos.map = pos.map
loopC = 0 'Comienza en los cuadrados mas cercanos

minX = maxi(pos.x - loopC, SV_Constantes.X_MINIMO_JUGABLE)
maxX = mini(pos.x + loopC, SV_Constantes.X_MAXIMO_JUGABLE)
minY = maxi(pos.y - loopC, SV_Constantes.Y_MINIMO_JUGABLE)
minY = mini(pos.y + loopC, SV_Constantes.Y_MAXIMO_JUGABLE)

'Voy a intentarlo mientras no encuentre una posicion y hayan fallado 13 intentos
Do While Notfound And loopC < 13

    For tY = pos.y - loopC To pos.y + loopC
        For tX = pos.x - loopC To pos.x + loopC
            'No es valida los lugares donde haya portales
            If LegalPos(nPos.map, tX, tY, personaje) Then
                If MapData(nPos.map, tX, tY).accion Is Nothing Then
                    nPos.x = tX
                    nPos.y = tY
                    Notfound = False
                    Exit Sub
                End If
            End If
        Next tX
    Next tY
    loopC = loopC + 1
Loop

If Notfound = True Then
    nPos.x = 0
    nPos.y = 0
    nPos.map = 0
End If
End Sub

Sub ClosestLegalPosNPC(pos As WorldPos, ByRef nPos As WorldPos, criatura As npc)
'*****************************************************************
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************
Dim Notfound As Boolean
Dim loopC As Integer
Dim tX As Integer
Dim tY As Integer

Notfound = True
nPos.map = pos.map
loopC = 0 'Comienza en los cuadrados mas cercanos

'Voy a intentarlo mientras no encuentre una posicion y hayan fallado 13 intentos
Do While Notfound And loopC < 13

    For tY = pos.y - loopC To pos.y + loopC
        For tX = pos.x - loopC To pos.x + loopC
            'No es valida los lugares donde haya portales
            If tY > 0 And tX > 0 Then
                If SV_PosicionesValidas.esPosicionJugable(CByte(tX), CByte(tY)) Then
                    If SV_PosicionesValidas.esPosicionUsableNPC(MapData(nPos.map, tX, tY), criatura) And MapData(nPos.map, tX, tY).accion Is Nothing Then
                        nPos.x = tX
                        nPos.y = tY
                        Notfound = False
                    End If
                End If
            End If
        Next tX
    Next tY
    
    loopC = loopC + 1
Loop

If Notfound = True Then
    nPos.x = 0
    nPos.y = 0
    nPos.map = 0
End If
End Sub

Sub HeadtoPos(Head As Byte, ByRef pos As WorldPos)
'*****************************************************************
'Toma una posicion y se mueve hacia donde esta perfilado
'*****************************************************************
Dim x As Integer
Dim y As Integer
Dim nX As Integer
Dim nY As Integer

x = pos.x
y = pos.y
If Head = eHeading.NORTH Then
    nX = x
    nY = y - 1
End If
If Head = eHeading.SOUTH Then
    nX = x
    nY = y + 1
End If
If Head = eHeading.EAST Then
    nX = x + 1
    nY = y
End If
If Head = eHeading.WEST Then
    nX = x - 1
    nY = y
End If
'Devuelve valores
pos.x = nX
pos.y = nY
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SendHelp
' DateTime  : 18/02/2007 19:16
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub SendHelp(ByVal index As Integer)
Dim NumHelpLines As Integer
Dim loopC As Integer

NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))
For loopC = 1 To NumHelpLines
    EnviarPaquete Paquetes.mensajeinfo, GetVar(DatPath & "Help.dat", "Help", "Line" & loopC), index
Next loopC

End Sub

'
Sub LookatTile(ByVal UserIndex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer)

Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Stat As String

'¿Posicion valida?

' ¿ Como funciona?
' Se fija si la posición es válida y actualizar el targetX, Y

' Prioridades
' Objeto
'   Objeto donde estoy haciendo clic
'   Objeto hacia la derecha de donde se esta haciendo clic
'   Objeto hacia abajo a la derecha
'   Objeto hacia abajo
' Sino hay Objeto
'   Se fija si hay una criatura en la posición Y + 1
'       Hace algo
'   Se fija si hay un personaje en la posición Y + 1
'       Si encontro un personaje le envía la información de ese personaje y lo pone en TargetUser
'
'
'

If SV_PosicionesValidas.existePosicionMundo(map, x, y) Then
    UserList(UserIndex).flags.TargetMap = map
    UserList(UserIndex).flags.TargetX = x
    UserList(UserIndex).flags.TargetY = y
    '¿Es un obj?
    If MapData(map, x, y).OBJInfo.ObjIndex > 0 Then
        'Informa el nombre
'        Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Name & FONTTYPE_INFO)
'        UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.ObjIndex
        UserList(UserIndex).flags.TargetObjMap = map
        UserList(UserIndex).flags.TargetObjX = x
        UserList(UserIndex).flags.TargetObjY = y
        FoundSomething = 1
    ElseIf MapData(map, x + 1, y).OBJInfo.ObjIndex > 0 Then
        'Informa el nombre
        If ObjData(MapData(map, x + 1, y).OBJInfo.ObjIndex).ObjType = OBJTYPE_PUERTAS Then
'            Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).Name & FONTTYPE_INFO)
'            UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y).OBJInfo.ObjIndex
            UserList(UserIndex).flags.TargetObjMap = map
            UserList(UserIndex).flags.TargetObjX = x + 1
            UserList(UserIndex).flags.TargetObjY = y
            FoundSomething = 1
        End If
    ElseIf MapData(map, x + 1, y + 1).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(map, x + 1, y + 1).OBJInfo.ObjIndex).ObjType = OBJTYPE_PUERTAS Then
            'Informa el nombre
'            Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).Name & FONTTYPE_INFO)
'            UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex
            UserList(UserIndex).flags.TargetObjMap = map
            UserList(UserIndex).flags.TargetObjX = x + 1
            UserList(UserIndex).flags.TargetObjY = y + 1
            FoundSomething = 1
        End If
    ElseIf MapData(map, x, y + 1).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(map, x, y + 1).OBJInfo.ObjIndex).ObjType = OBJTYPE_PUERTAS Then
            'Informa el nombre
'            Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).Name & FONTTYPE_INFO)
'            UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y + 1).OBJInfo.ObjIndex
            UserList(UserIndex).flags.TargetObjMap = map
            UserList(UserIndex).flags.TargetObjX = x
            UserList(UserIndex).flags.TargetObjY = y + 1
            FoundSomething = 1
        End If
    End If
    
    If FoundSomething = 1 Then
        UserList(UserIndex).flags.TargetObj = MapData(map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex
        
        If UserList(UserIndex).flags.Privilegios > 1 Or UserList(UserIndex).Stats.UserSkills(Supervivencia) = 100 Then
            EnviarPaquete Paquetes.ClickObjeto, ITS(UserList(UserIndex).flags.TargetObj) & ITS(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.Amount), UserIndex
        Else
            EnviarPaquete Paquetes.ClickObjeto, ITS(UserList(UserIndex).flags.TargetObj), UserIndex
        End If
    End If
    
    '¿Es un personaje?
    If y + 1 <= Y_MAXIMO_USABLE Then
    
        If MapData(map, x, y + 1).UserIndex > 0 Then
            TempCharIndex = MapData(map, x, y + 1).UserIndex
            FoundChar = 1
        End If
        
        If MapData(map, x, y + 1).npcIndex > 0 Then
            TempCharIndex = MapData(map, x, y + 1).npcIndex
            FoundChar = 2
        End If
        
    End If
    
    '¿Es un personaje?
    If FoundChar = 0 Then
        If MapData(map, x, y).UserIndex > 0 Then
            TempCharIndex = MapData(map, x, y).UserIndex
            FoundChar = 1
        End If
        If MapData(map, x, y).npcIndex > 0 Then
            TempCharIndex = MapData(map, x, y).npcIndex
            FoundChar = 2
        End If
    End If
    'Reaccion al personaje
    
  If FoundChar = 1 Then '  ¿Encontro un Usuario?
          If UserList(TempCharIndex).flags.AdminInvisible = 0 Or UserList(UserIndex).flags.Privilegios > 1 Then
            If UserList(TempCharIndex).flags.Muerto = 1 Then
                If EsNewbie(TempCharIndex) Then Stat = "1"
                EnviarPaquete Paquetes.VeUser, Stat & ITS(UserList(TempCharIndex).Char.charIndex), UserIndex
                FoundSomething = 1
                UserList(UserIndex).flags.TargetUser = TempCharIndex
                UserList(UserIndex).flags.TargetNPC = 0
                UserList(UserIndex).flags.TargetNpcTipo = 0
            Else
                
                If UserList(TempCharIndex).flags.Mimetizado = 1 Then
                    Stat = Stat & "9"
                ElseIf UserList(TempCharIndex).flags.Privilegios > 0 Then
                    Stat = Stat & UserList(TempCharIndex).flags.Privilegios + 1
                Else
                    If UserList(TempCharIndex).flags.PertAlCons = 1 Then
                        Stat = Stat & "6"
                    ElseIf UserList(TempCharIndex).flags.PertAlConsCaos = 1 Then
                        Stat = Stat & "7"
                    ElseIf UserList(TempCharIndex).faccion.alineacion = eAlineaciones.caos Then
                        Stat = Stat & "1"
                    ElseIf UserList(TempCharIndex).faccion.alineacion = eAlineaciones.Real Then
                        Stat = Stat & "8"
                    ElseIf UserList(TempCharIndex).faccion.alineacion = eAlineaciones.Neutro Then
                        Stat = Stat & "9"
                    End If
                End If
                
                
                'ARMADA RANGO CONSEJO, CAOS RANGO CONSEJO,CLAN,DESC
                If UserList(TempCharIndex).faccion.ArmadaReal <> 0 Then
                    Stat = ITS(Stat & 1 & UserList(TempCharIndex).faccion.RecompensasReal) & ITS(UserList(TempCharIndex).Char.charIndex) & "Ç" & UserList(TempCharIndex).desc
                ElseIf UserList(TempCharIndex).faccion.FuerzasCaos <> 0 Then
                    Stat = ITS(Stat & 2 & UserList(TempCharIndex).faccion.RecompensasCaos) & ITS(UserList(TempCharIndex).Char.charIndex) & "Ç" & UserList(TempCharIndex).desc
                Else
                    Stat = ITS(Stat & 8 & 0) & ITS(UserList(TempCharIndex).Char.charIndex) & "Ç" & UserList(TempCharIndex).desc
                End If
                
                If EsNewbie(TempCharIndex) Then Stat = Stat & "1"
                
                EnviarPaquete Paquetes.VeUser, Stat, UserIndex
                FoundSomething = 1
                UserList(UserIndex).flags.TargetUser = TempCharIndex
                UserList(UserIndex).flags.TargetNPC = 0
                UserList(UserIndex).flags.TargetNpcTipo = 0
       
            End If

        End If
    End If
    
    If FoundChar = 2 Then '¿Encontro un NPC?
            Dim estatus As String
          

            If Len(NpcList(TempCharIndex).desc) > 1 Then
                 'Clickeo a un centinela
                 If NpcList(TempCharIndex).Name = "Centinela" And UserList(UserIndex).CentinelaID > 0 Then
                    EnviarPaquete Paquetes.DescNpc2, ITS(NpcList(TempCharIndex).Char.charIndex) & "Hola " & UserList(UserIndex).Name & ", yo soy el centinela anti-macros por favor tipea el comando '/CENTINELA " & Centinelas(UserList(UserIndex).CentinelaID).codigo & "'.", UserIndex
                 Else
                    EnviarPaquete Paquetes.DescNpc2, ITS(NpcList(TempCharIndex).Char.charIndex) & NpcList(TempCharIndex).desc, UserIndex
                 End If
            Else
                If UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 60 Or UserList(UserIndex).flags.Privilegios >= 1 Then
                    estatus = " [" & NpcList(TempCharIndex).Stats.minHP & "/" & NpcList(TempCharIndex).Stats.MaxHP & "]"
                End If
            
               '[Wizard 04/09/05]
                If NpcList(TempCharIndex).MaestroUser > 0 Then
                    EnviarPaquete Paquetes.mensajeinfo, NpcList(TempCharIndex).Name & " es mascota de " & UserList(NpcList(TempCharIndex).MaestroUser).Name & estatus & IIf(NpcList(TempCharIndex).flags.Paralizado = 1, " [Paralizado]", IIf(NpcList(TempCharIndex).flags.Inmovilizado = 1, " [Inmovilizado]", "")), UserIndex
                Else
                   ' Esta luchando contra este
                   If Not (NpcList(TempCharIndex).UserIndexLucha = UserIndex) Then
                        If estaLuchando(NpcList(TempCharIndex)) Then
                        estatus = estatus & "[Luchando con " & UserList(NpcList(TempCharIndex).UserIndexLucha).Name & "]"
                        End If
                   End If
                    EnviarPaquete Paquetes.mensajeinfo, NpcList(TempCharIndex).Name & "." & estatus & IIf(NpcList(TempCharIndex).flags.Paralizado = 1, " [Paralizado]", IIf(NpcList(TempCharIndex).flags.Inmovilizado = 1, " [Inmovilizado]", "")), UserIndex
                End If
            End If
               '[/Wizard]
            FoundSomething = 1
            UserList(UserIndex).flags.TargetNpcTipo = NpcList(TempCharIndex).NPCtype
            UserList(UserIndex).flags.TargetNPC = TempCharIndex
            UserList(UserIndex).flags.TargetUser = 0
            UserList(UserIndex).flags.TargetObj = 0
    End If
    If FoundChar = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
    End If
    '*** NO ENCOTRO NADA ***
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
    End If
Else
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
    End If
End If

End Sub

Function FindDirection(pos As WorldPos, Target As WorldPos) As Byte
'*****************************************************************
'Devuelve la direccion en la cual el target se encuentra
'desde pos, 0 si la direc es igual
'*****************************************************************
Dim x As Integer
Dim y As Integer

x = pos.x - Target.x
y = pos.y - Target.y
'NE
If Sgn(x) = -1 And Sgn(y) = 1 Then
    FindDirection = eHeading.NORTH
    Exit Function
End If
'NW
If Sgn(x) = 1 And Sgn(y) = 1 Then
    FindDirection = eHeading.WEST
    Exit Function
End If
'SW
If Sgn(x) = 1 And Sgn(y) = -1 Then
    FindDirection = eHeading.WEST
    Exit Function
End If
'SE
If Sgn(x) = -1 And Sgn(y) = -1 Then
    FindDirection = eHeading.SOUTH
    Exit Function
End If
'Sur
If Sgn(x) = 0 And Sgn(y) = -1 Then
    FindDirection = eHeading.SOUTH
    Exit Function
End If
'norte
If Sgn(x) = 0 And Sgn(y) = 1 Then
    FindDirection = eHeading.NORTH
    Exit Function
End If
'oeste
If Sgn(x) = 1 And Sgn(y) = 0 Then
    FindDirection = eHeading.WEST
    Exit Function
End If
'este
If Sgn(x) = -1 And Sgn(y) = 0 Then
    FindDirection = eHeading.EAST
    Exit Function
End If
'misma
If Sgn(x) = 0 And Sgn(y) = 0 Then
    FindDirection = 0
    Exit Function
End If
End Function

Sub LookatTileII(ByVal UserIndex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer)
'Responde al click del usuario sobre el mapa
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer

' Esta funcion lo que hace es


' Objetos
'   Posicion donde hizo clic
'   A la derecha
'   A la derecha arriba
'   Abajo
' Personaje
'   Abajo de donde se hizo clic
' Criatura
'
'¿Posicion valida?

If SV_PosicionesValidas.existePosicionMundo(map, x, y) Then
    UserList(UserIndex).flags.TargetMap = map
    UserList(UserIndex).flags.TargetX = x
    UserList(UserIndex).flags.TargetY = y
    '¿Es un obj?
    If MapData(map, x, y).OBJInfo.ObjIndex > 0 Then
        'Informa el nombre
'        Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Name & FONTTYPE_INFO)
'        UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.ObjIndex
        UserList(UserIndex).flags.TargetObjMap = map
        UserList(UserIndex).flags.TargetObjX = x
        UserList(UserIndex).flags.TargetObjY = y
        FoundSomething = 1
    ElseIf MapData(map, x + 1, y).OBJInfo.ObjIndex > 0 Then
        'Informa el nombre
        If ObjData(MapData(map, x + 1, y).OBJInfo.ObjIndex).ObjType = OBJTYPE_PUERTAS Then
'            Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).Name & FONTTYPE_INFO)
'            UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y).OBJInfo.ObjIndex
            UserList(UserIndex).flags.TargetObjMap = map
            UserList(UserIndex).flags.TargetObjX = x + 1
            UserList(UserIndex).flags.TargetObjY = y
            FoundSomething = 1
        End If
    ElseIf MapData(map, x + 1, y + 1).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(map, x + 1, y + 1).OBJInfo.ObjIndex).ObjType = OBJTYPE_PUERTAS Then
            'Informa el nombre
'            Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).Name & FONTTYPE_INFO)
'            UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex
            UserList(UserIndex).flags.TargetObjMap = map
            UserList(UserIndex).flags.TargetObjX = x + 1
            UserList(UserIndex).flags.TargetObjY = y + 1
            FoundSomething = 1
        End If
    ElseIf MapData(map, x, y + 1).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(map, x, y + 1).OBJInfo.ObjIndex).ObjType = OBJTYPE_PUERTAS Then
            'Informa el nombre
'            Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).Name & FONTTYPE_INFO)
'            UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y + 1).OBJInfo.ObjIndex
            UserList(UserIndex).flags.TargetObjMap = map
            UserList(UserIndex).flags.TargetObjX = x
            UserList(UserIndex).flags.TargetObjY = y + 1
            FoundSomething = 1
        End If
    End If
    If FoundSomething = 1 Then
        UserList(UserIndex).flags.TargetObj = MapData(map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex
    End If
    '¿Es un personaje?
    If y + 1 <= Y_MAXIMO_USABLE Then
        If MapData(map, x, y + 1).UserIndex > 0 Then
            TempCharIndex = MapData(map, x, y + 1).UserIndex
            FoundChar = 1
        End If
        If MapData(map, x, y + 1).npcIndex > 0 Then
            TempCharIndex = MapData(map, x, y + 1).npcIndex
            FoundChar = 2
        End If
    End If
    '¿Es un personaje?
    If FoundChar = 0 Then
        If MapData(map, x, y).UserIndex > 0 Then
            TempCharIndex = MapData(map, x, y).UserIndex
            FoundChar = 1
        End If
        If MapData(map, x, y).npcIndex > 0 Then
            TempCharIndex = MapData(map, x, y).npcIndex
            FoundChar = 2
        End If
    End If
    'Reaccion al personaje
    
  If FoundChar = 1 Then '  ¿Encontro un Usuario?
        If UserList(TempCharIndex).flags.AdminInvisible = 0 Or UserList(UserIndex).flags.Privilegios > 1 Then
            If UserList(TempCharIndex).flags.Muerto = 1 Then
                FoundSomething = 1
                UserList(UserIndex).flags.TargetUser = TempCharIndex
                UserList(UserIndex).flags.TargetNPC = 0
                UserList(UserIndex).flags.TargetNpcTipo = 0
            Else
                FoundSomething = 1
                UserList(UserIndex).flags.TargetUser = TempCharIndex
                UserList(UserIndex).flags.TargetNPC = 0
                UserList(UserIndex).flags.TargetNpcTipo = 0
            End If
        End If
  End If
    
    If FoundChar = 2 Then '¿Encontro un NPC?
            FoundSomething = 1
            UserList(UserIndex).flags.TargetNpcTipo = NpcList(TempCharIndex).NPCtype
            UserList(UserIndex).flags.TargetNPC = TempCharIndex
            UserList(UserIndex).flags.TargetUser = 0
            UserList(UserIndex).flags.TargetObj = 0
    End If
    If FoundChar = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
    End If
    '*** NO ENCOTRO NADA ***
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
    End If
Else
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
    End If
End If

End Sub

'Obtiene la posicion más cerca desde la cual podría llegar caminando.
Public Function ObtenerPosicionMasCercana(ByRef personaje As User) As Boolean

'Me fijo si el norte es legal.
If LegalPos(personaje.pos.map, personaje.pos.x, personaje.pos.y - 1, personaje) Then
    personaje.pos.y = personaje.pos.y - 1
    ObtenerPosicionMasCercana = True
    Exit Function
End If

'Me fijo si el Sur es legal.
If LegalPos(personaje.pos.map, personaje.pos.x, personaje.pos.y + 1, personaje) Then
    personaje.pos.y = personaje.pos.y + 1
    ObtenerPosicionMasCercana = True
    Exit Function
End If

'Me fijo si el este es legal.
If LegalPos(personaje.pos.map, personaje.pos.x + 1, personaje.pos.y, personaje) Then
    personaje.pos.x = personaje.pos.x + 1
    ObtenerPosicionMasCercana = True
    Exit Function
End If

'Me fijo si el oeste es legal.
If LegalPos(personaje.pos.map, personaje.pos.x - 1, personaje.pos.y, personaje) Then
    personaje.pos.x = personaje.pos.x - 1
    ObtenerPosicionMasCercana = True
    Exit Function
End If

'Si ninguna es legal
ObtenerPosicionMasCercana = False

End Function

Public Function obtener_hora_fraccion(ByVal fraccion As Byte) As String
    Dim minutos As Integer
    Dim Hora As Integer
    minutos = fraccion - 1
    minutos = minutos * 15
    
    Hora = minutos \ 60
    minutos = minutos Mod 60
    
    obtener_hora_fraccion = IIf(Hora < 10, "0", "") & Hora & ":" & IIf(minutos < 10, "0", "") & minutos
End Function

Public Function obtenerCoordenadas(dato As String) As Variant
    Dim infoCoordenada As Variant
    Dim comienzoInfoCoordenada As Integer
    comienzoInfoCoordenada = InStr(1, dato, "(")
            
    infoCoordenada = Split((mid(dato, comienzoInfoCoordenada + 1, Len(dato) - comienzoInfoCoordenada - 1)), ",")
    
    obtenerCoordenadas = infoCoordenada
End Function
