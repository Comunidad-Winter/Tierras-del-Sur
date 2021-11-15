Attribute VB_Name = "Acciones"

'Argentum Online 0.9.0.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez
Option Explicit



Sub Accion(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error Resume Next

'�Posicion valida?
If InMapBounds(Map, X, Y) Then
   
    Dim FoundChar As Byte
    Dim FoundSomething As Byte
    Dim TempCharIndex As Integer
       
    '�Es un obj?
    If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.ObjIndex
        
        Select Case ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).ObjType
            
            Case OBJTYPE_PUERTAS 'Es una puerta
                Call AccionParaPuerta(Map, X, Y, UserIndex)
            Case OBJTYPE_CARTELES 'Es un cartel
                Call AccionParaCartel(Map, X, Y, UserIndex)
            Case OBJTYPE_FOROS 'Foro
                Call AccionParaForo(Map, X, Y, UserIndex)
            Case OBJTYPE_LE�A 'Le�a
                If MapData(Map, X, Y).OBJInfo.ObjIndex = FOGATA_APAG Then
                    Call AccionParaRamita(Map, X, Y, UserIndex)
                End If
            
        End Select
    '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
    ElseIf MapData(Map, X + 1, Y).OBJInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y).OBJInfo.ObjIndex
        Call SendData(ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).ObjType & "," & ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).Name & "," & "OBJ")
        Select Case ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).ObjType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, X + 1, Y, UserIndex)
            
        End Select
    ElseIf MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex
        Call SendData(ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).ObjType & "," & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).Name & "," & "OBJ")
        Select Case ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).ObjType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)
            
        End Select
    ElseIf MapData(Map, X, Y + 1).OBJInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y + 1).OBJInfo.ObjIndex
        Call SendData(ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).ObjType & "," & ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).Name & "," & "OBJ")
        Select Case ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).ObjType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, X, Y + 1, UserIndex)
            
        End Select
        
    ElseIf MapData(Map, X, Y).NpcIndex > 0 Then
        If Npclist(MapData(Map, X, Y).NpcIndex).Comercia = 1 Then
              If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 3 Then
                  Call SendData(ToIndex, UserIndex, 0, "Y7")
                  Exit Sub
              End If
              'Iniciamos la rutina pa' comerciar.
              Call IniciarCOmercioNPC(UserIndex)
'<Gorlok> - 2005-03-28
        ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = 4 Then 'NPC Banquero
            'Comando /BOVEDA
            '�Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "Y3")
                Exit Sub
            End If
            '�El target es un NPC valido?
            If Distancia(Npclist(MapData(Map, X, Y).NpcIndex).Pos, UserList(UserIndex).Pos) > 3 Then
                Call SendData(ToIndex, UserIndex, 0, "Y7")
                Exit Sub
            End If
            Call IniciarDeposito(UserIndex)
        ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = 1 Then 'NPC Sacerdote
            If UserList(UserIndex).flags.Muerto = 1 Then 'El usuario est� muerto
                'Comando /RESUCITAR
                If Distancia(UserList(UserIndex).Pos, Npclist(MapData(Map, X, Y).NpcIndex).Pos) > 10 Then
                    Call SendData(ToIndex, UserIndex, 0, "Y8")
                    Exit Sub
                End If
                Call RevivirUsuario(UserIndex)
                Call SendData(ToIndex, UserIndex, 0, "||��H�s sido resucitado!!" & FONTTYPE_INFO)
                Exit Sub
            Else
                'Comando /CURAR
                If Distancia(UserList(UserIndex).Pos, Npclist(MapData(Map, X, Y).NpcIndex).Pos) > 10 Then
                    Call SendData(ToIndex, UserIndex, 0, "Y8")
                    Exit Sub
                End If
                If (UserList(UserIndex).Stats.MaxHP <= UserList(UserIndex).Stats.MinHP) Then
                    Call SendData(ToIndex, UserIndex, 0, "||No est�s herido." & FONTTYPE_INFO)
                    Exit Sub
                End If
                UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
                Call SendUserStatsBox(val(UserIndex))
                Call SendData(ToIndex, UserIndex, 0, "Y17")
                Exit Sub
            End If
        End If
'</Gorlok> - 2005-03-28
        
    Else
        UserList(UserIndex).flags.TargetNpc = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
    End If
    
End If

End Sub

Sub AccionParaRamita(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj
Dim raise As Integer

Dim Pos As WorldPos
Pos.Map = Map
Pos.X = X
Pos.Y = Y

If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
    Call SendData(ToIndex, UserIndex, 0, "Y5")
    Exit Sub
End If


If UserList(UserIndex).Stats.UserSkills(Supervivencia) > 1 And UserList(UserIndex).Stats.UserSkills(Supervivencia) < 6 Then
            Suerte = 3
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(Supervivencia) <= 10 Then
            Suerte = 2
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 10 And UserList(UserIndex).Stats.UserSkills(Supervivencia) Then
            Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    Obj.ObjIndex = FOGATA
    Obj.Amount = 1
    
    Call SendData(ToIndex, UserIndex, 0, "Y28")
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "FO")
    
    Call MakeObj(ToMap, 0, Map, Obj, Map, X, Y)
    
    
Else
    Call SendData(ToIndex, UserIndex, 0, "Y29")
End If

'Sino tiene hambre o sed quizas suba el skill supervivencia
If UserList(UserIndex).flags.Hambre = 0 And UserList(UserIndex).flags.Sed = 0 Then
    Call SubirSkill(UserIndex, Supervivencia)
End If

End Sub

Sub AccionParaForo(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim Pos As WorldPos
Pos.Map = Map
Pos.X = X
Pos.Y = Y

If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
    Call SendData(ToIndex, UserIndex, 0, "Y5")
    Exit Sub
End If

'�Hay mensajes?
Dim f As String, tit As String, men As String, base As String, auxcad As String
f = App.Path & "\foros\" & UCase$(ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).ForoID) & ".for"
If FileExist(f, vbNormal) Then
    Dim num As Integer
    num = val(GetVar(f, "INFO", "CantMSG"))
    base = Left$(f, Len(f) - 4)
    Dim i As Integer
    Dim N As Integer
    For i = 1 To num
        N = FreeFile
        f = base & i & ".for"
        Open f For Input Shared As #N
        Input #N, tit
        men = ""
        auxcad = ""
        Do While Not EOF(N)
            Input #N, auxcad
            men = men & vbCrLf & auxcad
        Loop
        Close #N
        Call SendData(ToIndex, UserIndex, 0, "FMSG" & tit & Chr(176) & men)
        
    Next
End If
Call SendData(ToIndex, UserIndex, 0, "MFOR")
End Sub


Sub AccionParaPuerta(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim MiObj As Obj
Dim wp As WorldPos

If Not (Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2) Then
    If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Llave = 0 Then
        If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Cerrada = 1 Then
                'Abre la puerta
                If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Llave = 0 Then
                          
                     MapData(Map, X, Y).OBJInfo.ObjIndex = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).IndexAbierta
                                  
                     Call MakeObj(ToMap, 0, Map, MapData(Map, X, Y).OBJInfo, Map, X, Y)
                     
                     'Desbloquea
                     MapData(Map, X, Y).Blocked = 0
                     MapData(Map, X - 1, Y).Blocked = 0
                     
                     'Bloquea todos los mapas
                     Call Bloquear(ToMap, 0, Map, Map, X, Y, 0)
                     Call Bloquear(ToMap, 0, Map, Map, X - 1, Y, 0)
                     
                       
                     'Sonido
                     SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_PUERTA
                    
                Else
                     Call SendData(ToIndex, UserIndex, 0, "Y20")
                End If
        Else
                'Cierra puerta
                MapData(Map, X, Y).OBJInfo.ObjIndex = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).IndexCerrada
                
                Call MakeObj(ToMap, 0, Map, MapData(Map, X, Y).OBJInfo, Map, X, Y)
                
                
                MapData(Map, X, Y).Blocked = 1
                MapData(Map, X - 1, Y).Blocked = 1
                
                
                Call Bloquear(ToMap, 0, Map, Map, X - 1, Y, 1)
                Call Bloquear(ToMap, 0, Map, Map, X, Y, 1)
                
                SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_PUERTA
        End If
        
        UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.ObjIndex
    Else
        Call SendData(ToIndex, UserIndex, 0, "Y20")
    End If
Else
    Call SendData(ToIndex, UserIndex, 0, "Y5")
End If

End Sub

Sub AccionParaCartel(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next


Dim MiObj As Obj

If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).ObjType = 8 Then
  
  If Len(ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).texto) > 0 Then
       Call SendData(ToIndex, UserIndex, 0, "MCAR" & _
        ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).texto & _
        Chr(176) & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).GrhSecundario)
  End If
  
End If

End Sub

