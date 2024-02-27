Attribute VB_Name = "AI"
Option Explicit

Public Const ESTATICO = 1
Public Const MUEVE_AL_AZAR = 2
Public Const NPC_MALO_ATACA_USUARIOS_BUENOS = 3
Public Const NPCDEFENSA = 4
'Public Const GUARDIAS_ATACAN_CRIMINALES = 5 Se incluye en NPC_ATACA
Public Const SIGUE_AMO = 8
'Public Const NPC_ATACA_NPC = 9 Se incluye en el npc defensa
Public Const NPC_PATHFINDING = 10

Public Const ELEMENTALFUEGO = 93
Public Const ELEMENTALTIERRA = 94
Public Const ELEMENTALAGUA = 92

Public Const ESPIRITU_INDOMABLE = 110
Public Const FUEGO_FACTUO = 11

'---------------------------------------------------------------------------------------
' Procedure : RestoreOldMovement
' DateTime  : 18/02/2007 19:05
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub RestoreOldMovement(ByVal npcIndex As Integer)

If NpcList(npcIndex).MaestroUser = 0 Then
    
    'If Npclist(NpcIndex).Movement = ESTATICO And Npclist(NpcIndex).flags.OldMovement <> ESTATICO Then
     '   Call quitarEstatico(NpcIndex)
    'End If
    
    NpcList(npcIndex).Movement = NpcList(npcIndex).flags.OldMovement
    NpcList(npcIndex).TargetUserID = 0
End If

End Sub

Public Sub NPCAI(ByVal npcIndex As Integer)

    Select Case NpcList(npcIndex).Movement
    
        Case MUEVE_AL_AZAR
            
            If NpcList(npcIndex).flags.Inmovilizado = 1 Then Exit Sub
            
            Call IA_BASICA.moverAlAzar(NpcList(npcIndex), npcIndex)
            
        Case NPC_MALO_ATACA_USUARIOS_BUENOS
        
            Call IA_BASICA.inteligenciaBasica(NpcList(npcIndex), npcIndex)
            
        Case NPCDEFENSA
        
            Call IA_BASICA.inteligenciaBasica_Seguir_Agresor(NpcList(npcIndex), npcIndex)
                           
        Case SIGUE_AMO
        
            If NpcList(npcIndex).flags.Inmovilizado = 1 Then Exit Sub
            
            Call IA_BASICA.inteligenciaBasica_Seguir_Amo(NpcList(npcIndex), npcIndex)
            
        Case NPC_PATHFINDING
        
            If NpcList(npcIndex).flags.Inmovilizado = 1 Then Exit Sub
            
            If ReCalculatePath(npcIndex) Then
                Call PathFindingAI(npcIndex)
                'Existe el camino?
                If NpcList(npcIndex).PFINFO.NoPath Then 'Si no existe nos movemos al azar
                    'Move randomly
                    Call MoveNPCChar(npcIndex, Int(RandomNumber(1, 4)))
                End If
            Else
                If Not PathEnd(npcIndex) Then
                    Call FollowPath(npcIndex)
                Else
                    NpcList(npcIndex).PFINFO.PathLenght = 0
                End If
            End If
            
        Case Else
            Debug.Print "No tiene movement valido " & NpcList(npcIndex).Name
    End Select

End Sub

'---------------------------------------------------------------------------------------
' Procedure : UserNear
' DateTime  : 18/02/2007 19:06
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function UserNear(ByVal npcIndex As Integer) As Boolean
'#################################################################
'Returns True if there is an user adjacent to the npc position.
'#################################################################

UserNear = Not Int(Distance(NpcList(npcIndex).pos.x, NpcList(npcIndex).pos.y, UserList(NpcList(npcIndex).PFINFO.TargetUser).pos.x, UserList(NpcList(npcIndex).PFINFO.TargetUser).pos.y)) > 1

End Function

'---------------------------------------------------------------------------------------
' Procedure : ReCalculatePath
' DateTime  : 18/02/2007 19:06
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function ReCalculatePath(ByVal npcIndex As Integer) As Boolean
'#################################################################
'Returns true if we have to seek a new path
'#################################################################
If NpcList(npcIndex).PFINFO.PathLenght = 0 Then
    ReCalculatePath = True
ElseIf Not UserNear(npcIndex) And NpcList(npcIndex).PFINFO.PathLenght = NpcList(npcIndex).PFINFO.CurPos - 1 Then
    ReCalculatePath = True
End If
End Function
'---------------------------------------------------------------------------------------
' Procedure : PathEnd
' DateTime  : 18/02/2007 19:07
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function PathEnd(ByVal npcIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock
'Returns if the npc has arrived to the end of its path
'#################################################################
PathEnd = NpcList(npcIndex).PFINFO.CurPos = NpcList(npcIndex).PFINFO.PathLenght
End Function

'---------------------------------------------------------------------------------------
' Procedure : FollowPath
' DateTime  : 18/02/2007 19:07
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function FollowPath(ByVal npcIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock
'Moves the npc.
'#################################################################
Dim tmpPos As WorldPos
Dim tHeading As Byte

tmpPos.map = NpcList(npcIndex).pos.map
tmpPos.x = NpcList(npcIndex).PFINFO.Path(NpcList(npcIndex).PFINFO.CurPos).y ' invertí las coordenadas
tmpPos.y = NpcList(npcIndex).PFINFO.Path(NpcList(npcIndex).PFINFO.CurPos).x
'Debug.Print "(" & tmpPos.X & "," & tmpPos.Y & ")"
tHeading = FindDirection(NpcList(npcIndex).pos, tmpPos)
MoveNPCChar npcIndex, tHeading
NpcList(npcIndex).PFINFO.CurPos = NpcList(npcIndex).PFINFO.CurPos + 1

End Function

'---------------------------------------------------------------------------------------
' Procedure : PathFindingAI
' DateTime  : 18/02/2007 19:07
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function PathFindingAI(ByVal npcIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock / 11-07-02
'www.geocities.com/gmorgolock
'morgolock@speedy.com.ar
'This function seeks the shortest path from the Npc
'to the user's location.
'#################################################################
Dim y As Integer
Dim x As Integer

For y = NpcList(npcIndex).pos.y - 10 To NpcList(npcIndex).pos.y + 10    'Makes a loop that looks at
     For x = NpcList(npcIndex).pos.x - 10 To NpcList(npcIndex).pos.x + 10   '5 tiles in every direction
         'Make sure tile is legal
         If x >= SV_Constantes.X_MINIMO_JUGABLE And x <= SV_Constantes.X_MAXIMO_JUGABLE And y >= SV_Constantes.Y_MINIMO_JUGABLE And y <= SV_Constantes.Y_MAXIMO_JUGABLE Then
             'look for a user
             If MapData(NpcList(npcIndex).pos.map, x, y).UserIndex > 0 Then
                 'Move towards user
                  Dim tmpUserIndex As Integer
                  tmpUserIndex = MapData(NpcList(npcIndex).pos.map, x, y).UserIndex
                  If UserList(tmpUserIndex).flags.Muerto = 0 And UserList(tmpUserIndex).flags.Invisible = 0 And UserList(tmpUserIndex).flags.Mimetizado = 0 Then
                    'We have to invert the coordinates, this is because
                    'ORE refers to maps in converse way of my pathfinding
                    'routines.
                    NpcList(npcIndex).PFINFO.Target.x = UserList(tmpUserIndex).pos.y
                    NpcList(npcIndex).PFINFO.Target.y = UserList(tmpUserIndex).pos.x 'ops!
                    NpcList(npcIndex).PFINFO.TargetUser = tmpUserIndex
                    Call SeekPath(npcIndex)
                    Exit Function
                  End If
             End If
         End If
     Next x
 Next y

End Function
