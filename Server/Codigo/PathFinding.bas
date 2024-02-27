Attribute VB_Name = "PathFinding"
Option Explicit

Private Const ROWS = 100
Private Const COLUMS = 100
Private Const MAXINT = 1000

Private Type tIntermidiateWork
    Known As Boolean
    DistV As Integer
    PrevV As tVertice
End Type

Dim TmpArray(1 To ROWS, 1 To COLUMS) As tIntermidiateWork

Private Function Limites(ByVal vfila As Integer, ByVal vcolu As Integer)
Limites = vcolu >= 1 And vcolu <= COLUMS And vfila >= 1 And vfila <= ROWS
End Function

Private Function IsWalkable(ByVal map As Integer, ByVal row As Integer, ByVal Col As Integer, ByVal npcIndex As Integer) As Boolean
' TODO
'IsWalkable = MapData(map, row, Col).Blocked = 0 And MapData(map, row, Col).npcIndex = 0
'If MapData(map, row, Col).UserIndex <> 0 Then
'     If MapData(map, row, Col).UserIndex <> NpcList(npcIndex).PFINFO.TargetUser Then IsWalkable = False
'End If
End Function

Private Sub ProcessAdjacents(ByVal MapIndex As Integer, ByRef t() As tIntermidiateWork, ByRef vfila As Integer, ByRef vcolu As Integer, ByVal npcIndex As Integer)
    Dim V As tVertice
    Dim j As Integer
    'Look to eHeading.NORTH
    j = vfila - 1
    If Limites(j, vcolu) Then
            If IsWalkable(MapIndex, j, vcolu, npcIndex) Then
                    'Nos aseguramos que no hay un camino más corto
                    If t(j, vcolu).DistV = MAXINT Then
                        'Actualizamos la tabla de calculos intermedios
                        t(j, vcolu).DistV = t(vfila, vcolu).DistV + 1
                        t(j, vcolu).PrevV.x = vcolu
                        t(j, vcolu).PrevV.y = vfila
                        'Mete el vertice en la cola
                        V.x = vcolu
                        V.y = j
                        Call Push(V)
                    End If
            End If
    End If
    j = vfila + 1
    'look to eHeading.SOUTH
    If Limites(j, vcolu) Then
            If IsWalkable(MapIndex, j, vcolu, npcIndex) Then
                'Nos aseguramos que no hay un camino más corto
                If t(j, vcolu).DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
                    t(j, vcolu).DistV = t(vfila, vcolu).DistV + 1
                    t(j, vcolu).PrevV.x = vcolu
                    t(j, vcolu).PrevV.y = vfila
                    'Mete el vertice en la cola
                    V.x = vcolu
                    V.y = j
                    Call Push(V)
                End If
            End If
    End If
    'look to eHeading.WEST
    If Limites(vfila, vcolu - 1) Then
            If IsWalkable(MapIndex, vfila, vcolu - 1, npcIndex) Then
                'Nos aseguramos que no hay un camino más corto
                If t(vfila, vcolu - 1).DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
                    t(vfila, vcolu - 1).DistV = t(vfila, vcolu).DistV + 1
                    t(vfila, vcolu - 1).PrevV.x = vcolu
                    t(vfila, vcolu - 1).PrevV.y = vfila
                    'Mete el vertice en la cola
                    V.x = vcolu - 1
                    V.y = vfila
                    Call Push(V)
                End If
            End If
    End If
    'look to eHeading.EAST
    If Limites(vfila, vcolu + 1) Then
            If IsWalkable(MapIndex, vfila, vcolu + 1, npcIndex) Then
                'Nos aseguramos que no hay un camino más corto
                If t(vfila, vcolu + 1).DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
                    t(vfila, vcolu + 1).DistV = t(vfila, vcolu).DistV + 1
                    t(vfila, vcolu + 1).PrevV.x = vcolu
                    t(vfila, vcolu + 1).PrevV.y = vfila
                    'Mete el vertice en la cola
                    V.x = vcolu + 1
                    V.y = vfila
                    Call Push(V)
                End If
            End If
    End If
End Sub

Public Sub SeekPath(ByVal npcIndex As Integer, Optional ByVal MaxSteps As Integer = 30)
'############################################################
'This Sub seeks a path from the npclist(npcindex).pos
'to the location NPCList(NpcIndex).PFINFO.Target.
'The optional parameter MaxSteps is the maximum of steps
'allowed for the path.
'############################################################
Dim cur_npc_pos As tVertice
Dim tar_npc_pos As tVertice
Dim V As tVertice
Dim NpcMap As Integer
Dim steps As Integer
NpcMap = NpcList(npcIndex).pos.map
steps = 0
cur_npc_pos.x = NpcList(npcIndex).pos.y
cur_npc_pos.y = NpcList(npcIndex).pos.x
tar_npc_pos.x = NpcList(npcIndex).PFINFO.Target.x '  UserList(NPCList(NpcIndex).PFINFO.TargetUser).Pos.X
tar_npc_pos.y = NpcList(npcIndex).PFINFO.Target.y '  UserList(NPCList(NpcIndex).PFINFO.TargetUser).Pos.Y
Call InitializeTable(TmpArray, cur_npc_pos)
Call InitQueue
'We add the first vertex to the Queue
Call Push(cur_npc_pos)
Do While (Not IsEmpty)
    If steps > MaxSteps Then Exit Do
    V = Pop
    If V.x = tar_npc_pos.x And V.y = tar_npc_pos.y Then Exit Do
    Call ProcessAdjacents(NpcMap, TmpArray, V.y, V.x, npcIndex)
Loop
Call MakePath(npcIndex)
End Sub

Private Sub MakePath(ByVal npcIndex As Integer)
'#######################################################
'Builds the path previously calculated
'#######################################################
Dim Pasos As Integer
Dim miV As tVertice
Dim i As Integer
Pasos = TmpArray(NpcList(npcIndex).PFINFO.Target.y, NpcList(npcIndex).PFINFO.Target.x).DistV
NpcList(npcIndex).PFINFO.PathLenght = Pasos
If Pasos = MAXINT Then
    'MsgBox "There is no path."
    NpcList(npcIndex).PFINFO.NoPath = True
    NpcList(npcIndex).PFINFO.PathLenght = 0
    Exit Sub
End If
ReDim NpcList(npcIndex).PFINFO.Path(0 To Pasos) As tVertice
miV.x = NpcList(npcIndex).PFINFO.Target.x
miV.y = NpcList(npcIndex).PFINFO.Target.y
For i = Pasos To 1 Step -1
    NpcList(npcIndex).PFINFO.Path(i) = miV
    miV = TmpArray(miV.y, miV.x).PrevV
Next i
NpcList(npcIndex).PFINFO.CurPos = 1
NpcList(npcIndex).PFINFO.NoPath = False
End Sub

Private Sub InitializeTable(ByRef t() As tIntermidiateWork, ByRef s As tVertice, Optional ByVal MaxSteps As Integer = 30)
'#########################################################
'Initialize the array where we calculate the path
'#########################################################
Dim j As Integer, k As Integer
Const anymap = 1
For j = s.y - MaxSteps To s.y + MaxSteps
    For k = s.x - MaxSteps To s.x + MaxSteps
        If SV_PosicionesValidas.existePosicionMundo(anymap, j, k) Then
            t(j, k).Known = False
            t(j, k).DistV = MAXINT
            t(j, k).PrevV.x = 0
            t(j, k).PrevV.y = 0
        End If
    Next
Next
t(s.y, s.x).Known = False
t(s.y, s.x).DistV = 0
End Sub
