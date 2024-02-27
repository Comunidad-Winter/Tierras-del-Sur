Attribute VB_Name = "Areas"
Option Explicit

'*** Rangos de actualización***'
Public Const RangoX = 13
Public Const RangoY = 13

Public Sub ActualizarArea(UserIndex As Integer, ByVal movimiento As eHeading)
Dim x As Integer
Dim y As Integer
Dim AddX As Integer
Dim AddY As Integer
Dim QuiX As Integer
Dim QuiY As Integer
Dim AuxIndex As Integer

Select Case movimiento
    Case eHeading.NORTH ' me fijo los cahrs de la linea de arriba
        AddY = -RangoY: QuiY = -RangoY
        AddX = -RangoX: QuiX = RangoX
    Case eHeading.SOUTH 'me fijo los chars de la linea de abajo
        AddY = RangoY: QuiY = RangoY
        AddX = -RangoX: QuiX = RangoX
    Case eHeading.EAST 'me fijo los chars de la linea derecha
        AddY = -RangoY: QuiY = RangoY
        AddX = RangoX: QuiX = RangoX
    Case eHeading.WEST 'me fijo los chars de la izquierda
        AddY = -RangoY: QuiY = RangoY
        AddX = -RangoX: QuiX = -RangoX
End Select
          
    With UserList(UserIndex)
       For y = .pos.y + AddY To .pos.y + QuiY
         If y <= SV_Constantes.Y_MAXIMO_JUGABLE And y >= SV_Constantes.Y_MINIMO_JUGABLE Then
            For x = .pos.x + AddX To .pos.x + QuiX
                If x <= SV_Constantes.X_MAXIMO_JUGABLE And x >= SV_Constantes.X_MINIMO_JUGABLE Then
                    AuxIndex = MapData(.pos.map, x, y).UserIndex
                    
                    If AuxIndex > 0 Then
                            ActualizarChar UserIndex, AuxIndex
                            ActualizarChar AuxIndex, UserIndex
                    ElseIf MapData(.pos.map, x, y).npcIndex > 0 Then
                            ActualizarNpc NpcList(MapData(.pos.map, x, y).npcIndex), UserIndex
                    End If
                End If
            Next
           End If
        Next
    End With
End Sub

Private Sub ActualizarChar(De As Integer, a As Integer)
With UserList(De)
    EnviarPaquete Paquetes.ActualizarAreaUser, ITS(.Char.charIndex) & ITS(.pos.x) & ITS(.pos.y) & Chr$(.Char.heading) & ByteToString(.Char.FX) & ITS(.Char.Body) & ITS(.Char.Head) & ByteToString(.Char.WeaponAnim) & ByteToString(.Char.ShieldAnim) & ByteToString(.Char.CascoAnim), a, ToIndex
End With
Debug.Print "actualizo char"
End Sub

Private Sub ActualizarNpc(De As npc, a As Integer)
EnviarPaquete Paquetes.ActualizarAreaNpc, ITS(De.Char.charIndex) & ITS(De.pos.x) & ITS(De.pos.y) & ByteToString(De.Char.heading), a, ToIndex
End Sub

Public Sub ActualizarTodaArea(UserIndex As Integer)
Dim AuxIndex As Integer
Dim y As Integer
Dim x As Integer

    With UserList(UserIndex)
    
        For y = .pos.y - RangoY To .pos.y + RangoY
         If y <= SV_Constantes.Y_MAXIMO_JUGABLE And y >= SV_Constantes.Y_MINIMO_JUGABLE Then
            For x = .pos.x - RangoX To .pos.x + RangoX
                If x <= SV_Constantes.X_MAXIMO_JUGABLE And x >= SV_Constantes.X_MINIMO_JUGABLE Then
                AuxIndex = MapData(.pos.map, x, y).UserIndex
                    If AuxIndex > 0 Then
                            If Not UserList(MapData(.pos.map, x, y).UserIndex).Name = "" Then
                                ActualizarChar UserIndex, AuxIndex
                                ActualizarChar AuxIndex, UserIndex
                            End If
                            Debug.Print "Actualize Area"
                    ElseIf MapData(.pos.map, x, y).npcIndex > 0 Then
                            ActualizarNpc NpcList(MapData(.pos.map, x, y).npcIndex), UserIndex
                            Debug.Print "Actualize Area"
                    End If
                End If
            Next
           End If
        Next
    End With

End Sub

Public Function estaEnArea(ByRef personaje As User, ByRef personaje2 As User) As Boolean
    If personaje.pos.map <> personaje2.pos.map Then
        estaEnArea = False
        Exit Function
    End If
    
    If Abs(personaje.pos.x - personaje2.pos.x) > BORDE_TILES_INUTILIZABLE Or Abs(personaje.pos.y - personaje2.pos.y) > BORDE_TILES_INUTILIZABLE Then
        estaEnArea = False
        Exit Function
    End If
    
    estaEnArea = True
End Function

