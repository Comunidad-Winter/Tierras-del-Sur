Attribute VB_Name = "CLI_Areas"
'       __________________
'      / __________  ____ \
'     / |_   _|    \/  __\ \
'    /    | | | |\  \ `--.  \
'   /     | | | | | |`--. \  \
'  /      | | | |/  /\__/ /   \
'  \      |_| |____/\____/    /
'   \________________________/



Option Explicit

Public Const ARangoX = 14
Public Const ARangoY = 14

Sub BorrarAreaB()
    Dim X As Integer, Y As Integer

    For X = UserPos.X - ARangoX To UserPos.X + ARangoX
        If X >= X_MINIMO_VISIBLE And X <= X_MAXIMO_VISIBLE Then
            For Y = UserPos.Y - ARangoY To UserPos.Y + ARangoY
                If Y >= Y_MINIMO_VISIBLE And Y <= Y_MAXIMO_VISIBLE Then
                    If CharMap(X, Y) > 0 Then
                        DeactivateChar CharList(CharMap(X, Y))
                    End If
                End If
            Next Y
        End If
    Next X

End Sub

Sub BorrarB(ByVal Movimiento As E_Heading)

    Dim MinX%, MinY%, MaxX%, MaxY%, posx As position
    
    posx = CharList(UserCharIndex).Pos
    
    MinX = maxl(-ARangoX + posx.X, X_MINIMO_VISIBLE)
    MinY = maxl(-ARangoY + posx.Y, Y_MINIMO_VISIBLE)
    MaxX = minl(ARangoX + posx.X, X_MAXIMO_VISIBLE)
    MaxY = minl(ARangoY + posx.Y, Y_MAXIMO_VISIBLE)
    
    
    
    Dim i As Long
    Select Case Movimiento
    Case SOUTH
        'parte superior
     '   Debug.Print "Elimino linea superior " & MinY
        For i = MinX To MaxX
            If CharMap(i, MinY) > 0 Then
                Call Dialogos.RemoveDialog(CharMap(i, MinY))
                Call DeactivateChar(CharList(CharMap(i, MinY)))
            End If
        Next
    Case NORTH
        'linea inferior
       ' Debug.Print "Elimino linea inferior " & MaxY
        For i = MinX To MaxX
            If CharMap(i, MaxY) > 0 Then
                Call Dialogos.RemoveDialog(CharMap(i, MaxY))
                DeactivateChar CharList(CharMap(i, MaxY))
                'CharMap(i, MaxY) = 0
            End If
        Next
    Case EAST
        'Linea izquierda
        For i = MinY To MaxY
            If CharMap(MinX, i) > 0 Then
                Call Dialogos.RemoveDialog(CharMap(MinX, i))
                DeactivateChar CharList(CharMap(MinX, i))
                'CharMap(MinX, i) = 0
            End If
        Next
    Case WEST
        'Linea derecha
        For i = MinY To MaxY
            If CharMap(MaxX, i) > 0 Then
                Call Dialogos.RemoveDialog(CharMap(MaxX, i))
                DeactivateChar CharList(CharMap(MaxX, i))
            End If
        Next
    End Select

End Sub
