VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlMouseReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IControl

Private cW%, cH%, CX%, cY%

Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer)
Event Click(ByVal x As Integer, ByVal y As Integer)

Private pColor As Long

Private Sub Class_Initialize()
pColor = &H7F222222
End Sub

Private Function IControl_Click(ByVal x As Integer, ByVal y As Integer) As Boolean
    IControl_Click = Colision(x, y, cY, cY + cH, CX, CX + cW)
    If IControl_Click Then
        RaiseEvent Click(x, y)
    End If
End Function

Private Property Get IControl_IsFocuseable() As Boolean

End Property

Private Property Get IControl_IsMouseFocus() As Boolean
    IControl_IsMouseFocus = False
End Property

Private Function IControl_KeyDown(KeyCode As Integer, Shift As Integer) As Boolean

End Function

Private Function IControl_KeyPress(KeyAscii As Integer) As Boolean

End Function

Private Function IControl_KeyUp(KeyCode As Integer, Shift As Integer) As Boolean

End Function

Private Function IControl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    IControl_MouseDown = Colision(x, y, cY, cY + cH, CX, CX + cW)
    If IControl_MouseDown Then
        RaiseEvent MouseDown(Button, Shift, x - CX, y - cY)
    End If
End Function

Private Function IControl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer) As Boolean
    IControl_MouseMove = Colision(x, y, cY, cY + cH, CX, CX + cW)
    If IControl_MouseMove Then
        RaiseEvent MouseMove(Button, Shift, x - CX, y - cY, z)
    End If
End Function

Private Function IControl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    IControl_MouseUp = Colision(x, y, cY, cY + cH, CX, CX + cW)
    If IControl_MouseUp Then
        RaiseEvent MouseUp(Button, Shift, x - CX, y - cY)
    End If
End Function

Private Sub IControl_Render(ByVal focus As Boolean)
    Render
End Sub

Private Sub IControl_SetPos(ByVal x As Integer, ByVal y As Integer)
    SetPos x, y
End Sub

Private Sub IControl_SetSize(ByVal ancho As Integer, ByVal alto As Integer)
    SetSize alto, ancho
End Sub

Public Sub SetPos(ByVal x As Integer, ByVal y As Integer)
    CX = x
    cY = y
End Sub

Public Sub SetSize(ByVal ancho As Integer, ByVal alto As Integer)
    cW = ancho
    cH = alto
End Sub

Public Sub SetColor(Color As Long)
    pColor = Color
End Sub

Public Sub Render()
    Grh_Render_Solid pColor, CX, cY, cW, cH
End Sub
