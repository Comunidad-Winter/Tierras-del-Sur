VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IControl

Private position As position
Private size As position

Private textura As Integer

Private originOfssetX As Integer
Private originOffsetY As Integer
Private originWidth As Integer
Private originHeight As Integer

Private textureWidth As Integer
Private textureHeight As Integer

Public Sub setOrigin(offsetX As Integer, offsetY As Integer, width As Integer, Height As Integer)
    originOfssetX = offsetX
    originOffsetY = offsetY
    originWidth = width
    originHeight = Height
End Sub

Public Sub setTextura(textureNumber As Integer, width As Integer, Height As Integer)
    textura = textureNumber
    textureWidth = width
    textureHeight = Height
End Sub
Public Sub SetPos(X As Integer, Y As Integer)
    position.X = X
    position.Y = Y
End Sub

Private Function IControl_Click(ByVal X As Integer, ByVal Y As Integer) As Boolean
End Function

Private Property Get IControl_IsFocuseable() As Boolean
    IControl_IsFocuseable = False
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

Private Function IControl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

End Function

Private Function IControl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal z As Integer) As Boolean

End Function

Private Function IControl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

End Function

Private Sub IControl_Render(ByVal focus As Boolean)

If textura > 0 Then
    Call Grh_Render_Simple_rect(textura, originOfssetX, originOffsetY, originWidth, originHeight, position.X, position.Y, -1, textureWidth, textureHeight, 0, size.X, size.Y)
End If
End Sub

Private Sub IControl_SetPos(ByVal X As Integer, ByVal Y As Integer)
    SetPos X, Y
End Sub

Public Sub IControl_SetSize(ByVal ancho As Integer, ByVal alto As Integer)
    size.X = ancho
    size.Y = alto
End Sub

Public Property Get IControl_Width() As Integer
    IControl_Width = size.X
End Property

Public Property Get IControl_Height() As Integer
    IControl_Height = size.Y
End Property

Public Property Get IControl_X() As Integer
    IControl_X = position.X
End Property

Public Property Get IControl_Y() As Integer
    IControl_Y = position.Y
End Property

