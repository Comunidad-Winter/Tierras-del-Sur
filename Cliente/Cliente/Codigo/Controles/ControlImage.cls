VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IControl

Private position As position
Private size As position

Public GrhId As Integer
Private recortar As Boolean

Public Sub SetPos(X As Integer, Y As Integer)
    position.X = X
    position.Y = Y
End Sub

Public Sub SetRecortar(recortarExcendente As Boolean)
    recortar = recortarExcendente
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
If GrhId > 0 Then
    Call Engine_GrhDraw.Grh_Render_size(GrhId, position.X, position.Y, 0, -1, 0, size.X, size.Y, recortar)
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
    If size.X > 0 Then
        IControl_Width = size.X
    Else
        IControl_Width = GrhData(GrhId).pixelWidth
    End If
End Property

Public Property Get IControl_Height() As Integer
    If size.Y > 0 Then
        IControl_Height = size.Y
    Else
        IControl_Height = GrhData(GrhId).pixelHeight
    End If
End Property

Public Property Get IControl_X() As Integer
    IControl_X = position.X
End Property

Public Property Get IControl_Y() As Integer
    IControl_Y = position.Y
End Property
