VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlVentana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IControl

Private cW%, cH%, cX%, cY%

Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal y As Integer)
Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal y As Integer)
Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal y As Integer, ByVal z As Integer)
Event KeyPress(ByVal KeyAscii As Integer, ContinuarPropagacion As Boolean)
Event KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ContinuarPropagacion As Boolean)
Event KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer, ContinuarPropagacion As Boolean)

Event Click(ByVal X As Integer, ByVal y As Integer)

Private pTitulo As String
Private pColor As Long

Private pventanaHost As vWindow

Private Sub Class_Initialize()
pColor = &H7F222222
End Sub

Private Function IControl_Click(ByVal X As Integer, ByVal y As Integer) As Boolean
    IControl_Click = Colision(X, y, cY, cY + cH, cX, cX + cW)
    If IControl_Click Then
        RaiseEvent Click(X, y)
    End If
End Function

Private Function IControl_KeyDown(KeyCode As Integer, Shift As Integer) As Boolean
    Dim ContinuarPropagacion As Boolean
    ContinuarPropagacion = True
    RaiseEvent KeyDown(KeyCode, Shift, ContinuarPropagacion)
    IControl_KeyDown = Not ContinuarPropagacion
End Function

Private Function IControl_KeyPress(KeyAscii As Integer) As Boolean
    Dim ContinuarPropagacion As Boolean
    ContinuarPropagacion = True
    RaiseEvent KeyPress(KeyAscii, ContinuarPropagacion)
    IControl_KeyPress = Not ContinuarPropagacion
End Function

Private Function IControl_KeyUp(KeyCode As Integer, Shift As Integer) As Boolean
    Dim ContinuarPropagacion As Boolean
    ContinuarPropagacion = True
    RaiseEvent KeyUp(KeyCode, Shift, ContinuarPropagacion)
    IControl_KeyUp = Not ContinuarPropagacion
End Function

Private Function IControl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal y As Integer) As Boolean
    IControl_MouseDown = Colision(X, y, cY, cY + cH, cX, cX + cW)
    If IControl_MouseDown Then
        RaiseEvent MouseDown(Button, Shift, X - cX, y - cY)
    End If
End Function

Private Function IControl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal y As Integer, ByVal z As Integer) As Boolean
    IControl_MouseMove = Colision(X, y, cY, cY + cH, cX, cX + cW)
    If IControl_MouseMove Then
        RaiseEvent MouseMove(Button, Shift, X - cX, y - cY, z)
    End If
End Function

Private Function IControl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal y As Integer) As Boolean
    IControl_MouseUp = Colision(X, y, cY, cY + cH, cX, cX + cW)
    If IControl_MouseUp Then
        RaiseEvent MouseUp(Button, Shift, X - cX, y - cY)
    End If
End Function

Private Sub IControl_Render(ByVal focus As Boolean)
    Render
End Sub

Private Sub IControl_SetPos(ByVal X As Integer, ByVal y As Integer)
    SetPos X, y
End Sub

Private Sub IControl_SetSize(ByVal ancho As Integer, ByVal alto As Integer)
    SetSize alto, ancho
End Sub

Public Sub SetPos(ByVal X As Integer, ByVal y As Integer)
    cX = X
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
    Grh_Render_Solid pColor, cX, cY, cW, cH
    GUI_RenderDialog cX, cY, cW, cH, pTitulo, pventanaHost
End Sub


Public Property Get Titulo() As String
Titulo = pTitulo
End Property

Public Property Let Titulo(ByVal vNewValue As String)
pTitulo = vNewValue
End Property

Public Property Get ventanaHost() As vWindow
    Set ventanaHost = pventanaHost
End Property

Public Property Let ventanaHost(ByVal vNewValue As vWindow)
    Set pventanaHost = vNewValue
End Property
