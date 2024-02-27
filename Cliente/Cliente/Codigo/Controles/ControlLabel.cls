VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IControl

Option Explicit

Private CX%, cY%, cH%, cW%

Private ControlRect As RECT

Private tBox As Box_Vertex

Private cLabel As clsGuiText2

Private cEnabled As Boolean

Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer)
Public Event KeyPress(ByVal KeyAscii As Integer)
Public Event KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
Public Event KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
Public Event Click(ByVal x As Integer, ByVal y As Integer)

Public Sub SetStyle(Font_ As Byte, Color_ As Long, Color2_ As Long, Gradiente_ As eGuiColorFill)
    cLabel.Font = Font_
    cLabel.Color = Color_
    cLabel.Color2 = Color2_
    cLabel.GradientMode = Gradiente_
End Sub

Public Sub Render(Optional ByVal focus As Boolean = False)
    If Not cLabel Is Nothing Then
        cLabel.Render
    End If
End Sub

Private Sub ReDoRECT()
    Dim textX As Integer
    Dim textY As Integer

    textX = CX
    textY = cY
    
    If Not cLabel Is Nothing Then
        textX = textX
        textY = textY
        cLabel.SetPos CX, cY
    End If

    With ControlRect
        .top = textY
        .bottom = textY
        .left = textX
        .right = textX
    End With
    
    tBox.x0 = ControlRect.left
    tBox.x1 = ControlRect.left
        
    tBox.x2 = ControlRect.right
    tBox.x3 = ControlRect.right
        
    tBox.y0 = ControlRect.bottom
    tBox.y1 = ControlRect.top
        
    tBox.y2 = ControlRect.bottom
    tBox.y3 = ControlRect.top
End Sub



Public Property Let Enabled(ByVal v As Boolean)
    cEnabled = v
End Property

Private Sub Class_Initialize()
    cEnabled = True
    cW = 128
    cH = 16
    
    With tBox
        .rhw0 = 1
        .rhw1 = 1
        .rhw2 = 1
        .rhw3 = 1
        .tu0 = 0
        .tu1 = 0
        .tu2 = 1
        .tu3 = 1
        .color0 = -1
        .Color1 = -1
        .Color2 = -1
        .color3 = -1
    End With
    
    Set cLabel = New clsGuiText2
    
    cLabel.Font = 1
    cLabel.Centrar = False
End Sub

Public Property Get textura() As Integer
    textura = cLabel.text
End Property

Public Property Let textura(ByVal v As Integer)
    cLabel.text = v
End Property

Public Property Get text() As String
    text = cLabel.text
End Property

Public Property Get Font() As Byte
    Font = cLabel.Font
End Property

Public Property Let Font(ByVal f As Byte)
    cLabel.Font = f
End Property

Public Property Let text(ByVal T As String)
    cLabel.text = T
End Property


Private Function IControl_Click(ByVal x As Integer, ByVal y As Integer) As Boolean
    IControl_Click = ColisionRect(x, y, ControlRect)
    If IControl_Click Then
        RaiseEvent Click(x, y)
    End If
End Function

Private Function IControl_GetPosX() As Integer
    IControl_GetPosX = CX
End Function

Private Function IControl_GetPosY() As Integer
    IControl_GetPosY = cY
End Function

Private Property Get IControl_IsFocuseable() As Boolean
    IControl_IsFocuseable = False
End Property

Private Property Get IControl_IsMouseFocus() As Boolean
    IControl_IsMouseFocus = False
End Property

Private Function IControl_KeyDown(KeyCode As Integer, Shift As Integer) As Boolean
    IControl_KeyDown = False
End Function

Private Function IControl_KeyPress(KeyAscii As Integer) As Boolean
    IControl_KeyPress = False
End Function

Private Function IControl_KeyUp(KeyCode As Integer, Shift As Integer) As Boolean
    IControl_KeyUp = False
End Function

Private Function IControl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    If cEnabled Then
        IControl_MouseDown = ColisionRect(x, y, ControlRect)
    End If
    
    If IControl_MouseDown Then
        RaiseEvent MouseDown(Button, Shift, x - CX, y - cY)
    End If
End Function

Private Function IControl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer) As Boolean
    If cEnabled Then
        IControl_MouseMove = ColisionRect(x, y, ControlRect)
    End If
        
    If IControl_MouseMove Then
        RaiseEvent MouseMove(Button, Shift, x - CX, y - cY, z)
    End If
End Function

Private Function IControl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    If cEnabled Then
        IControl_MouseUp = ColisionRect(x, y, ControlRect)
    End If
    
    If IControl_MouseUp Then
        RaiseEvent MouseUp(Button, Shift, x - CX, y - cY)
    End If
End Function

Private Sub IControl_Render(ByVal focus As Boolean)
    Render focus
End Sub

Private Sub IControl_SetPos(ByVal x As Integer, ByVal y As Integer)
    CX = x
    cY = y
    ReDoRECT
End Sub

Public Sub SetPos(ByVal x As Integer, ByVal y As Integer)
    IControl_SetPos x, y
End Sub

Private Sub IControl_SetSize(ByVal w As Integer, ByVal h As Integer)
    cW = w
    cH = h
    cLabel.SetWidth w
    ReDoRECT
End Sub

Public Sub SetSize(ByVal w As Integer, ByVal h As Integer)
    IControl_SetSize w, h
End Sub

Public Sub SetCentrado(ByVal centrado As Boolean)
    cLabel.Centrar = centrado
End Sub

Public Property Get Gradient() As eGuiColorFill
    Gradient = cLabel.GradientMode
End Property

Public Property Let Gradient(ByVal vNewValue As eGuiColorFill)
    cLabel.GradientMode = vNewValue
End Property

Public Property Get Color2() As Long
    Color = cLabel.Color
End Property

Public Property Let Color2(ByVal vNewValue As Long)
    cLabel.Color2 = vNewValue
End Property

Public Property Get Color() As Long
    Color = cLabel.Color
End Property

Public Property Let Color(ByVal vNewValue As Long)
    cLabel.Color = vNewValue
End Property

Public Property Get Width() As Integer
    Dim w As Integer
    Dim h As Integer
    cLabel.GetSize w, h
    Width = w
End Property
