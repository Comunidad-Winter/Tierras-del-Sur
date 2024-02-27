VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IControl

Option Explicit

Private WithEvents BotonSiguiente As ControlBoton
Attribute BotonSiguiente.VB_VarHelpID = -1
Private WithEvents BotonAnterior As ControlBoton
Attribute BotonAnterior.VB_VarHelpID = -1

Private NombreElemento As clsGuiText2

Private valorMinimo As Integer
Private valorMaximo As Integer
Private ValorActual As Integer

Private posicioninicial As position

Public Event Change(valor As Integer)

Private cW As Integer
Private cH As Integer

Private vwControles As vControles

Public Property Let Nombre(ByVal v As String)
    NombreElemento.text = v
End Property

Public Sub SetStyle(Font_ As Byte, Color_ As Long, Color2_ As Long, Gradiente_ As eGuiColorFill)
    Font = Font_
    Color = Color_
    Color2 = Color2_
    Gradiente = Gradiente_
End Sub

Public Property Let Font(ByVal v As Byte)
    NombreElemento.Font = v
End Property

Public Property Let Color(ByVal v As Long)
    NombreElemento.Color = v
End Property

Public Property Let Color2(ByVal v As Long)
    NombreElemento.Color2 = v
End Property

Public Property Let Gradiente(ByVal v As eGuiColorFill)
    NombreElemento.GradientMode = v
End Property

Public Property Let min(ByVal v As Integer)
    valorMinimo = v
End Property

Public Property Let max(ByVal v As Integer)
    valorMaximo = v
End Property

Public Property Let valor(ByVal v As Integer)
    ValorActual = v
End Property

Public Property Get valor() As Integer
    valor = ValorActual
End Property


Private Sub BotonAnterior_Click()
    ValorActual = ValorActual - 1
    
    If ValorActual < valorMinimo Then
        ValorActual = valorMaximo
    End If
    
    RaiseEvent Change(ValorActual)
End Sub

Private Sub BotonSiguiente_Click()
    ValorActual = ValorActual + 1
    
    If ValorActual > valorMaximo Then
        ValorActual = valorMinimo
    End If
    
    RaiseEvent Change(ValorActual)
End Sub

Private Sub Class_Initialize()
    Set vwControles = New vControles
    
    Set NombreElemento = New clsGuiText2
    Set BotonSiguiente = vwControles.Add(New ControlBoton)
    Set BotonAnterior = vwControles.Add(New ControlBoton)
    
    BotonSiguiente.SetSize 20, 20
    BotonSiguiente.textura = 3361
    BotonSiguiente.text = ""
    
    BotonAnterior.SetSize 20, 20
    BotonAnterior.textura = 3362
    BotonAnterior.text = ""
    
    NombreElemento.Centrar = True
End Sub

Private Function IControl_Click(ByVal X As Integer, ByVal Y As Integer) As Boolean
End Function

Private Property Get IControl_IsFocuseable() As Boolean
    IControl_IsFocuseable = True
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
    IControl_MouseDown = vwControles.MouseDown(Button, Shift, X, Y)
End Function

Private Function IControl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal z As Integer) As Boolean
    IControl_MouseMove = vwControles.MouseMove(Button, Shift, X, Y, z)
End Function

Private Function IControl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    IControl_MouseUp = vwControles.MouseUp(Button, Shift, X, Y)
End Function

Private Sub IControl_Render(ByVal focus As Boolean)
    vwControles.Render
    NombreElemento.Render
End Sub

Public Sub IControl_SetPos(ByVal X As Integer, ByVal Y As Integer)
    posicioninicial.X = X
    posicioninicial.Y = Y
    
    Call actualizarPosiciones
End Sub

Public Sub IControl_SetSize(ByVal ancho As Integer, ByVal alto As Integer)
    cW = ancho
    cH = alto
End Sub

Private Sub actualizarPosiciones()
    BotonAnterior.SetPos posicioninicial.X, posicioninicial.Y + cH / 2
    
    NombreElemento.SetPos posicioninicial.X + cW / 2, posicioninicial.Y + cH / 2
    
    BotonSiguiente.SetPos posicioninicial.X + cW - 20, posicioninicial.Y + cH / 2
End Sub
