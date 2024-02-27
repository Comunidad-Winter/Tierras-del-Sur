VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlBarra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IControl

Option Explicit

Private NombreElemento As clsGuiText2
Private ValorElemento As clsGuiText2
Private TexturaBarrita As Integer  ' Textura de la Barrita


Private tamanoBarra As Long
Private valorMinimo As Integer
Private valorMaximo As Integer

Private posicioninicial As position

Private Progreso As Single
Private InicioTick As Long
Private FinTick As Long
Private AnteriorTamano As Long


Public Property Let Nombre(ByVal v As String)
    NombreElemento.text = v
End Property

Public Sub SetStyle(textura_ As Long, Font_ As Byte, Color_ As Long, Color2_ As Long, Gradiente_ As eGuiColorFill)
    Font = Font_
    Color = Color_
    Color2 = Color2_
    Gradiente = Gradiente_
    textura = textura_
End Sub

Public Property Let Font(ByVal v As Byte)
    NombreElemento.Font = v
    ValorElemento.Font = v
End Property

Public Property Let Color(ByVal v As Long)
    NombreElemento.Color = v
    ValorElemento.Color = v
End Property


Public Property Let Color2(ByVal v As Long)
    NombreElemento.Color2 = v
    ValorElemento.Color2 = v
End Property

Public Property Let Gradiente(ByVal v As eGuiColorFill)
    NombreElemento.GradientMode = v
    ValorElemento.GradientMode = v
End Property

Public Property Let min(ByVal v As Integer)
    valorMinimo = v
    refrescarTamano
End Property

Public Property Let max(ByVal v As Integer)
    valorMaximo = v
    refrescarTamano
End Property

Public Property Let valor(ByVal v As Integer)
    ValorElemento.text = CStr(v)
    refrescarTamano
End Property

Private Sub refrescarTamano()
    If valorMaximo - valorMinimo <= 0 Or ValorElemento.text = "" Then
        tamanoBarra = 0
        Exit Sub
    End If
    AnteriorTamano = tamanoBarra
    tamanoBarra = 10 + Math.Round(50 * val(ValorElemento.text - valorMinimo) / (valorMaximo - valorMinimo))
    
    InicioTick = GetTimer
    FinTick = GetTimer + 500
End Sub
Public Property Let textura(ByVal v As Integer)
    TexturaBarrita = v
End Property


Private Sub Class_Initialize()
    Set NombreElemento = New clsGuiText2
    Set ValorElemento = New clsGuiText2
End Sub

Private Function IControl_Click(ByVal x As Integer, ByVal y As Integer) As Boolean
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

Private Function IControl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean

End Function

Private Function IControl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer) As Boolean

End Function

Private Function IControl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean

End Function

Private Sub IControl_Render(ByVal focus As Boolean)
    Dim TmpTick As Long
    Dim longitud As Integer
    
    TmpTick = GetTimer
    
    Dim pepe As Single
    
    If FinTick > TmpTick Then
        Progreso = CosInterp(0.25, 1, (TmpTick - InicioTick) / (FinTick - InicioTick))
        longitud = AnteriorTamano + Math.Round((tamanoBarra - AnteriorTamano) * Progreso)
    Else
        longitud = tamanoBarra
    End If
    
    Call Engine_GrhDraw.Grh_Render_size(TexturaBarrita, posicioninicial.x, posicioninicial.y, 0, -1, 0, longitud, 8, True)
    
    NombreElemento.Render
    ValorElemento.Render
End Sub

Public Sub IControl_SetPos(ByVal x As Integer, ByVal y As Integer)
    posicioninicial.x = x + 80
    posicioninicial.y = y + 7
    
    NombreElemento.SetPos x, y
    ValorElemento.SetPos x + 60 + 85, y
End Sub

Private Sub IControl_SetSize(ByVal ancho As Integer, ByVal alto As Integer)

End Sub
