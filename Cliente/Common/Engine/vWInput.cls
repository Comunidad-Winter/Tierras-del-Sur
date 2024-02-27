VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IControl

Option Explicit

Private pText As New clsGuiText2
Private BloqMayus As clsGuiText2

Private CX%, cY%, cH%, cW%

Private ControlRect As RECT

Private cEnabled As Boolean

Private pTextura As Integer

Private tBox As Box_Vertex

Private EsContraseña As Boolean
Private texto As String

Private tipo As eGuiTiposTexto

Private textColor2 As Long

Private MaxLen As Integer

Public CantidadDecimales As Integer

Public Filtro As String

Private cLabel As clsGuiText2

Private alpha As Byte

Private isFocus As Boolean

Public Event Change()
Public Event ChangeValue(NuevoValor As Single)

Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal z As Integer)
Public Event KeyPress(ByVal KeyAscii As Integer)
Public Event KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
Public Event KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
Public Event Click(ByVal X As Integer, ByVal Y As Integer)

Public Sub SetStyle(Font_ As Byte, Color_ As Long, Color2_ As Long, Gradiente_ As eGuiColorFill)
    cLabel.Font = Font_
    cLabel.Color = Color_
    cLabel.Color2 = Color2_
    cLabel.GradientMode = Gradiente_
    
    textColor2 = Color2_
End Sub

Private Sub Draw(parte As Single, partes As Single)
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(pTextura)
    With tBox
        .tv0 = (parte + 1) / partes
        .tv1 = parte / partes
        .tv2 = .tv0
        .tv3 = .tv1
    End With
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size

End Sub

Public Sub setAlpha(alpha_ As Byte)
    alpha = alpha_
    cLabel.setAlpha alpha_
    pText.setAlpha alpha_
    
    With tBox
        .color0 = &HE6E7EA Or Alphas(alpha)
        .Color1 = &HE6E7EA Or Alphas(alpha)
        .Color2 = &HE6E7EA Or Alphas(alpha)
        .color3 = &HE6E7EA Or Alphas(alpha)
    End With
End Sub

Public Sub Render(Optional ByVal focus As Boolean = False)
    Static Timer_hover As Long
    Static Barrita_visible As Boolean
    
    Static ultimo_hover As Boolean
    
    Dim aux As Long
    
    If Not cLabel Is Nothing Then
        If focus Then
            cLabel.Color2 = textColor2 Or &HFF8B8989
        Else
            cLabel.Color2 = textColor2
        End If
        cLabel.Render
    End If
    
    If focus Then
        If Timer_hover + 500 < GetTickCount Then
            Barrita_visible = Not Barrita_visible
            Timer_hover = GetTickCount
            actualizar_texto
        End If
    Else
        If Barrita_visible Then
            actualizar_texto
            Barrita_visible = False
        End If
    End If

    If cEnabled Then
        If focus Then
            Draw 1, 4
        Else
            Draw 0, 4
        End If
    Else
        Draw 3, 4
    End If
    
    If tipo = Contraseña And focus Then
        If GetKeyState(vbKeyCapital) = 0 Then
            If Not BloqMayus Is Nothing Then
                Set BloqMayus = Nothing
            End If
        Else
            If BloqMayus Is Nothing Then
                Set BloqMayus = New clsGuiText2
                
                BloqMayus.SetStyle 3, mzInterfaceColor1, mzInterfaceColor2, dVertical
                BloqMayus.SetWidth 150
                BloqMayus.text = "Bloq Mayús"
                BloqMayus.SetPos ControlRect.right + 5, ControlRect.top
            End If
            BloqMayus.Render
        End If
    End If
        
    pText.Render
End Sub

Private Sub ReDoRECT()

    Dim textX As Integer
    Dim textY As Integer

    textX = CX
    textY = cY
    
    If Not cLabel Is Nothing Then
        textX = textX
        textY = textY + 18
        cLabel.SetPos CX, cY
    End If

    With ControlRect
        .top = textY
        .bottom = textY + cH
        .left = textX
        .right = textX + cW
    End With
    
    tBox.x0 = ControlRect.left
    tBox.x1 = ControlRect.left
        
    tBox.x2 = ControlRect.right
    tBox.x3 = ControlRect.right
        
    tBox.y0 = ControlRect.bottom
    tBox.y1 = ControlRect.top
        
    tBox.y2 = ControlRect.bottom
    tBox.y3 = ControlRect.top
        
    pText.SetPos textX + 4, textY + 5
End Sub

Public Property Get Enabled() As Boolean
    Enabled = cEnabled
End Property

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

    pTextura = 1100
    pText.Color = -1
    pText.GradientMode = dVertical
    
    CantidadDecimales = 2
    
    Set cLabel = New clsGuiText2
End Sub

Public Property Get textura() As Integer
    textura = pTextura
End Property

Public Property Let textura(ByVal v As Integer)
    pTextura = v
End Property

Public Property Get text() As String
    text = texto
End Property

Public Property Let text(ByVal T As String)
    Dim distinto As Boolean
    
    distinto = texto <> T
    texto = T
    
    If tipo And eGuiTiposTexto.numerico Then
        texto = Trim$(Str$(Round(CCVal(texto), CantidadDecimales)))
    End If
    
    If MaxLen <> 0 Then
        If Len(texto) > MaxLen Then
            texto = left$(texto, MaxLen)
        End If
    End If
    
    actualizar_texto
    
    If distinto Then
    RaiseEvent Change
    End If
End Property

Public Property Get TipoTexto() As eGuiTiposTexto
    TipoTexto = tipo
End Property

Public Property Let TipoTexto(ByVal T As eGuiTiposTexto)
    tipo = T
    actualizar_texto
End Property

Public Property Get MaxLenght() As Integer
    MaxLen = tipo
End Property

Public Property Let MaxLenght(ByVal T As Integer)
    MaxLen = T
    texto = Trim$(left$(texto, MaxLen))
    actualizar_texto
End Property

Private Sub actualizar_texto()
    Dim backup As String

    backup = pText.text
    If tipo And eGuiTiposTexto.Contraseña Then
        Dim i As Integer
        Dim tmp_txt As String
        For i = 1 To Len(texto)
            tmp_txt = tmp_txt & "•"
        Next i
        pText.text = tmp_txt
    Else
        pText.text = texto
    End If
    
    RaiseEvent Change

End Sub


Private Function IControl_Click(ByVal X As Integer, ByVal Y As Integer) As Boolean
    IControl_Click = ColisionRect(X, Y, ControlRect)
    If IControl_Click Then
        RaiseEvent Click(X, Y)
    End If
End Function

Private Function IControl_GetPosX() As Integer
    IControl_GetPosX = CX
End Function

Private Function IControl_GetPosY() As Integer
    IControl_GetPosY = cY
End Function

Private Property Get IControl_IsFocuseable() As Boolean
    IControl_IsFocuseable = True
End Property

Private Property Get IControl_IsMouseFocus() As Boolean
    IControl_IsMouseFocus = True
End Property

Private Function IControl_KeyDown(KeyCode As Integer, Shift As Integer) As Boolean
    IControl_KeyDown = True
    If cEnabled Then
        RaiseEvent KeyDown(KeyCode, Shift)
    End If
End Function

Private Function IControl_KeyPress(KeyAscii As Integer) As Boolean
Dim nuevoVal As Single
If cEnabled Then
        If KeyAscii <> 13 And (Len(Filtro) = 0 Or CBool(InStr(1, Filtro, Chr$(KeyAscii))) Or KeyAscii = 8) Then
            If KeyAscii = 8 And Len(texto) > 0 Then
                texto = mid(text, 1, Len(texto) - 1)
                If tipo And eGuiTiposTexto.numerico Then
                    nuevoVal = val(texto)
                    RaiseEvent ChangeValue(nuevoVal)
                    texto = Trim$(Str$(nuevoVal))
                Else
                    RaiseEvent Change
                End If
            ElseIf KeyAscii = 8 Then
                Beep
            ElseIf KeyAscii <> 8 Then
                If Len(texto) < MaxLen Or MaxLen = 0 Then
                    If tipo And eGuiTiposTexto.numerico Then
                        If InStr(1, "1234567890-+" & DecimalSeparator, Chr$(KeyAscii)) And Len(texto) > 0 Then
                            If Chr$(KeyAscii) <> DecimalSeparator Or (Chr$(KeyAscii) = DecimalSeparator And Not CBool(InStr(1, DecimalSeparator, texto))) Then
                                texto = texto & Chr$(KeyAscii)
                                
                                If Chr$(KeyAscii) <> DecimalSeparator Then
                                    texto = Str$(Round(CCVal(texto), CantidadDecimales))
                                End If
                                nuevoVal = val(texto)
                                RaiseEvent ChangeValue(nuevoVal)
                                texto = Trim$(Str$(nuevoVal))
                            Else
                                Beep
                            End If
                        Else
                            Beep
                        End If
                    Else
                        texto = texto & Chr$(KeyAscii)
                        RaiseEvent Change
                    End If
                End If
            End If
            
            
        End If
        actualizar_texto
        RaiseEvent KeyPress(KeyAscii)
    Else
        Beep
    End If
    IControl_KeyPress = True
End Function

Private Function IControl_KeyUp(KeyCode As Integer, Shift As Integer) As Boolean
IControl_KeyUp = True
    If cEnabled Then RaiseEvent KeyUp(KeyCode, Shift)
End Function

Private Function IControl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    If cEnabled Then
        IControl_MouseDown = ColisionRect(X, Y, ControlRect)
    End If
    If IControl_MouseDown Then
        RaiseEvent MouseDown(Button, Shift, X - CX, Y - cY)
    End If
End Function

Private Function IControl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal z As Integer) As Boolean
    If cEnabled Then
        IControl_MouseMove = ColisionRect(X, Y, ControlRect)
    End If
    
    If tipo = eGuiTiposTexto.numerico And IControl_MouseMove = True Then
        If z <> 0 Then
            If CantidadDecimales <> 0 Then
                texto = Str$(CCVal(texto) + z / CantidadDecimales)
            Else
                texto = Str$(CCVal(texto) + z)
            End If
            Dim nuevoVal As Single
            nuevoVal = val(texto)
            RaiseEvent ChangeValue(nuevoVal)
            texto = Trim$(Str$(nuevoVal))
            actualizar_texto
        End If
    End If
    
    If IControl_MouseMove Then
        RaiseEvent MouseMove(Button, Shift, X - CX, Y - cY, z)
    End If
End Function

Private Function IControl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    If cEnabled Then
        IControl_MouseUp = ColisionRect(X, Y, ControlRect)
    End If
    
    If IControl_MouseUp Then
        RaiseEvent MouseUp(Button, Shift, X - CX, Y - cY)
    End If
End Function

Private Sub IControl_Render(ByVal focus As Boolean)
    If isFocus = False And focus Then
        RaiseEvent Click(CX, cY)
    End If
    isFocus = focus
    Render focus
End Sub

Private Sub IControl_SetPos(ByVal X As Integer, ByVal Y As Integer)
    CX = X
    cY = Y
    ReDoRECT
End Sub

Public Sub SetPos(ByVal X As Integer, ByVal Y As Integer)
    IControl_SetPos X, Y
End Sub

Private Sub IControl_SetSize(ByVal w As Integer, ByVal h As Integer)
    cW = w
    cH = h
    ReDoRECT
End Sub

Public Sub SetSize(ByVal w As Integer, ByVal h As Integer)
    IControl_SetSize w, h
End Sub

Public Property Get label() As String
If Not cLabel Is Nothing Then
    label = cLabel.text
End If
End Property

Public Property Let label(ByVal v As String)
If cLabel Is Nothing Then
    Set cLabel = New clsGuiText2
    ReDoRECT
End If
cLabel.text = v

End Property

Public Property Get Color() As Long
    Color = pText.Color
End Property

Public Property Let Color(ByVal vNewValue As Long)
    pText.Color = vNewValue
End Property

