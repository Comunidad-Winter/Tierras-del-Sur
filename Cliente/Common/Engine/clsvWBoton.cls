VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlBoton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IControl

Option Explicit

Event Click()

Private cLabel As New clsGuiText2

Private CX%, cY%, cH%, cW%

Private ControlRect As RECT

Private cHover As Boolean
Private cPressed As Boolean

Private cEnabled As Boolean

Private pTextura As Integer

Private tBox As Box_Vertex

Private isFocus As Boolean

Private alpha As Byte

Private TextToolTip As clsGuiText2

Private tooltipData As String

Private show As Boolean

Public Function setAlpha(alpha_ As Byte)
    alpha = alpha_
    cLabel.setAlpha alpha_
End Function

Public Property Let Visible(v As Boolean)
    show = v
End Property

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

Public Sub Render()

    If Not show Then Exit Sub
    
    With tBox
        .color0 = &HE6E7EA Or Alphas(alpha)
        .Color1 = &HE6E7EA Or Alphas(alpha)
        .Color2 = &HE6E7EA Or Alphas(alpha)
        .color3 = &HE6E7EA Or Alphas(alpha)
    End With
    
    If cEnabled Then
        If cHover Or isFocus Then
            If cPressed Then
                Draw 2, 4
            Else
                Draw 1, 4
            End If
        Else
            Draw 0, 4
        End If
    Else
        Draw 3, 4
    End If
    
    
    Dim tH%, tW%
    cLabel.GetSize tW, tH
    cLabel.SetPos (cW - tW) / 2 + CX + 4, (cH - tH) / 2 + cY
    cLabel.Render
    
    If cHover And Not tooltipData = "" Then
        If TextToolTip Is Nothing Then
            Set TextToolTip = New clsGuiText2
            TextToolTip.text = tooltipData
            TextToolTip.Centrar = True
            TextToolTip.SetPos CX + cW / 2, cY + cH + 5
            TextToolTip.Color = mzCTalkMuertos
        End If
        TextToolTip.Render
    Else
        Set TextToolTip = Nothing
    End If
End Sub

Public Function Click(ByVal X As Integer, ByVal Y As Integer) As Boolean

End Function

Public Function MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    cHover = ColisionRect(X, Y, ControlRect)
    MouseMove = cHover
End Function

Public Function MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    cHover = ColisionRect(X, Y, ControlRect)
    If cEnabled Then
        If Button = vbLeftButton And cHover = True And cPressed = True Then
            RaiseEvent Click
        End If
    Else
        'Beep
    End If
    MouseUp = cHover
    cPressed = False
End Function

Public Function MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    cHover = ColisionRect(X, Y, ControlRect)
    If cEnabled Then
        If Button = vbLeftButton Then
            cPressed = True
        Else
            Beep
        End If
    Else
        'Beep
    End If
    MouseDown = cHover
End Function

Public Sub SetPos(ByVal X As Integer, ByVal Y As Integer)
    CX = X
    cY = Y
    ReDoRECT
End Sub

Public Sub SetSize(ByVal w As Integer, ByVal h As Integer)
    cW = w
    cH = h
    ReDoRECT
End Sub

Private Sub ReDoRECT()
    With ControlRect
        .top = cY
        .bottom = cY + cH
        .left = CX
        .right = CX + cW
        tBox.x0 = .left
        tBox.x1 = .left
        
        tBox.x2 = .right
        tBox.x3 = .right
        
        tBox.y0 = .bottom
        tBox.y1 = .top
        
        tBox.y2 = .bottom
        tBox.y3 = .top
        
    End With
    cLabel.SetPos CX + (cH - 16) / 2, cY + (cH - 16) / 2
    
End Sub

Public Property Get Enabled() As Boolean
    Enabled = cEnabled
End Property

Public Property Let Enabled(ByVal v As Boolean)
    cEnabled = v
End Property

Public Property Get Tooltip() As String
    Tooltip = tooltipData
End Property

Public Property Let Tooltip(text As String)
    tooltipData = text
End Property

Public Property Get text() As String
    text = cLabel.text
End Property

Public Property Let text(a As String)
    cLabel.text = a
End Property

Private Sub Class_Initialize()
    cEnabled = True
    cW = 128
    cH = 32
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
    pTextura = 1099
    cLabel.Color = -1
    cLabel.GradientMode = dVertical
    cLabel.Font = 1
    tooltipData = ""
    setAlpha 255
    Visible = True
End Sub

Public Property Get textura() As Integer
textura = pTextura
End Property

Public Property Let textura(ByVal v As Integer)
pTextura = v
End Property

Private Sub Class_Terminate()
    Set cLabel = Nothing
    Set TextToolTip = Nothing
End Sub

Private Function IControl_Click(ByVal X As Integer, ByVal Y As Integer) As Boolean
IControl_Click = Click(X, Y)
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
    If KeyAscii = vbKeyReturn Then
        RaiseEvent Click
    End If
    IControl_KeyPress = True
End Function

Private Function IControl_KeyUp(KeyCode As Integer, Shift As Integer) As Boolean

End Function

Private Function IControl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
IControl_MouseDown = MouseDown(Button, Shift, X, Y)
End Function

Private Function IControl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal z As Integer) As Boolean
IControl_MouseMove = MouseMove(Button, Shift, X, Y)
End Function

Private Function IControl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
IControl_MouseUp = MouseUp(Button, Shift, X, Y)
End Function

Private Sub IControl_Render(ByVal focus As Boolean)
    isFocus = focus
    Render
End Sub

Private Sub IControl_SetPos(ByVal X As Integer, ByVal Y As Integer)
SetPos X, Y
End Sub

Private Sub IControl_SetSize(ByVal w As Integer, ByVal h As Integer)
SetSize w, h
End Sub

Public Property Get GetPosX() As Integer
    GetPosX = CX
End Property

Public Property Get GetPosY() As Integer
    GetPosY = cY
End Property

Public Property Get IControl_Width() As Integer
    IControl_Width = cW
End Property

Public Property Get IControl_Height() As Integer
    IControl_Height = cH
End Property
