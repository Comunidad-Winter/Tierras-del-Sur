VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IControl

Option Explicit

Event Click()

Private text As New clsGuiText2

Private CX%, cY%, cH%, cW%

Private ControlRect As RECT

Private cHover As Boolean
Private cPressed As Boolean

Private cEnabled As Boolean

Private cValue As Boolean

Private pTextura As Integer

Private tBox As Box_Vertex

Private isFocus As Boolean

Private Vertical_ As Boolean

Private alpha As Byte

Public Sub SetStyle(Font_ As Byte, Color_ As Long, Color2_ As Long, Gradiente_ As eGuiColorFill)
    text.Font = Font_
    text.Color = Color_
    text.Color2 = Color2_
    text.GradientMode = Gradiente_
End Sub

Public Sub SetVertical(Vertical As Boolean)
    Vertical_ = Vertical
End Sub

Public Sub setAlpha(alpha_ As Byte)
    alpha = alpha_
    text.setAlpha alpha_
    
    With tBox
        .color0 = &HE6E7EA Or Alphas(alpha)
        .Color1 = &HE6E7EA Or Alphas(alpha)
        .Color2 = &HE6E7EA Or Alphas(alpha)
        .color3 = &HE6E7EA Or Alphas(alpha)
    End With
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

Public Sub Render()
    
    If cEnabled Then
        If cHover Or isFocus Then
            If Not cPressed Then
                If cValue Then
                    Draw 3, 4
                Else
                    Draw 2, 4
                End If
            Else
                If cValue Then
                    Draw 2, 4
                Else
                    Draw 3, 4
                End If
            End If
        Else
            If cValue Then
                Draw 1, 4
            Else
                Draw 0, 4
            End If
        End If
    Else
        Draw 2, 4
    End If
    
    text.Render
End Sub

Public Function Click(ByVal X As Integer, ByVal Y As Integer) As Boolean
    Click = ColisionRect(X, Y, ControlRect)
End Function

Public Function MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    cHover = ColisionRect(X, Y, ControlRect)
    MouseMove = cHover
End Function

Public Function MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    cHover = ColisionRect(X, Y, ControlRect)
    If cEnabled Then
        If Button = vbLeftButton And cHover = True And cPressed = True Then
            cValue = Not cValue
            RaiseEvent Click
        End If
    Else
        Beep
    End If
    MouseUp = cHover
    cPressed = False
End Function

Public Function MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    cHover = ColisionRect(X, Y, ControlRect)
    If cEnabled Then
        If Button = vbLeftButton Then
            cPressed = True
        End If
    End If
    MouseDown = cHover
End Function

Public Sub SetPos(ByVal X As Integer, ByVal Y As Integer)
    CX = X
    cY = Y
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
        
        tBox.x2 = .left + cW
        tBox.x3 = .left + cW
        
        tBox.y0 = .top + cH
        tBox.y1 = .top
        
        tBox.y2 = .top + cH
        tBox.y3 = .top
    End With
    
    Dim hh%, ww%
    
    text.GetSize ww, hh
        
    If Vertical_ Then
        text.SetPos CX + cW / 2, cY + cH
        text.Centrar = True
        
        With ControlRect
            .top = cY
            .bottom = cY + cH + hh
            .left = CX
            .right = CX + cW
        End With
    Else
        text.SetPos CX + cW + 4, cY
        
        With ControlRect
            .top = cY
            .bottom = cY + cH
            .left = CX
            .right = CX + cW + 4 + ww
        End With
    End If
    
End Sub


Public Property Let texto(ByVal value As String)
    text.text = value
End Property

Public Property Get Enabled() As Boolean
    Enabled = cEnabled
End Property

Public Property Let Enabled(ByVal v As Boolean)
    cEnabled = v
End Property

Public Property Get value() As Boolean
Attribute value.VB_UserMemId = 0
Attribute value.VB_MemberFlags = "200"
    value = cValue
End Property

Public Property Let value(ByVal v As Boolean)
    cValue = v
End Property

Private Sub Class_Initialize()
    cEnabled = True
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
    
    pTextura = 1098
    
    text.Color = -1
    text.GradientMode = dVertical
    
    cH = 16
    cW = 16
End Sub

Public Property Get textura() As Integer
textura = pTextura
End Property

Public Property Let textura(ByVal v As Integer)
pTextura = v
End Property

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
    If KeyAscii = vbKeySpace Then
        cValue = Not cValue
    End If
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

Public Sub IControl_SetSize(ByVal w As Integer, ByVal h As Integer)
    cH = h
    cW = w
    ReDoRECT
End Sub

Public Property Get IControl_Width()
    IControl_Width = cW
End Property

Public Property Get IControl_Height()
    IControl_Height = cH
End Property

Public Property Get IControl_X()
    IControl_X = CX
End Property

Public Property Get IControl_Y()
    IControl_Y = cY
End Property

