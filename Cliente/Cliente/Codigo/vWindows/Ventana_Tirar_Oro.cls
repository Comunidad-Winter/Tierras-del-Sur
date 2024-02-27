VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ventana_Tirar_Oro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements vWindow

Private vX As Integer
Private vY As Integer
Private vW As Integer
Private vH As Integer

Private alpha As Byte

Private Progreso As Single

Private InicioTick As Long
Private FinTick As Long

Private Matar As Boolean

Private I1 As Single
Private I2 As Single

Private WithEvents Boton As ControlBoton
Attribute Boton.VB_VarHelpID = -1
Private WithEvents Textbox As ControlInput
Attribute Textbox.VB_VarHelpID = -1
Private WithEvents vControles As vControles
Attribute vControles.VB_VarHelpID = -1

Private Titulo As String

Private ObjetoFocus As Object

'Implementacion de ventanas. No modificar
    Private pNext As vWindow
    Private pPrev As vWindow

Private Property Get vWindow_GetControl() As vControles
    Set vWindow_GetControl = vControles
End Property

    Private Function vWindow_GetNext(vWindow As vWindow) As Boolean: Set vWindow = pNext: vWindow_GetNext = Not (pNext Is Nothing): End Function
    Private Function vWindow_GetPrev(vWindow As vWindow) As Boolean: Set vWindow = pPrev: vWindow_GetPrev = Not (pPrev Is Nothing): End Function
    Private Sub vWindow_SetNext(vWindow As vWindow): Set pNext = vWindow: End Sub
    Private Sub vWindow_SetPrev(vWindow As vWindow): Set pPrev = vWindow: End Sub
'/Implementacion

Public Function setTitulo(ByRef texto As String)
Titulo = texto
End Function

Private Sub Class_Terminate()
    Set wLabel = Nothing

End Sub

Private Sub Class_Initialize()
Set vControles = New vControles
Set Boton = vControles.Add(New ControlBoton)
Set Textbox = vControles.Add(New ControlInput)

Boton.text = "Aceptar"
Boton.SetSize 128, 32


Textbox.text = "1"
Textbox.Enabled = True
Textbox.TipoTexto = numerico


InicioTick = timeGetTime
FinTick = InicioTick + 300
I1 = 0.25
I2 = 1
End Sub


Private Sub Boton_Click()
    I1 = 1
    I2 = 0
    InicioTick = timeGetTime
    FinTick = InicioTick + 300
    Matar = True
    
    If val(Textbox.text) > 0 Then
        If val(Textbox.text) > 100000 Then
            EnviarPaquete Tirar, Chr$(254) & Codify(100000)
        Else
            EnviarPaquete Tirar, Chr$(254) & Codify(val(Textbox.text))
        End If
    End If
    
End Sub

Private Sub vWindow_Hide()

End Sub

Private Function vWindow_Render() As Boolean

    Dim TmpTick As Long
    
    TmpTick = timeGetTime
    'Decimos "Si, la ventana sigue viva."
    vWindow_Render = True
    
    If TmpTick < FinTick Then
        Progreso = CosInterp(I1, I2, (TmpTick - InicioTick) / (FinTick - InicioTick))
    Else
        Progreso = 1
        If I2 = 0 Then
            'Decimos "La ventana acaba de morirse wachin"
            vWindow_Render = Not Matar
            If vWindow_Render = False Then Exit Function
        End If
    End If
    


    vW = 192 * Progreso
    vH = 128 * Progreso
    vX = (D3DWindow.BackBufferWidth - vW) / 2
    vY = CosInterp(200, (D3DWindow.BackBufferHeight - vH) / 2, Progreso)
    alpha = Progreso * 128

    'Oscurezco la pantalla
    Engine.Draw_FilledBox 0, 0, D3DWindow.BackBufferWidth, D3DWindow.BackBufferHeight, Alphas(Progreso * 64), 0, 0

    'Dibujo la ventana
    GUI_RenderDialog vX, vY, vW, vH, "Tirar oro", Me, Progreso
    
    If Progreso = 1 Then
        Boton.SetPos vX + 16, vY + vH - 48
        Boton.SetSize 128, 32
        Boton.Render
                
        Textbox.SetPos vX + 16, vY + vH - 48 - 32
        Textbox.SetSize vW - 32, 16
        Textbox.Render (ObjetoFocus Is Textbox)
    End If
    
End Function

Private Sub vWindow_SetPos(ByVal X As Integer, ByVal Y As Integer)
End Sub

Private Sub vWindow_Show()
End Sub

Public Property Get vWindow_X() As Integer
    vWindow_X = vX
End Property

Public Property Get vWindow_Y() As Integer
    vWindow_Y = vY
End Property

Public Property Get vWindow_Width() As Integer
    vWindow_Width = vW
End Property

Public Property Get vWindow_Height() As Integer
    vWindow_Height = vH
End Property


