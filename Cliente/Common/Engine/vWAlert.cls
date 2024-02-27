VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vWAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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

Private wLabel                  As clsGUIText
Private WithEvents Boton        As ControlBoton
Attribute Boton.VB_VarHelpID = -1

Private WithEvents Controles As vControles
Attribute Controles.VB_VarHelpID = -1

Private FinalPos As mzVECTOR2

Private Titulo As String

'Implementacion de ventanas. No modificar
    Private pNext As vWindow
    Private pPrev As vWindow

    Private Function vWindow_GetNext(vWindow As vWindow) As Boolean: Set vWindow = pNext: vWindow_GetNext = Not (pNext Is Nothing): End Function
    Private Function vWindow_GetPrev(vWindow As vWindow) As Boolean: Set vWindow = pPrev: vWindow_GetPrev = Not (pPrev Is Nothing): End Function
    Private Sub vWindow_SetNext(vWindow As vWindow): Set pNext = vWindow: End Sub
    Private Sub vWindow_SetPrev(vWindow As vWindow): Set pPrev = vWindow: End Sub
'/Implementacion

Private Property Get vWindow_GetControl() As vControles
    Set vWindow_GetControl = Controles
End Property


Public Function setTitulo(ByRef texto As String)
Titulo = texto
End Function

Private Sub Class_Terminate()
    Set wLabel = Nothing
    Set Controles = Nothing
End Sub

Private Sub Class_Initialize()
Set Controles = New vControles


Set Boton = Controles.Add(New ControlBoton)
Set wLabel = New clsGUIText

Boton.text = "Aceptar"
Boton.SetSize 128, 32

FinalPos.X = D3DWindow.BackBufferWidth / 2
FinalPos.Y = D3DWindow.BackBufferHeight / 2

InicioTick = GetTimer
FinTick = InicioTick + 300
I1 = 0.25
I2 = 1

wLabel.text = vTextoAlerta
Titulo = vAlerteTitle
End Sub

Private Sub Boton_Click()
    I1 = 1
    I2 = 0
    InicioTick = GetTimer
    FinTick = InicioTick + 300
    Matar = True
End Sub

Private Sub vWindow_Hide()

End Sub

Private Function Animar(ancho As Integer, alto As Integer) As Boolean
    Animar = True
    Dim TmpTick As Long
    TmpTick = GetTimer
    
    If TmpTick < FinTick Then
        Progreso = CosInterp(I1, I2, (TmpTick - InicioTick) / (FinTick - InicioTick))
    Else
        Progreso = 1
        If I2 = 0 Then
            'Decimos "La ventana acaba de morirse wachin"
            Animar = Not Matar
            If Animar = False Then Exit Function
        End If
    End If
    
    vW = ancho * Progreso
    vH = alto * Progreso
    vX = FinalPos.X - (vW / 2)
    vY = CosInterp(200, FinalPos.Y - (vH / 2), Progreso)
    
    GUI_RenderDialog vX, vY, vW, vH, Titulo, Me, Progreso
End Function

Private Function vWindow_Render() As Boolean
    Dim w%, h%
    
    wLabel.GetSize w, h
    
    If w = 0 Then wLabel.Render
    
    Engine.Draw_FilledBox 0, 0, D3DWindow.BackBufferWidth, D3DWindow.BackBufferHeight, Alphas(Progreso * 128), 0, 0
    
    vWindow_Render = Animar(w + 32, h + 128 + 16)

    'Oscurezco la pantalla
    If Progreso = 1 Then
        Boton.SetPos vX + vW - 32 - 128, vY + vH - 48
        Boton.SetSize 128, 32
        Boton.Render
        
        wLabel.SetPos vX + 16, vY + 16 + 24
        wLabel.Render
    End If
    
End Function

Private Sub vWindow_SetPos(ByVal X As Integer, ByVal Y As Integer)
    FinalPos.X = X
    FinalPos.Y = Y
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

