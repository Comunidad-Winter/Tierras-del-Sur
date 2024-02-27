VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vwMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements vWindow

Private vX As Integer
Private vY As Integer
Private vW As Integer
Private vH As Integer

' Superior
Private WithEvents BotonSalir As ControlBoton
Attribute BotonSalir.VB_VarHelpID = -1
Private WithEvents BotonWeb As ControlBoton
Attribute BotonWeb.VB_VarHelpID = -1
Private WithEvents BotonOpciones As ControlBoton
Attribute BotonOpciones.VB_VarHelpID = -1
Private WithEvents BotonAyuda As ControlBoton
Attribute BotonAyuda.VB_VarHelpID = -1

Private WithEvents vControles As vControles
Attribute vControles.VB_VarHelpID = -1

'Implementacion de ventanas. No modificar
    Private pNext As vWindow
    Private pPrev As vWindow

Private Sub BotonAyuda_Click()
    Call Sonido_Play(SND_CLICK)
    Call openUrl("https://wiki.tierrasdelsur.cc")
End Sub

Private Sub BotonOpciones_Click()
    Call Sonido_Play(SND_CLICK)
    Call frmOpciones.Show
End Sub

Private Sub BotonSalir_Click()
    Call Sonido_Play(SND_CLICK)
    prgRun = False
End Sub

Private Sub BotonWeb_Click()
    Call Sonido_Play(SND_CLICK)
    Call openUrl("https://tierrasdelsur.cc")
End Sub

Private Property Get vWindow_GetControl() As vControles
    Set vWindow_GetControl = vControles
End Property

    Private Function vWindow_GetNext(vWindow As vWindow) As Boolean: Set vWindow = pNext: vWindow_GetNext = Not (pNext Is Nothing): End Function
    Private Function vWindow_GetPrev(vWindow As vWindow) As Boolean: Set vWindow = pPrev: vWindow_GetPrev = Not (pPrev Is Nothing): End Function

Private Property Get vWindow_IsVisible() As Boolean
     vWindow_IsVisible = Visible
End Property

    Private Sub vWindow_SetNext(vWindow As vWindow): Set pNext = vWindow: End Sub
    Private Sub vWindow_SetPrev(vWindow As vWindow): Set pPrev = vWindow: End Sub

'/Implementacion

Private Sub Class_Terminate()
    Set vControles = Nothing
End Sub

Private Sub Class_Initialize()
Set vControles = New vControles

Set BotonWeb = vControles.Add(New ControlBoton)
Set BotonOpciones = vControles.Add(New ControlBoton)
Set BotonAyuda = vControles.Add(New ControlBoton)
Set BotonSalir = vControles.Add(New ControlBoton)

updateControles
End Sub

Private Sub updateControles()
    Dim botoneraX As Integer
    Dim botoneraY As Integer
    
    botoneraY = vY
    botoneraX = vX
    
    BotonOpciones.textura = 3247
    BotonOpciones.SetPos botoneraX, botoneraY
    BotonOpciones.SetSize 22, 22
    BotonOpciones.Enabled = True
    BotonOpciones.Tooltip = "Opciones"
    
    botoneraX = botoneraX + 32
    
    BotonAyuda.textura = 3246
    BotonAyuda.SetPos botoneraX, botoneraY
    BotonAyuda.SetSize 22, 22
    BotonAyuda.Enabled = True
    BotonAyuda.Tooltip = "Manual"
    
    botoneraX = botoneraX + 32
    
    BotonWeb.textura = 3248
    BotonWeb.SetPos botoneraX, botoneraY
    BotonWeb.SetSize 22, 22
    BotonWeb.Enabled = True
    BotonWeb.Tooltip = "Ingresar a la Web"
    
    botoneraX = botoneraX + 32
    
    BotonSalir.textura = 3235
    BotonSalir.SetPos botoneraX, botoneraY
    BotonSalir.SetSize 22, 22
    BotonSalir.Enabled = True
    BotonSalir.Tooltip = "Salir"
End Sub
Private Sub vWindow_Hide()
    
End Sub

Private Function vWindow_Render() As Boolean

    vWindow_Render = True

    vControles.Render
End Function

Private Sub vWindow_SetPos(ByVal X As Integer, ByVal Y As Integer)
    vX = X
    vY = Y
    Call updateControles
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

