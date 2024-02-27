VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vwMiniMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vWindow

Private vX As Integer
Private vY As Integer

Private alpha As Byte

Private WithEvents Fondo        As ControlMouseReader
Attribute Fondo.VB_VarHelpID = -1

Private WithEvents Controles As vControles
Attribute Controles.VB_VarHelpID = -1

'Implementacion de ventanas. No modificar
    Private pNext As vWindow
    Private pPrev As vWindow


Private Sub moverPantalla(ByVal x As Integer, ByVal y As Integer)
    Dim puedeMoverse As Boolean
    
    puedeMoverse = False
    
    If WalkMode = False Then
        puedeMoverse = True
    ElseIf PuedoCaminar(x, y, E_Heading.NONE, False, False) = True Then
        puedeMoverse = True
    End If
    
    If puedeMoverse Then
        ' Cambiamos la posición
        UserPos.x = x
        UserPos.y = y
    
        ' Forzamos a actualizar la Vista
        Call rm2a
        Cachear_Tiles = True
    Else
        Beep
    End If
    
End Sub
Private Sub Fondo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)

    If Not esPosicionJugable(x, y) Then Exit Sub

    Call moverPantalla(x, y)
    
End Sub

Private Sub Fondo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer)
    If Not esPosicionJugable(x, y) Or Button = 0 Then Exit Sub

    Call moverPantalla(x, y)

End Sub

    Private Function vWindow_GetNext(vWindow As vWindow) As Boolean: Set vWindow = pNext: vWindow_GetNext = Not (pNext Is Nothing): End Function
    Private Function vWindow_GetPrev(vWindow As vWindow) As Boolean: Set vWindow = pPrev: vWindow_GetPrev = Not (pPrev Is Nothing): End Function

    Private Sub vWindow_SetNext(vWindow As vWindow): Set pNext = vWindow: End Sub

    Private Sub vWindow_SetPrev(vWindow As vWindow): Set pPrev = vWindow: End Sub
'/Implementacion

Private Property Get vWindow_GetControl() As vControles
    Set vWindow_GetControl = Controles
End Property

Private Sub Class_Initialize()
Set Controles = New vControles

Set Fondo = Controles.Add(New ControlMouseReader)

vX = D3DWindow.BackBufferWidth - 10 - 200
vY = D3DWindow.BackBufferHeight - 10 - 355

Fondo.SetSize SV_Constantes.ANCHO_MAPA, SV_Constantes.ALTO_MAPA

Fondo.SetPos vX + 1, vY + 1
End Sub

Private Sub vWindow_SetPos(ByVal x As Integer, ByVal y As Integer)
    vX = x
    vY = y
    Fondo.SetPos vX + 1, vY + 1
End Sub

Private Sub vWindow_Show()

End Sub

Private Sub vWindow_Hide()

End Sub

Private Function vWindow_Render() As Boolean

    miniMap_Render vX, vY, Colision(frmMain.MouseX, frmMain.MouseY, vY, vY + 220, vX, vX + 220)
    
    
    vWindow_Render = True
End Function

