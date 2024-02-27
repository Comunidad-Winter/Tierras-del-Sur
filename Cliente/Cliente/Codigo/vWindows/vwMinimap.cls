VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vwMinimap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements vWindow

Option Explicit

Private vX As Integer
Private vY As Integer
Private vW As Integer
Private vH As Integer

Private Visible As Boolean
Private vControles As vControles

Private FondoVentana As ControlImage

'Implementacion de ventanas. No modificar
Private pNext As vWindow
Private pPrev As vWindow

Private numeroMapa As Integer

Private WithEvents BotonCerrar As ControlBoton
Attribute BotonCerrar.VB_VarHelpID = -1
Private WithEvents BotonMaximizar As ControlBoton
Attribute BotonMaximizar.VB_VarHelpID = -1
Private maximizado As Boolean
Private tamanoMapa As Integer
Private tamanoMapaReal As Integer

Private posRelativaX As Integer
Private posRelativaY As Integer
    
Public Event Cerrar(maximizado As Boolean)

Public Sub SetMapa(mapa As Integer)
    numeroMapa = mapa
End Sub

Private Sub BotonCerrar_Click()
    RaiseEvent Cerrar(maximizado)
End Sub

Private Sub BotonMaximizar_Click()
   Call SetMaximizado(Not maximizado)
End Sub

Public Sub SetMaximizado(max As Boolean)
    maximizado = max
    If maximizado Then
        tamanoMapa = tamanoMapaReal
        Call actualizarPosiciones
        BotonMaximizar.text = "Minimizar"
    Else
        tamanoMapa = 256
        Call actualizarPosiciones
        BotonMaximizar.text = "Maximizar"
    End If
End Sub

Private Sub Class_Initialize()
    Set vControles = New vControles
    Set FondoVentana = vControles.Add(New ControlImage)
    
    FondoVentana.grhId = 20352
    
    Set BotonCerrar = vControles.Add(New ControlBoton)
    Set BotonMaximizar = vControles.Add(New ControlBoton)
    
    BotonCerrar.text = "Cerrar"
    BotonMaximizar.text = "Maximizar"
    
    tamanoMapaReal = 512
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

Public Sub vWindow_Hide()
    Visible = False
End Sub

Public Sub vWindow_Show()
    Visible = True
End Sub

Private Function vWindow_Render() As Boolean
'<EhHeader>
On Error GoTo vWindow_Render_Err
'</EhHeader>
    Dim pixelPorTile As Single
    Dim tTop As Integer
    Dim tLeft As Integer
        
    If Visible Then vControles.Render
    
    vWindow_Render = True

    tTop = posRelativaY + 4
    tLeft = posRelativaX + 4
         
    pixelPorTile = tamanoMapa / ALTO_MAPA
        
    If maximizado = False Then
        Call Grh_Render_Simple_box_offset(BASE_TEXTURA_MAPAS + numeroMapa, mini(384, (UserPos.X - HalfWindowTileWidth) * pixelPorTile), mini(384, (UserPos.Y - HalfWindowTileWidth) * pixelPorTile), tamanoMapa, tamanoMapa, tLeft, tTop, -1, tamanoMapaReal)
    Else
        Call Grh_Render_Simple_box(BASE_TEXTURA_MAPAS + numeroMapa, tLeft, tTop, -1, tamanoMapaReal)
    End If
    
    tTop = tTop + (UserPos.Y - HalfWindowTileHeight) * pixelPorTile
    tLeft = tLeft + (UserPos.X - HalfWindowTileWidth) * pixelPorTile
        
    Call Grh_Render_Solid(mzWhite And &H7FFFFFFF, tLeft, tTop, WindowTileWidth * pixelPorTile, WindowTileHeight * pixelPorTile)
    '<EhFooter>
    Exit Function

vWindow_Render_Err:
    LogError Err.Description & _
           "in vWindow_Render "
    '</EhFooter>
End Function


Public Sub vWindow_SetPos(ByVal X As Integer, ByVal Y As Integer)
    vX = X
    vY = Y
    
    Call actualizarPosiciones
End Sub

Public Sub vWindow_SetSize(ByVal width As Integer, ByVal Height As Integer)
    vW = width
    vH = Height
    
    Call actualizarPosiciones
End Sub

Private Sub actualizarPosiciones()
    vW = tamanoMapa + 8
    vH = tamanoMapa + 8
    
    posRelativaX = vX / 2 - ((tamanoMapa / 2) + 4 + 20)
    posRelativaY = vY - (tamanoMapa + 4)
    
    FondoVentana.IControl_SetSize vW, vH
    
    FondoVentana.SetPos posRelativaX, posRelativaY
    BotonCerrar.SetPos posRelativaX + vW + 1, posRelativaY + vH - 30
    
    BotonMaximizar.SetPos posRelativaX + vW + 1, posRelativaY + vH - 60
    BotonCerrar.SetSize 100, 25
    BotonMaximizar.SetSize 100, 25
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


