VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VWOpenPakMap"
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

Private MousePos As D3DVECTOR2


Private TmpMapMouse As Integer
Private MapSelected As Integer

Private WithEvents BotonAbrir   As ControlBoton
Attribute BotonAbrir.VB_VarHelpID = -1
Private WithEvents BotonCerrar  As ControlBoton
Attribute BotonCerrar.VB_VarHelpID = -1

Private WithEvents Fondo        As ControlMouseReader
Attribute Fondo.VB_VarHelpID = -1

Private WithEvents Controles    As vControles
Attribute Controles.VB_VarHelpID = -1

Private TamanioTileMapa         As New clsAlpha

'''
Private MapIH As INFOHEADER

'Implementacion de ventanas. No modificar
    Private pNext As vWindow
    Private pPrev As vWindow


Private Property Get vWindow_GetControl() As vControles
    Set vWindow_GetControl = Controles
End Property

    Private Function vWindow_GetNext(vWindow As vWindow) As Boolean: Set vWindow = pNext: vWindow_GetNext = Not (pNext Is Nothing): End Function
    Private Function vWindow_GetPrev(vWindow As vWindow) As Boolean: Set vWindow = pPrev: vWindow_GetPrev = Not (pPrev Is Nothing): End Function

    Private Sub vWindow_SetNext(vWindow As vWindow): Set pNext = vWindow: End Sub
    Private Sub vWindow_SetPrev(vWindow As vWindow): Set pPrev = vWindow: End Sub
'/Implementacion

Private Sub Class_Initialize()

InicioTick = GetTimer
FinTick = InicioTick + 300
I1 = 0.25
I2 = 1

Set Controles = New vControles

Set BotonAbrir = Controles.Add(New ControlBoton)
Set BotonCerrar = Controles.Add(New ControlBoton)
Set Fondo = Controles.Add(New ControlMouseReader)

Fondo.SetPos 10, 10
Fondo.SetSize D3DWindow.BackBufferWidth - 20, D3DWindow.BackBufferHeight - 20

BotonAbrir.Enabled = False

BotonAbrir.text.text = "Abrir"
BotonCerrar.text.text = "Cancelar"

BotonAbrir.SetSize 90, 32
BotonCerrar.SetSize 90, 32

TamanioTileMapa.InitialValue = 48
TamanioTileMapa.Speed = 3000
TamanioTileMapa.value = 48
End Sub

Private Sub vWindow_Hide()

End Sub

Private Sub Fondo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal z As Integer)
    MousePos.X = X
    MousePos.Y = Y
    TamanioTileMapa.value = TamanioTileMapa.value + z * 3
End Sub

Private Sub cerrar()
    I1 = 1
    I2 = 0
    InicioTick = GetTimer
    FinTick = InicioTick + 300
    Matar = True
End Sub

Private Sub Fondo_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    If Colision(X, Y, 0, 512, 0, 702) Then
        If TmpMapMouse <> MapSelected Then
            MapSelected = TmpMapMouse
            BotonAbrir.Enabled = True
        End If
    End If
End Sub

Private Sub BotonAbrir_Click()
    If MapSelected = 0 Then
        Beep
    Else

        Call frmMain.ABRIR_Mapa(MapSelected)
        
        cerrar
    End If
End Sub

Private Sub BotonCerrar_Click()
    cerrar
End Sub


Private Function vWindow_Render() As Boolean

    Dim TmpTick As Long
    
    TmpTick = GetTimer
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
    
    vW = 250 * Progreso
    vH = 416 * Progreso
    vX = CosInterp(0, D3DWindow.BackBufferWidth - vW - 10, Progreso)
    vY = CosInterp(0, 48, Progreso)
    alpha = Progreso * 128

    'Oscurezco la pantalla
    'Engine.Draw_FilledBox 0, 0, D3DWindow.BackBufferWidth, D3DWindow.BackBufferHeight, Alphas(Progreso * 200), 0, 0
    
    Fondo.SetColor Alphas(Progreso * 200)
    Fondo.Render
    
    'Dibujo la ventana
    
    Dim X%, Y%, MX%, MY%, MMX%, MMY%
    
    Dim TamanioTiles As Single
    Dim TamanioTilesMapa As Single
    TamanioTilesMapa = Round(TamanioTileMapa.value)
    
    TamanioTiles = TamanioTilesMapa - 1
    
    MX = (MousePos.X / (702 / 2)) * ME_Mundo.cantidadMapasX * 24
    MY = (MousePos.Y / (D3DWindow.BackBufferHeight / 2)) * ME_Mundo.cantidadMapasY * 24
    
    Dim escalaTexto As Single
    escalaTexto = TamanioTilesMapa / 48
    
    If MapasArrayCargado Then
        alpha = Progreso * 255
        For X = 1 To ME_Mundo.cantidadMapasX
            For Y = 1 To ME_Mundo.cantidadMapasY
                    If ME_Mundo.MapasArray(X, Y).Numero > 0 Then
                        MMX = (X * TamanioTilesMapa - MX + 1)
                        MMY = (Y * TamanioTilesMapa - MY + 1)
                        
                        If MMX < -TamanioTilesMapa Then GoTo NextLine
                        If MMX > D3DWindow.BackBufferWidth Then GoTo NextLine
                        If MMY < -TamanioTilesMapa Then GoTo NextAlto
                        If MMY > D3DWindow.BackBufferHeight Then GoTo NextAlto
                        
                        If ME_Mundo.MapasArray(X, Y).existe And pakMapasME.Cabezal_GetFileSize(ME_Mundo.MapasArray(X, Y).Numero) > 0 Then
                            If Not Colision(MousePos.X, MousePos.Y, MMY, MMY + TamanioTiles, MMX, MMX + TamanioTiles) Then
                                Engine.Draw_FilledBox MMX, MMY, TamanioTiles, TamanioTiles, ME_Mundo.MapasArray(X, Y).Color And &HFFFFFF Or Alphas(alpha), 0, 0
                            Else
                                Engine.Draw_FilledBox MMX - 1, MMY - 1, TamanioTiles + 2, TamanioTiles + 2, ME_Mundo.MapasArray(X, Y).Color And &HFFFFFF Or Alphas(alpha), mzYellow, 4
                                TmpMapMouse = ME_Mundo.MapasArray(X, Y).Numero
                            End If
                        Else
                            Engine.Draw_FilledBox MMX, MMY, TamanioTilesMapa, TamanioTilesMapa, (ME_Mundo.MapasArray(X, Y).Color) And &H4FFFFFFF, 0, 0
                        End If
                        
                        If TamanioTilesMapa > 16 Then
                            If ME_Mundo.MapasArray(X, Y).Puedo Then
                                Engine.text_render_graphic Str(ME_Mundo.MapasArray(X, Y).Numero), MMX + 6 * escalaTexto, MMY + 15 * escalaTexto, Alphas(alpha) Or &HFFFFFF
                            Else
                                Engine.text_render_graphic Str(ME_Mundo.MapasArray(X, Y).Numero), MMX + 6 * escalaTexto, MMY + 15 * escalaTexto, Alphas(alpha) Or &HCF0000
                            End If
                        End If
                    End If
NextAlto:
            Next Y
NextLine:
        Next X
    End If
    alpha = Progreso * 128
    
    Engine_GUI_Manager.GUI_RenderDialog vX, vY, vW, vH, "Abrir mapa " & MapSelected, Me, Progreso
    
    BotonAbrir.SetPos vX + vW - 108, vY + vH - 48
    BotonCerrar.SetPos vX + vW - 216, vY + vH - 48
    
    BotonAbrir.Render
    BotonCerrar.Render
        
End Function



Private Sub vWindow_SetPos(ByVal X As Integer, ByVal Y As Integer)

End Sub


Private Sub vWindow_Show()

End Sub
