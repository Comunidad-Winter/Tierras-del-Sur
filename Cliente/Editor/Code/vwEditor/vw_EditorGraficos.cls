VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vw_EditorGraficos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vWindow

Private vX As Integer
Private vY As Integer

Private Const DEFAULT_FONDO_VISTA_PREVIA As Long = &H7F7F7F

Private VisibilidadFondo                As Long ' 0 .. 255

Private WithEvents Fondo                As ControlMouseReader
Attribute Fondo.VB_VarHelpID = -1

Private WithEvents Controles            As vControles
Attribute Controles.VB_VarHelpID = -1

' Sombra Radio
Private WithEvents ShadowSize           As ControlInput
Attribute ShadowSize.VB_VarHelpID = -1

' Offset del Gráfico en la Grilla
Private WithEvents offsetX              As ControlInput
Attribute offsetX.VB_VarHelpID = -1
Private WithEvents offsetY              As ControlInput
Attribute offsetY.VB_VarHelpID = -1

' Check para Vista Previa del Gráfico que estoy modificando
Private WithEvents chkVistaPrevia       As ControlCheckBox
Attribute chkVistaPrevia.VB_VarHelpID = -1
' Color de fondo para la vista previa
Private WithEvents txtFondoVistaPrevia  As ControlInput
Attribute txtFondoVistaPrevia.VB_VarHelpID = -1

' Botonera
Private WithEvents BotonAplicar         As ControlBoton
Attribute BotonAplicar.VB_VarHelpID = -1
Private WithEvents BotonRestablecer     As ControlBoton
Attribute BotonRestablecer.VB_VarHelpID = -1

' Inforacion del grafico con el cual estoy trabajando
Private GrhDataItemOriginal         As GrhData
Private pGrhNumber                  As Long

' Auxiliar para dibujar el grafic
Private TmpGrh                      As Grh

' La vista previa del grafico (incluye grilla) se pone a partir de esta parte de la pantalla
Private MARGEN_VISTA_PREVIA_X As Integer
Private MARGEN_VISTA_PREVIA_Y As Integer

Public Event Aplicar()


Private LightSrcPos                 As D3DVECTOR ' NO SE USA
Private ShadowTBOX(6)               As TLVERTEX

'Implementacion de ventanas. No modificar
Private pNext As vWindow
Private pPrev As vWindow
    


Private Sub actualizarOffset()

Dim x As Integer
Dim y As Integer

x = maxi(-63, mini(63, val(offsetX.text)))
y = maxi(-63, mini(63, val(offsetY.text)))

Call Me_indexar_Graficos.establecerOffsetBruto(GrhData(pGrhNumber), x, y)

End Sub

Private Sub BotonAplicar_Click()
    RaiseEvent Aplicar
End Sub

Private Sub offsetX_ChangeValue(NuevoValor As Single)
    If NuevoValor < -63 Then
        offsetX = 63
    ElseIf NuevoValor > 63 Then
        offsetX = 63
    Else
        actualizarOffset
    End If
End Sub

Private Sub offsetY_ChangeValue(NuevoValor As Single)
    If NuevoValor < -63 Then
        offsetY = 63
    ElseIf NuevoValor > 63 Then
        offsetY = 63
    Else
        actualizarOffset
    End If
End Sub

Private Sub txtFondoVistaPrevia_change()
    Fondo.SetColor val("&H" & txtFondoVistaPrevia.text) Or Alphas(VisibilidadFondo)
End Sub

Private Sub txtFondoVistaPrevia_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
Dim bk As String
Dim c As RGBCOLOR

If Button = 2 Then
    bk = right$("000000" + txtFondoVistaPrevia, 6)
    
    frmMain.ColorDialog.flags = cdlCCRGBInit
    frmMain.ColorDialog.Color = RGB(CLng("&H" + mid$(bk, 1, 2)), CLng("&H" + mid$(bk, 3, 2)), CLng("&H" + mid$(bk, 5, 2)))
    frmMain.ColorDialog.ShowColor
    
    VBC2RGBC frmMain.ColorDialog.Color, c
    txtFondoVistaPrevia.text = Hex(CLng(c.r) * &H10000 + CLng(c.g) * &H100 + CLng(c.b))
    txtFondoVistaPrevia.Color = CLng("&H" & txtFondoVistaPrevia.text)
End If
End Sub

    Private Function vWindow_GetNext(vWindow As vWindow) As Boolean: Set vWindow = pNext: vWindow_GetNext = Not (pNext Is Nothing): End Function
    Private Function vWindow_GetPrev(vWindow As vWindow) As Boolean: Set vWindow = pPrev: vWindow_GetPrev = Not (pPrev Is Nothing): End Function

    Private Sub vWindow_SetNext(vWindow As vWindow): Set pNext = vWindow: End Sub

    Private Sub vWindow_SetPrev(vWindow As vWindow): Set pPrev = vWindow: End Sub
    
    Private Property Get vWindow_GetControl() As vControles: Set vWindow_GetControl = Controles: End Property

Private Sub ShadowSize_change()
    If CCVal(ShadowSize) < 0 Then ShadowSize = 0
    If CCVal(ShadowSize) > 255 Then ShadowSize = 255
    
    If pGrhNumber Then
        GrhData(pGrhNumber).SombrasSize = maxl(0, mini(val(ShadowSize.text), 255)) ' Enre 0 y 255
    End If
End Sub

Private Sub Fondo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
    If pGrhNumber <> 0 And Shift <> 0 Then
        offsetX = MARGEN_VISTA_PREVIA_X - x + offsetX
        offsetY = MARGEN_VISTA_PREVIA_Y - y + offsetY
        actualizarOffset
    End If
End Sub

Private Sub Fondo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer)
    LightSrcPos.x = x
    LightSrcPos.y = y
    
    If pGrhNumber Then
        GenerateLightBox ShadowTBOX, MARGEN_VISTA_PREVIA_X, MARGEN_VISTA_PREVIA_Y, x, y, 200, GrhData(pGrhNumber).SombrasSize
    End If
    
    If z = 0 Then Exit Sub
    
    If Shift Then
        VisibilidadFondo = VisibilidadFondo + z * 10
    Else
        VisibilidadFondo = VisibilidadFondo + z * 50
    End If
    
    If VisibilidadFondo > 255 Then VisibilidadFondo = 255
    If VisibilidadFondo < 0 Then VisibilidadFondo = 0
    
    Fondo.SetColor val("&H" & txtFondoVistaPrevia.text) Or Alphas(VisibilidadFondo)
End Sub

Private Sub Class_Initialize()
    Set Controles = New vControles
    
    
    Set ShadowSize = Controles.Add(New ControlInput)
    Set offsetX = Controles.Add(New ControlInput)
    Set offsetY = Controles.Add(New ControlInput)
    
    ' Vista Pevia
    Set chkVistaPrevia = Controles.Add(New ControlCheckBox)
    Set txtFondoVistaPrevia = Controles.Add(New ControlInput)
    
    ' Botones
    Set BotonAplicar = Controles.Add(New ControlBoton)
    Set BotonRestablecer = Controles.Add(New ControlBoton)
    
    ' Fondo
    Set Fondo = Controles.Add(New ControlMouseReader)
    
    VisibilidadFondo = 120
    
    Fondo.SetColor DEFAULT_FONDO_VISTA_PREVIA Or Alphas(VisibilidadFondo)
    Fondo.SetSize D3DWindow.BackBufferWidth, D3DWindow.BackBufferHeight
    
    ' Campos de Texto
    With ShadowSize
        .TipoTexto = eGuiTiposTexto.numerico
        .Label = "Radio sombra"
        .CantidadDecimales = 0
        .SetPos 15, 70
    End With

    With offsetX
        .TipoTexto = eGuiTiposTexto.numerico
        .Label = "Ajuste Grilla X"
        .CantidadDecimales = 0
        .SetPos 15, 110
    End With

    With offsetY
        .TipoTexto = eGuiTiposTexto.numerico
        .Label = "Ajuste Grilla Y"
        .CantidadDecimales = 0
        .SetPos 15, 130
    End With

    With chkVistaPrevia
        .text = "Vista Previa"
        .value = True
        .SetPos 15, 150
    End With

    With txtFondoVistaPrevia
        .text = UCase$(Hex$(DEFAULT_FONDO_VISTA_PREVIA))
        .Label = "Color del fondo:"
        .SetPos 120, 150
        .MaxLenght = 8
        .Filtro = "0123456789ABCDEFabcdef"
        .Color = DEFAULT_FONDO_VISTA_PREVIA
    End With
    
    ' Botones
    With BotonAplicar
        .text = "Aplicar"
        .SetPos 15, 30
    End With
    
    With BotonRestablecer
        .text = "Restablecer"
        .SetPos 170, 30
    End With


    
    ' Calculamos cual es el centro del Render
    MARGEN_VISTA_PREVIA_X = D3DWindow.BackBufferWidth / 2
    MARGEN_VISTA_PREVIA_Y = D3DWindow.BackBufferHeight / 2
End Sub

Private Sub vWindow_SetPos(ByVal x As Integer, ByVal y As Integer)

End Sub

Private Sub vWindow_Show()
    pGrhNumber = 0
End Sub

Private Sub vWindow_Hide()
    If pGrhNumber Then
        GrhData(pGrhNumber) = GrhDataItemOriginal
    End If
    pGrhNumber = 0
End Sub

Private Function vWindow_Render() As Boolean
    
    Controles.Render
    
    ' Esto no se que hace
    Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(TexturaSombra)
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLEFAN, 5, ShadowTBOX(0), TL_size
    
    ' Si no quiero vista previa o no tengo grafico, salgo
    If Not chkVistaPrevia.value Or pGrhNumber = 0 Then
        vWindow_Render = True
        Exit Function
    End If
    
    ' Para ayudarnos. Dibujamos una grilla
    Grh_Render_Simple_box GRILLA_TEXTURA, MARGEN_VISTA_PREVIA_X - 256 - 16, MARGEN_VISTA_PREVIA_Y - 256 - 16, &HFF009900, 512
     
    ' Dibujamos el Gráfico. Como uso la misma función que el juego, está ya incluye el agregado del offset
    Engine.Draw_Grh TmpGrh, MARGEN_VISTA_PREVIA_X - 16, MARGEN_VISTA_PREVIA_Y - 16, 1, UserPos.x, UserPos.y
    
    If Controles.ObjetoFocus Is ShadowSize And GrhData(pGrhNumber).SombrasSize > 0 Then
        If pGrhNumber Then
            ' Dibujo un cuadrado que simboliza el tamaño de la sombra generada
            Grh_Render_Solid &H7F000000, MARGEN_VISTA_PREVIA_X - GrhData(pGrhNumber).SombrasSize / 2, MARGEN_VISTA_PREVIA_Y - GrhData(pGrhNumber).SombrasSize / 2, GrhData(pGrhNumber).SombrasSize, GrhData(pGrhNumber).SombrasSize
        End If
    End If

    ' Dibujamos el punto en donde se genera la sombra
    Grh_Render_Solid &H7FFF00FF, MARGEN_VISTA_PREVIA_X - 1, MARGEN_VISTA_PREVIA_Y - 1, 3, 3
    
    
    
    vWindow_Render = True
End Function

Public Sub SetGrafico(ByVal GrhNumber As Long)
    Dim offsetNeto As Position
    
    ' Guardamos una copia para trabajar sobre ella
    GrhDataItemOriginal = GrhData(GrhNumber)
    pGrhNumber = GrhNumber
    
    InitGrh TmpGrh, GrhNumber
    
    ' Establecemos los datos en el formualrio
    Call Me_indexar_Graficos.calcularOffsetNeto(GrhData(GrhNumber), offsetNeto.x, offsetNeto.y)
    ShadowSize.text = GrhDataItemOriginal.SombrasSize
    
    offsetX.text = offsetNeto.x
    offsetY.text = offsetNeto.y
End Sub

Private Sub BotonRestablecer_Click()
    ' Volvemos a cargar la info del grafico original
    If pGrhNumber Then SetGrafico (pGrhNumber)
End Sub
