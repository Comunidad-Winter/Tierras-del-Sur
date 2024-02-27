VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vw_Part_Stream"
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

Private WithEvents BotonGuardar     As ControlBoton
Attribute BotonGuardar.VB_VarHelpID = -1
Private WithEvents BotonCancelar    As ControlBoton
Attribute BotonCancelar.VB_VarHelpID = -1
Private WithEvents BotonRestaurar   As ControlBoton
Attribute BotonRestaurar.VB_VarHelpID = -1

Private WithEvents BotonReiniciar   As ControlBoton
Attribute BotonReiniciar.VB_VarHelpID = -1

'Cosas del emisor

Private WithEvents Titulo           As ControlInput
Attribute Titulo.VB_VarHelpID = -1

Private WithEvents textura          As ControlInput
Attribute textura.VB_VarHelpID = -1

Private WithEvents varX             As ControlInput
Attribute varX.VB_VarHelpID = -1
Private WithEvents varY             As ControlInput
Attribute varY.VB_VarHelpID = -1
Private WithEvents spdX             As ControlInput
Attribute spdX.VB_VarHelpID = -1
Private WithEvents spdY             As ControlInput
Attribute spdY.VB_VarHelpID = -1
Private WithEvents accX             As ControlInput
Attribute accX.VB_VarHelpID = -1
Private WithEvents accY             As ControlInput
Attribute accY.VB_VarHelpID = -1

Private WithEvents OffX             As ControlInput
Attribute OffX.VB_VarHelpID = -1
Private WithEvents OffY             As ControlInput
Attribute OffY.VB_VarHelpID = -1

Private WithEvents fricc            As ControlInput
Attribute fricc.VB_VarHelpID = -1
Private WithEvents ancho            As ControlInput
Attribute ancho.VB_VarHelpID = -1

Private WithEvents Cantidad         As ControlInput
Attribute Cantidad.VB_VarHelpID = -1
Private WithEvents ModTimer         As ControlInput
Attribute ModTimer.VB_VarHelpID = -1
Private WithEvents FactorTiempo     As ControlInput
Attribute FactorTiempo.VB_VarHelpID = -1

Private WithEvents rndA             As ControlInput
Attribute rndA.VB_VarHelpID = -1
Private WithEvents af               As ControlInput
Attribute af.VB_VarHelpID = -1

Private WithEvents size             As ControlInput
Attribute size.VB_VarHelpID = -1

Private WithEvents VidaParticula    As ControlInput
Attribute VidaParticula.VB_VarHelpID = -1
Private WithEvents VidaParticulaRnd As ControlInput
Attribute VidaParticulaRnd.VB_VarHelpID = -1

Private WithEvents CantidadAEmitir  As ControlInput
Attribute CantidadAEmitir.VB_VarHelpID = -1




Private WithEvents AlphaCoseno      As ControlCheckBox
Attribute AlphaCoseno.VB_VarHelpID = -1

Private WithEvents Linea            As ControlCheckBox
Attribute Linea.VB_VarHelpID = -1
Private WithEvents Crece            As ControlCheckBox
Attribute Crece.VB_VarHelpID = -1

Private WithEvents AccNormal        As ControlCheckBox
Attribute AccNormal.VB_VarHelpID = -1
Private WithEvents AccSin           As ControlCheckBox
Attribute AccSin.VB_VarHelpID = -1
Private WithEvents AccCos           As ControlCheckBox
Attribute AccCos.VB_VarHelpID = -1

Private WithEvents PosNormal        As ControlCheckBox
Attribute PosNormal.VB_VarHelpID = -1
Private WithEvents PosSin           As ControlCheckBox
Attribute PosSin.VB_VarHelpID = -1
Private WithEvents PosCos           As ControlCheckBox
Attribute PosCos.VB_VarHelpID = -1

Private WithEvents VelSinCos        As ControlCheckBox
Attribute VelSinCos.VB_VarHelpID = -1
Private WithEvents movsin           As ControlCheckBox
Attribute movsin.VB_VarHelpID = -1

Private WithEvents BlendOne         As ControlCheckBox
Attribute BlendOne.VB_VarHelpID = -1
Private WithEvents blendadd         As ControlCheckBox
Attribute blendadd.VB_VarHelpID = -1

Private WithEvents ColorInicial     As ControlInput
Attribute ColorInicial.VB_VarHelpID = -1
Private WithEvents ColorFinal       As ControlInput
Attribute ColorFinal.VB_VarHelpID = -1


'/COsas del emisor

Private WithEvents MouseFrame       As ControlMouseReader
Attribute MouseFrame.VB_VarHelpID = -1


Private datos As ParticleStream
Private datosOriginales As ParticleStream

Public Emisor As Engine_Particle_Stream
Private TituloOriginal As String

Private Controles As vControles


Private OffsetYReal As New clsAlpha


'Implementacion de ventanas. No modificar
' LA PUTA QUE LO PARIOOOOOOO PORQUE NO LO PUEDO EXTENDER Y YA!?!?. LA VIDA SERIA MEJOR SI SE PUDIERA EXTENDER UNA CLASE EN ESTE VB DE MIERDA
    Private pNext As vWindow
    Private pPrev As vWindow

Private Sub BotonReiniciar_Click()
    Emisor.reset
End Sub

Private Sub CantidadAEmitir_ChangeValue(NuevoValor As Single)
    datos.CantidadAEmitir = NuevoValor
    SetDatos
End Sub

Private Sub Class_Initialize()
    Set Controles = New vControles
    
    Set BotonGuardar = Controles.Add(New ControlBoton)
    Set BotonCancelar = Controles.Add(New ControlBoton)
    Set BotonRestaurar = Controles.Add(New ControlBoton)
    
    Set BotonReiniciar = Controles.Add(New ControlBoton)
    
    Set Titulo = Controles.Add(New ControlInput)
    
    Set textura = Controles.Add(New ControlInput)
    
    Set varX = Controles.Add(New ControlInput)
    Set varY = Controles.Add(New ControlInput)
    Set spdX = Controles.Add(New ControlInput)
    Set spdY = Controles.Add(New ControlInput)
    Set accX = Controles.Add(New ControlInput)
    Set accY = Controles.Add(New ControlInput)
    
    Set OffX = Controles.Add(New ControlInput)
    Set OffY = Controles.Add(New ControlInput)
    
    Set size = Controles.Add(New ControlInput)
    
    Set Cantidad = Controles.Add(New ControlInput)
    Set ModTimer = Controles.Add(New ControlInput)
    Set FactorTiempo = Controles.Add(New ControlInput)
    
    Set ancho = Controles.Add(New ControlInput)
    Set fricc = Controles.Add(New ControlInput)
        
    Set rndA = Controles.Add(New ControlInput)
    Set af = Controles.Add(New ControlInput)
    
    Set VidaParticula = Controles.Add(New ControlInput)
    Set VidaParticulaRnd = Controles.Add(New ControlInput)
    
    Set AlphaCoseno = Controles.Add(New ControlCheckBox)
    Set Linea = Controles.Add(New ControlCheckBox)
    Set Crece = Controles.Add(New ControlCheckBox)
    
    Set AccNormal = Controles.Add(New ControlCheckBox)
    Set AccSin = Controles.Add(New ControlCheckBox)
    Set AccCos = Controles.Add(New ControlCheckBox)
    
    Set PosNormal = Controles.Add(New ControlCheckBox)
    Set PosSin = Controles.Add(New ControlCheckBox)
    Set PosCos = Controles.Add(New ControlCheckBox)
    
    Set VelSinCos = Controles.Add(New ControlCheckBox)
    Set movsin = New ControlCheckBox 'Controles.Add(New ControlCheckBox)

    Set BlendOne = Controles.Add(New ControlCheckBox)
    Set blendadd = Controles.Add(New ControlCheckBox)
    
    Set ColorInicial = Controles.Add(New ControlInput)
    Set ColorFinal = Controles.Add(New ControlInput)
    Set CantidadAEmitir = Controles.Add(New ControlInput)
    
    Set MouseFrame = Controles.Add(New ControlMouseReader)
    
    
    AlphaCoseno.text = "Desvanecer"
    Linea.text = "Lineas"
    Crece.text = "Tamaño creciente"
    
    AccNormal.text = "Mov. Normal"
    AccSin.text = "Mov. IN"
    AccCos.text = "Mov. OUT"
    
    PosNormal.text = "Pos. Normal"
    PosSin.text = "Pos. Anillo"
    PosCos.text = "Pos. Circulo"
    
    VelSinCos.text = "Vel Dinámica"
    
    movsin.text = "Mov Dinámico"
    
    BlendOne.text = "Colorea fondo"
    blendadd.text = "Brilla fondo"
    
    CantidadAEmitir.text = "Cantidad a emitir"
    CantidadAEmitir.TipoTexto = eGuiTiposTexto.NUMERICO
    
    
    Titulo.Label = "Titulo"
    
    textura.TipoTexto = eGuiTiposTexto.NUMERICO
    
    textura.Label = "Textura"
    
    varX.TipoTexto = eGuiTiposTexto.NUMERICO
    
    varX.Label = "Random X"
    
    varY.TipoTexto = eGuiTiposTexto.NUMERICO
    
    varY.Label = "Random Y"
    spdX.TipoTexto = eGuiTiposTexto.NUMERICO
    spdX.Label = "Velocidad X"
    spdY.TipoTexto = eGuiTiposTexto.NUMERICO
    spdY.Label = "Velocidad Y"
    accX.TipoTexto = eGuiTiposTexto.NUMERICO
    accX.Label = "Acel. X"
    accY.TipoTexto = eGuiTiposTexto.NUMERICO
    accY.Label = "Acel. Y"
    
    OffX.TipoTexto = eGuiTiposTexto.NUMERICO
    OffX.Label = "Offset X"
    OffY.TipoTexto = eGuiTiposTexto.NUMERICO
    OffY.Label = "Offset Y"
    
    size.TipoTexto = eGuiTiposTexto.NUMERICO
    size.Label = "Tamaño"
    
    Cantidad.TipoTexto = eGuiTiposTexto.NUMERICO
    Cantidad.Label = "Cantidad"
    
    ModTimer.TipoTexto = eGuiTiposTexto.NUMERICO
    ModTimer.Label = "Mod timer"
    
    rndA.TipoTexto = eGuiTiposTexto.NUMERICO
    rndA.Label = "Random vida"
    
    af.TipoTexto = eGuiTiposTexto.NUMERICO
    af.Label = "Factor Alpha"
    
    FactorTiempo.TipoTexto = eGuiTiposTexto.NUMERICO
    FactorTiempo.Label = "Factor Progreso"
    
    FactorTiempo.CantidadDecimales = 4
    ModTimer.CantidadDecimales = 4
    
    rndA.CantidadDecimales = 4
    ModTimer.CantidadDecimales = 4
    
    Cantidad.CantidadDecimales = 0
    textura.CantidadDecimales = 0
    OffX.CantidadDecimales = 0
    OffY.CantidadDecimales = 0
    
    VidaParticula.CantidadDecimales = 0
    
    VidaParticula.TipoTexto = eGuiTiposTexto.NUMERICO
    VidaParticulaRnd.TipoTexto = eGuiTiposTexto.NUMERICO
    
    VidaParticula.Label = "Vida media"
    VidaParticulaRnd.Label = "Modif vida"
    
    ancho.TipoTexto = eGuiTiposTexto.NUMERICO
    fricc.TipoTexto = eGuiTiposTexto.NUMERICO
    
    ancho.Label = "Ancho"
    fricc.Label = "Fricción"
    
    
    vX = 0
    vY = 0
    
    vW = 256 + 128
    vH = 512
    
    BotonGuardar.SetPos vX + vW - 75, vY + vH - 72
    BotonRestaurar.SetPos vX + vW - 75, vY + vH - 52
    BotonCancelar.SetPos vX + vW - 75, vY + vH - 32
    BotonReiniciar.SetPos vX + 11, vY + vH - 32
    
    BotonCancelar.SetSize 64, 16
    BotonGuardar.SetSize 64, 16
    BotonRestaurar.SetSize 64, 16
    BotonReiniciar.SetSize 64, 16
    
    BotonCancelar.text = "Cancelar"
    BotonGuardar.text = "Guardar"
    BotonRestaurar.text = "Restaurar"
    BotonReiniciar.text = "Aplicar"
    
    ColorInicial.Label = "Color Inicial"
    ColorFinal.Label = "Color Final"
    
    OffsetYReal.Speed = 400
    OffsetYReal.InitialValue = 0
    
    ColorInicial.MaxLenght = 8
    ColorFinal.MaxLenght = 8
    
    ColorInicial.Filtro = "0123456789ABCDEFabcdef"
    ColorFinal.Filtro = "0123456789ABCDEFabcdef"
End Sub

Private Sub ColorFinal_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal y As Integer)
If Button = 2 Then
    Dim bk As String
    bk = Right$("000000" + ColorFinal, 6)
    frmMain.ColorDialog.flags = cdlCCRGBInit
    frmMain.ColorDialog.Color = RGB(CLng("&H" + mid$(bk, 1, 2)), CLng("&H" + mid$(bk, 3, 2)), CLng("&H" + mid$(bk, 5, 2)))
    frmMain.ColorDialog.ShowColor
    Dim c As RGBCOLOR
    VBC2RGBC frmMain.ColorDialog.Color, c
    ColorFinal = Hex(CLng(c.r) * &H10000 + CLng(c.g) * &H100 + CLng(c.b))
End If
End Sub

Private Sub ColorInicial_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal y As Integer)
If Button = 2 Then
    
    Dim bk As String
    bk = Right$("000000" + ColorInicial, 6)
    frmMain.ColorDialog.flags = cdlCCRGBInit
    frmMain.ColorDialog.Color = RGB(CLng("&H" + mid$(bk, 1, 2)), CLng("&H" + mid$(bk, 3, 2)), CLng("&H" + mid$(bk, 5, 2)))
    frmMain.ColorDialog.ShowColor
    Dim c As RGBCOLOR
    VBC2RGBC frmMain.ColorDialog.Color, c
    ColorInicial = Hex(CLng(c.r) * &H10000 + CLng(c.g) * &H100 + CLng(c.b))
End If
End Sub

Private Sub MouseFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal y As Integer, ByVal z As Integer)
If z Then
    'OffsetYReal = mins(OffsetYReal + z * 64, 15)
End If
End Sub

Private Sub movsin_Click()
    If Not movsin Then
        datos.tipo = datos.tipo And Not Caracteristicas.movimiento_sinoudal
    Else
        datos.tipo = datos.tipo Or Caracteristicas.movimiento_sinoudal
    End If
    SetDatos
End Sub

Private Sub VelSinCos_Click()
    If Not VelSinCos Then
        datos.tipo = datos.tipo And Not Caracteristicas.spd_trig
    Else
        datos.tipo = datos.tipo Or Caracteristicas.spd_trig
    End If
    SetDatos
End Sub

    Private Property Get vWindow_GetControl() As vControles
        Set vWindow_GetControl = Controles
    End Property

    Private Function vWindow_GetNext(vWindow As vWindow) As Boolean: Set vWindow = pNext: vWindow_GetNext = Not (pNext Is Nothing): End Function
    Private Function vWindow_GetPrev(vWindow As vWindow) As Boolean: Set vWindow = pPrev: vWindow_GetPrev = Not (pPrev Is Nothing): End Function
    Private Sub vWindow_SetNext(vWindow As vWindow): Set pNext = vWindow: End Sub
    Private Sub vWindow_SetPrev(vWindow As vWindow): Set pPrev = vWindow: End Sub
'/Implementacion

Public Sub BindEmisor(e As Engine_Particle_Stream)
    Set Emisor = e
    
    TituloOriginal = Trim$(e.GetNombre)
    Titulo = TituloOriginal
    
    datos = Emisor.GetEmisor
    datosOriginales = datos
    
    LlenarCampos
End Sub

Friend Function GetDatos() As ParticleStream
    GetDatos = datos
End Function

Private Sub vWindow_Hide()
    GUI_Quitar Me
End Sub

Private Sub vWindow_SetPos(ByVal X As Integer, ByVal y As Integer)

End Sub

Private Sub vWindow_Show()
    GUI_SetFocus Me
End Sub

Private Function vWindow_Render() As Boolean
    vWindow_Render = True
    
    MouseFrame.SetColor &HFF222222
    MouseFrame.SetPos vX, vY
    MouseFrame.SetSize vW, vH
    
    Dim offsetY As Single
    offsetY = 15 'OffsetYReal
    
    Titulo.SetPos vX + 11, vY + offsetY
    
    textura.SetPos vX + 11, vY + offsetY + 1 * 20
    
    varX.SetPos vX + 11, vY + offsetY + 2 * 20
    varY.SetPos vX + 11, vY + offsetY + 3 * 20
    spdX.SetPos vX + 11, vY + offsetY + 4 * 20
    spdY.SetPos vX + 11, vY + offsetY + 5 * 20
    accX.SetPos vX + 11, vY + offsetY + 6 * 20
    accY.SetPos vX + 11, vY + offsetY + 7 * 20
    
    OffX.SetPos vX + 11, vY + offsetY + 8 * 20
    OffY.SetPos vX + 11, vY + offsetY + 9 * 20
    
    size.SetPos vX + 11, vY + offsetY + 10 * 20
    
    Cantidad.SetPos vX + 11, vY + offsetY + 11 * 20
    ModTimer.SetPos vX + 11, vY + offsetY + 12 * 20
    FactorTiempo.SetPos vX + 11, vY + offsetY + 13 * 20
    
    ancho.SetPos vX + 11, vY + offsetY + 14 * 20
    fricc.SetPos vX + 11, vY + offsetY + 15 * 20
    
    rndA.SetPos vX + 11, vY + offsetY + 16 * 20
    af.SetPos vX + 11, vY + offsetY + 17 * 20
    
    VidaParticula.SetPos vX + 11, vY + offsetY + 18 * 20
    VidaParticulaRnd.SetPos vX + 11, vY + offsetY + 19 * 20
    
    ColorFinal.SetPos vX + 11, vY + offsetY + 20 * 20
    ColorInicial.SetPos vX + 11, vY + offsetY + 21 * 20
    
    vX = vX + (300 - 64)
    
    
    AlphaCoseno.SetPos vX + 11, vY + offsetY + 0 * 20
    Linea.SetPos vX + 11, vY + offsetY + 1 * 20
    Crece.SetPos vX + 11, vY + offsetY + 2 * 20
    
    AccNormal.SetPos vX + 11, vY + offsetY + 4 * 20
    AccSin.SetPos vX + 11, vY + offsetY + 5 * 20
    AccCos.SetPos vX + 11, vY + offsetY + 6 * 20
    
    PosNormal.SetPos vX + 11, vY + offsetY + 8 * 20
    PosSin.SetPos vX + 11, vY + offsetY + 9 * 20
    PosCos.SetPos vX + 11, vY + offsetY + 10 * 20
    
    VelSinCos.SetPos vX + 11, vY + offsetY + 12 * 20
    
    movsin.SetPos vX + 11, vY + offsetY + 14 * 20
    
    BlendOne.SetPos vX + 11, vY + offsetY + 16 * 20
    blendadd.SetPos vX + 11, vY + offsetY + 17 * 20
    
    CantidadAEmitir.SetPos vX + 11, vY + offsetY + 18 * 20
    

    
    vX = vX - (300 - 64)
    
    Controles.Render
    
End Function

Private Sub SetDatos()
    datos.AlphaInicial = 1
    'datos.alpha_coseno = 1
        
    'datos.StartColor = &HFFFFFFFF
    'datos.EndColor = &HFFFFAACC

    Emisor.SetNombre Titulo
    Emisor.SetEmisor datos
End Sub

Private Sub LlenarCampos()
    textura = datos.texture
    
    varX = datos.rnd_x
    varY = datos.rnd_y
    
    rndA = datos.rnd_alpha_factor
    af = datos.alpha_factor
    
    spdX = datos.spdX
    spdY = datos.spdY
    accX = datos.accX
    accY = datos.accY
    
    OffX = datos.offset_x
    OffY = datos.offset_y
    
    size = datos.size
    
    Cantidad = datos.NumOfParticles
    ModTimer = datos.mod_timer
    FactorTiempo = datos.progress_factor
    
    CantidadAEmitir = datos.CantidadAEmitir
    
    ancho = datos.VarZ
    fricc = datos.fricc
    
    If datos.tipo And Caracteristicas.acc_SS Then
        AccSin = True
        AccCos = False
        AccNormal = False
    Else
        If datos.tipo And Caracteristicas.acc_trig Then
            AccNormal = False
            AccSin = False
            AccCos = True
        Else
            AccNormal = True
            AccSin = False
            AccCos = False
        End If
    End If
    
    If datos.tipo And Caracteristicas.pos_SS Then
        PosSin = True
        PosCos = False
        PosNormal = False
    Else
        If datos.tipo And Caracteristicas.pos_trig Then
            PosNormal = False
            PosSin = False
            PosCos = True
        Else
            PosNormal = True
            PosSin = False
            PosCos = False
        End If
    End If
    
    Crece = datos.Crece <> 0
    
    Linea = (datos.Line = 1)
    
    ColorInicial = UCase$(Hex$(datos.StartColor))
    ColorFinal = UCase$(Hex$(datos.EndColor))
    
    blendadd = (datos.blend_mode And 1)
    
    BlendOne = (datos.blend_mode And 2)
    
    movsin = datos.tipo And Caracteristicas.movimiento_sinoudal
    VelSinCos = datos.tipo And Caracteristicas.spd_trig
    
    AlphaCoseno = datos.alpha_coseno = 1
    
    VidaParticula = datos.life
    VidaParticulaRnd = datos.RandomLife
End Sub

Private Sub Cantidad_ChangeValue(NuevoValor As Single)
    NuevoValor = bounds(1, 10000, NuevoValor)
    If NuevoValor <= 0 Then
        datos.NumOfParticles = 1
        Cantidad = "1"
    Else
        datos.NumOfParticles = NuevoValor
    End If
    SetDatos
End Sub

Private Sub fricc_ChangeValue(NuevoValor As Single)
    NuevoValor = boundsf(0, 10, NuevoValor)
    datos.fricc = NuevoValor
    SetDatos
End Sub

Private Sub ModTimer_ChangeValue(NuevoValor As Single)
    NuevoValor = boundsf(0.01, 50, NuevoValor)
    datos.mod_timer = NuevoValor
    SetDatos
End Sub

Private Sub offX_ChangeValue(NuevoValor As Single)
    datos.offset_x = NuevoValor
    SetDatos
End Sub

Private Sub offY_ChangeValue(NuevoValor As Single)
    datos.offset_y = NuevoValor
    SetDatos
End Sub

Private Sub Size_ChangeValue(NuevoValor As Single)
    datos.size = NuevoValor
    SetDatos
End Sub

Private Sub accX_ChangeValue(NuevoValor As Single)
    datos.accX = NuevoValor
    SetDatos
End Sub

Private Sub accY_ChangeValue(NuevoValor As Single)
    datos.accY = NuevoValor
    SetDatos
End Sub

Private Sub spdX_ChangeValue(NuevoValor As Single)
    datos.spdX = NuevoValor
    SetDatos
End Sub

Private Sub spdY_ChangeValue(NuevoValor As Single)
    datos.spdY = NuevoValor
    SetDatos
End Sub

Private Sub textura_ChangeValue(NuevoValor As Single)
    NuevoValor = bounds(1, &H7FFF, NuevoValor)
    datos.texture = NuevoValor
    SetDatos
End Sub

Private Sub varX_ChangeValue(NuevoValor As Single)
    datos.rnd_x = NuevoValor
    SetDatos
End Sub

Private Sub varY_ChangeValue(NuevoValor As Single)
    datos.rnd_y = NuevoValor
    SetDatos
End Sub

Private Sub FactorTiempo_ChangeValue(NuevoValor As Single)
    datos.progress_factor = NuevoValor
    SetDatos
End Sub


Private Sub rndA_ChangeValue(NuevoValor As Single)
    NuevoValor = boundsf(0, 1, NuevoValor)
    datos.rnd_alpha_factor = NuevoValor
    SetDatos
End Sub

Private Sub af_ChangeValue(NuevoValor As Single)
    NuevoValor = boundsf(0, 1, NuevoValor)
    datos.alpha_factor = NuevoValor
    SetDatos
End Sub

Private Sub VidaParticula_ChangeValue(NuevoValor As Single)
    NuevoValor = bounds(0, 16000, NuevoValor)
    datos.life = NuevoValor
    SetDatos
End Sub

Private Sub VidaParticulaRnd_ChangeValue(NuevoValor As Single)
    NuevoValor = boundsf(0, 2, NuevoValor)
    datos.RandomLife = NuevoValor
    SetDatos
End Sub

Private Sub AlphaCoseno_Click()
    datos.alpha_coseno = IIf(AlphaCoseno.Value, 1, 0)
    SetDatos
End Sub

Private Sub PosCos_Click()
    PosSin = False
    PosNormal = False
    datos.tipo = datos.tipo And (Not Caracteristicas.pos_SS) Or Caracteristicas.pos_trig
    SetDatos
End Sub

Private Sub PosSin_Click()
    PosCos = False
    PosNormal = False
    datos.tipo = datos.tipo Or Caracteristicas.pos_SS
    SetDatos
End Sub

Private Sub PosNormal_Click()
    PosSin = False
    PosCos = False
    datos.tipo = datos.tipo And (Not (Caracteristicas.pos_SS Or Caracteristicas.pos_trig))
    SetDatos
End Sub

Private Sub AccCos_Click()
    AccSin = False
    AccNormal = False
    datos.tipo = datos.tipo And (Not Caracteristicas.acc_SS) Or Caracteristicas.acc_trig
    SetDatos
End Sub

Private Sub AccNormal_Click()
    If AccNormal Then
        AccCos = False
        AccSin = False
        datos.tipo = datos.tipo And (Not (Caracteristicas.acc_SS Or Caracteristicas.acc_trig))
        SetDatos
    End If
End Sub

Private Sub AccSin_Click()
    If AccSin Then
        AccCos = False
        AccNormal = False
        datos.tipo = datos.tipo Or Caracteristicas.acc_SS
        SetDatos
    End If
End Sub

Private Sub ancho_ChangeValue(NuevoValor As Single)
    NuevoValor = Abs(NuevoValor)
    datos.VarZ = NuevoValor
    SetDatos
End Sub

Private Sub ColorFinal_Change()
On Error Resume Next
    datos.EndColor = CLng("&H" & ColorFinal)
    ColorFinal.Color = datos.EndColor Or &HFF000000
    SetDatos
End Sub

Private Sub ColorInicial_Change()
On Error Resume Next
    datos.StartColor = CLng("&H" & ColorInicial)
    ColorInicial.Color = datos.StartColor Or &HFF000000
    SetDatos
End Sub

Private Sub Crece_Click()
    datos.Crece = IIf(Crece, 1, 0)
    SetDatos
End Sub

Private Sub Linea_Click()
    datos.Line = IIf(Linea, 1, 0)
    SetDatos
End Sub

Private Sub blendadd_Click()
    datos.blend_mode = IIf(BlendOne, 2, 0) Or IIf(blendadd, 1, 0)
    SetDatos
End Sub

Private Sub BlendOne_Click()
    datos.blend_mode = IIf(BlendOne, 2, 0) Or IIf(blendadd, 1, 0)
    SetDatos
End Sub







' BOTONES !



Private Sub BotonCancelar_Click()
    Emisor.CancelarEdicion
End Sub

Private Sub BotonGuardar_Click()
    SetDatos
    Emisor.GuardarEdicion
    'GUI_Quitar Me
End Sub

Private Sub BotonRestaurar_Click()
    Titulo = TituloOriginal
    datos = datosOriginales
    SetDatos
    LlenarCampos
End Sub
