VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VwCrearPersonajePaso2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements vWindow

Option Explicit

Private visible As Boolean
Private vControles As vControles

Private WithEvents ImagenHombre As ControlCheckBox
Attribute ImagenHombre.VB_VarHelpID = -1
Private WithEvents ImagenMujer As ControlCheckBox
Attribute ImagenMujer.VB_VarHelpID = -1
Private RadioButtons(1 To 2) As ControlCheckBox

Private WithEvents BotonSiguienteRaza As ControlBoton
Attribute BotonSiguienteRaza.VB_VarHelpID = -1
Private WithEvents BotonAnteriorRaza As ControlBoton
Attribute BotonAnteriorRaza.VB_VarHelpID = -1

Private FondoVentana As ControlImage
Private FondoRaza As ControlImage
Private ImagenRaza As ControlImage
Private NombreRaza As ControlLabel
Private TituloRaza As ControlLabel
Private FondoTituloRaza As ControlImage
Private FondoTituloGenero As ControlImage

Private lineaSeparadora As ControlImage


Private TituloGenero As ControlLabel
Private TextoHombre As ControlLabel
Private TextoMujer As ControlLabel

Private numRazaActual As Integer
Private razaActual As tRaza

Private BarraFuerza As ControlBarra
Private BarraAgilidad As ControlBarra
Private BarraInteligencia As ControlBarra
Private BarraCarisma As ControlBarra
Private BarraConsitucion As ControlBarra

Private LabelError As ControlLabel

Private WithEvents BotonSiguiente As ControlBoton
Attribute BotonSiguiente.VB_VarHelpID = -1
Private WithEvents BotonAtras As ControlBoton
Attribute BotonAtras.VB_VarHelpID = -1

Public Event Atras()
Public Event siguiente(Genero As eGenero, Raza As eRaza)

Private vX As Integer
Private vY As Integer
Private vW As Integer
Private vH As Integer

'Implementacion de ventanas. No modificar
Private pNext As vWindow
Private pPrev As vWindow

Public Sub setDatos(Genero As eGenero, ByVal Raza As eRaza)
   
    If Raza > 0 Then
       numRazaActual = getRazaNumero(Raza)
       Call CambiaSeleccion
    End If
    
    If Genero > 0 Then
        If Genero = Hombre Then
            ImagenHombre.value = True
            Call changeRadioButton(RadioButtons, ImagenMujer)
        Else
            ImagenMujer.value = True
            Call changeRadioButton(RadioButtons, ImagenMujer)
        End If
    End If

End Sub

Private Sub BotonAtras_Click()
    Call Sonido_Play(SND_CLICK)
    RaiseEvent Atras
End Sub

Private Sub BotonSiguiente_Click()
    Call Sonido_Play(SND_CLICK)
    
    If Not validar Then Exit Sub
    Dim Genero As eGenero
    Dim Raza As eRaza
    
    If ImagenHombre.value Then
        Genero = Hombre
    Else
        Genero = Mujer
    End If
    
    RaiseEvent siguiente(Genero, razaActual.id)
End Sub

Private Sub Class_Initialize()
    Call modCrearPersonaje.initRazas
    
    Set vControles = New vControles
    
    Set BotonAtras = vControles.Add(New ControlBoton)
    Set BotonSiguiente = vControles.Add(New ControlBoton)
    
    Set BarraFuerza = vControles.Add(New ControlBarra)
    Set BarraAgilidad = vControles.Add(New ControlBarra)
    Set BarraInteligencia = vControles.Add(New ControlBarra)
    Set BarraCarisma = vControles.Add(New ControlBarra)
    Set BarraConsitucion = vControles.Add(New ControlBarra)
    Set ImagenRaza = vControles.Add(New ControlImage)
    Set FondoRaza = vControles.Add(New ControlImage)
    
    Set BotonAnteriorRaza = vControles.Add(New ControlBoton)
    Set BotonSiguienteRaza = vControles.Add(New ControlBoton)
    
    Set NombreRaza = vControles.Add(New ControlLabel)
    Set TituloRaza = vControles.Add(New ControlLabel)
    
    Set TituloGenero = vControles.Add(New ControlLabel)
    
    Set ImagenHombre = vControles.Add(New ControlCheckBox)
    Set ImagenMujer = vControles.Add(New ControlCheckBox)
    
    Set RadioButtons(1) = ImagenHombre
    Set RadioButtons(2) = ImagenMujer
    
    Set TextoMujer = vControles.Add(New ControlLabel)
    Set TextoHombre = vControles.Add(New ControlLabel)
    
    Set lineaSeparadora = vControles.Add(New ControlImage)
    
    Set FondoTituloRaza = vControles.Add(New ControlImage)
    Set FondoTituloGenero = vControles.Add(New ControlImage)
     
    Set FondoVentana = vControles.Add(New ControlImage)
    
    FondoVentana.grhId = 20352
    
    TextoMujer.text = "Mujer"
    TextoMujer.SetStyle 3, mzInterfaceColor1, mzInterfaceColor2, dVertical
    TextoMujer.SetCentrado True
    
    TextoHombre.text = "Hombre"
    TextoHombre.SetStyle 3, mzInterfaceColor1, mzInterfaceColor2, dVertical
    TextoHombre.SetCentrado True
    
    FondoRaza.grhId = 20351
    FondoTituloRaza.grhId = 21868
    FondoTituloGenero.grhId = 21868

    BotonSiguienteRaza.textura = 3361
    BotonSiguienteRaza.SetSize 20, 20
    
    BotonAnteriorRaza.textura = 3362
    BotonAnteriorRaza.SetSize 20, 20
     
    NombreRaza.SetCentrado True
    NombreRaza.SetStyle 2, mzInterfaceColor1, mzInterfaceColor2, dVertical
 
    ImagenHombre.textura = 3407
    ImagenHombre.IControl_SetSize 64, 64
       
    ImagenMujer.textura = 3408
    ImagenMujer.IControl_SetSize 64, 64
    
    lineaSeparadora.grhId = 21867
    
    BarraFuerza.Nombre = "Fuerza"
    BarraAgilidad.Nombre = "Agilidad"
    BarraInteligencia.Nombre = "Inteligencia"
    BarraCarisma.Nombre = "Carisma"
    BarraConsitucion.Nombre = "Constitución"
    
    TituloRaza.text = "Elegi tu Raza"
    TituloRaza.SetStyle 4, mzInterfaceColor1, mzInterfaceColor2, dVertical
    TituloRaza.SetCentrado True
    
    TituloGenero.SetStyle 4, mzInterfaceColor1, mzInterfaceColor2, dVertical
    TituloGenero.SetCentrado True
    TituloGenero.text = "Elegi tu Genero"
    BarraFuerza.min = 15
    BarraFuerza.max = 22
    
    BarraAgilidad.min = 15
    BarraAgilidad.max = 22
        
    BarraInteligencia.min = 15
    BarraInteligencia.max = 22
    
    BarraCarisma.min = 15
    BarraCarisma.max = 22
        
    BarraConsitucion.min = 15
    BarraConsitucion.max = 22
    
    BarraConsitucion.SetStyle 20461, 3, mzInterfaceColor1, mzInterfaceColor2, dVertical
    BarraInteligencia.SetStyle 20461, 3, mzInterfaceColor1, mzInterfaceColor2, dVertical
    BarraCarisma.SetStyle 20461, 3, mzInterfaceColor1, mzInterfaceColor2, dVertical
    BarraAgilidad.SetStyle 20461, 3, mzInterfaceColor1, mzInterfaceColor2, dVertical
    BarraFuerza.SetStyle 20461, 3, mzInterfaceColor1, mzInterfaceColor2, dVertical
    
    BotonAtras.text = "Volver"
    BotonSiguiente.text = "Siguiente"
    
    numRazaActual = 1
    
    Call CambiaSeleccion
    Call actualizarPosiciones
End Sub

Private Sub ImagenHombre_Click()
    changeRadioButton RadioButtons, ImagenHombre
End Sub

Private Sub ImagenMujer_Click()
    changeRadioButton RadioButtons, ImagenMujer
End Sub

Private Property Get vWindow_GetControl() As vControles
    Set vWindow_GetControl = vControles
End Property
Private Function vWindow_GetNext(vWindow As vWindow) As Boolean: Set vWindow = pNext: vWindow_GetNext = Not (pNext Is Nothing): End Function
Private Function vWindow_GetPrev(vWindow As vWindow) As Boolean: Set vWindow = pPrev: vWindow_GetPrev = Not (pPrev Is Nothing): End Function

Private Property Get vWindow_IsVisible() As Boolean
   vWindow_IsVisible = visible
End Property

Private Sub vWindow_SetNext(vWindow As vWindow): Set pNext = vWindow: End Sub
Private Sub vWindow_SetPrev(vWindow As vWindow): Set pPrev = vWindow: End Sub

Public Sub vWindow_Hide()
    visible = False
End Sub

Public Sub vWindow_Show()
    visible = True
End Sub

Private Function vWindow_Render() As Boolean
    If visible Then vControles.Render
    
    vWindow_Render = True
End Function

Public Sub vWindow_SetPos(ByVal X As Integer, ByVal Y As Integer)
    vX = X
    vY = Y
    
    Call actualizarPosiciones
End Sub

Private Sub CambiaSeleccion()
    razaActual = getRaza(numRazaActual)
    
    ImagenRaza.grhId = razaActual.grhId
    NombreRaza.text = razaActual.Nombre
    
    BarraFuerza.valor = razaActual.atributos(eAtributos.Fuerza)
    BarraAgilidad.valor = razaActual.atributos(eAtributos.Agilidad)
    BarraConsitucion.valor = razaActual.atributos(eAtributos.Constitucion)
    BarraInteligencia.valor = razaActual.atributos(eAtributos.Inteligencia)
    BarraCarisma.valor = razaActual.atributos(eAtributos.Inteligencia)
End Sub
Private Sub BotonAnteriorRaza_Click()
    Call Sonido_Play(SND_CLICK)

    numRazaActual = numRazaActual - 1
    
    If numRazaActual < 1 Then
        numRazaActual = modCrearPersonaje.getCantidadRazas()
    End If
    
    Call CambiaSeleccion
End Sub

Private Sub BotonSiguienteRaza_Click()
    Call Sonido_Play(SND_CLICK)

    numRazaActual = numRazaActual + 1
    
    If numRazaActual > modCrearPersonaje.getCantidadRazas() Then
        numRazaActual = 1
    End If
    
    Call CambiaSeleccion
End Sub

Public Sub vWindow_SetSize(ByVal width As Integer, ByVal Height As Integer)
    vW = width
    vH = Height
    Call actualizarPosiciones
End Sub

Private Sub actualizarPosiciones()
    FondoVentana.IControl_SetSize vW, vH
    FondoVentana.SetPos vX, vY
    
    ' Genero
    Dim generoX As Integer
    Dim generoY As Integer
    Dim tamSeleccion As Integer
    Dim margenGeneros As Integer
    Dim posXHombre As Integer
    Dim posXMujer As Integer
    
    margenGeneros = 30
    
    tamSeleccion = ImagenHombre.IControl_Width + margenGeneros + ImagenMujer.IControl_Width
    
    generoY = vY + 8
    
    TituloGenero.SetPos vX + vW / 2, generoY
    
    FondoTituloGenero.SetPos vX + vW / 2 - TituloGenero.width / 2, vY + 1
    FondoTituloGenero.IControl_SetSize TituloGenero.width, 32
      
    
    generoY = generoY + 35 ' Margen entre titulo y contenido
    
    posXHombre = vX + (vW - tamSeleccion) / 2
    posXMujer = vX + (vW - tamSeleccion) / 2 + ImagenHombre.IControl_Width + margenGeneros
    
    ImagenHombre.SetPos posXHombre, generoY
    TextoHombre.SetPos posXHombre + (ImagenHombre.IControl_Width / 2), generoY + ImagenHombre.IControl_Height
    
    ImagenMujer.SetPos posXMujer, generoY
    TextoMujer.SetPos posXMujer + (ImagenMujer.IControl_Width / 2), generoY + ImagenMujer.IControl_Height
    
 
    ' Raza
    Dim razaX As Integer
    Dim razaY As Integer
    Dim lineaSeparadoraY As Integer
    
    ' Margen entre secciones
    razaY = generoY + ImagenMujer.IControl_Height + 50
    
    lineaSeparadoraY = generoY + ImagenMujer.IControl_Height + 40
    lineaSeparadora.SetPos vX, lineaSeparadoraY
    lineaSeparadora.IControl_SetSize vW, 10
    
    TituloRaza.SetPos vX + vW / 2, razaY
    
    FondoTituloRaza.SetPos vX + vW / 2 - TituloRaza.width / 2, lineaSeparadoraY + 5
    FondoTituloRaza.IControl_SetSize TituloRaza.width, 32
    
    FondoRaza.SetPos vX + 40, razaY + 35
    ImagenRaza.SetPos vX + 40 + FondoRaza.IControl_Width / 2 - ImagenRaza.IControl_Width / 2, razaY + 35 + 10
    
    BotonAnteriorRaza.SetPos vX + 40, razaY + 35 + ImagenRaza.IControl_Height
    NombreRaza.SetPos vX + 40 + (FondoRaza.IControl_Width / 2), razaY + 35 + FondoRaza.IControl_Height
    BotonSiguienteRaza.SetPos vX + 40 + FondoRaza.IControl_Width - 20, razaY + 35 + FondoRaza.IControl_Height
    
    Dim posicionYBarras As Integer
    Dim posicionXBarras As Integer
    Dim margenBarras As Integer
    
    posicionYBarras = razaY + 50
    margenBarras = 20
    posicionXBarras = FondoRaza.IControl_X + FondoRaza.IControl_Width + margenBarras
     
    BarraConsitucion.IControl_SetPos posicionXBarras, posicionYBarras
    posicionYBarras = posicionYBarras + 20
    
    BarraFuerza.IControl_SetPos posicionXBarras, posicionYBarras
    posicionYBarras = posicionYBarras + 20
    
    BarraAgilidad.IControl_SetPos posicionXBarras, posicionYBarras
    posicionYBarras = posicionYBarras + 20
    
    BarraInteligencia.IControl_SetPos posicionXBarras, posicionYBarras
    posicionYBarras = posicionYBarras + 20
    
    BarraCarisma.IControl_SetPos posicionXBarras, posicionYBarras
    posicionYBarras = posicionYBarras + 20
    
    '  Botones
    If Not BotonSiguiente Is Nothing Then BotonSiguiente.SetPos vX + vW - 128, vY + vH + 10
    
    If Not BotonAtras Is Nothing Then BotonAtras.SetPos vX, vY + vH + 10
    
    If Not LabelError Is Nothing Then
        LabelError.SetPos vX + vW / 2, vY + vH - 25
    End If
End Sub

Private Function validar() As Boolean
    If getIndexSelected(RadioButtons) = -1 Then
        mostrarError 0, "Debes seleccionar un género para tu personaje."
        Exit Function
    End If

    validar = True
End Function

Private Sub changeRadioButton(group() As ControlCheckBox, check As ControlCheckBox)
    Dim loopElemento As Byte
    
    If check.value = False Then Exit Sub
    
    For loopElemento = LBound(group) To UBound(group)
        If group(loopElemento) Is check Then
            group(loopElemento).value = True
        Else
            group(loopElemento).value = False
        End If
    Next
End Sub

Private Function getIndexSelected(group() As ControlCheckBox) As Integer
    Dim loopElemento As Byte
    For loopElemento = LBound(group) To UBound(group)
        If group(loopElemento).value = True Then
            getIndexSelected = loopElemento
            Exit Function
        End If
    Next
    
    getIndexSelected = -1
End Function

Public Sub mostrarError(error As Byte, errordesc As String)
    If error = modLogin.JUEGO_DESACTUALIZADO Then
        Call juegoDesactualizado
    End If
    
    If LabelError Is Nothing Then
        Set LabelError = vControles.AddFirst(New ControlLabel)
    End If
   
    LabelError.SetPos vX + vW / 2, vY + vH - 25
    LabelError.Color = mzRed
    LabelError.SetCentrado True
    LabelError.SetSize vW - 100, vY
    
    ' Y el texto
    If errordesc = "" Then
        LabelError.text = modLogin.erroresDescripcion(error)
    ElseIf error > 0 Then
        LabelError.text = modLogin.erroresDescripcion(error) & " Error: " & errordesc
    Else
        LabelError.text = errordesc
    End If
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

