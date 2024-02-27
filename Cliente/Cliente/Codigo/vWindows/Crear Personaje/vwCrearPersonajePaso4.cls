VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vwCrearPersonajePaso4"
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

Private WithEvents BotonCrearPersonaje As ControlBoton
Attribute BotonCrearPersonaje.VB_VarHelpID = -1
Private WithEvents BotonAtras As ControlBoton
Attribute BotonAtras.VB_VarHelpID = -1

Private FondoVentana As ControlImage

Private FondoTitulo As ControlImage
Private TituloSeccion As ControlLabel

Private LabelAdvertencia As ControlLabel

Private LabelError As ControlLabel

Private FondoPersonaje As ControlImage

Private WithEvents SeleccionPelo As ControlSelector
Attribute SeleccionPelo.VB_VarHelpID = -1
Private WithEvents SeleccionBarba As ControlSelector
Attribute SeleccionBarba.VB_VarHelpID = -1
Private WithEvents SeleccionRopaInterior As ControlSelector
Attribute SeleccionRopaInterior.VB_VarHelpID = -1
Private WithEvents SeleccionTonoPiel As ControlSelector
Attribute SeleccionTonoPiel.VB_VarHelpID = -1
Private WithEvents SeleccionColorPelo As ControlSelector
Attribute SeleccionColorPelo.VB_VarHelpID = -1


Private ImagenBarba As ControlPicture
Private ImagenPelo As ControlPicture
Private ImagenRopaInterior As ControlPicture
Private ImagenCuerpo As ControlPicture
Private ImagenCara As ControlPicture

Private personalizacion As tRazaPartes

'Implementacion de ventanas. No modificar
Private pNext As vWindow
Private pPrev As vWindow

Private headId As Integer           ' Identificador de la cabeza seleccionada
Private bodyId As Integer           ' Identificador del cuerpo seleccionado
Private ropaInteriorId As Integer   ' Identificador de la ropa seleccionada
Private barbaId As Integer          ' Identificador de la barba selecionada
Private peloId As Integer           ' Identificiador del pelo seleccionado

Public Event Atras()
Public Event siguiente(headId As Integer, bodyId As Integer, ropaInteriorId As Integer, barbaId As Integer, peloId As Integer)

Private Sub BotonAtras_Click()
    Call Sonido_Play(SND_CLICK)
    RaiseEvent Atras
End Sub

Private Sub BotonCrearPersonaje_Click()
    If Not validar Then Exit Sub

    Call Sonido_Play(SND_CLICK)
    
    RaiseEvent siguiente(headId, bodyId, ropaInteriorId, barbaId, peloId)
End Sub

Private Function validar() As Boolean
    Dim valorColorPelo As Integer
    Dim valorBarba As Integer
    Dim valorPelo As Integer
    
    valorColorPelo = SeleccionColorPelo.valor
    
    valorBarba = ((SeleccionBarba.valor - 1) * 13 + valorColorPelo)
    valorPelo = ((SeleccionPelo.valor - 1) * 13 + valorColorPelo)
    
    If Not MiCuenta.cuenta.premium Then
        If valorBarba > 35 Or valorPelo > 35 Then
            validar = False
            Call mostrarError(0, "Este estilo es válido solo para Cuentas Premium.")
            Exit Function
        End If
    End If
        
    Call ocultarError
    
    validar = True
End Function

Private Sub ocultarError()
    If LabelError Is Nothing Then
        Exit Sub
    End If
    
    Dim Control As IControl
    Set Control = LabelError
    Call vControles.RemoveControl(LabelError)
    
    Set LabelError = Nothing
End Sub

Private Sub Class_Initialize()
    Set vControles = New vControles
        
    Set BotonCrearPersonaje = vControles.Add(New ControlBoton)
    Set BotonAtras = vControles.Add(New ControlBoton)
        
    Set TituloSeccion = vControles.Add(New ControlLabel)
    
    Set FondoTitulo = vControles.Add(New ControlImage)
         
    Set SeleccionPelo = vControles.Add(New ControlSelector)
    Set SeleccionTonoPiel = vControles.Add(New ControlSelector)
    Set SeleccionBarba = vControles.Add(New ControlSelector)
    Set SeleccionRopaInterior = vControles.Add(New ControlSelector)
    Set SeleccionColorPelo = vControles.Add(New ControlSelector)
    
    ' El personaje
    Set ImagenBarba = vControles.Add(New ControlPicture)
    Set ImagenPelo = vControles.Add(New ControlPicture)
    Set ImagenRopaInterior = vControles.Add(New ControlPicture)
    Set ImagenCuerpo = vControles.Add(New ControlPicture)
    Set ImagenCara = vControles.Add(New ControlPicture)
    
    Set FondoPersonaje = vControles.Add(New ControlImage)
    Set FondoVentana = vControles.Add(New ControlImage)
    
    ' Titulo
    FondoTitulo.grhId = 21868
        
    TituloSeccion.text = "Personaliza tu personaje"
    TituloSeccion.SetStyle 4, mzInterfaceColor1, mzInterfaceColor2, dVertical
    TituloSeccion.SetCentrado True
    
    ' Lienzo del Personaje
    FondoPersonaje.grhId = 20351
        
    ' Barras de seleccion
    SeleccionPelo.Nombre = "Peinado"
    SeleccionPelo.SetStyle 3, mzInterfaceColor1, mzInterfaceColor2, dVertical
    
    SeleccionBarba.Nombre = "Barba"
    SeleccionBarba.SetStyle 3, mzInterfaceColor1, mzInterfaceColor2, dVertical
    
    SeleccionTonoPiel.Nombre = "Tono de Piel"
    SeleccionTonoPiel.SetStyle 3, mzInterfaceColor1, mzInterfaceColor2, dVertical
    
    SeleccionRopaInterior.Nombre = "Ropa Interior"
    SeleccionRopaInterior.SetStyle 3, mzInterfaceColor1, mzInterfaceColor2, dVertical
    
    SeleccionColorPelo.Nombre = "Color de Pelo"
    SeleccionColorPelo.SetStyle 3, mzInterfaceColor1, mzInterfaceColor2, dVertical
    
    ' Fin Barras
    personalizacion = modCrearPersonaje.getRazaById(DatosCreacion.Raza).personalizacion(DatosCreacion.Genero)
    
    SeleccionTonoPiel.min = 1
    SeleccionTonoPiel.max = UBound(personalizacion.cuerpos)
    SeleccionTonoPiel.valor = 1
    
    SeleccionPelo.min = 1
    SeleccionPelo.max = 5
    SeleccionPelo.valor = 1
    
    SeleccionBarba.min = 1
    SeleccionBarba.max = 5
    SeleccionBarba.valor = 1
    
    SeleccionRopaInterior.min = 1
    SeleccionRopaInterior.max = UBound(personalizacion.ropaInterior)
    SeleccionRopaInterior.valor = 1
    
    SeleccionColorPelo.min = 1
    SeleccionColorPelo.max = 13
    SeleccionColorPelo.valor = 1
    ' Fin Barras
    
    Call setTonoPiel(SeleccionTonoPiel.valor)
    Call setColorPelo(SeleccionColorPelo.valor)
    Call setRopaInterior(SeleccionRopaInterior.valor)
        
    FondoVentana.grhId = 20352
       
    ' Botones de accion
    BotonAtras.text = "Volver"
    BotonCrearPersonaje.text = "Crear"
    
    Call actualizarEstetica
End Sub

Private Sub Class_Terminate()
    Debug.Print "Destruyo Paso4"
End Sub

Private Sub SeleccionBarba_Change(valor As Integer)
    Dim valorPelo As Integer
    Dim valorBarba As Integer
    valorPelo = SeleccionColorPelo.valor
    
    valorBarba = ((valor - 1) * 13 + valorPelo)
    
    If valorBarba > UBound(personalizacion.barbas) Then
        barbaId = 0
    Else
        barbaId = personalizacion.barbas((valor - 1) * 13 + valorPelo)
    End If
    
    
    Call actualizarEstetica
    
    Call Sonido_Play(SND_CLICK)
End Sub

Private Sub setColorPelo(valor As Integer)
    Dim posicion As Integer
    Dim barbaValor As Integer
    
    peloId = personalizacion.pelos((SeleccionPelo.valor - 1) * 13 + valor)
    
    barbaValor = (SeleccionBarba.valor - 1) * 13 + valor
    
    If barbaValor > UBound(personalizacion.barbas) Then
        barbaId = 0
    Else
        barbaId = personalizacion.barbas((SeleccionBarba.valor - 1) * 13 + valor)
    End If
    
    Call actualizarEstetica
End Sub

Private Sub SeleccionColorPelo_Change(valor As Integer)
    Call setColorPelo(valor)
    
    Call Sonido_Play(SND_CLICK)
End Sub

Private Sub setPelo(valor As Integer)
    Dim valorPelo As Integer
    valorPelo = SeleccionColorPelo.valor
    
    peloId = personalizacion.pelos((valor - 1) * 13 + valorPelo)
    
    Call actualizarEstetica
End Sub
Private Sub SeleccionPelo_Change(valor As Integer)
    Call setPelo(valor)
    
    Call Sonido_Play(SND_CLICK)
End Sub

Private Sub setRopaInterior(valor As Integer)
    ropaInteriorId = personalizacion.ropaInterior(valor)
    
    Call actualizarEstetica
End Sub
Private Sub SeleccionRopaInterior_Change(valor As Integer)
    Call setRopaInterior(valor)
    
    Call Sonido_Play(SND_CLICK)
End Sub

Private Sub actualizarEstetica()
    Dim grhId As Integer

    Dim centro As Integer
    Dim posicionCabeza As Integer
    
    centro = FondoPersonaje.IControl_X + FondoPersonaje.IControl_Width / 2

    Call validar
   
    If bodyId > 0 Then
        grhId = BodyData(bodyId).Walk(E_Heading.NORTH).GrhIndex
        
        ImagenCuerpo.setOrigin 32, 0, 32, 64
        ImagenCuerpo.setTextura (CInt(GrhData(GrhData(grhId).frames(1)).filenum)), 256, 256
        ImagenCuerpo.IControl_SetSize 64, 128
        ImagenCuerpo.SetPos centro - 32, FondoPersonaje.IControl_Y
        
        If BodyData(bodyId).HeadOffset.Y = -28 Then
            posicionCabeza = FondoPersonaje.IControl_Y + BodyData(bodyId).HeadOffset.Y + 38
        Else
            posicionCabeza = FondoPersonaje.IControl_Y + BodyData(bodyId).HeadOffset.Y + 26
        End If
    End If
    
    If posicionCabeza > 0 Then
     If headId > 0 Then
         ImagenCara.setTextura (GrhData(HeadData(headId).Head(E_Heading.NORTH).GrhIndex).filenum), 128, 32
         ImagenCara.setOrigin 0, 0, 32, 32
         ImagenCara.IControl_SetSize 64, 64
         ImagenCara.SetPos centro - 32, posicionCabeza
     End If
     
     If barbaId > 0 Then
         ImagenBarba.setOrigin 0, 0, 32, 32
         ImagenBarba.setTextura barbaId, 128, 32
         ImagenBarba.SetPos centro - 32, posicionCabeza
         ImagenBarba.IControl_SetSize 64, 64
    End If
     
    If ropaInteriorId > 0 Then
        ImagenRopaInterior.setOrigin 32, 0, 32, 64
        ImagenRopaInterior.setTextura ropaInteriorId, 256, 256
        ImagenRopaInterior.IControl_SetSize 64, 128
        ImagenRopaInterior.SetPos centro - 32, FondoPersonaje.IControl_Y
     End If
     
     If peloId > 0 Then
          ImagenPelo.setOrigin 0, 0, 32, 32
          ImagenPelo.setTextura peloId, 128, 32
          ImagenPelo.SetPos centro - 32, posicionCabeza
          ImagenPelo.IControl_SetSize 64, 64
     End If
    End If
End Sub

Private Sub setTonoPiel(valor As Integer)
    headId = personalizacion.cabezas(valor)
    bodyId = personalizacion.cuerpos(valor)
    
    Call actualizarEstetica
End Sub

Private Sub SeleccionTonoPiel_Change(valor As Integer)
    Call setTonoPiel(valor)
    Call Sonido_Play(SND_CLICK)
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
    If Visible Then vControles.Render
    
    vWindow_Render = True
End Function


Public Sub mostrarError(error As Byte, errordesc As String)
    If error = modLogin.JUEGO_DESACTUALIZADO Then
        Call juegoDesactualizado
    End If
    
    If LabelError Is Nothing Then
        Set LabelError = vControles.AddFirst(New ControlLabel)
    End If
   
    LabelError.SetPos vX + vW / 2, vY + vH - 35
    LabelError.Color = mzRed
    LabelError.SetCentrado True
    LabelError.SetSize vW - 50, vY
    
    ' Y el texto
    If errordesc = "" Then
        LabelError.text = modLogin.erroresDescripcion(error)
    ElseIf error > 0 Then
        LabelError.text = modLogin.erroresDescripcion(error) & " Error: " & errordesc
    Else
        LabelError.text = errordesc
    End If
End Sub

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
    Dim margenDerecho As Integer
    Dim margenIzquierdo As Integer
    Dim margenSuperior As Integer
    
    FondoVentana.IControl_SetSize vW, vH
    FondoVentana.SetPos vX, vY
    
    ' Labels
    TituloSeccion.SetPos vX + vW / 2, vY + 10
    
    FondoTitulo.SetPos vX + vW / 2 - TituloSeccion.width / 2, vY + 3
    FondoTitulo.IControl_SetSize TituloSeccion.width, 32
        
    margenDerecho = vX + vW - 180
    margenSuperior = vY + 50
    
    ' FondoPersonaje
    FondoPersonaje.SetPos vX + 30, margenSuperior + 15
    
    ' Barras de Seleccion
    SeleccionPelo.IControl_SetPos margenDerecho, margenSuperior
    SeleccionPelo.IControl_SetSize 150, 20

    SeleccionTonoPiel.IControl_SetPos margenDerecho, 30 + margenSuperior
    SeleccionTonoPiel.IControl_SetSize 150, 20
    
    SeleccionBarba.IControl_SetPos margenDerecho, 60 + margenSuperior
    SeleccionBarba.IControl_SetSize 150, 20

    SeleccionColorPelo.IControl_SetPos margenDerecho, 90 + margenSuperior
    SeleccionColorPelo.IControl_SetSize 150, 20

    SeleccionRopaInterior.IControl_SetPos margenDerecho, 120 + margenSuperior
    SeleccionRopaInterior.IControl_SetSize 150, 20
    
    ' Botones finales
    BotonCrearPersonaje.SetPos vX + vW - 128, vY + vH + 10
    BotonAtras.SetPos vX, vY + vH + 10
    
    Call actualizarEstetica
    
    If Not LabelError Is Nothing Then
        LabelError.SetPos vX + vW / 2, vY + vH - 35
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

