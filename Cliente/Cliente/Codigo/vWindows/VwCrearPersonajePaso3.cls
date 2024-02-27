VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VwCrearPersonajePaso3"
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

Private InputNombre As ControlInput
Private FondoVentana As ControlImage

Private TituloAlineacion As ControlLabel
Private TituloNombre As ControlLabel

Private FondoTituloAlineacion As ControlImage
Private FondoTituloNombre As ControlImage
Private lineaSeparadora As ControlImage

Private WithEvents CheckAzul As ControlCheckBox
Attribute CheckAzul.VB_VarHelpID = -1
Private WithEvents CheckRojo As ControlCheckBox
Attribute CheckRojo.VB_VarHelpID = -1
Private WithEvents CheckNeutral As ControlCheckBox
Attribute CheckNeutral.VB_VarHelpID = -1
Private RadioButtons(1 To 3) As ControlCheckBox

Private LabelAdvertencia As ControlLabel

Private LabelError As ControlLabel

Public Event Atras()
Public Event siguiente(Alineacion As eAlineaciones, Nombre As String)

'Implementacion de ventanas. No modificar
Private pNext As vWindow
Private pPrev As vWindow


Public Sub setDatos(Nombre As String, Alineacion As eAlineaciones)
    InputNombre.text = Nombre
    Dim radioSeleccionado As Byte
    
    If Alineacion = indefinido Then Exit Sub
    
    If Alineacion = eAlineaciones.Real Then
        radioSeleccionado = 1
    ElseIf Alineacion = eAlineaciones.Neutro Then
        radioSeleccionado = 2
    ElseIf Alineacion = eAlineaciones.caos Then
        radioSeleccionado = 3
    End If

    RadioButtons(radioSeleccionado).value = True
    Call changeAlineacion(RadioButtons(radioSeleccionado))
End Sub

Private Sub BotonAtras_Click()
    Call Sonido_Play(SND_CLICK)
    RaiseEvent Atras
End Sub

Private Function validar() As Boolean
    Dim UserName As String
    Dim UserEmail As String
    Dim UserPassword As String
    
    UserName = Trim$(InputNombre.text)
    
    If Len(UserName) < 3 Then
        mostrarError 0, "El nombre debe tener al menos tres letras."
        Exit Function
    End If
    
    If Len(UserName) > 12 Then
        mostrarError 0, "El nombre de tu personaje no puede tener más de 12 letras."
        Exit Function
    End If
    
    If DobleEspacios(UserName) Then
        mostrarError 0, "¡El nombre no puede contener dos o más espacios seguidos!."
        Exit Function
    End If
    
    If Not AsciiValidos(UserName) Then
        mostrarError 0, "¡El nick ingresado no es válido!."
        Exit Function
    End If
    
    If getAlineacionSeleccionada = -1 Then
        mostrarError 0, "Debes seleccionar una alineación."
        Exit Function
    End If

    validar = True
End Function

Private Sub BotonCrearPersonaje_Click()
    Dim Alineacion As eAlineaciones
    Dim seleccionada As Integer
    
    Call Sonido_Play(SND_CLICK)

    If Not validar Then Exit Sub
   
    seleccionada = getAlineacionSeleccionada
    
    If seleccionada = 1 Then
        Alineacion = eAlineaciones.Real
    ElseIf seleccionada = 2 Then
        Alineacion = eAlineaciones.Neutro
    ElseIf seleccionada = 3 Then
        Alineacion = eAlineaciones.caos
    End If
        
    RaiseEvent siguiente(Alineacion, InputNombre.text)
End Sub

Private Sub CheckAzul_Click()
    Call changeAlineacion(CheckAzul)
End Sub

Private Sub changeAlineacion(tipo As ControlCheckBox)
    Dim loopAlineacion As Byte
    
    If tipo.value = False Then Exit Sub
    
    For loopAlineacion = LBound(RadioButtons) To UBound(RadioButtons)
        If RadioButtons(loopAlineacion) Is tipo Then
            RadioButtons(loopAlineacion).value = True
        Else
            RadioButtons(loopAlineacion).value = False
        End If
    Next
End Sub

Private Function getAlineacionSeleccionada() As Integer
    Dim loopAlineacion As Byte
    For loopAlineacion = LBound(RadioButtons) To UBound(RadioButtons)
        If RadioButtons(loopAlineacion).value = True Then
            getAlineacionSeleccionada = loopAlineacion
            Exit Function
        End If
    Next
    
    getAlineacionSeleccionada = -1
End Function

Private Sub CheckNeutral_Click()
    Call changeAlineacion(CheckNeutral)
End Sub

Private Sub CheckRojo_Click()
    Call changeAlineacion(CheckRojo)
End Sub

Private Sub Class_Initialize()
    Set vControles = New vControles
    
    Set BotonCrearPersonaje = vControles.Add(New ControlBoton)
    Set BotonAtras = vControles.Add(New ControlBoton)
    
    Set InputNombre = vControles.Add(New ControlInput)
    Set LabelAdvertencia = vControles.Add(New ControlLabel)
    Set TituloAlineacion = vControles.Add(New ControlLabel)
    Set TituloNombre = vControles.Add(New ControlLabel)
    
    Set CheckAzul = vControles.Add(New ControlCheckBox)
    Set CheckNeutral = vControles.Add(New ControlCheckBox)
    Set CheckRojo = vControles.Add(New ControlCheckBox)
    Set RadioButtons(1) = CheckAzul
    Set RadioButtons(2) = CheckNeutral
    Set RadioButtons(3) = CheckRojo
    
    Set lineaSeparadora = vControles.Add(New ControlImage)
    
    Set FondoTituloAlineacion = vControles.Add(New ControlImage)
    Set FondoTituloNombre = vControles.Add(New ControlImage)
    
    Set FondoVentana = vControles.Add(New ControlImage)
    
    TituloAlineacion.text = "Elegi tu Alineacion"
    TituloAlineacion.SetStyle 4, mzInterfaceColor1, mzInterfaceColor2, dVertical
    TituloAlineacion.SetCentrado True
    
    TituloNombre.text = "Elegi un Nombre"
    TituloNombre.SetStyle 4, mzInterfaceColor1, mzInterfaceColor2, dVertical
    TituloNombre.SetCentrado True
    
    FondoVentana.grhId = 20352
    lineaSeparadora.grhId = 21867
    
    FondoTituloAlineacion.grhId = 21868
    FondoTituloNombre.grhId = 21868
    
    ' Seleccion de alineacion
    CheckAzul.textura = 3365
    CheckAzul.IControl_SetSize 64, 128
    CheckAzul.texto = "Ejército Índigo"
    CheckAzul.SetVertical True
    CheckAzul.SetStyle 3, mzInterfaceColor1, mzInterfaceColor2, dVertical
    
    CheckNeutral.textura = 3366
    CheckNeutral.IControl_SetSize 64, 128
    CheckNeutral.texto = "Rebelde"
    CheckNeutral.SetVertical True
    CheckNeutral.SetStyle 3, mzInterfaceColor1, mzInterfaceColor2, dVertical
    
    CheckRojo.textura = 3367
    CheckRojo.IControl_SetSize 64, 128
    CheckRojo.texto = "Ejército Escarlata"
    CheckRojo.SetVertical True
    CheckRojo.SetStyle 3, mzInterfaceColor1, mzInterfaceColor2, dVertical
    
    ' Campos finales
    LabelAdvertencia.text = "Recuerda que el nombre debe ir en tono con la época medieval."
    LabelAdvertencia.SetStyle 3, mzInterfaceColor1, mzInterfaceColor2, dVertical
    LabelAdvertencia.SetCentrado True
    LabelAdvertencia.SetSize 1000, 30
    
    InputNombre.Filtro = "ABCDEFGHIJLKMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz "
    InputNombre.MaxLenght = 12
    
    ' Botones de accion
    BotonAtras.text = "Volver"
    BotonCrearPersonaje.text = "Siguiente"
End Sub

Private Sub Class_Terminate()
    Debug.Print "Destruyo Paso3"
    MsgBox "terminado"
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
   
    LabelError.SetPos vX + vW / 2, vY + vH - 30
    LabelError.Color = mzRed
    LabelError.SetCentrado True
    LabelError.SetSize vW - 100, vY
    
    ' Y el texto
    If errordesc = "" Then
        LabelError.text = modLogin.erroresDescripcion(error)
    ElseIf error <> 0 Then
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

    FondoVentana.IControl_SetSize vW, vH
    FondoVentana.SetPos vX, vY
    
    TituloAlineacion.SetPos vX + vW / 2, vY + 10
    
    FondoTituloAlineacion.SetPos vX + vW / 2 - TituloAlineacion.width / 2, vY + 3
    FondoTituloAlineacion.IControl_SetSize TituloAlineacion.width, 32
    
    Dim tamSeleccionAlineacion As Byte
    Dim margenAlineaciones As Byte
    
    margenAlineaciones = 30
    
    tamSeleccionAlineacion = CheckAzul.IControl_Width * UBound(RadioButtons) + margenAlineaciones * (UBound(RadioButtons) - 1)
        
    Dim posXChecks As Integer
    Dim posYChecks As Integer
    
    posXChecks = vX + (vW - tamSeleccionAlineacion) / 2
    posYChecks = vY + 10 + 30
    
    CheckAzul.SetPos posXChecks, posYChecks
    CheckNeutral.SetPos posXChecks + CheckAzul.IControl_Width + margenAlineaciones, posYChecks
    CheckRojo.SetPos posXChecks + (CheckAzul.IControl_Width + margenAlineaciones) * 2, posYChecks
    
    Dim seccionNombreY As Integer
    Dim lineaSeparadoraY As Integer
    
    lineaSeparadoraY = posYChecks + CheckAzul.IControl_Height + 30
    lineaSeparadora.SetPos vX, lineaSeparadoraY
    lineaSeparadora.IControl_SetSize vW, 10
    
    seccionNombreY = posYChecks + CheckAzul.IControl_Height + 40
    
    TituloNombre.SetPos vX + vW / 2, seccionNombreY
    FondoTituloNombre.SetPos vX + vW / 2 - TituloNombre.width / 2, lineaSeparadoraY + 5
    FondoTituloNombre.IControl_SetSize TituloNombre.width, 32
    
    LabelAdvertencia.SetPos vX + vW / 2, seccionNombreY + 35

    InputNombre.SetPos vX + vW / 2 - 300 / 2, seccionNombreY + 40
    InputNombre.SetSize 300, 26 '
    
    BotonCrearPersonaje.SetPos vX + vW - 128, vY + vH + 10
    BotonAtras.SetPos vX, vY + vH + 10
    
    If Not LabelError Is Nothing Then
        LabelError.SetPos vX + vW / 2, vY + vH - 30
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

