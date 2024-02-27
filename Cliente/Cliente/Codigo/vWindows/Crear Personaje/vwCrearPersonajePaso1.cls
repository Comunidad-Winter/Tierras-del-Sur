VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vwCrearPersonajePaso1"
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

Private fondo As ControlImage
Attribute fondo.VB_VarHelpID = -1
Private WithEvents BotonSiguiente As ControlBoton
Attribute BotonSiguiente.VB_VarHelpID = -1
Private WithEvents BotonAtras As ControlBoton
Attribute BotonAtras.VB_VarHelpID = -1
Private WithEvents BotonSiguienteClase As ControlBoton
Attribute BotonSiguienteClase.VB_VarHelpID = -1
Private WithEvents BotonAnteriorClase As ControlBoton
Attribute BotonAnteriorClase.VB_VarHelpID = -1
Private ImagenClase As ControlImage
Private NombreClase As ControlLabel
Private DescripcionClase As ControlLabel

Private tituloClase As ControlLabel
Private FondoTituloClase As ControlImage

Private WithEvents vControles As vControles
Attribute vControles.VB_VarHelpID = -1

Private ObjetoFocus As Object

Private numClaseActual As Integer
Private claseActual As tClase

Private visible As Boolean

Public Event Cerrar()
Public Event siguiente(idClase As eClass)

'Implementacion de ventanas. No modificar
    Private pNext As vWindow
    Private pPrev As vWindow

Public Sub setDatos(Clase As eClass)
    
    If Clase > 0 Then
        numClaseActual = getClaseNumero(Clase)
        Call CambiaSeleccion
    End If
    
End Sub

Private Sub CambiaSeleccion()
    claseActual = getClase(numClaseActual)
    
    ImagenClase.grhId = claseActual.grhId
    DescripcionClase.text = claseActual.descripcion
    NombreClase.text = claseActual.Nombre
End Sub
Private Sub BotonAnteriorClase_Click()
    Call Sonido_Play(SND_CLICK)

    numClaseActual = numClaseActual - 1
    
    If numClaseActual < 1 Then
        numClaseActual = modCrearPersonaje.getCantidadClases()
    End If
    
    Call CambiaSeleccion
End Sub

Private Sub BotonAtras_Click()
    Call Sonido_Play(SND_CLICK)
    RaiseEvent Cerrar
End Sub

Private Sub BotonSiguiente_Click()
    Call Sonido_Play(SND_CLICK)
    RaiseEvent siguiente(claseActual.id)
End Sub

Private Sub BotonSiguienteClase_Click()
    Call Sonido_Play(SND_CLICK)

    numClaseActual = numClaseActual + 1
    
    If numClaseActual > modCrearPersonaje.getCantidadClases() Then
        numClaseActual = 1
    End If
    
    Call CambiaSeleccion
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

'/Implementacion

Private Sub Class_Terminate()
End Sub

Private Sub Class_Initialize()
Set vControles = New vControles

modCrearPersonaje.initAlineaciones

modCrearPersonaje.initClases

numClaseActual = 1

claseActual = getClase(numClaseActual)

Set BotonSiguienteClase = vControles.Add(New ControlBoton)
Set BotonAnteriorClase = vControles.Add(New ControlBoton)

Set BotonSiguiente = vControles.Add(New ControlBoton)
Set BotonAtras = vControles.Add(New ControlBoton)

Set ImagenClase = vControles.Add(New ControlImage)

Set DescripcionClase = vControles.Add(New ControlLabel)
Set NombreClase = vControles.Add(New ControlLabel)

Set tituloClase = vControles.Add(New ControlLabel)

Set FondoTituloClase = vControles.Add(New ControlImage)
Set fondo = vControles.Add(New ControlImage)

fondo.grhId = 20352

FondoTituloClase.grhId = 21868

tituloClase.text = "Elegi tu clase"

tituloClase.Font = 4
tituloClase.Color = mzInterfaceColor1
tituloClase.Color2 = mzInterfaceColor2
tituloClase.Gradient = dVertical
tituloClase.SetCentrado True

BotonAtras.text = "Volver"

BotonSiguiente.text = "Siguiente"

BotonSiguienteClase.SetSize 20, 20
BotonSiguienteClase.textura = 3361
BotonSiguienteClase.text = ""

BotonAnteriorClase.SetSize 20, 20
BotonAnteriorClase.textura = 3362
BotonAnteriorClase.text = ""

ImagenClase.grhId = claseActual.grhId

Call DescripcionClase.SetStyle(3, mzInterfaceColor1, mzInterfaceColor2, dVertical)

NombreClase.SetStyle 2, mzInterfaceColor1, mzInterfaceColor2, dVertical
NombreClase.SetCentrado True

Call CambiaSeleccion
Call actualizarPosiciones
End Sub

Public Sub vWindow_Hide()
    visible = False
End Sub

Private Function vWindow_Render() As Boolean
    If Not visible Then
        vWindow_Render = True
        Exit Function
    End If
    
    vWindow_Render = True
    vControles.Render
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
    Dim tempX As Integer
    Dim tempY As Integer
    Dim alto As Integer
    
    tempX = vX
    tempY = vY
    alto = 300
    
    If Not fondo Is Nothing Then
        fondo.SetPos tempX, tempY
        fondo.IControl_SetSize vW, alto
        tempX = tempX + 20
        tempY = tempY + 8
    End If

    If Not tituloClase Is Nothing Then
        FondoTituloClase.SetPos vX + vW / 2 - tituloClase.width / 2, vY + 1
        tituloClase.SetPos vX + vW / 2, tempY
        FondoTituloClase.IControl_SetSize tituloClase.width, 32
    End If

    ' Clase
    Dim posicionX As Integer
    Dim posicionY As Integer
    Dim tamClase As Integer
    tamClase = 128
    posicionX = vX + 30
    posicionY = tempY + 76
    
    If Not ImagenClase Is Nothing Then ImagenClase.SetPos posicionX, posicionY
    
    If Not NombreClase Is Nothing Then
        NombreClase.SetPos posicionX + tamClase / 2, posicionY + tamClase + 5
    End If
    
    If Not DescripcionClase Is Nothing Then
        DescripcionClase.SetPos posicionX + tamClase + 10, tempY + 60
        DescripcionClase.SetSize 280, 0
    End If
        
    If Not BotonSiguienteClase Is Nothing Then BotonSiguienteClase.SetPos posicionX + tamClase - 20, posicionY + tamClase + 5
    
    If Not BotonAnteriorClase Is Nothing Then BotonAnteriorClase.SetPos posicionX, posicionY + tamClase + 5
    
    ' Botones inferiores
    If Not BotonSiguiente Is Nothing Then BotonSiguiente.SetPos vX + vW - 128, tempY + alto + 10
    
    If Not BotonAtras Is Nothing Then BotonAtras.SetPos vX, tempY + alto + 10
End Sub

Public Sub vWindow_Show()
    visible = True
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

