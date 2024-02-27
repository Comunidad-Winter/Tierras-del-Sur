VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vwCuenta"
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

Private visible As Boolean

Private LabelError As ControlLabel

' Superior
Private WithEvents BotonSalir As ControlBoton
Attribute BotonSalir.VB_VarHelpID = -1
Private WithEvents BotonCrearPersonaje As ControlBoton
Attribute BotonCrearPersonaje.VB_VarHelpID = -1
Private WithEvents BotonSiguientePagina As ControlBoton
Attribute BotonSiguientePagina.VB_VarHelpID = -1
Private WithEvents BotonAnteriorPagina As ControlBoton
Attribute BotonAnteriorPagina.VB_VarHelpID = -1

Private linea1 As ControlLabel
Attribute linea1.VB_VarHelpID = -1
Private linea2 As ControlLabel

Private fondo As ControlImage
Attribute fondo.VB_VarHelpID = -1

Private WithEvents vControles As vControles
Attribute vControles.VB_VarHelpID = -1

Private personajesControles As vControles

Private ObjetoFocus As Object

Public Event conectar(usuario As String, Password As String)
Public Event crear()
Public Event cerrarSesion()

'Implementacion de ventanas. No modificar
Private pNext As vWindow
Private pPrev As vWindow

Private Progreso As clsProgreso

Private CANTIDAD_PERSONAJES_ANCHO As Integer

Private paginaActual As Integer

Private Sub BotonAnteriorPagina_Click()
    paginaActual = paginaActual - 1
    mostrarPagina paginaActual
    actualizarPosiciones
End Sub

Private Sub BotonCrearPersonaje_Click()
    RaiseEvent crear
End Sub

Private Sub BotonSalir_Click()
    RaiseEvent cerrarSesion
End Sub

Private Sub BotonCrear_Click()
    Call Sonido_Play(SND_CLICK)
        
    RaiseEvent crear
End Sub

Private Sub BotonSiguientePagina_Click()
    paginaActual = paginaActual + 1
    mostrarPagina paginaActual
    actualizarPosiciones
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
 
Set LabelError = Nothing
Set BotonSalir = Nothing
Set BotonCrearPersonaje = Nothing
Set vControles = Nothing
Set Progreso = Nothing

End Sub

Private Sub Class_Initialize()
    Set vControles = New vControles
    
    Set BotonSalir = vControles.Add(New ControlBoton)
    Set BotonCrearPersonaje = vControles.Add(New ControlBoton)
    Set BotonSiguientePagina = vControles.Add(New ControlBoton)
    Set BotonAnteriorPagina = vControles.Add(New ControlBoton)
    Set fondo = vControles.Add(New ControlImage)
    Set linea1 = vControles.Add(New ControlLabel)
    Set linea2 = vControles.Add(New ControlLabel)
    Set personajesControles = New vControles
    
    fondo.grhId = 20352
    
    BotonSalir.text = "Salir"
    BotonCrearPersonaje.text = "Crear Personaje"
    
    BotonAnteriorPagina.SetSize 20, 20
    BotonAnteriorPagina.textura = 3362
    BotonAnteriorPagina.text = ""
    
    BotonSiguientePagina.SetSize 20, 20
    BotonSiguientePagina.textura = 3361
    BotonSiguientePagina.text = ""
    
    linea1.Color = mzInterfaceColor1
    linea2.Color = mzInterfaceColor1
    paginaActual = 0
End Sub

Public Sub refrescarPersonajes()

    linea1.text = MiCuenta.cuenta.Nombre
    linea2.text = "Liga: etapa de clasificación. Puntos: 0"
    
    Debug.Print "Tiene " & MiCuenta.personajes.count
    
    Call eliminarFichas
    
    Call cargarFichas(0, CANTIDAD_PERSONAJES_ANCHO * 2)
    
    Call actualizarPosiciones
End Sub

Private Sub cargarFichas(ByVal skip As Integer, cantidad As Integer)
    Dim personaje As personaje
    Dim fichaPersonaje As ControlFichaPersonaje
    Dim numeroPersonaje As Integer
    
    numeroPersonaje = 1
    
    For Each personaje In MiCuenta.personajes
        If skip = 0 Then
            Set fichaPersonaje = New ControlFichaPersonaje
            fichaPersonaje.personaje = personaje
            fichaPersonaje.IControl_SetSize 128, 128
            fichaPersonaje.setCallBack Me
            
            vControles.AddFirst fichaPersonaje
            personajesControles.Add fichaPersonaje
        
            numeroPersonaje = numeroPersonaje + 1
        Else
            skip = skip - 1
        End If
        If numeroPersonaje > cantidad Then Exit Sub
    Next
End Sub

Private Sub eliminarFichas()
    Dim Control As IControl

    
    For Each Control In personajesControles
        vControles.RemoveControl Control
    Next
    
    personajesControles.Clear
End Sub

Public Sub setPersonajeSeleccionado(personaje As personaje)
  ' TODO ¿Tiene una clave adicional?
  RaiseEvent conectar(personaje.Nick, personaje.Password)
End Sub

Public Sub vWindow_Hide()
    visible = False
End Sub

Private Function vWindow_Render() As Boolean
    vWindow_Render = True
        
    If visible Then
        vControles.Render
    End If
End Function

Public Sub vWindow_SetPos(ByVal X As Integer, ByVal Y As Integer)
    vX = X
    vY = Y
    
    Call actualizarPosiciones
End Sub

Public Sub vWindow_SetSize(ByVal width As Integer, ByVal Height As Integer)
    vW = width
    vH = Height
    
    CANTIDAD_PERSONAJES_ANCHO = (vW - 20 - 20) \ 140
    
    Call refrescarPersonajes
    Call actualizarPosiciones
End Sub

Public Sub vWindow_Show()
    visible = True
    Set Progreso = New clsProgreso
    Progreso.SetRango 0, 255
    Progreso.SetTicks GetTimer + 1000
End Sub

Private Sub ocultarError()
    If LabelError Is Nothing Then
        Exit Sub
    End If
    
    Dim Control As IControl
    Set Control = LabelError
    Call vControles.RemoveControl(LabelError)
    
    Set LabelError = Nothing
End Sub
Public Sub mostrarError(error As Byte, errordesc As String)
    If LabelError Is Nothing Then
        Set LabelError = vControles.Add(New ControlLabel)
    End If
    
    LabelError.Color = mzRed
    LabelError.SetCentrado False
    LabelError.SetSize 200, vY
    
    ' Y el texto
    If errordesc = "" Then
        LabelError.text = modLogin.erroresDescripcion(error)
    ElseIf error > 0 Then
        LabelError.text = modLogin.erroresDescripcion(error) & " Error: " & errordesc
    Else
        LabelError.text = errordesc
    End If
    
    Call actualizarPosiciones
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

Private Sub mostrarPagina(pagina As Integer)
    Call eliminarFichas
    Call cargarFichas(pagina * CANTIDAD_PERSONAJES_ANCHO * 2, CANTIDAD_PERSONAJES_ANCHO * 2)
End Sub
Private Sub actualizarPosiciones()

linea1.SetPos vX + 10, vY - 30
linea2.SetPos vX + 10, vY - 15

BotonSalir.SetSize 128, 32
BotonCrearPersonaje.SetSize 128, 32

BotonSalir.SetPos vX, vY + vH + 10
BotonCrearPersonaje.SetPos vX + vW - 128, vY + vH + 10
    
fondo.SetPos vX, vY
fondo.IControl_SetSize vW, vH

Dim personajeControl As IControl
Dim fichaX As Integer
Dim fichaY As Integer



Dim numeroControl As Integer
numeroControl = 0

If CANTIDAD_PERSONAJES_ANCHO > 0 And Not personajesControles Is Nothing Then
    For Each personajeControl In personajesControles
        fichaX = vX + 20 + (numeroControl Mod CANTIDAD_PERSONAJES_ANCHO) * 135
        fichaY = vY + 20 + (numeroControl \ CANTIDAD_PERSONAJES_ANCHO) * 145
                
        personajeControl.SetPos fichaX, fichaY
        
        numeroControl = numeroControl + 1
    Next
End If
exitFor:

BotonSiguientePagina.visible = False
BotonSiguientePagina.Enabled = False

BotonAnteriorPagina.visible = False
BotonAnteriorPagina.Enabled = False


If paginaActual > 0 Then
    BotonAnteriorPagina.visible = True
    BotonAnteriorPagina.Enabled = True
End If

If (paginaActual + 1) * CANTIDAD_PERSONAJES_ANCHO * 2 <= personajes.count Then
    BotonSiguientePagina.visible = True
    BotonSiguientePagina.Enabled = True
End If


BotonAnteriorPagina.SetPos vX + vW - 60, vY + vH - 25
BotonSiguientePagina.SetPos vX + vW - 30, vY + vH - 25
    
If Not LabelError Is Nothing Then
    LabelError.SetPos vX, BotonCrearPersonaje.GetPosY + BotonCrearPersonaje.IControl_Height + 10
End If

End Sub

