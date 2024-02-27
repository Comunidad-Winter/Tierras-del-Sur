VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VwCrearPersonaje"
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

Private vControles As New vControles

Private WithEvents Paso0 As vwCuenta
Attribute Paso0.VB_VarHelpID = -1
Private WithEvents Paso1 As vwCrearPersonajePaso1
Attribute Paso1.VB_VarHelpID = -1
Private WithEvents Paso2 As VwCrearPersonajePaso2
Attribute Paso2.VB_VarHelpID = -1
Private WithEvents Paso3 As VwCrearPersonajePaso3
Attribute Paso3.VB_VarHelpID = -1
Private WithEvents Paso4 As vwCrearPersonajePaso4
Attribute Paso4.VB_VarHelpID = -1

Private ultimoPaso As Integer

Private claseActual As tClase

Private Visible As Boolean

Public Event Cerrar()
Public Event crear(datos As CrearPersonajeDTO)
Public Event conectar(personaje As String, Password As String)

Private NombreClase As ControlLabel
Private NombreRazaGenero As ControlLabel
Private AlineacionLabel As ControlLabel

Private DatosCreacion As CrearPersonajeDTO

Private progress As clsProgreso
Private progressSaliente As clsProgreso
Private vWindowEntrante As vWindow
Private vWindowSaliente As vWindow

'Implementacion de ventanas. No modificar
    Private pNext As vWindow
    Private pPrev As vWindow

Private WINDOW_TOP As Integer

Private Sub Class_Initialize()
    Set NombreClase = vControles.Add(New ControlLabel)
    NombreClase.text = ""
    NombreClase.SetPos 10, 100
    NombreClase.Color = mzCTalkMuertos
    
    Set NombreRazaGenero = vControles.Add(New ControlLabel)
    NombreRazaGenero.text = ""
    NombreRazaGenero.SetPos 10, 120
    NombreRazaGenero.Color = mzCTalkMuertos
    
    Set AlineacionLabel = vControles.Add(New ControlLabel)
    AlineacionLabel.text = ""
    AlineacionLabel.SetPos 10, 140
    AlineacionLabel.Color = mzCTalkMuertos
    
    Set DatosCreacion = New CrearPersonajeDTO
    
    WINDOW_TOP = Engine_Resolution.pixelesAlto / 2 - 300 / 2
End Sub

Private Sub Paso0_cerrarSesion()
    RaiseEvent Cerrar
End Sub

Private Sub Paso0_conectar(usuario As String, Password As String)
    RaiseEvent conectar(usuario, Password)
End Sub

Private Sub Paso0_crear()
    Call OcultarPaso0(False)
    Call MostrarPaso1(False)
End Sub

Private Sub Paso1_Cerrar()
    NombreClase.text = ""
    NombreRazaGenero.text = ""
    AlineacionLabel.text = ""
    Call OcultarPaso1(True)
    Call MostrarPaso0(True)
End Sub

Private Sub Paso1_Siguiente(idClase As eClass)
    claseActual = getClaseById(idClase)
    
    NombreClase.text = claseActual.Nombre
            
    DatosCreacion.Clase = idClase
    Call OcultarPaso1(False)
    Call MostrarPaso2(False)
End Sub

Private Sub Paso2_Atras()
    Call OcultarPaso2(True)
    Call MostrarPaso1(True)
End Sub

Private Sub Paso2_Siguiente(Genero As eGenero, Raza As eRaza)
    Dim razaData As tRaza
    razaData = getRazaById(Raza)
    
    DatosCreacion.Genero = Genero
    DatosCreacion.Raza = Raza
    
    NombreRazaGenero.text = razaData.Nombre
    Call MostrarPaso3(False)
    Call OcultarPaso2(False)
End Sub

Private Sub Paso3_Atras()
    Call OcultarPaso3(True)
    Call MostrarPaso2(True)
End Sub

Private Sub Paso3_Siguiente(Alineacion As eAlineaciones, Nombre As String)
    DatosCreacion.Nombre = Nombre
    DatosCreacion.Alineacion = Alineacion
        
    AlineacionLabel.text = alineaciones(Alineacion)
    
    Set modCrearPersonaje.DatosCreacion = DatosCreacion
    
    Call OcultarPaso3
    Call MostrarPaso4
End Sub

Private Sub Paso4_Atras()
    Call OcultarPaso4(True)
    Call MostrarPaso3(True)
End Sub

Public Sub refrescarPersonajes()
    If Not Paso0 Is Nothing Then
        Paso0.refrescarPersonajes
    End If
End Sub

Private Sub Paso4_Siguiente(headId As Integer, bodyId As Integer, ropaInteriorId As Integer, barbaId As Integer, peloId As Integer)
    DatosCreacion.Contraseña = RandomString(12)
    
    DatosCreacion.headId = headId
    DatosCreacion.bodyId = bodyId
    DatosCreacion.ropaInteriorId = ropaInteriorId
    DatosCreacion.barbaId = barbaId
    DatosCreacion.peloId = peloId
    
    RaiseEvent crear(DatosCreacion)
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

Private Sub Class_Terminate()
    Set vControles = Nothing
    Set Paso0 = Nothing
    Set Paso1 = Nothing
    Set Paso2 = Nothing
    Set Paso3 = Nothing
    Set Paso4 = Nothing
    Set NombreClase = Nothing
    Set NombreRazaGenero = Nothing
    Set AlineacionLabel = Nothing
    Set DatosCreacion = Nothing
    Set progress = Nothing
    Set progressSaliente = Nothing
    Set vWindowEntrante = Nothing
    Set vWindowSaliente = Nothing
End Sub

Public Sub vWindow_Hide()
    Visible = False
    
    If Not Paso0 Is Nothing Then
        Call GUI_Quitar(Paso0)
        Set Paso0 = Nothing
    End If
    
    If Not Paso1 Is Nothing Then
        Call GUI_Quitar(Paso1)
        Set Paso1 = Nothing
    End If
    
    If Not Paso2 Is Nothing Then
        Call GUI_Quitar(Paso2)
        Set Paso2 = Nothing
    End If
    
    If Not Paso3 Is Nothing Then
        Call GUI_Quitar(Paso3)
        Set Paso3 = Nothing
    End If
    
    If Not Paso4 Is Nothing Then
        Call GUI_Quitar(Paso4)
        Set Paso4 = Nothing
    End If


End Sub

Private Function vWindow_Render() As Boolean

    If Not Visible Then
        vWindow_Render = True
        Exit Function
    End If
    
    If Not progress Is Nothing Then
        vWindowEntrante.SetPos progress.Calcular, vWindowEntrante.Y
        If progress.isCompleto Then
            Set vWindowEntrante = Nothing
            Set progress = Nothing
        End If
    End If
    
    If Not progressSaliente Is Nothing Then
        vWindowSaliente.SetPos progressSaliente.Calcular, vWindowSaliente.Y
        If progressSaliente.isCompleto Then
            Call GUI_Quitar(vWindowSaliente)
            Set vWindowSaliente = Nothing
            Set progressSaliente = Nothing
        End If
    End If
    
    vWindow_Render = True
    vControles.Render
End Function

Public Sub vWindow_SetPos(ByVal X As Integer, ByVal Y As Integer)
    vX = X
    vY = Y
End Sub

Public Sub vWindow_Show()
    Visible = True
    
    If Not Paso4 Is Nothing Then
        Paso4.vWindow_Show
        Exit Sub
    End If
    
    If MiCuenta.personajes.count = 0 Then
        Call MostrarPaso1(False)
    Else
        Call MostrarPaso0(False)
    End If
End Sub

Private Sub OcultarPaso1(Optional ByVal reversa As Boolean = False)
    Call OcultarPasoX(Paso1, reversa)
    Set Paso1 = Nothing
End Sub

Private Sub OcultarPaso2(Optional ByVal reversa As Boolean = False)
    Call OcultarPasoX(Paso2, reversa)
    Set Paso2 = Nothing
End Sub
Private Sub OcultarPaso3(Optional ByVal reversa As Boolean = False)
    Call OcultarPasoX(Paso3, reversa)
    Set Paso3 = Nothing
End Sub

Private Sub OcultarPaso0(Optional ByVal reversa As Boolean = False)
    Call OcultarPasoX(Paso0, reversa)
    Set Paso0 = Nothing
End Sub


Private Sub OcultarPaso4(Optional ByVal reversa As Boolean = False)
    Call OcultarPasoX(Paso4, reversa)
    Set Paso4 = Nothing
End Sub

Private Sub MostrarPaso0(Optional ByVal reversa As Boolean = False, Optional ByVal sinAnimacion = False)
    
    If Not Paso0 Is Nothing Then Exit Sub
    
    Set Paso0 = New vwCuenta
    
    ultimoPaso = 0
    
    Paso0.vWindow_SetSize 600, 350
            
    Call MostrarPasoX(Paso0, reversa, sinAnimacion)
End Sub


Private Sub MostrarPaso1(Optional ByVal reversa As Boolean = False)
    Set Paso1 = New vwCrearPersonajePaso1
    
    Paso1.vWindow_SetSize 460, 350
    Paso1.setDatos DatosCreacion.Clase
            
    Call MostrarPasoX(Paso1, reversa)
End Sub

Private Sub MostrarPaso2(Optional ByVal reversa As Boolean = False)
    Set Paso2 = New VwCrearPersonajePaso2
    
    Paso2.vWindow_SetSize 460, 370
    Paso2.setDatos DatosCreacion.Genero, DatosCreacion.Raza
        
    Call MostrarPasoX(Paso2, reversa)
End Sub

Private Sub MostrarPaso3(Optional ByVal reversa As Boolean = False)
    Set Paso3 = New VwCrearPersonajePaso3
    
    Paso3.vWindow_SetSize 460, 340
    
    Paso3.setDatos DatosCreacion.Nombre, DatosCreacion.Alineacion
    
    Call MostrarPasoX(Paso3, reversa)
End Sub

Private Sub MostrarPaso4(Optional ByVal reversa As Boolean = False, Optional ByVal sinAnimacion = False)
    
    If Not Paso4 Is Nothing Then Exit Sub
    
    Set Paso4 = New vwCrearPersonajePaso4
   
    Paso4.vWindow_SetSize 380, 260
        
    ultimoPaso = 4
    
    Call MostrarPasoX(Paso4, reversa, sinAnimacion)
End Sub


Private Sub MostrarPasoX(paso As vWindow, reversa As Boolean, Optional ByVal sinAnimacion = False)
    Dim posicionInicialX As Integer
    
    posicionInicialX = Engine_Resolution.pixelesAncho / 2 - paso.width / 2  ' Siempre la centramos
    
    If sinAnimacion = False Then
        If reversa Then
            setPantallaEntrante paso, -460, posicionInicialX, WINDOW_TOP
        Else
            setPantallaEntrante paso, Engine_Resolution.pixelesAncho, posicionInicialX, WINDOW_TOP
        End If
    Else
        paso.SetPos posicionInicialX, WINDOW_TOP
        paso.Show
    End If
    
    Call GUI_Load(paso)
End Sub

Private Sub OcultarPasoX(paso As vWindow, reversa As Boolean)
    If reversa Then
        setPantallaSaliente paso, paso.X, Engine_Resolution.pixelesAncho, WINDOW_TOP
    Else
        setPantallaSaliente paso, paso.X, -paso.width, WINDOW_TOP
    End If
End Sub


Public Sub mostrarError(error As Byte, errordesc As String)
    If ultimoPaso = 4 Then
        Call MostrarPaso4(False, True)
        Call Paso4.mostrarError(error, errordesc)
    Else
        Call MostrarPaso0(False, True)
        Call Paso0.mostrarError(error, errordesc)
    End If
End Sub


Private Sub setPantallaEntrante(pantalla As vWindow, xInicial As Integer, xFinal As Integer, Y As Integer)
    Set progress = New clsProgreso
    progress.SetRango xInicial, xFinal
    progress.SetTicks GetTimer + 500
    pantalla.Show
    Set vWindowEntrante = pantalla
    vWindowEntrante.SetPos xInicial, Y
End Sub

Private Sub setPantallaSaliente(pantalla As vWindow, xInicial As Integer, xFinal As Integer, Y As Integer)
    Set progressSaliente = New clsProgreso
    progressSaliente.SetRango xInicial, xFinal
    progressSaliente.SetTicks GetTimer + 500
    Set vWindowSaliente = pantalla
    vWindowSaliente.SetPos xInicial, Y
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
