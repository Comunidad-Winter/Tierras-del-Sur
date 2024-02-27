VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vwLogin"
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

Private WithEvents vControles As vControles
Attribute vControles.VB_VarHelpID = -1

Private ObjetoFocus As Object

'Implementacion de ventanas. No modificar
Private pNext As vWindow
Private pPrev As vWindow

Private WithEvents formConectar As vwConectar
Attribute formConectar.VB_VarHelpID = -1
Private WithEvents formCuenta As VwCrearPersonaje
Attribute formCuenta.VB_VarHelpID = -1
Private WithEvents formCrear As vwCrearCuenta
Attribute formCrear.VB_VarHelpID = -1

Private menuSuperior As vWindow
Private label As clsGUIText
Private cargando As ControlCargando

Private Const WINDOW_TOP As Integer = 260

Private Visible As Boolean

Public Sub mostrarCuenta()

If Not formCuenta Is Nothing Then
    formCuenta.refrescarPersonajes
    Exit Sub
End If

Call ocultarCargando

If Not formConectar Is Nothing Then
    Call GUI_Quitar(formConectar)
    Set formConectar = Nothing
End If

If Not formCrear Is Nothing Then
    Call GUI_Quitar(formCrear)
    Set formCrear = Nothing
End If

Call mostrarFormCuenta

End Sub

Public Sub mostrarError(error As Byte, errordesc As String)
    If error = modLogin.JUEGO_DESACTUALIZADO Then
        Call juegoDesactualizado
    End If
    
    Call ocultarCargando
    
    If Not formConectar Is Nothing Then
        Call formConectar.mostrarError(error, errordesc)
        Call formConectar.vWindow_Show
    ElseIf Not formCrear Is Nothing Then
        Call formCrear.mostrarError(error, errordesc)
        Call formCrear.vWindow_Show
    ElseIf Not formCuenta Is Nothing Then
        Call formCuenta.mostrarError(error, errordesc)
        Call formCuenta.vWindow_Show
    End If
End Sub

Private Sub ocultarFormCuenta()
    If Not formCuenta Is Nothing Then
        Call formCuenta.vWindow_Hide
        Call GUI_Quitar(formCuenta)
        Set formCuenta = Nothing
    End If
End Sub

Private Sub formCrear_Atras()
    Call GUI_Quitar(formCrear)

    Set formCrear = Nothing
        
    Call MostrarConectar
End Sub

Private Sub formCrear_CrearCuenta(nombreCuenta As String, Password As String, Email As String)
    Call formCrear.vWindow_Hide
    
    Call mostrarCargando("Creando cuenta")
    
    Call frmConnect.CrearCuenta(nombreCuenta, Password, Email)
End Sub

Private Sub formCuenta_Cerrar()
    Call MiCuenta.cerrarSesion
    
    Call ocultarFormCuenta
    
    Call MostrarConectar
End Sub

Private Sub formCuenta_conectar(personaje As String, Password As String)
    Call formCuenta.vWindow_Hide
        
    Call mostrarCargando("Ingresando")
    
    EstadoLogin = E_MODO.IngresarPersonaje
    
    Call frmConnect.conectarPersonaje(personaje, Password)
End Sub

Private Sub formCuenta_crear(datos As CrearPersonajeDTO)
    EstadoLogin = E_MODO.PantallaCreacion
        
    Set modCrearPersonaje.DatosCreacion = datos
    
    Call formCuenta.vWindow_Hide
    
    Call mostrarCargando("Iniciando una nueva aventura")
    
    Call frmConnect.crearPersonaje
End Sub

Private Property Get vWindow_IsVisible() As Boolean

End Property

Public Sub vWindow_SetPos(ByVal X As Integer, ByVal Y As Integer)
    vX = X
    vY = Y
End Sub

Public Sub vWindow_SetSize(ByVal width As Integer, ByVal Height As Integer)
    vW = width
    vH = Height
End Sub

Public Sub vWindow_Show()
    ' Menu superior
    Set menuSuperior = New vwMenu
    Call menuSuperior.SetPos(Engine_Resolution.pixelesAncho - 35 * 4, 15)
    Call GUI_Load(menuSuperior)
    
    If Not MiCuenta.cuenta Is Nothing Then
        frmConnect.cargarPersonajes
        Call mostrarCuenta
    Else
        Call MostrarConectar
    End If
    
    Visible = True
End Sub

Private Sub BotonSalir_Click()
    prgRun = False
End Sub

Private Sub formConectar_Conectar(usuario As String, Password As String, recordarClave As Boolean)
    Call formConectar.vWindow_Hide
    
    Call mostrarCargando("Ingresando")
    
    Call frmConnect.conectar(usuario, Password)
End Sub

Private Sub formConectar_crear()
    Call GUI_Quitar(formConectar)
    Set formConectar = Nothing
    Set formCrear = New vwCrearCuenta

    formCrear.vWindow_SetSize 400, 290
    formCrear.vWindow_SetPos Engine_Resolution.pixelesAncho / 2 - formCrear.vWindow_Width / 2, WINDOW_TOP
    
    Call GUI_Load(formCrear)
    formCrear.vWindow_Show
End Sub

Private Sub formCrear_cerrar()
    ' Desconectamos
    frmMain.Socket1.Disconnect
    frmMain.Socket1.Cleanup
    EstadoConexion = E_Estado.Ninguno
    EstadoLogin = E_MODO.Ninguno
    
    Call GUI_Quitar(formCrear)
    Set formCrear = Nothing
    
    Call MostrarConectar
End Sub

Private Sub ocultarCargando()
    If Not cargando Is Nothing Then
        Set cargando = Nothing
    End If
End Sub
Private Sub mostrarCargando(texto As String)
    If cargando Is Nothing Then
        Set cargando = New ControlCargando
    End If
        
    cargando.SetPos Engine_Resolution.pixelesAncho / 2, Engine_Resolution.pixelesAlto / 2
    cargando.text = texto
    cargando.SetStyle 2, mzInterfaceColor1, mzInterfaceColor2, dVertical
End Sub

Private Property Get vWindow_GetControl() As vControles
    Set vWindow_GetControl = vControles
End Property

    Private Function vWindow_GetNext(vWindow As vWindow) As Boolean: Set vWindow = pNext: vWindow_GetNext = Not (pNext Is Nothing): End Function
    Private Function vWindow_GetPrev(vWindow As vWindow) As Boolean: Set vWindow = pPrev: vWindow_GetPrev = Not (pPrev Is Nothing): End Function
    Private Sub vWindow_SetNext(vWindow As vWindow): Set pNext = vWindow: End Sub
    Private Sub vWindow_SetPrev(vWindow As vWindow): Set pPrev = vWindow: End Sub

'/Implementacion

Private Sub Class_Terminate()
    Set cargando = Nothing
    Set vControles = Nothing
End Sub

Private Sub Class_Initialize()
    Set vControles = New vControles
End Sub

Private Sub vWindow_Hide()
    If Not menuSuperior Is Nothing Then
        Call GUI_Quitar(menuSuperior)
        Set menuSuperior = Nothing
    End If
    
    If Not formConectar Is Nothing Then
        Call GUI_Quitar(formConectar)
        Set formConectar = Nothing
    End If
    
    If Not formCuenta Is Nothing Then
        Call GUI_Quitar(formCuenta)
        Set formCuenta = Nothing
    End If
    
    Visible = False
End Sub

Private Function vWindow_Render() As Boolean
    vWindow_Render = True

    If Not label Is Nothing Then
        label.Render
    End If
    
    If Not cargando Is Nothing Then
        cargando.Render
    End If
    vControles.Render
End Function

Private Sub mostrarFormCuenta()
    ' Arrancamos con el conectar
    Set formCuenta = New VwCrearPersonaje
    Call GUI_Load(formCuenta)
    Call formCuenta.vWindow_Show
    
  '  formCuenta.vWindow_SetSize 400, 400
    formCuenta.vWindow_SetPos Engine_Resolution.pixelesAncho / 2 - formCuenta.vWindow_Width / 2, WINDOW_TOP
End Sub

Private Sub MostrarConectar()
    ' Arrancamos con el conectar
    Set formConectar = New vwConectar
    Call GUI_Load(formConectar)
    Call formConectar.vWindow_Show
    
    formConectar.vWindow_SetSize 150, 250
    formConectar.vWindow_SetPos Engine_Resolution.pixelesAncho / 2 - formConectar.vWindow_Width / 2, WINDOW_TOP
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

