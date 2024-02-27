VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Juego Tierras del Sur"
   ClientHeight    =   11235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14265
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   720
   ScaleMode       =   0  'User
   ScaleWidth      =   1280
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin InetCtlsObjects.Inet inetConectarCuenta 
      Left            =   2160
      Top             =   8160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer tmrConectarCuenta 
      Left            =   2400
      Top             =   9240
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6720
      Left            =   1800
      ScaleHeight     =   448
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   656
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   9840
   End
   Begin VB.Timer tmrInnetConnect 
      Left            =   1560
      Top             =   6480
   End
   Begin InetCtlsObjects.Inet inetConectorWeb 
      Left            =   2280
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
      RequestTimeout  =   5
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   36
      ImageHeight     =   36
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConnect.frx":0E42
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private navegadorWeb As navegadorWeb

Public Sub CrearCuenta(nombreCuenta As String, Password As String, Email As String)

' Conectamos con la cuenta
Dim callbackCrearCuenta As CallbakCrearCuenta

Dim request As CHTTPRequest
Dim base64Converter As New base64Converter

Set callbackCrearCuenta = New CallbakCrearCuenta
Set request = New CHTTPRequest
Set navegadorWeb = New navegadorWeb

Dim body As Dictionary

Set body = New Dictionary

Debug.Print base64Converter.Encode(Password)

Call body.Add("nombre", nombreCuenta)
Call body.Add("email", Email)
Call body.Add("password", base64Converter.Encode(Password))

request.Host = WEB_API
request.Path = "usuario"
request.method = eHttpMethod.httppost
request.UserAgent = "TDSExternalUser"
request.body = JSON.toString(body)

Call callbackCrearCuenta.setDatos(nombreCuenta, Password)

Call navegadorWeb.ejecutarConsulta(Me.inetConectarCuenta, Me.tmrConectarCuenta, request, callbackCrearCuenta)

End Sub


Public Sub cargarPersonajes()

' Conectamos con la cuenta
Dim callBackCOnectar As CallBackObtenerPersonajes

Dim request As CHTTPRequest
Dim base64Converter As New base64Converter

Set callBackCOnectar = New CallBackObtenerPersonajes
Set request = New CHTTPRequest
Set navegadorWeb = New navegadorWeb

request.Host = WEB_API
request.Path = "juego/personajes"
request.method = eHttpMethod.httpGET
request.UserAgent = "TDSExternalUser"

Call request.addHeader("Authorization", "Bearer " & MiCuenta.cuenta.Token)

Call navegadorWeb.ejecutarConsulta(Me.inetConectarCuenta, Me.tmrConectarCuenta, request, callBackCOnectar)

End Sub

Public Sub crearPersonaje()
    Dim infoLogin As modConectar.retornoInfo
        
    ' Evitamos que se llame dos veces
    If Not EstadoConexion = E_Estado.Ninguno Then
        Call LoginInit
        Exit Sub
    End If
    
    UserName = DatosCreacion.Nombre
    
    Call modMiPersonaje.iniciar
    
    infoLogin = modConectar.conectarParaCrear

    If (infoLogin.error > 0) Then
        EstadoConexion = E_Estado.Ninguno
        Call modDibujarInterface.mostrarError(infoLogin.error, infoLogin.errordesc)
    Else
        EstadoConexion = E_Estado.conectado
        Call modConectar.conectar(infoLogin.datos)
    End If
    
End Sub
Public Sub conectarPersonaje(nombrePersonaje As String, clavePersonaje As String)

Dim infoLogin As modConectar.retornoInfo

UserName = nombrePersonaje
UserPassword = clavePersonaje

' Evitamos que se llame dos veces
If Not EstadoConexion = E_Estado.Ninguno Then
    Call LoginInit
    Exit Sub
End If

Call modMiPersonaje.iniciar

EstadoConexion = E_Estado.Conectando

infoLogin = modConectar.conectarPersonaje()

If (infoLogin.error > 0) Then
    EstadoConexion = E_Estado.Ninguno
    Call modDibujarInterface.mostrarError(infoLogin.error, infoLogin.errordesc)
Else
    EstadoConexion = E_Estado.conectado
    Call modConectar.conectar(infoLogin.datos)
End If

End Sub
Public Sub conectar(nombreCuenta As String, clave As String)

' Conectamos con la cuenta
Dim callBackCOnectar As CallBackConectarCuenta

Dim request As CHTTPRequest
Dim base64Converter As New base64Converter

Set callBackCOnectar = New CallBackConectarCuenta
Set request = New CHTTPRequest
Set navegadorWeb = New navegadorWeb

request.Host = WEB_API
request.Path = "sesion"
request.method = eHttpMethod.httpGET
request.UserAgent = "TDSExternalUser"

Debug.Print base64Converter.Encode(nombreCuenta & ":" & clave)
Call request.addHeader("Authorization", "Basic " & base64Converter.Encode(nombreCuenta & ":" & clave))

Call navegadorWeb.ejecutarConsulta(Me.inetConectarCuenta, Me.tmrConectarCuenta, request, callBackCOnectar)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    prgRun = False
End If
   

End Sub

Private Sub Form_Load()

' Seteamos tamaño
Me.width = Engine_Resolution.pixelesAncho * Screen.TwipsPerPixelX
Me.Height = Engine_Resolution.pixelesAlto * Screen.TwipsPerPixelY

Me.picInv.top = 0
Me.picInv.left = 0
Me.picInv.width = Engine_Resolution.pixelesAncho * Screen.TwipsPerPixelX
Me.picInv.Height = Engine_Resolution.pixelesAlto * Screen.TwipsPerPixelY

If CursorPer = 1 Then
    Call CambiarCursor(frmConnect)
End If

End Sub

Private Sub picInv_KeyDown(KeyCode As Integer, Shift As Integer)
   Call GUI_KeyDown(KeyCode, Shift)
End Sub

Private Sub picInv_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyTab Then
        Call GUI_AdvanceFoucs
        Exit Sub
    End If
    
    Call GUI_Keypress(KeyAscii)
End Sub

Private Sub picInv_KeyUp(KeyCode As Integer, Shift As Integer)
    Call GUI_KeyUp(KeyCode, Shift)
End Sub

Private Sub picInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call GUI_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call GUI_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call GUI_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub tmrInnetConnect_Timer()
    'EL DNS no se pudo resolver aun?
    If Me.inetConectorWeb.tag = "icResolvingHost" Or Me.inetConectorWeb.tag = "icConnecting" Then
        Me.inetConectorWeb.tag = "icConnectingTimeOut"
        Call inetConectorWeb.Cancel
    End If
End Sub
