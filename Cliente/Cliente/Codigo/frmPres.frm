VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmPres 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Juego Tierras del Sur"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1024
   ScaleMode       =   0  'User
   ScaleWidth      =   768
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1320
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer tmrUpdater 
      Left            =   1200
      Top             =   3960
   End
   Begin VB.Label lblActualizaciones 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscando actualizaciones...."
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   5280
      TabIndex        =   1
      Top             =   8520
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Label lblCargando 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pulsa la tecla BLOQ NÚM para mantenerte caminando"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   480
      Left            =   5040
      TabIndex        =   0
      Top             =   6720
      Width           =   3210
      WordWrap        =   -1  'True
   End
   Begin VB.Image logo 
      Height          =   3840
      Left            =   4680
      Top             =   2640
      Width           =   3840
   End
End
Attribute VB_Name = "frmPres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private navegador As navegadorWeb
Private callback As CallbackUpdater

Private Sub Form_Load()
    ' Seteamos tamaño
    frmPres.width = Engine_Resolution.pixelesAncho * Screen.TwipsPerPixelX
    frmPres.Height = Engine_Resolution.pixelesAlto * Screen.TwipsPerPixelY
    
    ' Imagen del logo
    frmPres.logo.Picture = LoadPicture(app.Path & "\Recursos\LOGO.jpg")
    
    ' Posicionamos el logo
    Me.logo.left = Me.ScaleWidth / 2 - logo.width / 2
    Me.logo.top = Me.ScaleHeight / 2 - logo.Height / 2 - Me.ScaleHeight / 10
    
    Me.lblCargando.left = Me.logo.left
    Me.lblCargando.top = Me.logo.top + Me.logo.Height + 20
    
    Me.lblCargando.width = logo.width
    
    Me.lblActualizaciones.top = Me.ScaleHeight - 30
    Me.lblActualizaciones.left = Me.logo.left
    Me.lblActualizaciones.width = Me.lblCargando.width
End Sub

Public Sub checkJuego()
    Dim request As CHTTPRequest
    Set request = New CHTTPRequest
    Set navegador = New navegadorWeb
    Set callback = New CallbackUpdater
    
    Me.lblActualizaciones.Visible = True
     
    request.Host = HTTP_URL
    request.UserAgent = "ExternalUser"
    request.method = httpGET
    request.Path = UPDATER_PATH & "?v=" & Configuracion_Usuario.versionActual
    
    'Ejecutamos
    Call navegador.ejecutarConsulta(Me.Inet1, Me.tmrUpdater, request, callback)
End Sub

Public Function isReady() As Boolean
    isReady = Not (callback.status = 0)
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set navegador = Nothing
    Set callback = Nothing
End Sub

