VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Borrar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Borrar Personaje"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   3960
      TabIndex        =   10
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Volver"
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Borrar"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   2775
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2160
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Caption         =   "Si borras el personaje de nuestra base de datos NO habra ninguna forma de recuperarlo. "
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "IMPORTANTE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo de Seguridad"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Mail con el cual fue creado el pj:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Nick del PJ:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "Borrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************Misery_Ezequiel 28/05/05********************'
Option Explicit

Public Conectado As Integer
Public IntentandoTirarElServer As Integer

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Dim enviar, Nickdelpj, mail As String
Dim falta As Long
If Me.Text1 = "" Then
falta = MsgBox("Ingrese el nick del personaje por favor.", vbExclamation, "Borrar Personaje")
Exit Sub
End If
        If Me.Text3 = "" Then
        falta = MsgBox("Ingrese el mail del personaje por favor.", vbExclamation, "Borrar Personaje")
        Exit Sub
        End If
   
   Nickdelpj = Me.Text1
   mail = Me.Text3
   Winsock1.SendData "$" & Nickdelpj & "$" & mail & "$" & Me.Text2
   Form1.enabled = False
   Me.Command2.Caption = "Borrando.. Aguarde unos instantes por favor"
   DoEvents
End Sub

Private Sub Command3_Click()
'Me.Winsock1.SendData "p" + Me.Text4
'Frame3.Visible = True
Me.Width = 5450
End Sub

Private Sub Command5_Click()
'Frame3.Visible = False
Me.Width = 3180
End Sub

Private Sub Form_Load()
Me.Winsock1.Close
Winsock1.RemoteHost = frmConnect.IPTxt

Me.Winsock1.RemotePort = 888
Winsock1.Connect

Me.Text1.enabled = False
Me.Text3.enabled = False
Me.Command2.enabled = False
Me.Label2.enabled = False
Me.Label2.enabled = False
Me.Label3.enabled = False
Conectado = 1
End Sub

Private Sub Winsock1_Connect()
Me.Text1.enabled = True
Me.Text3.enabled = True
Me.Text2.enabled = True
Me.Label1.enabled = True
Me.Label2.enabled = True
Me.Label3.enabled = True
Me.Command2.enabled = True
Me.Caption = "Conectado"
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim datos As String
Dim er As Long
   Winsock1.GetData datos

Select Case datos
Case "0"
er = MsgBox("No se ha encontrado el personaje " & Me.Text1 & " en nuestra base de datos.", vbCritical, "Borrar Personaje")
Me.enabled = True

Case "1"
er = MsgBox("No es el mismo mail con el que se registro el pj", vbCritical, "Borrar Personaje")
Me.enabled = True

Case "2"
er = MsgBox("Se ha originado un nuevo codigo de seguridad y se ha enviado a: " + Me.Text3, vbInformation, "Borrar Personaje")
Me.enabled = True
Exit Sub

Case "3"
er = MsgBox("El Codigo de Seguridad no es el correcto, por favor verifiquelo e intente luego.", vbCritical, "Borrar Personaje")
Me.enabled = True

Case "4"
er = MsgBox("El personaje es mas de nivel 14. Por medidas de seguridad solo se permiten borrar personajes que sean de un nivel superior a 14.", vbCritical, "Borrar Personaje")
Me.enabled = True

Case "5"
er = MsgBox("El Personaje " & Me.Text1 & "ha sido borrado de nuestra base de datos.", vbInformation, "Borrar Personaje")
Me.enabled = True
Exit Sub

Case "6"
Conectado = 1

Case "7"
er = MsgBox("No ha pedido un codigo de seguridad con anticipacion o bien han pasado 24hs desde que lo hizo, por ende su personaje no ha sido borrado. Para borrarlo complete los campos 'Nick' y 'Mail' dejando en blanco el campo 'Codigo de Seguridad'. Aguarde unos instantes y le llegara una clave de seguridad a su mail. Complete los campos 'Nick' y 'Mail' y ponga su nuevo codigo de seguridad y podra borrar su personaje." & vbCrLf & vbCrLf & "Staff de Tierras Del Sur.", vbCritical, "Borrar Personaje")
Me.enabled = True

End Select
DoEvents

Me.Command2.Caption = "Borrar"
IntentandoTirarElServer = IntentandoTirarElServer + 1
If IntentandoTirarElServer > 10 Then
er = MsgBox("A intentado borrar fallidamente mas de 10 veces. Por cuestiones de seguridad su IP ha sido guardada e informada a los administradores. Disculpe las molestias", vbCritical, "Borrar Personaje")
Else
End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim gr As Long
gr = MsgBox("No se ha podido conectar con el server, por favor intente mas tarde.", vbCritical, "Recuperar Contraseña")
Me.Caption = "No se pudo Conectar"
End Sub
'********************Misery_Ezequiel 28/05/05********************'
