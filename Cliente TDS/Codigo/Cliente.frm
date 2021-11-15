VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Recuperar Contraseña"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3060
   FillStyle       =   6  'Cross
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Left            =   3120
      Top             =   1200
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   11.25
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   11.25
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3120
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recuperar"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   14.25
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000009&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFC0&
      BorderWidth     =   4
      DrawMode        =   3  'Not Merge Pen
      FillColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Volver"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   14.25
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000009&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFC0&
      BorderWidth     =   4
      DrawMode        =   3  'Not Merge Pen
      FillColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nick del PJ:"
      BeginProperty Font 
         Name            =   "Modern"
         Size            =   11.25
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mail con el cual fue creado:"
      BeginProperty Font 
         Name            =   "Modern"
         Size            =   11.25
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************Misery_Ezequiel 28/05/05********************'
Option Explicit

Public IntentandoTirarElServer As Integer
Public Conectado As Integer

Private Sub Form_Load()
Me.Winsock1.Close
Winsock1.RemoteHost = frmConnect.IPTxt
Me.Winsock1.RemotePort = 888
Winsock1.Connect
Me.Timer1.enabled = True
Me.Timer1.Interval = 200
Me.Text1.enabled = False
Me.Text3.enabled = False
Me.Label1.enabled = False
Me.Label2.enabled = False
Me.Label3.enabled = False
Conectado = 1
End Sub

Private Sub Label1_Click()
Dim enviar, Nickdelpj, mail As String
Dim er As Long
If Me.Text1 = "" Then
er = MsgBox("Falta ingresar el Mail del personaje.", vbExclamation, "Recuperar Contraseña")

Exit Sub
End If
If Me.Text3 = "" Then
er = MsgBox("Falta ingresar el Mail del personaje.", vbExclamation, "Recuperar Contraseña")
Exit Sub
End If
Me.enabled = False

   Nickdelpj = Me.Text1
   mail = Me.Text3
   Winsock1.SendData "€" & Nickdelpj & "€" & mail
    DoEvents
Me.Caption = "Recuperando..."
End Sub

Private Sub Label4_Click()
Unload Me
End Sub

Private Sub Winsock1_Connect()
Me.Text1.enabled = True
Me.Text3.enabled = True
Me.Label1.enabled = True
Label2.enabled = True
Me.Label3.enabled = True
Me.Caption = "Conectado"
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim datos As String
Dim er As Long
Winsock1.GetData datos

Select Case datos
Case "N"
er = MsgBox("No se ha encontrado el personaje en nuestra base de datos.", vbCritical, "Recuperar Contraseña")
Me.enabled = True

Case "M"
er = MsgBox("No es el mismo mail con el que se registro el pj", vbCritical, "Recuperar Contraseña")
Me.enabled = True


Case "S"
er = MsgBox("Se ha originado un nuevo password y se ha enviado a: " + Me.Text3, vbInformation, "Recuperar Contraseña")
Me.enabled = True

Case "O"
Conectado = 1
End Select

Me.Caption = "Conectado"
DoEvents
Unload Me
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Me.Caption = "No se pudo Conectar"
End Sub
'********************Misery_Ezequiel 28/05/05********************'
