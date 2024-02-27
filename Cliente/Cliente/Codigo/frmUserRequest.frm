VERSION 5.00
Begin VB.Form frmUserRequest 
   BorderStyle     =   0  'None
   Caption         =   "Peticion"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4800
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   1575
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   4335
   End
   Begin VB.Image Command1 
      Height          =   375
      Left            =   1860
      Top             =   2370
      Width           =   1095
   End
End
Attribute VB_Name = "frmUserRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Public Sub recievePeticion(ByVal p As String)
Text1 = Replace(p, "º", vbCrLf)
Call MostrarFormulario(Me, frmMain)
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Command1.tag <> "1" Then
    Command1.tag = "1"
    Call DameImagen(Command1, 51)
    End If
End Sub

Private Sub Form_Load()
Call CambiarCursor(frmUserRequest)
DameImagenForm Me, 112
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Command1.tag = "1" Then
    Command1.tag = "0"
    Command1.Picture = Nothing
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
End Sub

Private Sub Text1_Change()
    If Command1.tag = "1" Then
    Command1.tag = "0"
    Command1.Picture = Nothing
    End If
End Sub
