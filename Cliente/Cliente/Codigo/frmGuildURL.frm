VERSION 5.00
Begin VB.Form frmGuildURL 
   BorderStyle     =   0  'None
   Caption         =   "Oficial Web Site"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6270
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
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      MaxLength       =   25
      TabIndex        =   0
      Top             =   840
      Width           =   5895
   End
   Begin VB.Image Command1 
      Height          =   315
      Left            =   2610
      Top             =   1125
      Width           =   945
   End
End
Attribute VB_Name = "frmGuildURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If Text1 <> "" And Len(Text1) <= 25 Then _
    EnviarPaquete Paquetes.URLChange, Replace(Text1, vbCrLf, " ")
Unload Me
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Command1.tag <> "1" Then
Command1.tag = 1
Call DameImagen(Command1, 50)
End If
End Sub

Private Sub Form_Load()
Call CambiarCursor(frmGuildURL)
DameImagenForm Me, 91
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Command1.tag = "1" Then
Command1.tag = "0"
Command1.Picture = Nothing
End If
End Sub
