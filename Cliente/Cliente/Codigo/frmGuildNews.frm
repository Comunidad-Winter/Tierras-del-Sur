VERSION 5.00
Begin VB.Form frmGuildNews 
   BorderStyle     =   0  'None
   Caption         =   "GuildNews"
   ClientHeight    =   6780
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   452
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   339
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox aliados 
      Appearance      =   0  'Flat
      Height          =   1005
      ItemData        =   "frmGuildNews.frx":0000
      Left            =   360
      List            =   "frmGuildNews.frx":0002
      TabIndex        =   2
      Top             =   5040
      Width           =   4335
   End
   Begin VB.ListBox guerra 
      Appearance      =   0  'Flat
      Height          =   1005
      ItemData        =   "frmGuildNews.frx":0004
      Left            =   360
      List            =   "frmGuildNews.frx":0006
      TabIndex        =   1
      Top             =   3480
      Width           =   4335
   End
   Begin VB.TextBox news 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   800
      Width           =   4335
   End
   Begin VB.Image Label1 
      Height          =   465
      Left            =   1965
      Top             =   6255
      Width           =   1125
   End
End
Attribute VB_Name = "frmGuildNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ParseGuildNews(ByVal s As String)

Me.news = s

Me.Show , frmMain
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Unload Me
frmMain.SetFocus
End If
End Sub

Private Sub Form_Load()
Call CambiarCursor(Me)
DameImagenForm Me, 106
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Label1.tag = "1" Then
Label1.tag = "0"
Label1.Picture = Nothing
End If
End Sub

Private Sub Label1_Click()
'on error Resume Next
Unload Me
frmMain.SetFocus
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Label1.tag <> "1" Then
    Label1.tag = "1"
    Call DameImagen(Label1, 8)
    End If
End Sub
