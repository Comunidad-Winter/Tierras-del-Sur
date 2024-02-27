VERSION 5.00
Begin VB.Form frmGuildAdm 
   BorderStyle     =   0  'None
   Caption         =   "Lista de Clanes Registrados"
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4140
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
   ScaleHeight     =   260
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   276
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox GuildsList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1980
      ItemData        =   "frmGuildAdm.frx":0000
      Left            =   405
      List            =   "frmGuildAdm.frx":0002
      TabIndex        =   0
      Top             =   1050
      Width           =   3195
   End
   Begin VB.Image Boton 
      Height          =   450
      Index           =   2
      Left            =   1515
      Top             =   3270
      Width           =   1065
   End
   Begin VB.Image Boton 
      Height          =   450
      Index           =   1
      Left            =   2715
      Top             =   3255
      Width           =   1065
   End
End
Attribute VB_Name = "frmGuildAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Selecionado As Byte
Public Sub ParseGuildList(ByVal Rdata As String)
Dim j As Integer, k As Integer
Dim informacion() As String


On Error Resume Next

informacion = Split(Rdata, ",")

For j = 0 To UBound(informacion) - 1
    guildslist.AddItem informacion(j)
Next j
Me.Show

End Sub

Private Sub Boton_Click(Index As Integer)
Select Case Index
Case 1
        If guildslist.ListIndex = -1 Then Exit Sub
        frmGuildBrief.EsLeader = False
        EnviarPaquete Paquetes.GuildDetail, guildslist.list(guildslist.ListIndex)
Case 2
        Unload Me
End Select
End Sub

Private Sub Boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Selecionado <> Index And Selecionado > 0 Then
    Boton(Selecionado).tag = "0"
    Boton(Selecionado).Picture = Nothing
    End If
    
    If Boton(Index).tag <> "1" Then
    Boton(Index).tag = "1"
    Selecionado = Index
    Call DameImagen(Boton(Index), Index + 75)
    End If
End Sub

Private Sub Form_Load()
Call CambiarCursor(frmGuildAdm)
DameImagenForm Me, 107
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Selecionado > 0 Then
    If Boton(Selecionado).tag = "1" Then
    Boton(Selecionado).tag = "0"
    Boton(Selecionado).Picture = Nothing
    End If
End If
End Sub
