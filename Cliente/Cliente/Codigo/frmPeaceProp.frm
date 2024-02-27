VERSION 5.00
Begin VB.Form frmPeaceProp 
   BorderStyle     =   0  'None
   Caption         =   "Ofertas de paz"
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5190
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
   ScaleHeight     =   226
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   346
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lista 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      ItemData        =   "frmPeaceProp.frx":0000
      Left            =   240
      List            =   "frmPeaceProp.frx":0002
      TabIndex        =   0
      Top             =   600
      Width           =   4695
   End
   Begin VB.Image Boton 
      Height          =   420
      Index           =   3
      Left            =   3795
      Top             =   2805
      Width           =   1125
   End
   Begin VB.Image Boton 
      Height          =   435
      Index           =   2
      Left            =   2535
      Top             =   2790
      Width           =   1155
   End
   Begin VB.Image Boton 
      Height          =   435
      Index           =   1
      Left            =   1290
      Top             =   2790
      Width           =   1140
   End
   Begin VB.Image Boton 
      Height          =   465
      Index           =   0
      Left            =   195
      Top             =   2775
      Width           =   1020
   End
End
Attribute VB_Name = "frmPeaceProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Selecionado As Byte

Public Sub ParsePeaceOffers(ByVal s As String)
Dim r As Integer
Dim tt As Variant
tt = Split(s, ",")

For r = 0 To UBound(tt) - 1
    Call lista.AddItem(tt(r))
Next r
Me.Show vbModeless, frmMain
End Sub

Private Sub Boton_Click(Index As Integer)
Select Case Index
Case 0
Unload Me
Case 1
EnviarPaquete Paquetes.PEACEDET, lista.list(lista.ListIndex)
Case 2
EnviarPaquete Paquetes.PeaceAccpt, lista.list(lista.ListIndex)
Unload Me
Case 3
End Select
End Sub

Private Sub Boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Selecionado <> Index Then
    Boton(Selecionado).tag = "0"
    Boton(Selecionado).Picture = Nothing
    End If
    
    If Boton(Index).tag <> "1" Then
    Boton(Index).tag = "1"
    Selecionado = Index
    Call DameImagen(Boton(Index), Index + 72)
    End If
End Sub

Private Sub Form_Load()
Call CambiarCursor(frmPeaceProp)
DameImagenForm Me, 109
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Boton(Selecionado).tag = "1" Then
    Boton(Selecionado).tag = "0"
    Boton(Selecionado).Picture = Nothing
    End If
End Sub
