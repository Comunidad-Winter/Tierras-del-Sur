VERSION 5.00
Begin VB.Form frmCommet 
   BorderStyle     =   0  'None
   Caption         =   "Oferta de paz"
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4785
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
   ScaleHeight     =   222
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   319
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   1935
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   4335
   End
   Begin VB.Image Boton 
      Height          =   480
      Index           =   1
      Left            =   2940
      Top             =   2610
      Width           =   1125
   End
   Begin VB.Image Boton 
      Height          =   480
      Index           =   0
      Left            =   570
      Top             =   2610
      Width           =   1140
   End
End
Attribute VB_Name = "frmCommet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Selecionado As Byte
Public Nombre As String
Private Sub Boton_Click(Index As Integer)
Select Case Index
Case 0
Unload Me
Case 1
If Text1 = "" Then
    MsgBox "Debes redactar un mensaje solicitando la paz al lider de " & Nombre
    Exit Sub
End If
EnviarPaquete Paquetes.EnvPeaceOffer, Nombre & "," & Replace(Text1, vbCrLf, "º")
Unload Me
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
    Call DameImagen(Boton(Index), Index + 46)
    End If
End Sub

Private Sub Form_Load()
DameImagenForm Me, 108
Call CambiarCursor(frmCommet)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Boton(Selecionado).tag = "1" Then
    Boton(Selecionado).tag = "0"
    Boton(Selecionado).Picture = Nothing
    End If
End Sub
Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Boton(Selecionado).tag = "1" Then
    Boton(Selecionado).tag = "0"
    Boton(Selecionado).Picture = Nothing
    End If
End Sub
