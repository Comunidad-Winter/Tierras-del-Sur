VERSION 5.00
Begin VB.Form frmGuildSol 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "Ingreso"
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4875
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
   ScaleHeight     =   269
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   1455
      Left            =   360
      MaxLength       =   150
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Image Boton 
      Height          =   375
      Index           =   1
      Left            =   3240
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Image Boton 
      Height          =   375
      Index           =   0
      Left            =   375
      Top             =   3315
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSolicitud.frx":0000
      ForeColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
End
Attribute VB_Name = "frmGuildSol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Selecionado As Byte
Dim CName As String
Public Sub RecieveSolicitud(ByVal GuildName As String)
CName = GuildName
End Sub

Private Sub Boton_Click(Index As Integer)

Select Case Index
    Case 0
        EnviarPaquete Paquetes.GuildSol, CName & "¬" & Replace(Text1, vbCrLf, " ")
End Select

Unload Me
End Sub

Private Sub Boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Selecionado <> Index Then
    Boton(Selecionado).tag = "0"
    Boton(Selecionado).Picture = Nothing
    End If
    
    If Boton(Index).tag <> "1" Then
    Boton(Index).tag = "1"
    Selecionado = Index
    Call DameImagen(Boton(Index), Index + 78)
    End If
End Sub

Private Sub Form_Load()
DameImagenForm Me, 105
Call CambiarCursor(frmGuildSol)
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
