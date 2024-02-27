VERSION 5.00
Begin VB.Form frmEntrenador 
   BorderStyle     =   0  'None
   Caption         =   "Seleccione la criatura"
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4305
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
   ScaleHeight     =   274
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   287
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCriaturas 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   2970
   End
   Begin VB.Image Boton 
      Height          =   375
      Index           =   1
      Left            =   840
      Top             =   3405
      Width           =   1095
   End
   Begin VB.Image Boton 
      Height          =   375
      Index           =   0
      Left            =   2235
      Top             =   3405
      Width           =   1095
   End
End
Attribute VB_Name = "frmEntrenador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Selecionado As Byte
Private Sub Boton_Click(Index As Integer)
Select Case Index
Case 1
EnviarPaquete Paquetes.entrenador, lstCriaturas.list(lstCriaturas.ListIndex)
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
    Call DameImagen(Boton(Index), Index + 48)
    End If
End Sub

Private Sub Form_Load()
Call CambiarCursor(frmEntrenador)
DameImagenForm Me, 100
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Boton(Selecionado).tag = "1" Then
    Boton(Selecionado).tag = "0"
    Boton(Selecionado).Picture = Nothing
    End If
End Sub

Private Sub lstCriaturas_ItemCheck(item As Integer)
Me.lstCriaturas.ListIndex = -1
End Sub

Private Sub lstCriaturas_KeyDown(KeyCode As Integer, Shift As Integer)
Me.lstCriaturas.ListIndex = -1
End Sub

Private Sub lstCriaturas_KeyPress(KeyAscii As Integer)
Me.lstCriaturas.ListIndex = -1
End Sub

Private Sub lstCriaturas_KeyUp(KeyCode As Integer, Shift As Integer)
Me.lstCriaturas.ListIndex = -1
End Sub

Private Sub lstCriaturas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Boton(Selecionado).tag = "1" Then
    Boton(Selecionado).tag = "0"
    Boton(Selecionado).Picture = Nothing
    End If
End Sub
