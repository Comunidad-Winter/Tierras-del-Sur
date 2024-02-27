VERSION 5.00
Begin VB.Form frmHerrero 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "Herrero"
   ClientHeight    =   4605
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   307
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   307
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "1"
      Top             =   3600
      Width           =   4095
   End
   Begin VB.ListBox lstArmas 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   4080
   End
   Begin VB.ListBox lstArmaduras 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   4080
   End
   Begin VB.Image Boton 
      Height          =   480
      Index           =   3
      Left            =   420
      Top             =   3990
      Width           =   1140
   End
   Begin VB.Image Boton 
      Height          =   450
      Index           =   2
      Left            =   2835
      Top             =   4020
      Width           =   1305
   End
   Begin VB.Image Boton 
      Height          =   420
      Index           =   1
      Left            =   2655
      Top             =   660
      Width           =   1455
   End
   Begin VB.Image Boton 
      Height          =   435
      Index           =   0
      Left            =   405
      Top             =   645
      Width           =   1155
   End
End
Attribute VB_Name = "frmHerrero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Selecionado As Byte
Private Sub Boton_Click(Index As Integer)
Select Case Index
Case 0
            lstArmaduras.Visible = False
            lstArmas.Visible = True
Case 1
            lstArmaduras.Visible = True
            lstArmas.Visible = False
Case 2
            Me.Text1 = val(Me.Text1)
            If Me.Text1 <= 0 Then Exit Sub
            
            If IScombate = True Then
            AddtoRichTextBox frmConsola.ConsolaFlotante, "No puedes trabajar en modo combate.", 255, 0, 0, True, False, False
            Else
                If frmHerrero.lstArmas.Visible Then
                    If frmHerrero.lstArmas.ListIndex = -1 Then Exit Sub
                EnviarPaquete Paquetes.CHerrero, ITS(Me.Text1) & Codify(ArmasHerrero(frmHerrero.lstArmas.ListIndex))
                Else
                    If frmHerrero.lstArmaduras.ListIndex = -1 Then Exit Sub
                EnviarPaquete Paquetes.CHerrero, ITS(Me.Text1) & Codify(ArmadurasHerrero(frmHerrero.lstArmaduras.ListIndex))
                End If
            Unload Me
            End If
Case 3
Unload Me
End Select
End Sub
Private Sub Boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Selecionado <> Index Then
Boton(Selecionado).tag = "0"
Boton(Selecionado).Picture = Nothing
End If
If Boton(Index).tag <> "1" Then
Boton(Index).tag = "1"
Selecionado = Index
Call DameImagen(Boton(Index), Index + 82)
End If
End Sub
Private Sub Form_Load()
Call CambiarCursor(frmHerrero)
DameImagenForm Me, 104
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Boton(Selecionado).tag = "1" Then
Boton(Selecionado).tag = "0"
Boton(Selecionado).Picture = Nothing
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
End Sub

Private Sub lstArmas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Boton(Selecionado).tag = "1" Then
Boton(Selecionado).tag = "0"
Boton(Selecionado).Picture = Nothing
End If
End Sub
Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Boton(Selecionado).tag = "1" Then
Boton(Selecionado).tag = "0"
Boton(Selecionado).Picture = Nothing
End If
End Sub
