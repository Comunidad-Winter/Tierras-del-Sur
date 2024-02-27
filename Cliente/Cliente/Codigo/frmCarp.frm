VERSION 5.00
Begin VB.Form frmCarp 
   BorderStyle     =   0  'None
   Caption         =   "Carpintero"
   ClientHeight    =   4365
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   291
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "1"
      Top             =   3240
      Width           =   4095
   End
   Begin VB.ListBox lstArmas 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   4080
   End
   Begin VB.Image Boton 
      Height          =   375
      Index           =   1
      Left            =   285
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Image Boton 
      Height          =   375
      Index           =   0
      Left            =   3105
      Top             =   3705
      Width           =   1215
   End
End
Attribute VB_Name = "frmCarp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Selecionado As Byte

Private Sub Boton_Click(index As Integer)
Select Case index
Case 0
    Me.Text1 = val(Me.Text1)
    If Me.Text1 <= 0 Or Me.Text1 > 9999 Or frmCarp.lstArmas.ListIndex < 0 Then Exit Sub
    If IScombate = True Then
    AddtoRichTextBox frmConsola.ConsolaFlotante, "No puedes trabajar en modo combate.", 255, 0, 0, True, False, False
    Else
    EnviarPaquete Paquetes.CCarpintero, ITS(Me.Text1) & Codify(ObjCarpintero(frmCarp.lstArmas.ListIndex).index)
    End If
End Select
Unload frmCarp
End Sub

Private Sub Boton_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Selecionado <> index Then
Boton(Selecionado).tag = "0"
Boton(Selecionado).Picture = Nothing
End If
If Boton(index).tag <> "1" Then
Boton(index).tag = "1"
Selecionado = index
Call DameImagen(Boton(index), index + 19)
End If
End Sub

Private Sub Form_Load()
Call CambiarCursor(frmCarp)
DameImagenForm Me, 94
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Boton(Selecionado).tag = "1" Then
Boton(Selecionado).tag = "0"
Boton(Selecionado).Picture = Nothing
End If
End Sub

