VERSION 5.00
Begin VB.Form cpassword 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   2955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
   Begin VB.Image Boton 
      Height          =   480
      Index           =   1
      Left            =   1590
      Top             =   2430
      Width           =   1230
   End
   Begin VB.Image Boton 
      Height          =   480
      Index           =   0
      Left            =   130
      Top             =   2430
      Width           =   1230
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password nuevo"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Re ingrese su pasword nuevo"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password viejo"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "cpassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Selecionado As Byte

Private Sub Boton_Click(Index As Integer)
Select Case Index
    Case 1
        If Text2 <> Text3 Then
            AddtoRichTextBox frmConsola.ConsolaFlotante, "Verificá haber ingresado correctamente la password nueva.", 65, 190, 156: Exit Sub
        Else
            Call sSendData(Paquetes.Comandos, Complejo.PASSWD, Me.Text1 & "@" & Me.Text2)
        End If
End Select
frmMain.Enabled = True
Unload Me
End Sub

Private Sub Boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Selecionado <> Index Then
    Boton(Selecionado).tag = "0"
    Boton(Selecionado).Picture = Nothing
    End If
    
    If Boton(Index).tag <> "1" Then
    Boton(Index).tag = "1"
    Selecionado = Index
    Call DameImagen(Boton(Index), 158 + Index)
    End If
End Sub

Private Sub Form_Load()
DameImagenForm Me, 157
Call CambiarCursor(cpassword)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Boton(Selecionado).tag = "0"
Boton(Selecionado).Picture = Nothing
End Sub

