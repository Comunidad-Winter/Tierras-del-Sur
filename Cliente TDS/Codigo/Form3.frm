VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lingotear"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2310
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   2310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "1"
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Trabajar"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad:"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If IScombate Then Exit Sub
If Val(Me.Text1) > 0 Then
    Cantidadlingo = 1
    If UserEstado = 1 Then Exit Sub
    Istrabajando = True
    Lingoteando = True
    frmMain.Macro.enabled = True
    AddtoRichTextBox frmMain.RecTxt, "Empiezas a trabajar.", 255, 0, 0, True, False, False
Else
End If
Me.Hide
End Sub


