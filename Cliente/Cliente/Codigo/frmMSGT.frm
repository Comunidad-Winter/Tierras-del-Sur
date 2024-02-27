VERSION 5.00
Begin VB.Form frmMSGT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuarios Trabajando"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2430
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3460.485
   ScaleMode       =   0  'User
   ScaleWidth      =   2430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   180
      TabIndex        =   1
      Top             =   425
      Width           =   1980
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   180
      MouseIcon       =   "frmMSGT.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2685
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuarios trabajando"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1410
   End
   Begin VB.Menu menU_usuario 
      Caption         =   "Usuario"
      Visible         =   0   'False
      Begin VB.Menu mnuIR 
         Caption         =   "Ir donde esta el usuario"
      End
      Begin VB.Menu mnutraer 
         Caption         =   "Traer usuario"
      End
   End
End
Attribute VB_Name = "frmMSGT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coded By Ezequiel Matías Montero
'****************************************************************
'****************************************************************
'****************************************************************
'NO TOCAR NI MODIFICAR, POSIBLES ERRORES SI LO HACES
'****************************************************************
'****************************************************************
'****************************************************************
Option Explicit

Private Const MAX_GM_MSG = 800

Private MisMSG(0 To MAX_GM_MSG) As String
Private Apunt(0 To MAX_GM_MSG) As Integer

Public Sub CrearGMmSg(Nick As String, msg As String)
If List1.ListCount < MAX_GM_MSG Then
        List1.AddItem Nick & "-" & List1.ListCount
        MisMSG(List1.ListCount - 1) = msg
        Apunt(List1.ListCount - 1) = List1.ListCount - 1
End If
End Sub

Private Sub Command1_Click()
Me.Visible = False
List1.Clear
End Sub

Private Sub Form_Deactivate()
Me.Visible = False
List1.Clear
End Sub

Private Sub List1_Click()
Dim ind As Integer
ind = Val(ReadField(2, List1.List(List1.ListIndex), Asc("-")))
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu menU_usuario
End If
End Sub

Private Sub mnuIR_Click()
EnviarPaquete Conse2.IRA, ReadField(1, List1.List(List1.ListIndex), Asc("-"))
End Sub

Private Sub mnutraer_Click()
EnviarPaquete SemiDios2.CSUM, ReadField(1, List1.List(List1.ListIndex), Asc("-"))
End Sub
