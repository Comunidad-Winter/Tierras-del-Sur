VERSION 5.00
Begin VB.Form frmMSG 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensajes de GMs"
   ClientHeight    =   3270
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   2355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   2355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      MouseIcon       =   "frmMSG.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2685
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   180
      TabIndex        =   1
      Top             =   450
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuarios"
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   615
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
      Begin VB.Menu mnuBorrar 
         Caption         =   "Borrar mensaje"
      End
   End
End
Attribute VB_Name = "frmMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_GM_MSG = 300

Private MisMSG(0 To MAX_GM_MSG) As String
Private Apunt(0 To MAX_GM_MSG) As Integer

Public Sub CrearGMmSg(Nick As String, Msg As String)
If List1.ListCount < MAX_GM_MSG Then
        List1.AddItem Nick & "-" & List1.ListCount
        MisMSG(List1.ListCount - 1) = Msg
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

Private Sub Form_Load()
List1.Clear
End Sub

Private Sub List1_Click()
Dim ind As Integer
ind = val(ReadField(2, List1.list(List1.ListIndex), Asc("-")))
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
    PopupMenu menU_usuario
End If
End Sub

Private Sub mnuBorrar_Click()
If List1.ListIndex < 0 Then Exit Sub
Call sSendData(Paquetes.ComandosConse, , List1.list(List1.ListIndex))
List1.RemoveItem List1.ListIndex
End Sub

Private Sub mnuIR_Click()
Call sSendData(Paquetes.ComandosConse, Conse2.IRA, List1.list(List1.ListIndex))
End Sub

Private Sub mnutraer_Click()
Call sSendData(Paquetes.ComandosSemi, SemiDios2.CSUM, List1.list(List1.ListIndex))
End Sub
