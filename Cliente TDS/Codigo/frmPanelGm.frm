VERSION 5.00
Begin VB.Form frmPanelGm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel GM"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4665
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Anticheat"
      Height          =   315
      Index           =   20
      Left            =   2460
      TabIndex        =   23
      Top             =   2600
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Banco"
      Height          =   315
      Index           =   19
      Left            =   180
      TabIndex        =   22
      Top             =   2600
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Ver S.O.S"
      Height          =   315
      Index           =   18
      Left            =   2460
      TabIndex        =   21
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Boveda"
      Height          =   315
      Index           =   17
      Left            =   180
      TabIndex        =   20
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Banear ip"
      Height          =   315
      Index           =   16
      Left            =   3600
      TabIndex        =   19
      Top             =   1400
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Penas"
      Height          =   315
      Index           =   15
      Left            =   180
      TabIndex        =   18
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Nicks del ip"
      Height          =   315
      Index           =   14
      Left            =   2460
      TabIndex        =   17
      Top             =   1000
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Ip del nick"
      Height          =   315
      Index           =   13
      Left            =   2460
      TabIndex        =   16
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Unbanear"
      Height          =   315
      Index           =   12
      Left            =   3600
      TabIndex        =   15
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Carcel"
      Height          =   315
      Index           =   11
      Left            =   1320
      TabIndex        =   14
      Top             =   1400
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Skillpoints"
      Height          =   315
      Index           =   10
      Left            =   180
      TabIndex        =   13
      Top             =   2200
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Inventario"
      Height          =   315
      Index           =   9
      Left            =   180
      TabIndex        =   12
      Top             =   1400
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Informacion"
      Height          =   315
      Index           =   8
      Left            =   180
      TabIndex        =   11
      Top             =   1000
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "N.Enemigos"
      Height          =   315
      Index           =   7
      Left            =   2460
      TabIndex        =   10
      Top             =   1400
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Donde"
      Height          =   315
      Index           =   6
      Left            =   1320
      TabIndex        =   9
      Top             =   2200
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Hora"
      Height          =   315
      Index           =   5
      Left            =   2460
      TabIndex        =   8
      Top             =   2200
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Comentar"
      Height          =   315
      Index           =   4
      Left            =   1320
      TabIndex        =   7
      Top             =   2600
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Ir hacia el"
      Height          =   315
      Index           =   3
      Left            =   1320
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Sumonear"
      Height          =   315
      Index           =   2
      Left            =   1320
      TabIndex        =   5
      Top             =   1000
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Banear"
      Height          =   315
      Index           =   1
      Left            =   3600
      TabIndex        =   4
      Top             =   2200
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Echar"
      Height          =   315
      Index           =   0
      Left            =   1320
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdActualiza 
      Caption         =   "Actualiza"
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   3120
      Width           =   3250
   End
   Begin VB.ComboBox cboListaUsus 
      Height          =   315
      ItemData        =   "frmPanelGm.frx":0000
      Left            =   180
      List            =   "frmPanelGm.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3290
   End
   Begin VB.Line Line9 
      X1              =   3480
      X2              =   3480
      Y1              =   2640
      Y2              =   1200
   End
   Begin VB.Line Line8 
      X1              =   3480
      X2              =   3480
      Y1              =   1320
      Y2              =   550
   End
   Begin VB.Line Line7 
      X1              =   4580
      X2              =   3480
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line6 
      X1              =   4600
      X2              =   4600
      Y1              =   2620
      Y2              =   1310
   End
   Begin VB.Line Line5 
      X1              =   3500
      X2              =   4600
      Y1              =   2615
      Y2              =   2615
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   120
      Y1              =   2980
      Y2              =   550
   End
   Begin VB.Line Line3 
      X1              =   3480
      X2              =   120
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line2 
      X1              =   3480
      X2              =   3480
      Y1              =   3000
      Y2              =   2600
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3500
      Y1              =   550
      Y2              =   550
   End
End
Attribute VB_Name = "frmPanelGm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[Wizard 03/09/05] Volvi True la opcion Sorted de la Lista para que los nicks se ordenen alfaveticamente.
Option Explicit
Private Sub cmdAccion_Click(Index As Integer)
Dim Ok As Boolean, Tmp As String, Tmp2 As String
Dim Nick As String

Nick = cboListaUsus.Text

Select Case Index
Case 0 '/ECHAR nick
    Call SendData("/ECHAR " & Nick)
Case 1 '/ban motivo@nick
    Tmp = InputBox("Motivo ?", "")
    If MsgBox("Esta seguro que desea banear al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
        Call SendData("/BAN " & Tmp & "@" & Nick)
    End If
Case 2 '/sum nick
    Call SendData("/SUM " & Nick)
Case 3 '/ira nick
    Call SendData("/IRA " & Nick)
Case 4 '/rem
    Tmp = InputBox("Comentario ?", "")
    Call SendData("/REM " & Tmp)
Case 5 '/hora
    Call SendData("/HORA")
Case 6 '/donde nick
    Call SendData("/DONDE " & Nick)
Case 7 '/nene
    Tmp = InputBox("Mapa ?", "")
    Call SendData("/NENE " & Trim(Tmp))
Case 8 '/info nick
    Call SendData("/INFO " & Nick)
Case 9 '/inv nick
    Call SendData("/INV " & cboListaUsus.Text)
Case 10 '/skills nick
    Call SendData("/SKILLS " & Nick)
Case 11 '/carcel minutos nick
    Tmp = InputBox("Minutos ? (hasta 30)", "")
    Tmp2 = InputBox("Razon ?", "")
    If MsgBox("Esta seguro que desea encarcelar al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
        Call SendData("/CARCEL " & Nick & "@" & Tmp2 & "@" & Tmp)
    End If
Case 12 '/unban nick
    If MsgBox("Esta seguro que desea removerle el ban al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
        Call SendData("/UNBAN " & Nick)
    End If
Case 13 '/nick2ip nick
    Call SendData("/NICK2IP " & Nick)
Case 14 '/ip2nick ip
    Call SendData("/IP2NICK " & Nick)
Case 15 '/penas
    Call SendData("/PENAS " & cboListaUsus.Text)
Case 16 'Ban X ip
    If MsgBox("Esta seguro que desea banear el (ip o personaje) " & Nick & "Por IP?", vbYesNo) = vbYes Then
    Call SendData("/BANIP " & Nick)
    End If
Case 17 ' MUESTA BOBEDA
    Call SendData("/BOV " & Nick)
Case 18 ' Sos
    Call SendData("/SHOW SOS")
Case 19 ' Balance
    Call SendData("/BAL " & cboListaUsus.Text)
Case 20 'Anticheat
    Call SendData("/ANTICHEAT " & cboListaUsus.Text)
End Select
End Sub

Private Sub cmdActualiza_Click()
Call SendData("LISTUSU")
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Show
Call cmdActualiza_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub
