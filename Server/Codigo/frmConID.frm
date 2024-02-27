VERSION 5.00
Begin VB.Form frmConID 
   Caption         =   "ConID"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Liberar todos los slots"
      Height          =   390
      Left            =   135
      TabIndex        =   3
      Top             =   3840
      Width           =   4290
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ver estado"
      Height          =   390
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   4290
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   180
      TabIndex        =   1
      Top             =   150
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Width           =   4290
   End
   Begin VB.Label Label1 
      Height          =   510
      Left            =   180
      TabIndex        =   4
      Top             =   2430
      Width           =   4230
   End
End
Attribute VB_Name = "frmConID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()

Dim c As Integer
Dim i As Integer

List1.Clear

For i = 1 To MaxUsers
    If UserList(i).flags.Saliendo = SaliendoForsozamente Then
        Call List1.AddItem("UserIndex " & i & " -- " & UserList(i).ConnID & "-- CERRANDO")
    Else
        Call List1.AddItem("UserIndex " & i & " -- " & UserList(i).ConnID & "--" & HelperIP.longToIP(UserList(i).ip))
    End If
    
    If Not UserList(i).ConnID = INVALID_SOCKET Then c = c + 1
Next i

If c = MaxUsers Then
    Label1.Caption = "No hay conexiones libres!"
Else
    Label1.Caption = "Hay " & MaxUsers - c & " conexiones libres.!"
End If

Label1.Caption = Label1.Caption & ". UserIndex Libres: " & modPersonajes.obtenerSlotsLibres

End Sub

Private Sub Command3_Click()
    Call Admin.liberarTodosSlots
End Sub

