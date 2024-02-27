VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sistema de lenguajes"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Lenguajes listos"
      Height          =   3735
      Left            =   2760
      TabIndex        =   3
      Top             =   240
      Width           =   5055
      Begin VB.FileListBox File2 
         Appearance      =   0  'Flat
         Height          =   2565
         Left            =   120
         Pattern         =   "*.leg*"
         TabIndex        =   10
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2760
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ver mensaje"
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton Command5 
         Caption         =   "buscar en mensajes"
         Height          =   375
         Left            =   2760
         TabIndex        =   6
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cargar lenguaje"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Lenguajes disponibles :"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Generar lenguaje"
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "Generar lenguaje"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   3240
         Width           =   2055
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         Height          =   2565
         Left            =   240
         Pattern         =   "*.txt*"
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Lenguajes disponibles :"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long


Private Sub Command1_Click()
Dim RutaOrigen As String
Dim RutaDestino As String
If Me.File1 = "" Then Exit Sub
RutaOrigen = App.Path & "\Lenguajes\" & Me.File1
RutaDestino = App.Path & "\Finales\" & Left(Me.File1, Len(Me.File1) - 4) & ".leg"

Dim numero(0 To 50) As Single

i = 0
jose = 0
Open RutaOrigen For Input As #1
    While Not EOF(1)
        Line Input #1, mensajes
        i = i + 1
        If InStr(1, mensajes, "////") = 0 Then
            aponer = mensajes
            jose = jose + 1
        Else
            ii = ii + 1
            If jose - numero(ii - 1) > 0 Then
            tuerco = tuerco + numero(ii - 1)
            numero(ii) = jose - tuerco
            Else
            ii = ii - 1
            End If
        End If
    Wend
Close #1

For A = 1 To 12
    aa = aa & numero(A) & ";"
Next


'/////////////////////////////
i = 0
Open RutaDestino For Append As #2
Print #2, aa
Open RutaOrigen For Input As #1

While Not EOF(1)
    Line Input #1, mensajes
    i = i + 1
    If InStr(1, mensajes, "////") = 0 Then
        aponer = Split(mensajes, "^")
        jose = jose + 1
    
        If UBound(aponer) > -1 Then
            Print #2, aponer(0)
        Else
            Print #2, ""
        End If
    Else
        ii = ii + 1
        numero(ii) = jose
    End If
Wend
Close #1
Close #2
File2.Refresh
MsgBox "Lenguajes compilados exitosamente"
End Sub

Private Sub Command2_Click()
i = 0
nacho = GetTickCount

Open App.Path & "\Finales\" & Me.File2 For Input As #1
For i = 1 To 380
Line Input #1, mensajes
mensaje(i) = mensajes
Next
Close #1

MsgBox GetTickCount - nacho
End Sub

Private Sub Command3_Click()
MsgBox mensaje(Val(Me.Text1) + 1)
End Sub
Private Sub Command5_Click()
    For i = 1 To UBound(mensaje)
    If InStr(1, mensaje(i), Me.Text2) > 0 Then MsgBox i - 1 & "-" & mensaje(i)
Next
End Sub

Private Sub Form_Load()
CryptoInit
Me.File1.Path = App.Path & "\Lenguajes\"
Me.File2.Path = App.Path & "\Finales\"
End Sub

Private Sub List1_Click()

End Sub
