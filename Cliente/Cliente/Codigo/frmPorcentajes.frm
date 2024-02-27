VERSION 5.00
Begin VB.Form frmPartyPorc 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   Caption         =   "Acomodar Porcentajes"
   ClientHeight    =   2985
   ClientLeft      =   4305
   ClientTop       =   3105
   ClientWidth     =   3270
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   199
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   218
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Porc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "100"
      Top             =   2010
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Porc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "100"
      Top             =   1650
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Porc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "100"
      Top             =   1290
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Porc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "100"
      Top             =   930
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Porc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "0"
      Top             =   570
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Boton 
      Height          =   375
      Index           =   1
      Left            =   150
      Top             =   2490
      Width           =   975
   End
   Begin VB.Image Boton 
      Height          =   375
      Index           =   0
      Left            =   2160
      Top             =   2505
      Width           =   975
   End
   Begin VB.Line Lin 
      BorderColor     =   &H00E0E0E0&
      Index           =   5
      Visible         =   0   'False
      X1              =   8
      X2              =   208
      Y1              =   152
      Y2              =   152
   End
   Begin VB.Line Lin 
      BorderColor     =   &H00E0E0E0&
      Index           =   4
      Visible         =   0   'False
      X1              =   8
      X2              =   208
      Y1              =   128
      Y2              =   128
   End
   Begin VB.Line Lin 
      BorderColor     =   &H00E0E0E0&
      Index           =   3
      Visible         =   0   'False
      X1              =   8
      X2              =   208
      Y1              =   104
      Y2              =   104
   End
   Begin VB.Line Lin 
      BorderColor     =   &H00E0E0E0&
      Index           =   2
      Visible         =   0   'False
      X1              =   8
      X2              =   208
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Line Lin 
      BorderColor     =   &H00E0E0E0&
      Index           =   1
      Visible         =   0   'False
      X1              =   8
      X2              =   208
      Y1              =   56
      Y2              =   56
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   2880
      TabIndex        =   11
      Top             =   240
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Pj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pj1"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Pj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pj1"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Pj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pj1"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Pj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pj1"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Pj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pj1"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "frmPartyPorc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SkillsL As Integer
Public Selecionado As Byte

Private Sub Boton_Click(Index As Integer)
Select Case Index
Case 1
Dim Suma As Integer
Dim MinP As Integer
Dim MaxP As Integer
Dim cadena As String
If SkillsL >= 60 And SkillsL < 75 Then
    MinP = 30
    MaxP = 70
ElseIf SkillsL >= 75 And SkillsL < 90 Then
    MinP = 20
    MaxP = 80
ElseIf SkillsL >= 90 Then
    MinP = 10
    MaxP = 90
Else
    MsgBox "Necesitas al menos 60 skills en liderazgo para acomodar los porcentajes.", vbInformation
    Unload Me
    Exit Sub
End If
For i = 1 To 5
    Suma = Suma + val(Porc(i).text)
    If Porc(i).Visible Then
        If Porc(2).Visible = True Then
        If val(Porc(i)) < MinP Then
            MsgBox "El porcentaje asignado al personaje " & i & " debe ser de al menos un " & MinP & "%"
            Exit Sub
        End If
        If val(Porc(i)) > MaxP Then
            MsgBox "El porcentaje asignado al personaje " & i & " no debe superar el " & MaxP & "%"
            Exit Sub
        End If
        cadena = cadena & val(Porc(i).text) & "|"
        Else
        If val(Porc(i)) <> 100 Then
            MsgBox "Al ser solamente uno el integrante de la party, le corresponde el 100% de la experiencia ganada"
            Exit Sub
        End If
        cadena = cadena & val(Porc(i).text) & "|"
        Exit For
        End If
    End If
Next i
If Suma <> 100 Then
    MsgBox "La suma de todos los porcentajes debe dar como valor 100.", vbCritical
    Exit Sub
End If
Unload Me
Call sSendData(Paquetes.Comandos, Complejo.AcomodarPorcentajesDeParty, cadena)
Case 0
Unload Me
End Select
End Sub

Private Sub Boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Selecionado <> Index Then
    Boton(Selecionado).tag = "0"
    Boton(Selecionado).Picture = Nothing
    End If
    
    If Boton(Index).tag <> "1" Then
    Boton(Index).tag = "1"
    Selecionado = Index
    Call DameImagen(Boton(Index), Index + 86)
    End If
End Sub

Private Sub Form_Load()
Call CambiarCursor(frmPartyPorc)
DameImagenForm Me, 111
Me.Porc(1) = Partym.Label8(0)
Me.Porc(2) = Partym.Label8(1)
Me.Porc(3) = Partym.Label8(2)
Me.Porc(4) = Partym.Label8(3)
Me.Porc(5) = Partym.Label8(4)

Me.Pj(1) = Partym.Label8(0)
Me.Pj(2) = Partym.Label8(1)
Me.Pj(3) = Partym.Label8(2)
Me.Pj(4) = Partym.Label8(3)
Me.Pj(5) = Partym.Label8(4)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Boton(Selecionado).tag = "1" Then
    Boton(Selecionado).tag = "0"
    Boton(Selecionado).Picture = Nothing
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
End Sub

Private Sub Porc_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If val(Porc(Index)) = 100 Then
Porc(Index).text = left(Porc(Index).text, 3)
ElseIf val(Porc(Index)) > 100 Then
Porc(Index).text = left(Porc(Index).text, 2)
End If
End Sub
