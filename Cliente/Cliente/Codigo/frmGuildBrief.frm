VERSION 5.00
Begin VB.Form frmGuildBrief 
   BorderStyle     =   0  'None
   Caption         =   "Detalles del Clan"
   ClientHeight    =   7755
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7605
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   517
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   507
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Desc 
      Appearance      =   0  'Flat
      Height          =   900
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   6240
      Width           =   6975
   End
   Begin VB.Image Backup 
      Height          =   15
      Left            =   10500
      Top             =   0
      Width           =   15
   End
   Begin VB.Image Boton 
      Height          =   360
      Index           =   4
      Left            =   5775
      Top             =   7320
      Width           =   1740
   End
   Begin VB.Image Boton 
      Height          =   360
      Index           =   3
      Left            =   4065
      Top             =   7320
      Width           =   1665
   End
   Begin VB.Image Boton 
      Height          =   390
      Index           =   2
      Left            =   2400
      Top             =   7305
      Width           =   1635
   End
   Begin VB.Image Boton 
      Height          =   375
      Index           =   1
      Left            =   1080
      Top             =   7305
      Width           =   1335
   End
   Begin VB.Image Boton 
      Height          =   390
      Index           =   0
      Left            =   90
      Top             =   7290
      Width           =   990
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   20
      Top             =   4320
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   19
      Top             =   3840
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   18
      Top             =   4080
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   17
      Top             =   4560
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   16
      Top             =   4800
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   15
      Top             =   5040
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   14
      Top             =   5280
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   13
      Top             =   5520
      Width           =   6735
   End
   Begin VB.Label nombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   810
      Width           =   6975
   End
   Begin VB.Label fundador 
      BackStyle       =   0  'Transparent
      Caption         =   "Fundador:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1280
      Width           =   6975
   End
   Begin VB.Label creacion 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de creacion:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1040
      Width           =   6975
   End
   Begin VB.Label lider 
      BackStyle       =   0  'Transparent
      Caption         =   "Lider:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1520
      Width           =   6975
   End
   Begin VB.Label web 
      BackStyle       =   0  'Transparent
      Caption         =   "Web site:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1760
      Width           =   6975
   End
   Begin VB.Label Miembros 
      BackStyle       =   0  'Transparent
      Caption         =   "Miembros:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2000
      Width           =   6975
   End
   Begin VB.Label eleccion 
      BackStyle       =   0  'Transparent
      Caption         =   "Dias para proxima eleccion de lider:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2240
      Width           =   6975
   End
   Begin VB.Label Oro 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2480
      Width           =   6975
   End
   Begin VB.Label Enemigos 
      BackStyle       =   0  'Transparent
      Caption         =   "Clanes Enemigos:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2710
      Width           =   6975
   End
   Begin VB.Label Aliados 
      BackStyle       =   0  'Transparent
      Caption         =   "Clanes Aliados:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2940
      Width           =   6975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineacion:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   3130
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   3120
      Width           =   2775
   End
End
Attribute VB_Name = "frmGuildBrief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public EsLeader As Boolean
Public Selecionado As Byte

Public Sub ParseGuildInfo(ByVal buffer As String)


If Not EsLeader Then
    Boton(1).Visible = False
    Boton(2).Visible = False
    Boton(3).Visible = False
Else
    Boton(1).Visible = True
    Boton(2).Visible = True
    Boton(3).Visible = True
End If

Debug.Print buffer


Dim informacion() As String

informacion = Split(buffer, "¬")


Nombre.Caption = "Nombre: " & informacion(0)
fundador.Caption = "Fundador: " & informacion(1)
creacion.Caption = "Fecha de creacion: " & informacion(2)
lider.Caption = "Lider: " & informacion(3)
web.Caption = "Web site: " & informacion(4)
Miembros.Caption = "Miembros: " & informacion(5)

If informacion(6) > 0 Then
    eleccion.Caption = "Dias para proxima eleccion de lider: " & informacion(6)
Else
    eleccion.Caption = "Elecciones de lider en curso."
End If

oro.Caption = "Oro: 0"
Enemigos.Caption = "Clanes enemigos: Ninguno"
aliados.Caption = "Clanes aliados: Ninguno"
'Copie lo de wizard que tanto
Select Case informacion(7)
    Case 1 'Neutro:)
            Label2.Caption = "Neutro"
            Label2.ForeColor = &H80000012
    Case 2 'Real
            Label2.Caption = "Armada Real"
            Label2.ForeColor = &H8000000D
    Case 3 'Caos
            Label2.Caption = "Fuerzas del Caos"
            Label2.ForeColor = &HFF&
    Case Else 'Comete un biscocho
        Call LogError("ERROR EN ACCEPTMEMBER; ALINEACION")
        Exit Sub
End Select

Dim i As Byte
For i = 0 To 7
    Codex(i).Caption = informacion(9 + i)
Next i

Desc = informacion(8)

Call Me.Show(vbModeless, frmMain)

End Sub

Private Sub Boton_Click(Index As Integer)
Select Case Index
Case 0
Unload Me
Case 1
frmCommet.Nombre = right(Nombre.Caption, Len(Nombre.Caption) - 7)
Call frmCommet.Show(vbModeless, frmGuildBrief)
Case 2
EnviarPaquete Paquetes.DeclararAlly, right$(Nombre, Len(Nombre) - 7)
Unload Me
Case 3
EnviarPaquete Paquetes.DeclararWar, right$(Nombre.Caption, Len(Nombre.Caption) - 7)
Unload Me
Case 4
Call frmGuildSol.RecieveSolicitud(right$(Nombre, Len(Nombre) - 7))
Call frmGuildSol.Show(vbModeless, frmGuildBrief)
End Select
End Sub


Private Sub Boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Selecionado <> Index Then
    Boton(Selecionado).tag = "0"
    Boton(Selecionado).Picture = backup.Picture
    End If
    
    If Boton(Index).tag <> "1" Then
    Boton(Index).tag = "1"
    Selecionado = Index
    backup.Picture = Boton(Selecionado).Picture
    Call DameImagen(Boton(Index), Index + 67)
    End If
End Sub

Private Sub Form_Load()
Call CambiarCursor(frmGuildBrief)
Call DameImagen(Boton(0), 62) '
Call DameImagen(Boton(1), 63) '
Call DameImagen(Boton(2), 64) '
Call DameImagen(Boton(3), 65) '
Call DameImagen(Boton(4), 66) '
backup.Picture = Boton(0).Picture
DameImagenForm Me, 101
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Boton(Selecionado).tag = "1" Then
    Boton(Selecionado).tag = "0"
    Boton(Selecionado).Picture = backup.Picture
    End If
End Sub
