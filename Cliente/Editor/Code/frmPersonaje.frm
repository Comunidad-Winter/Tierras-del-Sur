VERSION 5.00
Begin VB.Form frmPersonaje 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurar Modo Caminata"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Zurdo"
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.ListBox s 
      Height          =   3180
      Left            =   4080
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
   Begin VB.ListBox w 
      Height          =   3180
      Left            =   2760
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.ListBox h 
      Height          =   3180
      Left            =   1440
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.ListBox b 
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Escudo:"
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Arma:"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Cabeza:"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Cuerpos 
      Caption         =   "Cuerpo:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub b_Click()
CharList(UserCharIndex).Body = BodyData(val(b.text))
End Sub

Private Sub Check1_Click()

End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    For i = 1 To NumWeaponAnims
        'If WeaponAnimData(i).grh Then
            w.AddItem i
        'End If
    Next i
    
    For i = 1 To NumShieldAnims
        'If ShieldAnimData(i).ShieldWalk(1).GrhIndex Then
            s.AddItem i
        'End If
    Next i
    
    For i = 1 To NumCuerpos
        'If BodyData(i).Walk(1).GrhIndex Then
            b.AddItem i
        'End If
    Next i
    
    For i = 1 To Numheads
        'If HeadData(i).Head(1).GrhIndex Then
            h.AddItem i
        'End If
    Next i
    
    For i = 1 To NumCascos
        'If CascoAnimData(i).Head(1).GrhIndex Then
            c.AddItem i
        'End If
    Next i
End Sub
