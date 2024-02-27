VERSION 5.00
Begin VB.Form frmGuildDetails 
   BorderStyle     =   0  'None
   Caption         =   "Detalles del Clan"
   ClientHeight    =   7050
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6945
   ClipControls    =   0   'False
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
   ScaleHeight     =   470
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   463
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   600
      MaxLength       =   50
      TabIndex        =   8
      Top             =   3480
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   600
      MaxLength       =   50
      TabIndex        =   7
      Top             =   3840
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   600
      MaxLength       =   50
      TabIndex        =   6
      Top             =   4200
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   600
      MaxLength       =   50
      TabIndex        =   5
      Top             =   4560
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   600
      MaxLength       =   50
      TabIndex        =   4
      Top             =   4920
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   600
      MaxLength       =   50
      TabIndex        =   3
      Top             =   5280
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   600
      MaxLength       =   50
      TabIndex        =   2
      Top             =   5640
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   7
      Left            =   600
      MaxLength       =   50
      TabIndex        =   1
      Top             =   6000
      Width           =   5655
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      Height          =   1245
      Left            =   360
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   790
      Width           =   6135
   End
   Begin VB.Image Boton 
      Height          =   375
      Index           =   1
      Left            =   5655
      Top             =   6525
      Width           =   975
   End
   Begin VB.Image Boton 
      Height          =   375
      Index           =   0
      Left            =   210
      Top             =   6510
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGuildDetails.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   6255
   End
End
Attribute VB_Name = "frmGuildDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Selecionado As Byte
Private Sub Boton_Click(Index As Integer)
Dim k, cantidadCodecs As Integer
Dim infoClan As String
Select Case Index
Case 1

    Dim descripcion As String
    
    descripcion = Replace(txtDesc, vbCrLf, " ", , , vbBinaryCompare)
    descripcion = Replace(txtDesc, "¬", " ", , , vbBinaryCompare)
    
 
    
    If LenB(descripcion) > 250 Then
            MsgBox "La descripción no puede tener más de 250 caracteres."
            Exit Sub
    End If
    
    
   cantidadCodecs = 0
    
    For k = 0 To txtCodex1.UBound
        If Len(Trim(txtCodex1(k).text)) > 0 Then cantidadCodecs = cantidadCodecs + 1
    Next k
    
    If cantidadCodecs < 4 Then
            MsgBox "Debes definir al menos cuatro mandamientos."
            Exit Sub
    End If
    
    Dim paquete As Byte
    
    If CreandoClan Then
        'Enviar la información para crear el clan
        paquete = Paquetes.GuildDSend
        infoClan = ClanName & "¬" & descripcion & "¬" & Site & "¬" & cantidadCodecs & "¬" & CAlineacion
    Else
        'Actualiza la descripción y los codecs
        paquete = Paquetes.GuildCode
        infoClan = descripcion & "¬" & cantidadCodecs
    End If
    
    For k = 0 To txtCodex1.UBound
        infoClan = infoClan & "¬" & txtCodex1(k)
    Next k
    
    EnviarPaquete paquete, infoClan
    CreandoClan = False
End Select
Unload Me
End Sub

Private Sub Boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Selecionado <> Index Then
Boton(Selecionado).tag = "0"
Boton(Selecionado).Picture = Nothing
End If
If Boton(Index).tag <> "1" Then
Boton(Index).tag = "1"
Selecionado = Index
Call DameImagen(Boton(Index), Index + 21)
End If
End Sub

Private Sub Form_Load()
Call CambiarCursor(frmGuildDetails)
DameImagenForm Me, 96
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
