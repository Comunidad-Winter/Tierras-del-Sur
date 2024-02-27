VERSION 5.00
Begin VB.Form frmGuildLeader 
   BorderStyle     =   0  'None
   Caption         =   "Administración del Clan"
   ClientHeight    =   6675
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6000
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
   ScaleHeight     =   445
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox solicitudes 
      Appearance      =   0  'Flat
      Height          =   1005
      ItemData        =   "frmGuildLeader.frx":0000
      Left            =   240
      List            =   "frmGuildLeader.frx":0002
      TabIndex        =   3
      Top             =   4290
      Width           =   2580
   End
   Begin VB.TextBox txtguildnews 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   240
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2650
      Width           =   5460
   End
   Begin VB.ListBox guildslist 
      Appearance      =   0  'Flat
      Height          =   1200
      ItemData        =   "frmGuildLeader.frx":0004
      Left            =   240
      List            =   "frmGuildLeader.frx":0006
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.ListBox members 
      Appearance      =   0  'Flat
      Height          =   1200
      ItemData        =   "frmGuildLeader.frx":0008
      Left            =   3120
      List            =   "frmGuildLeader.frx":000A
      TabIndex        =   0
      Top             =   480
      Width           =   2550
   End
   Begin VB.Image Boton 
      Height          =   465
      Index           =   7
      Left            =   945
      Top             =   5535
      Width           =   1095
   End
   Begin VB.Image Boton 
      Height          =   465
      Index           =   6
      Left            =   930
      Top             =   1875
      Width           =   1095
   End
   Begin VB.Image Boton 
      Height          =   465
      Index           =   5
      Left            =   3930
      Top             =   1860
      Width           =   1095
   End
   Begin VB.Image Boton 
      Height          =   465
      Index           =   4
      Left            =   2385
      Top             =   3375
      Width           =   1095
   End
   Begin VB.Image Boton 
      Height          =   345
      Index           =   3
      Left            =   3045
      Top             =   4170
      Width           =   2655
   End
   Begin VB.Image Boton 
      Height          =   345
      Index           =   2
      Left            =   3060
      Top             =   4665
      Width           =   2655
   End
   Begin VB.Image Boton 
      Height          =   345
      Index           =   1
      Left            =   3060
      Top             =   5175
      Width           =   2655
   End
   Begin VB.Image Boton 
      Height          =   465
      Index           =   0
      Left            =   3915
      Top             =   5925
      Width           =   1095
   End
   Begin VB.Label Miembros 
      BackStyle       =   0  'Transparent
      Caption         =   "El clan cuenta con x miembros"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   6120
      Width           =   2535
   End
End
Attribute VB_Name = "frmGuildLeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Selecionado As Byte

Public Sub ParserInfoSolicitudes(datos As String)
Dim informacion() As String
Dim i As Integer

informacion = Split(datos, ",")

Call solicitudes.Clear
'Agrego las solicitudes
For i = LBound(informacion) To UBound(informacion) - 1
    Call solicitudes.AddItem(informacion(i))
Next i
End Sub

Public Sub ParserInfoMiembros(datos As String)
Dim informacion() As String
Dim i As Integer

informacion = Split(datos, ",")

Call members.Clear
'Agrego los miembros
For i = LBound(informacion) To UBound(informacion)
    Call members.AddItem(informacion(i))
Next i

Miembros.Caption = "El clan cuenta con " & UBound(informacion) + 1 & " miembros."

End Sub

Public Sub ParserInfoNovedades(datos As String)
'Novedades
Me.txtguildnews = datos
End Sub

Public Sub ParseLeaderInfo(ByVal Data As String)
Dim informacion() As String
Dim i As Integer

'Para que no se solape la información
If Me.Visible Then Exit Sub

informacion = Split(Data, ",")

Call Me.guildslist.Clear

'Agrego los clanes
For i = LBound(informacion) To UBound(informacion) - 1
   Call guildslist.AddItem(informacion(i))
Next i

Me.Show

End Sub

Private Sub Boton_Click(Index As Integer)
Select Case Index
Case 0
    Unload Me
Case 1
    'EnviarPaquete Paquetes.PeaceProp
Case 2
    Call frmGuildURL.Show(vbModeless, frmGuildLeader)
Case 3
    Call MostrarFormulario(frmGuildDetails, frmGuildLeader)
Case 4
    EnviarPaquete Paquetes.ActualizarGNews, Replace(txtguildnews, vbCrLf, " ")
Case 5
    If members.ListIndex <> -1 Then
        frmCharInfo.frmmiembros = True
        frmCharInfo.frmsolicitudes = False
        EnviarPaquete Paquetes.MemberInfo, members.list(members.ListIndex)
    End If
Case 6
    frmGuildBrief.EsLeader = True
    EnviarPaquete Paquetes.GuildDetail, guildslist.list(guildslist.ListIndex)
Case 7
    If solicitudes.ListIndex >= 0 Then
        frmCharInfo.frmsolicitudes = True
        frmCharInfo.frmmiembros = False
        EnviarPaquete Paquetes.MemberInfo, solicitudes.list(solicitudes.ListIndex)
    End If
End Select
End Sub

Private Sub Boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Selecionado >= 0 Then
    Boton(Selecionado).tag = "0"
    Boton(Selecionado).Picture = Nothing
End If

If Boton(Index).tag <> "1" Then
    Boton(Index).tag = "1"
    Selecionado = Index
    Call DameImagen(Boton(Index), Index + 11)
End If
End Sub

Private Sub Form_Load()
Call CambiarCursor(frmGuildLeader)
DameImagenForm Me, 95

Me.solicitudes.Clear
Me.solicitudes.AddItem "Cargando...."

Me.txtguildnews.text = "Cargando..."

Me.members.Clear
Me.members.AddItem "Cargando..."

EnviarPaquete Paquetes.obtClanMiembros
EnviarPaquete Paquetes.obtClanSolicitudes
EnviarPaquete Paquetes.obtClanNews
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Selecionado >= 0 Then
    Boton(Selecionado).tag = "0"
    Boton(Selecionado).Picture = Nothing
End If
End Sub

