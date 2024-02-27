VERSION 5.00
Begin VB.Form frmGuildFoundation 
   BorderStyle     =   0  'None
   Caption         =   "Creación de un Clan"
   ClientHeight    =   5160
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4155
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   344
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   277
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtClanName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      MaxLength       =   15
      TabIndex        =   5
      Top             =   2040
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      MaxLength       =   25
      TabIndex        =   4
      Top             =   3120
      Width           =   3495
   End
   Begin VB.CheckBox ChkReal 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   360
      TabIndex        =   3
      Top             =   4140
      Width           =   195
   End
   Begin VB.CheckBox ChkNeutro 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   1560
      TabIndex        =   2
      Top             =   4140
      Width           =   195
   End
   Begin VB.CheckBox ChkCaos 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2880
      TabIndex        =   1
      Top             =   4140
      Width           =   195
   End
   Begin VB.Label lblArmadaReal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ejército ïndigo"
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   630
      TabIndex        =   8
      Top             =   4080
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblNeutro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rebelde"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1830
      TabIndex        =   7
      Top             =   4155
      Width           =   690
   End
   Begin VB.Label lblLegionOscura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ejército Escarlata"
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   3120
      TabIndex        =   6
      Top             =   4080
      Width           =   825
      WordWrap        =   -1  'True
   End
   Begin VB.Image Boton 
      Height          =   465
      Index           =   1
      Left            =   150
      Top             =   4635
      Width           =   1140
   End
   Begin VB.Image Boton 
      Height          =   465
      Index           =   0
      Left            =   2625
      Top             =   4620
      Width           =   1305
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGuildFoundation.frx":0000
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
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
End
Attribute VB_Name = "frmGuildFoundation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Boton_Click(Index As Integer)
Select Case Index
Case 0
        txtClanName.text = Trim(txtClanName.text)
        
        If Len(txtClanName.text) = 0 Then Exit Sub
        
        'Chequeo el nombre del clan
        If Len(txtClanName.text) <= 25 Then
            If Not AsciiValidos(txtClanName) Then
                MsgBox "Nombre invalido."
                Exit Sub
            End If
        Else
            MsgBox "Nombre demasiado extenso."
            Exit Sub
        End If
        
        'Chequeo la direccion del sitio web
        If Len(Text2.text) <= 25 Then
            If Not AsciiValidos(txtClanName) Then
                MsgBox "Direccion de web invalida."
                Exit Sub
            End If
        Else
            MsgBox "Direccion de clan demasiado extenso."
            Exit Sub
        End If
                
        ClanName = txtClanName
        Site = Text2
        
        If CAlineacion = 0 Then MsgBox "Debes seleccionar una alineacion para el clan": Exit Sub
        
        Unload Me
        frmGuildDetails.Show , Me
Case 1
        Unload Me
End Select
End Sub

Private Sub ChkCaos_Click()
If ChkNeutro.value = 1 Then ChkNeutro.value = 0
If ChkReal.value = 1 Then ChkReal.value = 0
CAlineacion = 3
End Sub

Private Sub ChkNeutro_Click()
If ChkCaos.value = 1 Then ChkCaos.value = 0
If ChkReal.value = 1 Then ChkReal.value = 0
CAlineacion = 1
End Sub

Private Sub ChkReal_Click()
If ChkCaos.value = 1 Then ChkCaos.value = 0
If ChkNeutro.value = 1 Then ChkNeutro.value = 0
CAlineacion = 2
End Sub
Private Sub Form_Deactivate()
Me.SetFocus
End Sub

Private Sub Form_Load()
Call CambiarCursor(frmGuildFoundation)
DameImagenForm Me, 99
End Sub

Private Sub lblArmadaReal_Click()
    Me.ChkReal.value = IIf(Me.ChkReal.value = 1, 0, 1)
End Sub

Private Sub lblLegionOscura_Click()
    Me.ChkCaos.value = IIf(Me.ChkCaos.value = 1, 0, 1)
End Sub

Private Sub lblNeutro_Click()
    Me.ChkNeutro.value = IIf(Me.ChkNeutro.value = 1, 0, 1)
End Sub
