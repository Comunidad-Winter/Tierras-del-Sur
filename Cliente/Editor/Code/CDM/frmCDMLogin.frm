VERSION 5.00
Begin VB.Form frmCDMLogin 
   BackColor       =   &H00A8BDB7&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tierras del Sur - Iniciar Sesion"
   ClientHeight    =   1815
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5430
   Icon            =   "frmCDMLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   121
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   362
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRecordarClave 
      Appearance      =   0  'Flat
      BackColor       =   &H00A8BDB7&
      Caption         =   "Recordar clave"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1560
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Frame frmMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00A8BDB7&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3615
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1140
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   435
         Width           =   2445
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   380
         Left            =   1080
         TabIndex        =   4
         Top             =   1200
         Width           =   1170
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Ingresar"
         Default         =   -1  'True
         Height          =   390
         Left            =   2400
         TabIndex        =   3
         Top             =   1200
         Width           =   1170
      End
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1140
         TabIndex        =   1
         Top             =   0
         Width           =   2445
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   0
         TabIndex        =   6
         Top             =   525
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   330
         TabIndex        =   5
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   4200
      Picture         =   "frmCDMLogin.frx":1CCA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1155
   End
End
Attribute VB_Name = "frmCDMLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub cmdCancel_Click()

    #If Testeo = 1 Then
        Call CDM.cerebro.LoginDummy
    #End If
    
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    '¿Ingreso el user?
    If Len(txtUserName.text) = 0 Then
        MsgBox "Ingrese tu usuario.", vbInformation, Me.caption
        Exit Sub
    End If
      
    '¿Ingreso la clave? Sino la ingreso vemos si esta recordada
    If Len(txtPassword) = 0 Then
        Call setearPasswordDesdeGuardado
        If Len(Me.txtPassword.text) = 0 Then
            MsgBox "Ingrese tu clave por favor.", vbInformation, Me.caption
            Exit Sub
        End If
    End If
    
    Me.caption = "Espere por favor... Ingresando..."
        
    'Desactivamos los botones y los campos
    Call modPosicionarFormulario.setEnabledHijos(False, Me.frmMain, Me)
        
    Call GuardarPassword(txtUserName, IIf(Me.chkRecordarClave.value = 1, Me.txtPassword, "X"))
            
    'Iniciamos session
    If CDM.cerebro.Login(txtUserName.text, txtPassword.text) Then
        Me.Visible = False
    Else
        Me.caption = CDM.cerebro.ultimoError
        Call modPosicionarFormulario.setEnabledHijos(True, Me.frmMain, Me)
    End If

End Sub

Private Sub Form_Load()
    Call modPosicionarFormulario.setEnabledHijos(True, Me.frmMain, Me)
End Sub

Private Sub setearPasswordDesdeGuardado()
    Me.txtPassword.text = BuscarPassword(Me.txtUserName)
    
    If Not Me.txtPassword.text = "" Then Me.chkRecordarClave.value = 1
End Sub

Private Sub txtPassword_GotFocus()
   setearPasswordDesdeGuardado
End Sub
