VERSION 5.00
Begin VB.Form Retos 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Retos"
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2895
   Icon            =   "Retos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   193
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TDS_1.UpDownText txtCantidadRojas 
      Height          =   285
      Left            =   1800
      TabIndex        =   18
      Top             =   1320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      MaxValue        =   0
      MinValue        =   0
      Value           =   0
      Enabled         =   -1  'True
      Blanqueado      =   0   'False
   End
   Begin VB.CheckBox chkSinCascoEscudo 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   3810
      Width           =   180
   End
   Begin VB.CheckBox chkValeResu 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      Caption         =   "Vale resu"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   360
      TabIndex        =   17
      Top             =   1620
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.CheckBox ckhPlantado 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   3360
      Width           =   180
   End
   Begin VB.CheckBox chkLImitarRojas 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      Caption         =   "Pociones rojas:"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   360
      TabIndex        =   15
      Top             =   1365
      Width           =   180
   End
   Begin VB.TextBox txtEnemigo3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      MaxLength       =   30
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtCompañero2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      MaxLength       =   30
      TabIndex        =   6
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1860
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   14
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   210
   End
   Begin VB.TextBox oro 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      MaxLength       =   8
      TabIndex        =   1
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtEnemigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      MaxLength       =   30
      TabIndex        =   2
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtEnemigo2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      MaxLength       =   30
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtCompañero 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      MaxLength       =   30
      TabIndex        =   5
      Top             =   4080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox Check4 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      Caption         =   "Por los items"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   1125
      Width           =   180
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   300
      TabIndex        =   9
      Top             =   720
      Value           =   -1  'True
      Width           =   210
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   210
   End
   Begin VB.Label lblPorLosItems 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Por los items"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   27
      Top             =   1125
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   195
      Left            =   720
      TabIndex        =   26
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label lblLImitarRojas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Limitar rojas:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   25
      Top             =   1365
      Width           =   870
   End
   Begin VB.Label lblValeResu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vale resucitar"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   24
      Top             =   1620
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lbl3vs3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3 vs 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2100
      TabIndex        =   23
      Top             =   720
      Width           =   435
   End
   Begin VB.Label lbl2vs2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2 vs 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1320
      TabIndex        =   22
      Top             =   720
      Width           =   435
   End
   Begin VB.Label lbl1vs1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1 vs 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   525
      TabIndex        =   21
      Top             =   720
      Width           =   435
   End
   Begin VB.Label lblPlantados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Plantado"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   20
      Top             =   3360
      Width           =   630
   End
   Begin VB.Label lblCascosYEscudos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No permitir el uso de Cascos y Escudos"
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   600
      TabIndex        =   19
      Top             =   3720
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblForOtrasOpciones 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Más opciones..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   360
      TabIndex        =   16
      Top             =   3120
      Width           =   1350
   End
   Begin VB.Image retar 
      Height          =   465
      Left            =   210
      Top             =   4845
      Width           =   1080
   End
   Begin VB.Image Label5 
      Height          =   450
      Left            =   1500
      Top             =   4865
      Width           =   1125
   End
   Begin VB.Label lblForAliados 
      BackStyle       =   0  'Transparent
      Caption         =   "Personajes aliados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   3840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblForEnemigos 
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje enemigo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad de oro en juego:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1920
      Width           =   2295
   End
End
Attribute VB_Name = "Retos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub chkLImitarRojas_Click()

    Me.txtCantidadRojas.Enabled = chkLImitarRojas.value = vbChecked

End Sub

Private Sub Form_Load()
Call CambiarCursor(Retos)
DameImagenForm Me, 92

Me.txtCantidadRojas.MinValue = 0
Me.txtCantidadRojas.MaxValue = 10000
Me.txtCantidadRojas.value = 0

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label5.tag = "1" Then
Label5.tag = "0"
Label5.Picture = Nothing
ElseIf retar.tag = "1" Then
retar.Picture = Nothing
retar.tag = "0"
End If
End Sub
Private Sub Label5_Click()
Unload Me
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If retar.tag = "1" Then
retar.Picture = Nothing
retar.tag = "0"
End If
If Label5.tag <> "1" Then
Label5.tag = 1
Call DameImagen(Label5, 9)
End If
End Sub

Private Sub lbl1vs1_Click()
    Me.Option1(0).value = Not Me.Option1(0).value
End Sub

Private Sub lbl2vs2_Click()
    Me.Option1(1).value = Not Me.Option1(1).value
End Sub

Private Sub lbl3vs3_Click()
    Me.Option1(2).value = Not Me.Option1(2).value
End Sub

Private Sub lblCascosYEscudos_Click()
    Me.chkSinCascoEscudo.value = IIf(Me.chkSinCascoEscudo.value = 1, 0, 1)
End Sub

Private Sub lblLImitarRojas_Click()
    Me.chkLImitarRojas.value = IIf(Me.chkLImitarRojas.value = 1, 0, 1)
End Sub

Private Sub lblPlantados_Click()
     Me.ckhPlantado.value = IIf(Me.ckhPlantado.value = 1, 0, 1)
End Sub

Private Sub lblPorLosItems_Click()
    Me.Check4.value = IIf(Me.Check4.value = 1, 0, 1)
End Sub

Private Sub lblValeResu_Click()
     Me.chkValeResu.value = IIf(Me.chkValeResu.value = 1, 0, 1)
End Sub

Private Sub Option1_Click(Index As Integer)

    If Index = 0 Then ' 1 vs 1
    
        lblForOtrasOpciones.Visible = True
        
        ckhPlantado.Visible = True
        lblPlantados.Visible = True
        chkValeResu.Visible = False
        lblValeResu.Visible = False
        
        chkSinCascoEscudo.Visible = True
        chkSinCascoEscudo.Enabled = True
        lblCascosYEscudos.Visible = True
        
        
        lblForEnemigos.Caption = "Personaje Enemigo:"
        txtEnemigo2.Visible = False
        txtEnemigo3.Visible = False
        lblForAliados.Visible = False
        txtCompañero.Visible = False
        txtCompañero2.Visible = False


    ElseIf Index = 1 Then ' 2 vs 2
        
        chkSinCascoEscudo.Visible = False
        chkSinCascoEscudo.Enabled = False
        lblCascosYEscudos.Visible = False
        
        chkValeResu.Visible = True
        lblValeResu.Visible = True
        
        lblForEnemigos.Caption = "Personajes Enemigos"
        txtEnemigo2.Visible = True
        txtEnemigo2.top = 208
        txtEnemigo3.Visible = False
        lblForAliados.Visible = True
        lblForAliados.Caption = "Personajes Aliados"
        lblForAliados.top = 232
        txtCompañero.Visible = True
        txtCompañero.top = 248
        txtCompañero2.Visible = False
        
        lblForOtrasOpciones.Visible = False
        ckhPlantado.Visible = False
        lblPlantados.Visible = False
        
    ElseIf Index = 2 Then ' 3 vs 3
    
        chkSinCascoEscudo.Visible = False
        chkSinCascoEscudo.Enabled = False
        lblCascosYEscudos.Visible = False
        
        chkValeResu.Visible = True
        lblValeResu.Visible = True
        lblForEnemigos.Caption = "Personajes Enemigos"
        txtEnemigo2.Visible = True
        txtEnemigo2.top = 208
        
        txtEnemigo3.Visible = True
        txtEnemigo3.top = 232

        ' 3vs3 Aliados
        lblForAliados.Visible = True
        lblForAliados.Caption = "Personajes Aliados"
        lblForAliados.top = 256

        txtCompañero.Visible = True
        txtCompañero.top = 272
        
        txtCompañero2.Visible = True
        txtCompañero2.top = 296

        lblForOtrasOpciones.Visible = False
        ckhPlantado.Visible = False
        lblPlantados.Visible = False
End If

End Sub

Private Sub retar_Click()

Dim cantidadOro As Long
Dim modo As Byte
Dim integrantesPorEquipo As Byte
Dim equipos As String
Dim cantidadRojas As Integer
Dim limitarRojas As Boolean
Dim plantado As Boolean
Dim valeResu As Boolean
Dim valeCascoyEscudo As Boolean

plantado = False

If Me.oro = "" And Me.Check4 = 0 Then
    MsgBox "Por favor inserte la cantidad de oro"
    Exit Sub
End If

If Me.Check4.value = vbUnchecked Then
    modo = 1 'Solo oro
Else
    modo = 3 'Oro e items
End If

If Me.chkLImitarRojas.value = vbChecked Then
    cantidadRojas = Me.txtCantidadRojas.value
    limitarRojas = True
Else
    limitarRojas = False
    cantidadRojas = 0
End If

cantidadOro = CLng(val(Me.oro))

If cantidadOro > 50000000 Then
    MsgBox "No puedes apostar tanto oro."
    Exit Sub
ElseIf cantidadOro = 0 Then
    MsgBox "No puedes apostar 0 monedas de oro."
    Exit Sub
End If

Me.txtEnemigo = Trim$(Me.txtEnemigo)
Me.txtEnemigo2 = Trim$(Me.txtEnemigo2)
Me.txtEnemigo3 = Trim$(Me.txtEnemigo3)

Me.txtCompañero = Trim$(Me.txtCompañero)
Me.txtCompañero2 = Trim$(Me.txtCompañero2)

' ¿Qué modo eligio?
If Me.Option1(0).value Then ' 1vs 1

    If Len(Me.txtEnemigo.text) = 0 Then
        MsgBox ("Tenes que ingresar el nombre de tu enemigo.")
    Exit Sub
    End If
    
    If UCase$(Me.txtEnemigo.text) = UCase$(UserName) Then
        MsgBox ("No puedes retarte a ti mismo.")
    Exit Sub
End If

    plantado = (Me.ckhPlantado.value = vbChecked)
    valeCascoyEscudo = (Me.chkSinCascoEscudo.value = vbChecked)
    valeResu = False
    
    integrantesPorEquipo = 1
    equipos = UserName & "|" & Me.txtEnemigo.text
        
ElseIf Me.Option1(1).value Then

    If Len(Me.txtCompañero.text) = 0 Or UCase$(Me.txtCompañero.text) = UCase$(UserName) Then
        MsgBox ("Tenes que ingresar el nombre del personaje que te van a acompañar en el combate.")
        Exit Sub
    End If
    
    If Len(Me.txtEnemigo.text) = 0 Or Len(Me.txtEnemigo2.text) = 0 Then
        MsgBox ("Tenes que completar los nombres de los personajes del equipo contrario.")
        Exit Sub
    End If
    
    If UCase$(Me.txtEnemigo.text) = UCase$(UserName) Or UCase$(Me.txtEnemigo2.text) = UCase$(UserName) Then
        MsgBox ("No pudes formar partes del equipo enemigo.")
        Exit Sub
    End If
    
    integrantesPorEquipo = 2
    valeResu = (Me.chkValeResu.value = vbChecked)
    valeCascoyEscudo = False
    plantado = False
    
    equipos = UserName & "-" & Me.txtCompañero.text & "|" & Me.txtEnemigo.text & "-" & Me.txtEnemigo2.text
    
ElseIf Me.Option1(2).value Then

    If Len(Me.txtCompañero.text) = 0 Or UCase$(Me.txtCompañero.text) = UCase$(UserName) Or Len(Me.txtCompañero2.text) = 0 Or UCase$(Me.txtCompañero2.text) = UCase$(UserName) Then
        MsgBox ("Tenes que ingresar el nombre de los personajes que te van a acompañar en el combate.")
        Exit Sub
    End If
    
    If Len(Me.txtEnemigo.text) = 0 Or Len(Me.txtEnemigo2.text) = 0 Or Len(Me.txtEnemigo3.text) = 0 Then
        MsgBox ("Tenes que completar los nombres de los personajes del equipo contrario.")
        Exit Sub
    End If
    
    If UCase$(Me.txtEnemigo.text) = UCase$(UserName) Or UCase$(Me.txtEnemigo2.text) = UCase$(UserName) Or UCase$(Me.txtEnemigo3.text) = UCase$(UserName) Then
        MsgBox ("No pudes formar partes del equipo enemigo.")
        Exit Sub
End If

    valeResu = (Me.chkValeResu.value = vbChecked)
    integrantesPorEquipo = 3
    plantado = False
    valeCascoyEscudo = False
    
    equipos = UserName & "-" & Me.txtCompañero.text & "-" & Me.txtCompañero2.text & "|" & Me.txtEnemigo.text & "-" & Me.txtEnemigo2.text & "-" & Me.txtEnemigo3.text & "|"
Else
    MsgBox ("Tenes que elegir la modalidad: 1vs1, 2vs2 o 3vs3.")
    Exit Sub
End If


Call sSendData(Paquetes.CrearReto, 0, ByteToString(integrantesPorEquipo) & ByteToString(modo) & ByteToString(IIf(plantado, 1, 0)) & LongToString(oro) & ByteToString(IIf(valeResu, 1, 0)) & ByteToString(IIf(limitarRojas, 1, 0)) & ITS(cantidadRojas) & ByteToString(IIf(valeCascoyEscudo, 1, 0)) & "|" & equipos)

Unload Me
End Sub

Private Sub retar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label5.tag = "1" Then
Label5.tag = "0"
Label5.Picture = Nothing
End If
If retar.tag <> "1" Then
retar.tag = 1
Call DameImagen(retar, 10)
End If
End Sub

