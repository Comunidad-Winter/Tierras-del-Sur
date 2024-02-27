VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpciones 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   ClientHeight    =   4425
   ClientLeft      =   3780
   ClientTop       =   3240
   ClientWidth     =   5085
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
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   295
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   339
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox lstResoluciones 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmOpciones.frx":0152
      Left            =   3240
      List            =   "frmOpciones.frx":0154
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   2490
      Width           =   1575
   End
   Begin VB.CheckBox chkSonidoDrogas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   31
      Top             =   240
      Width           =   195
   End
   Begin VB.CheckBox chkVSync 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   26
      Top             =   1020
      Width           =   195
   End
   Begin VB.CheckBox checkFullScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3240
      TabIndex        =   18
      Top             =   1920
      Width           =   195
   End
   Begin VB.OptionButton BMP 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3480
      TabIndex        =   15
      Top             =   1080
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.OptionButton JPG 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3480
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.CheckBox Rpassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3240
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ComboBox lstlenguajes 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CheckBox musi 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   960
      Width           =   195
   End
   Begin VB.CheckBox cursoresnuevos 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3240
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox efectos 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   600
      Width           =   195
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   1695
      Left            =   3960
      TabIndex        =   20
      Top             =   1560
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   2990
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   10
      SmallChange     =   10
      Max             =   100
      TickStyle       =   2
      TickFrequency   =   10
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   1695
      Left            =   3240
      TabIndex        =   21
      Top             =   1560
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   2990
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   10
      SmallChange     =   10
      Max             =   100
      SelStart        =   1
      TickStyle       =   2
      TickFrequency   =   10
      Value           =   1
      TextPosition    =   1
   End
   Begin VB.Label lblResolucion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resolución"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3240
      TabIndex        =   33
      Top             =   2250
      Visible         =   0   'False
      Width           =   1035
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSonidoDopa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Efecto drogas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3480
      TabIndex        =   30
      Top             =   270
      Visible         =   0   'False
      Width           =   1755
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblConfigurarTeclas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Configurar teclas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   3300
      Visible         =   0   'False
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblNecesitaReiniciar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para que estos cambios se apliquen deberás re-iniciar el juego"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   2760
      TabIndex        =   28
      Top             =   2880
      Visible         =   0   'False
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblVSync 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sinc Vertical (Vsync). Recomendado para PCs modernas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   3480
      TabIndex        =   27
      Top             =   720
      Visible         =   0   'False
      Width           =   1470
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblVolumen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Volumen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   25
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblLenguage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lenguaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3240
      TabIndex        =   24
      Top             =   2520
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblEfectos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Efectos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3240
      TabIndex        =   23
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblMusica 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Musica"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3960
      TabIndex        =   22
      Top             =   1320
      Width           =   585
   End
   Begin VB.Label cNores 
      BackStyle       =   0  'Transparent
      Caption         =   "Pantalla Completa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      TabIndex        =   19
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label cBMP 
      BackStyle       =   0  'Transparent
      Caption         =   ".BMP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cJPG 
      BackStyle       =   0  'Transparent
      Caption         =   ".JPG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   16
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cCapture 
      BackStyle       =   0  'Transparent
      Caption         =   "Capturar imagen:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3240
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblRecordarClave 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Recordar Clave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3420
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label OAceptar 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label OAudio 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   360
      TabIndex        =   8
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Ovideo 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Ogeneral 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Efectos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   630
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Musica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   1005
      Width           =   975
   End
   Begin VB.Label lblCursoresTDS 
      BackStyle       =   0  'Transparent
      Caption         =   "Cursores TDS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub efectos_Click()
If Me.efectos.value = 1 Then
    EfectosSonidoActivados = True
Else
    EfectosSonidoActivados = False
End If
End Sub

Private Sub Form_Load()
If Connected = False Then frmConnect.Enabled = False

CargarOpcionesC
Call CambiarCursor(frmOpciones)
DameImagenForm Me, 132
End Sub

Private Sub Label16_Click()
'cpmensaje.Caption = "Estamos buscando la mejor configuración para su pc. Aguarde... "
'Frame2.Visible = True
'FistCheckUp
'Frame2.Visible = False
End Sub

Private Sub memvideo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

MsgBox "Es necesario que reinicie el cliente para que esta opción haga efecto."

End Sub

Private Sub lblConfigurarTeclas_Click()
    Call MostrarFormulario(frmConfigurarTeclas, Me)
End Sub

Private Sub musi_Click()
If Me.musi.value = 1 Then
    Call CLI_Audio.activarMusica
 Else
    Call CLI_Audio.desactivarMusica
End If
End Sub


Private Sub OAceptar_Click()

Call GuardarOpcionesC
Call guardarConfiguracion

'Me.Frame2.Visible = False

frmMain.SoundFX.Enabled = EfectosSonidoActivados

If Connected = False Then frmConnect.Enabled = True
Unload Me
Sonido_Play (SND_CLICK)
End Sub

Private Sub OAceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.OAceptar.tag = "0" Then Call Sonido_Play(SND_OVER): Me.OAceptar.tag = ""
End Sub

Private Sub OAudio_Click()
MostrarAudio
OcultarVideo
OcultarGeneral
Call Sonido_Play(SND_CLICK)
End Sub

Private Sub Ogeneral_Click()

OcultarAudio
OcultarVideo
MostrarGeneral

Call Sonido_Play(SND_CLICK)
End Sub

Private Sub Ovideo_Click()

OcultarAudio
OcultarGeneral
MostrarVideo

Call Sonido_Play(SND_CLICK)
End Sub

Private Sub Slider1_Change()
    volumenMusica = (100 - Slider1) / 100
    Me.Slider1.text = volumenMusica * 100 & "%"

    Call actualizarVolumen(volumenMusica)
End Sub


Public Sub CargarOpcionesC()
'Mostramos por defecto las del audio
'''
MostrarAudio
OcultarVideo
OcultarGeneral
'''''
'Sonido

Me.chkSonidoDrogas.value = IIf(SonidoFinalizacionDopa, 1, 0)

Me.musi = IIf(Musica, 1, 0)
Me.efectos = IIf(EfectosSonidoActivados, 1, 0)

Me.Slider1 = 100 - volumenMusica * 100
Me.Slider2 = 100 - VolumenF * 100

Me.checkFullScreen = IIf(forzarFullScreen, 1, 0)
Me.chkVSync = IIf(UsarVSync, 1, 0)

'Varios
Me.cursoresnuevos = IIf(CursorPer, 1, 0)

Me.Rpassword.value = Recpassword

If oJPG = 1 Then
Me.JPG = True
Else
Me.BMP = True
End If

With lstlenguajes
.AddItem "es"
End With

Dim i As Integer
For i = 0 To lstlenguajes.ListCount
    If lenguaje = lstlenguajes.list(i) Then
    lstlenguajes.ListIndex = i
    End If
Next


Call Me.lstResoluciones.AddItem("4:3 (1024x768", 0)
Me.lstResoluciones.itemData(0) = RESOLUCION_43

Call Me.lstResoluciones.AddItem("16:9 (1280x720", 1)
Me.lstResoluciones.itemData(1) = RESOLUCION_169

Me.lstResoluciones.ListIndex = ResolucionJuego - 1

End Sub

Public Sub GuardarOpcionesC()
Musica = IIf(Me.musi.value = 1, True, False)
EfectosSonidoActivados = IIf(Me.efectos.value = 1, True, False)

'Video
forzarFullScreen = IIf(Me.checkFullScreen.value = 1, True, False)

UsarVSync = IIf(Me.chkVSync.value = 1, True, False)

SonidoFinalizacionDopa = (Me.chkSonidoDrogas.value = 1)

If Not Me.lstResoluciones.ListIndex = -1 Then
    ResolucionJuego = Me.lstResoluciones.ListIndex + 1
End If

'Varios
Recpassword = Me.Rpassword

If Me.BMP.value = True Then
    oJPG = 0
Else
    oJPG = 1
End If

CursorPer = Me.cursoresnuevos

If CursorPer = 1 Then
    Call CambiarCursor(frmMain)
ElseIf Me.cursoresnuevos = 0 Then
    If Not Connected Then frmConnect.MousePointer = 1
    frmMain.MousePointer = 1
End If

If lstlenguajes <> lenguaje Then
    If CLI_Lenguajes.CargarLenguaje(lstlenguajes) Then
        lenguaje = lstlenguajes
    Else
        MsgBox "El lenguaje solicitado no se encuentra disponible.", vbCritical, "Tierras del Sur"
    End If
End If

End Sub

Private Sub OcultarGeneral()
Me.lblCursoresTDS.Visible = False
Me.cursoresnuevos.Visible = False
Me.lstlenguajes.Visible = False
Me.lblLenguage.Visible = False
Me.Rpassword.Visible = False
Me.JPG.Visible = False
Me.BMP.Visible = False
Me.cBMP.Visible = False
Me.cJPG.Visible = False
Me.cCapture.Visible = False
Me.lblRecordarClave.Visible = False
Me.lblConfigurarTeclas.Visible = False
End Sub

Private Sub MostrarGeneral()
Me.lblCursoresTDS.Visible = True
Me.cursoresnuevos.Visible = True
Me.lstlenguajes.Visible = True
Me.lblLenguage.Visible = True
Me.Rpassword.Visible = True
Me.JPG.Visible = True
Me.BMP.Visible = True
Me.cBMP.Visible = True
Me.cJPG.Visible = True
Me.cCapture.Visible = True
Me.lblRecordarClave.Visible = True
Me.lblConfigurarTeclas.Visible = True
End Sub

Private Sub OcultarAudio()
Me.Label1.Visible = False
Me.Label2.Visible = False
Me.efectos.Visible = False
Me.musi.Visible = False
Me.Slider1.Visible = False
Me.Slider2.Visible = False
Me.lblMusica.Visible = False
Me.lblEfectos.Visible = False
Me.lblVolumen.Visible = False
Me.lblSonidoDopa.Visible = False
Me.chkSonidoDrogas.Visible = False
End Sub

Private Sub MostrarAudio()
Me.Label1.Visible = True
Me.Label2.Visible = True
Me.efectos.Visible = True
Me.musi.Visible = True
Me.Slider1.Visible = True
Me.Slider2.Visible = True
Me.lblMusica.Visible = True
Me.lblEfectos.Visible = True
Me.lblVolumen.Visible = True
Me.lblSonidoDopa.Visible = True
Me.chkSonidoDrogas.Visible = True
End Sub

Private Sub OcultarVideo()
Me.cNores.Visible = False
Me.checkFullScreen.Visible = False
Me.chkVSync.Visible = False
Me.lblVSync.Visible = False
Me.lblNecesitaReiniciar.Visible = False
Me.lstResoluciones.Visible = False
Me.lblResolucion.Visible = False
End Sub

Private Sub MostrarVideo()
Me.cNores.Visible = True
Me.checkFullScreen.Visible = True
Me.chkVSync.Visible = True
Me.lblVSync.Visible = True
Me.lblNecesitaReiniciar.Visible = True
Me.lstResoluciones.Visible = True
Me.lblResolucion.Visible = True
End Sub

Private Sub Slider2_Scroll()
    VolumenF = (100 - Slider2) / 100
    Me.Slider2.text = VolumenF * 100 & "%"
    Call Sonido_Play(6)
End Sub
