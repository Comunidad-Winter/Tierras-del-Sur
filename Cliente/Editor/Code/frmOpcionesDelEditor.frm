VERSION 5.00
Begin VB.Form frmOpcionesDelEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tierras del Sur - Opciones del Editor"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpcionesDelEditor.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmVisualizacion 
      Caption         =   "Visualizacion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4815
      Begin VB.CheckBox chkSombras 
         Appearance      =   0  'Flat
         Caption         =   "Sombras activadas *"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   19
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox chkOcultarBarras 
         Appearance      =   0  'Flat
         Caption         =   "Ocultar Barra de Herramientas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   2775
      End
      Begin VB.CheckBox chkSprites 
         Appearance      =   0  'Flat
         Caption         =   "Sprites activados *"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin EditorTDS.UpDownText sldEditorTilesAncho 
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         MaxValue        =   0
         MinValue        =   0
         Enabled         =   -1  'True
      End
      Begin EditorTDS.UpDownText sldEditorTilesAlto 
         Height          =   375
         Left            =   2400
         TabIndex        =   15
         Top             =   1440
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         MaxValue        =   0
         MinValue        =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblOriginalResoluciones 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   855
         Left            =   3360
         TabIndex        =   16
         Top             =   960
         Width           =   1290
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblEditorTilesAlto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de tiles a lo alto: *"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   2010
      End
      Begin VB.Label lblEditorTilesAncho 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de tiles a lo ancho: * "
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   2220
      End
   End
   Begin VB.Frame frmTamano 
      Caption         =   "Tamaño de la pantalla"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   4815
      Begin VB.OptionButton optTamanoPantalla 
         Appearance      =   0  'Flat
         Caption         =   "Personalizada (X * Y)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   4455
      End
      Begin VB.OptionButton optTamanoPantalla 
         Appearance      =   0  'Flat
         Caption         =   "1366 x 768"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   1335
      End
      Begin VB.OptionButton optTamanoPantalla 
         Appearance      =   0  'Flat
         Caption         =   "1280 * 720"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   1335
      End
      Begin VB.OptionButton optTamanoPantalla 
         Appearance      =   0  'Flat
         Caption         =   "800 x 600"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton optTamanoPantalla 
         Appearance      =   0  'Flat
         Caption         =   "1024 x 600"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton optTamanoPantalla 
         Appearance      =   0  'Flat
         Caption         =   "1024 x 768"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblResolucionMonitor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resolucion del Monitor: X * X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   1560
      TabIndex        =   1
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   3240
      TabIndex        =   0
      Top             =   5160
      Width           =   1590
   End
   Begin VB.Label lblAclaracionReiniciar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Es necesario reiniciar el Editor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   4800
      Width           =   2265
   End
End
Attribute VB_Name = "frmOpcionesDelEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const OPT_RES_PERSONALIZADA = 20

Private Sub cmdAceptar_Click()
    Dim tilesAlto As Byte
    Dim tilesAncho As Byte
        
    ' Opciones Generales
    Call actualizarPreferenciaWorkSpace("SPRITES", IIf(Me.chkSprites.value, "SI", "NO"))
    Call actualizarPreferenciaWorkSpace("SOMBRAS", IIf(Me.chkSombras.value, "SI", "NO"))
    
    ' Pantalla
    tilesAlto = Me.sldEditorTilesAlto.value
    tilesAncho = Me.sldEditorTilesAncho.value

    If (tilesAncho / 2) = Int((tilesAncho / 2)) Or (tilesAlto / 2) = Int((tilesAlto / 2)) Then
        Call MsgBox("Para poder emular en el Editor como se ve en el cliente con 'CONTROL + J' es necesario que la cantidad de tiles de ancho y la cantidad de tiles de alto sea impar.", vbInformation + vbOKOnly)
    End If
   
    ' Modificamos
    modPantalla.TilesPantalla.Y = Me.sldEditorTilesAlto.value
    modPantalla.TilesPantalla.X = Me.sldEditorTilesAncho.value
    
    modPantalla.mostrarBarraHerramientas = (Me.chkOcultarBarras.value = 0)
    
    ' Guardamos
    modPantalla.Pantalla_Guardar
    
    ME_General.ModificandoOpcionesEditor = False
    
    Call frmMain.Form_Resize
    frmMain.Hide
    frmMain.Show
    frmMain.Form_Resize
    
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    ME_General.ModificandoOpcionesEditor = False
    Unload Me
End Sub

Private Sub Form_Load()
    ME_General.ModificandoOpcionesEditor = True

    ' Generales
    Me.chkSprites.value = IIf(ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("SPRITES") = "SI", vbChecked, vbUnchecked)
    Me.chkSombras.value = IIf(ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("SOMBRAS") = "SI", vbChecked, vbUnchecked)
        
        
    ' Resolucon de Pantalla
    Me.lblResolucionMonitor = "Resolución del monitor " & (Screen.width \ Screen.TwipsPerPixelX) & " * " & Screen.height \ Screen.TwipsPerPixelY

    Me.lblOriginalResoluciones = "Original: " & vbCrLf & " Cliente: 21 * 21. " & vbCrLf & "Editor 32 * 21."
    
    Me.sldEditorTilesAlto.MinValue = 10
    Me.sldEditorTilesAlto.MaxValue = 42
    
    Me.sldEditorTilesAncho.MinValue = 10
    Me.sldEditorTilesAncho.MaxValue = 42
    
    Me.optTamanoPantalla(OPT_RES_PERSONALIZADA).caption = "Personalizada (" & frmMain.ScaleWidth & " * " & frmMain.ScaleHeight & ")"
    
    Me.sldEditorTilesAlto.value = modPantalla.TilesPantalla.Y
    Me.sldEditorTilesAncho.value = modPantalla.TilesPantalla.X
    
    Me.chkOcultarBarras.value = IIf(modPantalla.mostrarBarraHerramientas, 0, 1)
End Sub

