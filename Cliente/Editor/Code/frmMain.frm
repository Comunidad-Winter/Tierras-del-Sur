VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Tierras del Sur - Editor del Mundo"
   ClientHeight    =   10560
   ClientLeft      =   -240
   ClientTop       =   480
   ClientWidth     =   15150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   704
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1010
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timerAnimarDia 
      Enabled         =   0   'False
      Interval        =   48
      Left            =   11520
      Top             =   10080
   End
   Begin VB.CommandButton cmdConfigurarVertex 
      Caption         =   "Configurar Verxter"
      Height          =   360
      Left            =   -70440
      TabIndex        =   167
      Top             =   1920
      Width           =   2055
   End
   Begin VB.ListBox lstPixelShader_Viejos 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   -74880
      TabIndex        =   155
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton cmdPixelShader_Borrar 
      Caption         =   "Borrar"
      Height          =   360
      Left            =   -68280
      TabIndex        =   154
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Frame frmEntidades 
      Caption         =   "Entidades"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   -74880
      TabIndex        =   112
      Top             =   60
      Width           =   10725
      Begin EditorTDS.ListaConBuscador lstConBuscadorAcciones 
         Height          =   1800
         Left            =   2760
         TabIndex        =   121
         Top             =   375
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   3175
      End
      Begin VB.CommandButton cmdBorrarEntidadMultiple 
         Caption         =   "Borrar"
         Height          =   360
         Left            =   9600
         TabIndex        =   120
         ToolTipText     =   "Borra la entidad del tile seleccionada"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdInsertarEntidadMultipl 
         Caption         =   "Insertar"
         Height          =   360
         Left            =   9600
         TabIndex        =   119
         ToolTipText     =   "Agrega una entidad abajo de la seleccionada"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ListBox lstEntidadesEnTile 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   6960
         TabIndex        =   117
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton cmdInsertarEntidad 
         Caption         =   "Insertar"
         Height          =   360
         Left            =   4800
         TabIndex        =   114
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdBorrarEntidad 
         Caption         =   "Borrar"
         Height          =   360
         Left            =   4800
         TabIndex        =   113
         Top             =   720
         Width           =   1935
      End
      Begin EditorTDS.ListaConBuscador lstEntidades 
         Height          =   1935
         Left            =   120
         TabIndex        =   115
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   3413
      End
      Begin VB.Label lblAccionEntidad 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acción que ejecuta al morir"
         Height          =   195
         Left            =   2760
         TabIndex        =   122
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label lblEntidadEnTile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entidades en (clic derecho en un tile)"
         Height          =   195
         Left            =   6960
         TabIndex        =   118
         Top             =   360
         Width           =   3690
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAclaracionEntidades 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Solo se pueden insertar las que su tipo de vida es por puntos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   4800
         TabIndex        =   116
         Top             =   1560
         Width           =   1860
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdCompilarShader 
      Caption         =   "Compilar"
      Height          =   360
      Left            =   -66240
      TabIndex        =   111
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtShader 
      Appearance      =   0  'Flat
      Height          =   1815
      Left            =   -68280
      MultiLine       =   -1  'True
      TabIndex        =   110
      Text            =   "frmMain.frx":1CCA
      Top             =   60
      Width           =   4095
   End
   Begin VB.Frame frmAcciones 
      Caption         =   "Acciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   -74880
      TabIndex        =   86
      Top             =   60
      Width           =   10695
      Begin VB.ListBox listTipoAccionesDisponibles 
         Appearance      =   0  'Flat
         Height          =   1590
         Left            =   120
         TabIndex        =   92
         Top             =   480
         Width           =   3375
      End
      Begin VB.ListBox listTileAccionActuales 
         Appearance      =   0  'Flat
         Height          =   1590
         ItemData        =   "frmMain.frx":1CDE
         Left            =   3840
         List            =   "frmMain.frx":1CE0
         TabIndex        =   91
         ToolTipText     =   "Doble Clic para modificar"
         Top             =   480
         Width           =   3135
      End
      Begin VB.CommandButton cmdInsertarAccion 
         Appearance      =   0  'Flat
         Caption         =   "Insertar"
         Height          =   375
         Left            =   7200
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   2175
      End
      Begin VB.CommandButton cmdEliminarAccion 
         Caption         =   "<   Eliminar"
         Height          =   375
         Left            =   7200
         TabIndex        =   89
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CheckBox ckbMostrarAcciones 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Mostrar Acciones"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7202
         TabIndex        =   88
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdBorrarAccion 
         Caption         =   "Borrar"
         Height          =   375
         Left            =   7200
         TabIndex        =   87
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblAccionesDisponibles 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Acciones Disponibles"
         Height          =   255
         Index           =   32
         Left            =   120
         TabIndex        =   94
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblAccionesDisponibles 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acciones Disponibles para Utilizar"
         Height          =   255
         Index           =   0
         Left            =   3840
         TabIndex        =   93
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame frmGraficos 
      Caption         =   "Graficos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   120
      TabIndex        =   76
      Top             =   60
      Width           =   10695
      Begin VB.CheckBox chkVer_Graficos 
         Appearance      =   0  'Flat
         Caption         =   "Ver tiles con graficos"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8880
         TabIndex        =   4
         Top             =   120
         Width           =   1790
      End
      Begin VB.OptionButton grh_capa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Capa 5 (adornos paredes)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   3240
         TabIndex        =   165
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CheckBox chkTransparentarTechos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "transparentar"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4800
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   101
         ToolTipText     =   "Hace transparentes los graficos de la Capa 4 para que se pueda ver a través de ellos"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.ListBox lstUltimosGraficosUsados 
         Appearance      =   0  'Flat
         Height          =   1590
         ItemData        =   "frmMain.frx":1CE2
         Left            =   6240
         List            =   "frmMain.frx":1CE4
         TabIndex        =   100
         Top             =   480
         Width           =   2535
      End
      Begin VB.ListBox lstGraficosCopiados 
         Appearance      =   0  'Flat
         Height          =   1005
         ItemData        =   "frmMain.frx":1CE6
         Left            =   8880
         List            =   "frmMain.frx":1CF9
         MultiSelect     =   2  'Extended
         TabIndex        =   84
         ToolTipText     =   "Seleccione una o mas lineas (capas) para copiar su contenido."
         Top             =   1050
         Width           =   1695
      End
      Begin VB.CommandButton cmdBorrarGrafico 
         Caption         =   "Borrar"
         Height          =   375
         Left            =   4920
         TabIndex        =   82
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdInsertarGrafico 
         Caption         =   "Insertar"
         Height          =   375
         Left            =   3240
         TabIndex        =   81
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton grh_capa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Capa 4 (techos)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   3240
         TabIndex        =   80
         Top             =   1920
         Width           =   1455
      End
      Begin VB.OptionButton grh_capa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Capa 2 (decorado piso con volumen)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   3240
         TabIndex        =   79
         Top             =   1200
         Width           =   3015
      End
      Begin VB.OptionButton grh_capa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Capa 1 (decorado piso sin volumen)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   78
         Top             =   960
         Width           =   2895
      End
      Begin VB.OptionButton grh_capa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Capa 3 (arboles, columnas, paredes)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   77
         Top             =   1440
         Width           =   3015
      End
      Begin EditorTDS.ListaConBuscador ListaConBuscadorGraficos 
         Height          =   1935
         Left            =   120
         TabIndex        =   83
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   3413
      End
      Begin VB.Label lblUltimosGraficosUsados 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Últimos usados:"
         Height          =   195
         Left            =   6240
         TabIndex        =   99
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lblGraficosEnPos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Graficos en (click derecho con el insertar)"
         Height          =   435
         Left            =   8880
         TabIndex        =   85
         Top             =   580
         Width           =   1755
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame frmOpcionesGeneralesDelMapa 
      Caption         =   "Herramientas de prueba generales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   -74880
      TabIndex        =   75
      Top             =   60
      Width           =   10695
      Begin VB.CommandButton cmdMasOpciones 
         BackColor       =   &H80000015&
         Caption         =   "Más opciones..."
         Height          =   360
         Left            =   9000
         MaskColor       =   &H80000015&
         TabIndex        =   163
         Top             =   1800
         Width           =   1575
      End
      Begin VB.HScrollBar scrlSangre_Altura 
         Height          =   255
         Left            =   3720
         Max             =   50
         Min             =   1
         TabIndex        =   161
         Top             =   1875
         Value           =   40
         Width           =   2175
      End
      Begin VB.HScrollBar scrSangre_Cantidad 
         Height          =   255
         Left            =   3720
         Max             =   500
         Min             =   1
         TabIndex        =   158
         Top             =   1560
         Value           =   30
         Width           =   2175
      End
      Begin VB.CommandButton cmdLastimarPersonajes 
         Caption         =   "Lastimar Personaje"
         Height          =   360
         Left            =   2880
         TabIndex        =   157
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CommandButton cmdPegarPersonaje 
         Cancel          =   -1  'True
         Caption         =   "Pegar con el personaje"
         Height          =   360
         Left            =   240
         TabIndex        =   156
         Top             =   1800
         Width           =   2415
      End
      Begin VB.ComboBox cmbClimaActual 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdProbarClima 
         Caption         =   "Probar clima"
         Height          =   375
         Left            =   240
         TabIndex        =   102
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton cmdIniciarMusicaAmbiente 
         Caption         =   "Iniciar Música Ambiente"
         Height          =   375
         Left            =   240
         TabIndex        =   95
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblSangre_Altura 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Altura:"
         Height          =   195
         Left            =   2880
         TabIndex        =   160
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lblSangre_Cantidad 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   2880
         TabIndex        =   159
         Top             =   1560
         Width           =   705
      End
      Begin VB.Label lblTamanoTile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tamaño de cada tile: 32 * 32"
         Height          =   315
         Left            =   8520
         TabIndex        =   149
         Top             =   120
         Width           =   2085
      End
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      Caption         =   "Criaturas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2220
      Index           =   12
      Left            =   -74880
      TabIndex        =   70
      Top             =   60
      Width           =   10695
      Begin VB.CheckBox chkVerZonaDondeNaceCraitura 
         Appearance      =   0  'Flat
         Caption         =   "Ver zona donde nace la criatura"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5040
         TabIndex        =   153
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CheckBox chkZonaNacimientoCriaturas 
         Appearance      =   0  'Flat
         Caption         =   "Marcar zonas de nacimiento"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5040
         TabIndex        =   152
         Top             =   1800
         Width           =   2055
      End
      Begin VB.ListBox lstZonaNacimientoCriaturas 
         Appearance      =   0  'Flat
         Height          =   1605
         Left            =   7200
         Style           =   1  'Checkbox
         TabIndex        =   151
         Top             =   480
         Width           =   3375
      End
      Begin VB.CommandButton cmdInsertarNpc 
         Caption         =   "Insertar"
         Height          =   375
         Left            =   5040
         TabIndex        =   73
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdBorrarNpc 
         Caption         =   "Borrar"
         Height          =   375
         Left            =   5040
         TabIndex        =   72
         Top             =   720
         Width           =   1815
      End
      Begin EditorTDS.ListaConBuscador ListaConBuscadorNpcs 
         Height          =   1935
         Left            =   120
         TabIndex        =   71
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   3413
      End
      Begin VB.Label lblZonaCriatura 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zona de nacimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7200
         TabIndex        =   150
         Top             =   240
         Width           =   1665
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Bloqueos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Index           =   10
      Left            =   -74880
      TabIndex        =   65
      Top             =   60
      Width           =   10695
      Begin VB.CheckBox Bloqueos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Mostrar Bloqueos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   69
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdInsertarBloqueo 
         Caption         =   "Insertar"
         Height          =   375
         Left            =   240
         TabIndex        =   68
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdInsertarDobleBloqueo 
         Caption         =   "Insertar Bloqueo Doble"
         Height          =   375
         Left            =   240
         TabIndex        =   67
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton cmdBorrarBloqueo 
         Caption         =   "Borrar"
         Height          =   375
         Left            =   240
         TabIndex        =   66
         Top             =   1200
         Width           =   2415
      End
   End
   Begin VB.CheckBox chkElMapa 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "El mapa tiene iluminacion fija"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   -67200
      TabIndex        =   62
      Top             =   360
      Width           =   2895
   End
   Begin VB.CheckBox chkMostrarLineas 
      Caption         =   "Mostrar lineas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -61562
      TabIndex        =   46
      Top             =   372
      Width           =   1815
   End
   Begin VB.CommandButton Commanddf 
      Caption         =   "Probar entidades"
      Height          =   375
      Left            =   -63960
      TabIndex        =   39
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Frame Frame 
      Caption         =   "Predefinidos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Index           =   14
      Left            =   -74880
      TabIndex        =   37
      Top             =   60
      Width           =   10695
      Begin VB.ListBox lstUltimosPredefinidosUtilizados 
         Appearance      =   0  'Flat
         Columns         =   2
         Height          =   1590
         IntegralHeight  =   0   'False
         Left            =   4440
         TabIndex        =   124
         Top             =   520
         Width           =   4455
      End
      Begin VB.CommandButton cmdAyudaPredefinidos 
         Caption         =   "?"
         Height          =   255
         Left            =   10200
         TabIndex        =   96
         Top             =   600
         Width           =   270
      End
      Begin VB.CommandButton cmdEliminarPreset 
         Caption         =   "< Eliminar"
         Height          =   375
         Left            =   9000
         TabIndex        =   49
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton cmdInsertarPreset 
         Caption         =   "Insertar"
         Height          =   375
         Left            =   9000
         TabIndex        =   38
         Top             =   120
         Width           =   1575
      End
      Begin EditorTDS.ListaConBuscador ListaConBuscadorPresets 
         Height          =   1935
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   3413
      End
      Begin VB.Label lblUltimosPredefinidosUtilizados 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Utimos utilizados"
         Height          =   195
         Left            =   4440
         TabIndex        =   123
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.CheckBox ctriggers 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Mostrar Triggers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   -65280
      TabIndex        =   30
      Top             =   960
      Width           =   1070
   End
   Begin VB.CommandButton hlpTilesets 
      Caption         =   "?"
      Height          =   375
      Left            =   -64680
      TabIndex        =   29
      Top             =   120
      Width           =   375
   End
   Begin VB.Frame Frame 
      Caption         =   "Agua debajo del Terreno"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Index           =   7
      Left            =   -74880
      TabIndex        =   20
      Top             =   120
      Width           =   10695
      Begin VB.CheckBox chkForzarPisoCorrectoOff 
         Appearance      =   0  'Flat
         Caption         =   "Insercción libre del piso"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7920
         TabIndex        =   168
         Tag             =   $"frmMain.frx":1D34
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CheckBox chkVerEfectosdeSonido 
         Appearance      =   0  'Flat
         Caption         =   "Ver efectos de sonido"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7920
         TabIndex        =   164
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CheckBox chkMostrarNumeroTileSet 
         Appearance      =   0  'Flat
         Caption         =   "Mostrar número de Tile del Piso"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7920
         TabIndex        =   162
         Top             =   1920
         Width           =   2655
      End
      Begin VB.CommandButton cmdSeleccionarAreaAguaTierra 
         Caption         =   "Seleccionar"
         Height          =   315
         Left            =   2520
         TabIndex        =   104
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox chkAgua 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Hay agua debajo del mapa"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton tilesets_area_sel_agua 
         Caption         =   "Usar area seleccionada"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   32
         Top             =   1320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.VScrollBar agua_profundidad 
         Height          =   1455
         Left            =   4680
         Max             =   -30
         Min             =   10
         TabIndex        =   21
         Top             =   600
         Value           =   -10
         Width           =   255
      End
      Begin VB.Label lblAguaY2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   1320
         TabIndex        =   109
         Top             =   1680
         Width           =   90
      End
      Begin VB.Label lblAguaY1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   1320
         TabIndex        =   108
         Top             =   1440
         Width           =   90
      End
      Begin VB.Label lblAguaX2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   600
         TabIndex        =   107
         Top             =   1680
         Width           =   90
      End
      Begin VB.Label lblAguaX1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   600
         TabIndex        =   106
         Top             =   1440
         Width           =   90
      End
      Begin VB.Label lblTexturaSeleccionadaAguaTierra 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ninguna"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   105
         Top             =   960
         Width           =   2115
      End
      Begin VB.Label lblTexturaAguaTierra 
         Caption         =   "Textura que se mostrará como ""agua"":"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label lblTexturaAguaTierraArea 
         Caption         =   "Área dentro de la textura:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblTexturaAguaTierraAreaEjes 
         Caption         =   "X1:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblTexturaAguaTierraAreaEjes 
         Caption         =   "X2:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblTexturaAguaTierraAreaEjes 
         Caption         =   "Y1:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   24
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblTexturaAguaTierraAreaEjes 
         Caption         =   "Y2:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   23
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label agua_profundidad_lbl 
         Caption         =   "Nivel del agua: 0"
         Height          =   255
         Left            =   4320
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Triggers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2280
      Index           =   6
      Left            =   -74882
      TabIndex        =   18
      Top             =   60
      Width           =   10695
      Begin VB.CommandButton cmdTriggers_AplicarATodo 
         Caption         =   "Llenar el mapa"
         Height          =   255
         Left            =   4800
         TabIndex        =   127
         Top             =   120
         Width           =   1470
      End
      Begin VB.CommandButton cmdResetearListaTriggers 
         Caption         =   "R"
         Height          =   255
         Left            =   6480
         TabIndex        =   44
         ToolTipText     =   "Destildar todos"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdBorrarTrigger 
         Caption         =   "Borrar"
         Height          =   375
         Left            =   7080
         TabIndex        =   43
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton cmdInsertarTrigger 
         Caption         =   "Insertar"
         Height          =   375
         Left            =   7080
         TabIndex        =   41
         Top             =   360
         Width           =   2175
      End
      Begin VB.ListBox lstTriggers 
         Appearance      =   0  'Flat
         Columns         =   2
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1830
         ItemData        =   "frmMain.frx":1DC4
         Left            =   120
         List            =   "frmMain.frx":1DCE
         Style           =   1  'Checkbox
         TabIndex        =   19
         Top             =   360
         Width           =   6720
      End
      Begin VB.CommandButton Command4 
         Caption         =   "?"
         Height          =   255
         Left            =   10320
         TabIndex        =   31
         Top             =   120
         Width           =   270
      End
      Begin VB.Label lblDescTrigger 
         Caption         =   "Descripción: --"
         Height          =   970
         Left            =   6960
         TabIndex        =   42
         Top             =   1250
         Width           =   3650
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame lblHora 
      Caption         =   "Hora: 00:00hs."
      Height          =   615
      Left            =   -67200
      TabIndex        =   16
      Top             =   1080
      Width           =   2895
      Begin VB.HScrollBar hora_scroll 
         Height          =   255
         LargeChange     =   4
         Left            =   120
         Max             =   96
         Min             =   1
         TabIndex        =   17
         Top             =   240
         Value           =   1
         Width           =   2655
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Luces"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Index           =   8
      Left            =   -74880
      TabIndex        =   10
      Top             =   60
      Width           =   10695
      Begin VB.CheckBox chkAnimarDia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "TimeLapse"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9480
         TabIndex        =   166
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdPonerAntorcha 
         Height          =   480
         Left            =   1560
         Picture         =   "frmMain.frx":1E07
         Style           =   1  'Graphical
         TabIndex        =   129
         ToolTipText     =   "Pone a modo de antorcha una luz sobre el personaje"
         Top             =   1680
         Width           =   495
      End
      Begin VB.CheckBox chkUtilizarBrillo 
         Appearance      =   0  'Flat
         Caption         =   "Utilizar brillo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2280
         TabIndex        =   126
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox forzar_dia_c 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Forzar luz del dia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7680
         TabIndex        =   125
         Top             =   1800
         Width           =   2895
      End
      Begin VB.CheckBox chkLuzSobrePersonaje 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Ver oscuro alrededor del personaje"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5520
         TabIndex        =   98
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox chkPuntosDondeHayLuces 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Tiles donde hay luces"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5520
         TabIndex        =   97
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton cmdBorrarLuz 
         Caption         =   "Borrar Luz"
         Height          =   375
         Left            =   5520
         TabIndex        =   64
         Top             =   650
         Width           =   1815
      End
      Begin VB.CommandButton cmdInsertarLuz 
         Caption         =   "Insertar luz"
         Height          =   375
         Left            =   5520
         TabIndex        =   63
         Top             =   240
         Width           =   1815
      End
      Begin VB.Frame FraBrillo 
         Caption         =   "Brillo"
         Height          =   840
         Left            =   2160
         TabIndex        =   58
         Top             =   240
         Width           =   3255
         Begin VB.HScrollBar luz_luminosidad 
            Height          =   255
            Left            =   120
            Max             =   254
            TabIndex        =   60
            Top             =   480
            Value           =   100
            Width           =   3015
         End
         Begin VB.CommandButton cmd 
            Caption         =   "?"
            Height          =   255
            Left            =   2880
            TabIndex        =   59
            Top             =   120
            Width           =   255
         End
         Begin VB.Label luz_luminosidad_lbl 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Brillo: 50%"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame FraHorario 
         Caption         =   "Horario"
         Height          =   855
         Left            =   2160
         TabIndex        =   52
         Top             =   1080
         Width           =   3255
         Begin VB.HScrollBar horaFinLuz 
            Enabled         =   0   'False
            Height          =   255
            LargeChange     =   4
            Left            =   1080
            Max             =   96
            Min             =   1
            TabIndex        =   55
            Top             =   480
            Value           =   1
            Width           =   2055
         End
         Begin VB.HScrollBar horaInicioLuz 
            Enabled         =   0   'False
            Height          =   255
            LargeChange     =   4
            Left            =   1080
            Max             =   96
            Min             =   1
            TabIndex        =   54
            Top             =   240
            Value           =   1
            Width           =   2055
         End
         Begin VB.CheckBox chkPrendeEn 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Prende en horarios"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   0
            Width           =   1695
         End
         Begin VB.Label lblFin00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fin: 00:00"
            Height          =   255
            Left            =   0
            TabIndex        =   57
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblInicio00 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio: 00:00"
            Height          =   255
            Left            =   0
            TabIndex        =   56
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.CheckBox chkAnimacionFuego 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Animacion fuego"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox chkLuzCuadrada 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Luz cuadrada"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1680
         Width           =   1815
      End
      Begin VB.HScrollBar luces_radio 
         Height          =   255
         Left            =   120
         Max             =   10
         Min             =   2
         TabIndex        =   12
         Top             =   720
         Value           =   3
         Width           =   1695
      End
      Begin VB.Label LabelCol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Color de la iluminacion fija del mapa"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7680
         TabIndex        =   74
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label luces_color 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label 
         Caption         =   "Color:"
         Height          =   255
         Index           =   30
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label luces_radio_label 
         Caption         =   "Radio: 3"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Montañas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Index           =   4
      Left            =   -74880
      TabIndex        =   2
      Top             =   60
      Width           =   10695
      Begin VB.Frame Frame 
         Caption         =   "Radio de la montaña"
         Height          =   1695
         Index           =   5
         Left            =   2400
         TabIndex        =   130
         Top             =   240
         Width           =   4215
         Begin VB.HScrollBar radio_montaña 
            Height          =   255
            Left            =   960
            Max             =   10
            Min             =   1
            TabIndex        =   132
            Top             =   600
            Value           =   3
            Width           =   2175
         End
         Begin VB.CheckBox modifica_alt_pie 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "La altura de la montaña afecta al personaje"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   131
            Top             =   1200
            Value           =   1  'Checked
            Width           =   3855
         End
         Begin VB.Label radio_montaña_lbl 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Radio: 3"
            Height          =   195
            Left            =   1080
            TabIndex        =   134
            Top             =   840
            Width           =   2100
         End
         Begin VB.Label LblTeclaRadioMenos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Utilizá el scroll del mouse para agrandar o achicar"
            Height          =   195
            Left            =   120
            TabIndex        =   133
            Top             =   240
            Width           =   3525
         End
      End
      Begin VB.OptionButton opt_montaña 
         Appearance      =   0  'Flat
         Caption         =   "Altura pie"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton opt_montaña 
         Appearance      =   0  'Flat
         Caption         =   "Suavizar"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton opt_montaña 
         Appearance      =   0  'Flat
         Caption         =   "Meseta"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton opt_montaña 
         Appearance      =   0  'Flat
         Caption         =   "Borrador"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton opt_montaña 
         Appearance      =   0  'Flat
         Caption         =   "Montaña slerp"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton opt_montaña 
         Appearance      =   0  'Flat
         Caption         =   "Montaña suma"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton opt_montaña 
         Appearance      =   0  'Flat
         Caption         =   "Montaña normal"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Objetos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Index           =   13
      Left            =   -74880
      TabIndex        =   33
      Top             =   60
      Width           =   10695
      Begin VB.CheckBox chkMostrarCantidadObjeto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Mostrar Catidad Puesta"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3000
         TabIndex        =   48
         Top             =   1680
         Width           =   2055
      End
      Begin EditorTDS.ListaConBuscador ListaConBuscadorObjetos 
         Height          =   1935
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   3413
      End
      Begin VB.CommandButton cmdBorrarObjeto 
         Caption         =   "Borrar"
         Height          =   360
         Left            =   3000
         TabIndex        =   45
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox obj_cantidad 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         MaxLength       =   4
         TabIndex        =   36
         Text            =   "1"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdInsertarObjeto 
         Caption         =   "Insertar"
         Height          =   375
         Left            =   3000
         TabIndex        =   34
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label 
         Caption         =   "Cantidad:"
         Height          =   255
         Index           =   31
         Left            =   3120
         TabIndex        =   35
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame frmBotonera 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   10950
      TabIndex        =   135
      Top             =   7695
      Width           =   4215
      Begin VB.CommandButton cmdSolapas 
         Appearance      =   0  'Flat
         Caption         =   "Graficos &R"
         Height          =   480
         Index           =   8
         Left            =   0
         TabIndex        =   148
         Top             =   0
         Width           =   1050
      End
      Begin VB.CommandButton cmdSolapas 
         Caption         =   "Triggers &Y"
         Height          =   480
         Index           =   5
         Left            =   2160
         TabIndex        =   147
         Top             =   0
         Width           =   1050
      End
      Begin VB.CommandButton cmdSolapas 
         Caption         =   "Presets &T"
         Height          =   480
         Index           =   1
         Left            =   1080
         TabIndex        =   146
         Top             =   0
         Width           =   1050
      End
      Begin VB.CommandButton cmdSolapas 
         Caption         =   "Criaturas &U"
         Height          =   480
         Index           =   11
         Left            =   3240
         TabIndex        =   145
         Top             =   0
         Width           =   1050
      End
      Begin VB.CommandButton cmdSolapas 
         Appearance      =   0  'Flat
         Caption         =   "Montaña &F"
         Height          =   480
         Index           =   3
         Left            =   0
         TabIndex        =   144
         Top             =   540
         Width           =   1050
      End
      Begin VB.CommandButton cmdSolapas 
         Appearance      =   0  'Flat
         Caption         =   "Piso &G"
         Height          =   480
         Index           =   2
         Left            =   1080
         TabIndex        =   143
         Top             =   540
         Width           =   1050
      End
      Begin VB.CommandButton cmdSolapas 
         Appearance      =   0  'Flat
         Caption         =   "Bloqueos  &H"
         Height          =   480
         Index           =   10
         Left            =   2160
         TabIndex        =   142
         Top             =   540
         Width           =   1050
      End
      Begin VB.CommandButton cmdSolapas 
         Appearance      =   0  'Flat
         Caption         =   "Luces &J"
         Height          =   480
         Index           =   7
         Left            =   3225
         MaskColor       =   &H8000000C&
         Style           =   1  'Graphical
         TabIndex        =   141
         Top             =   540
         Width           =   1050
      End
      Begin VB.CommandButton cmdSolapas 
         Appearance      =   0  'Flat
         Caption         =   "Acciones &C"
         Height          =   480
         Index           =   9
         Left            =   0
         TabIndex        =   140
         Top             =   1080
         Width           =   1050
      End
      Begin VB.CommandButton cmdSolapas 
         Appearance      =   0  'Flat
         Caption         =   "Particulas &V"
         Height          =   480
         Index           =   4
         Left            =   1080
         TabIndex        =   139
         Top             =   1080
         Width           =   1050
      End
      Begin VB.CommandButton cmdSolapas 
         Appearance      =   0  'Flat
         Caption         =   "Objetos &B"
         Height          =   480
         Index           =   6
         Left            =   2160
         TabIndex        =   138
         Top             =   1080
         Width           =   1050
      End
      Begin VB.CommandButton cmdSolapas 
         Appearance      =   0  'Flat
         Caption         =   "Mapa &N"
         Height          =   480
         Index           =   0
         Left            =   3225
         TabIndex        =   137
         Top             =   1080
         Width           =   1050
      End
      Begin VB.CommandButton cmdSolapas 
         Appearance      =   0  'Flat
         Caption         =   "Entidades &E"
         Height          =   480
         Index           =   12
         Left            =   0
         TabIndex        =   136
         Top             =   1620
         Width           =   1050
      End
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   12360
      Top             =   10080
   End
   Begin VB.PictureBox renderer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   7680
      Left            =   0
      ScaleHeight     =   512
      ScaleMode       =   0  'User
      ScaleWidth      =   1017
      TabIndex        =   1
      Top             =   0
      Width           =   15255
      Begin MSComDlg.CommonDialog ColorDialog 
         Left            =   3000
         Top             =   2880
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Label lblTapaSolapas 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Generando imagen..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   7095
      Left            =   0
      TabIndex        =   128
      Top             =   0
      Visible         =   0   'False
      Width           =   15210
   End
   Begin VB.Image Image1 
      Height          =   180
      Index           =   2
      Left            =   11880
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   2010
   End
   Begin VB.Image Image1 
      Height          =   180
      Index           =   0
      Left            =   8400
      MousePointer    =   99  'Custom
      Top             =   8760
      Width           =   1005
   End
   Begin VB.Image PicResu 
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   10200
      Stretch         =   -1  'True
      Top             =   5880
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Archivo"
      Begin VB.Menu mnuNewMap 
         Caption         =   "Nuevo Mapa"
      End
      Begin VB.Menu iii 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenPak 
         Caption         =   "Abrir"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Guardar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAndCompile 
         Caption         =   "Guardar en cliente y compilar para server"
      End
      Begin VB.Menu mnuCambiarZonaTrabajo 
         Caption         =   "Cambiar zona de trabajo"
         WindowList      =   -1  'True
         Begin VB.Menu mnuZonaTrabajo 
            Caption         =   "No hay zonas cargadas"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuExportarZonaDeTrabajo 
         Caption         =   "Exportar zona de trabajo a Imagen"
      End
      Begin VB.Menu dsdaskdasdkaksd 
         Caption         =   "-"
      End
      Begin VB.Menu sdasdasdas 
         Caption         =   "Mapas de prueba"
         Begin VB.Menu mnuAbrir 
            Caption         =   "Abrir archivo..."
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Guardar archivo..."
         End
      End
      Begin VB.Menu separador2Archivo 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalirDelEditor 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "Edición"
      Begin VB.Menu mnuDeshacer 
         Caption         =   "Deshacer"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRehacer 
         Caption         =   "Rehacer"
         Shortcut        =   ^Y
      End
      Begin VB.Menu separadorEdicion1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCortar 
         Caption         =   "Cortar"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopiar 
         Caption         =   "Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCopiarPortapapeles 
         Caption         =   "Copiar desde el portapeles"
         Begin VB.Menu mnuPortapapeles 
            Caption         =   "1: < Vacio >"
            Enabled         =   0   'False
            Index           =   1
         End
      End
      Begin VB.Menu mnuPegar 
         Caption         =   "Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuTrabajarCon 
         Caption         =   "Trabajar con..."
         Begin VB.Menu mnuTrabajarCon_Todo 
            Caption         =   "Todo"
         End
         Begin VB.Menu mnuTrabajarCon_Nada 
            Caption         =   "Nada"
         End
         Begin VB.Menu mnuTrabajarConSeparador 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTrabajarCon_Elemento 
            Caption         =   "Acciones"
            Index           =   0
         End
         Begin VB.Menu mnuTrabajarCon_Elemento 
            Caption         =   "Criaturas"
            Index           =   1
         End
         Begin VB.Menu mnuTrabajarCon_Elemento 
            Caption         =   "Gráficos"
            Index           =   2
         End
         Begin VB.Menu mnuTrabajarCon_Elemento 
            Caption         =   "Luces"
            Index           =   3
         End
         Begin VB.Menu mnuTrabajarCon_Elemento 
            Caption         =   "Triggers y Boqueos"
            Index           =   4
         End
         Begin VB.Menu mnuTrabajarCon_Elemento 
            Caption         =   "Objetos"
            Index           =   5
         End
         Begin VB.Menu mnuTrabajarCon_Elemento 
            Caption         =   "Particulas"
            Index           =   6
         End
         Begin VB.Menu mnuTrabajarCon_Elemento 
            Caption         =   "Piso"
            Index           =   7
         End
      End
      Begin VB.Menu separadorEdicion2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBorrarSeleccion 
         Caption         =   "Borrar"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu separadorEdicion4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCrearPredefinido 
         Caption         =   "Crear Predefinido"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuCrearZonaCriatura 
         Caption         =   "Crear Zona de Nacimiento de Criatura"
         Shortcut        =   ^N
      End
      Begin VB.Menu separadorEdicion3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopiarBordes 
         Caption         =   "Copiar bordes a este mapa"
      End
      Begin VB.Menu mnuInsertarPisoEnMapa 
         Caption         =   "Insertar piso en área (piso)"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuInsertarPisoEnMapaBloqueos 
         Caption         =   "Insertar piso en área (bloqueos)"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuInsertarBloqueoArea 
         Caption         =   "Insertar bloqueo en área"
         Shortcut        =   ^{F3}
      End
   End
   Begin VB.Menu mnumapa 
      Caption         =   "Mapa"
      Begin VB.Menu mnuSetNum 
         Caption         =   "Definir número"
      End
      Begin VB.Menu cmdPropiedadesMapa 
         Caption         =   "Propiedades"
      End
      Begin VB.Menu separadorMapa0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenerarInformes 
         Caption         =   "Generar informe"
      End
      Begin VB.Menu mnuExportarAImagen 
         Caption         =   "Exportar a imagen"
      End
   End
   Begin VB.Menu mnu_ver 
      Caption         =   "Ver"
      Begin VB.Menu mnuMiniMapa 
         Caption         =   "Mini mapa"
         Begin VB.Menu chkMiniMapa 
            Caption         =   "Mini mapa"
            Shortcut        =   {F12}
         End
         Begin VB.Menu separadorMiniMapa 
            Caption         =   "-"
         End
         Begin VB.Menu chkMiniMapaBloqueos 
            Caption         =   "Bloqueos"
            Shortcut        =   {F1}
         End
         Begin VB.Menu chkMiniMapaLuces 
            Caption         =   "Luces"
            Shortcut        =   {F2}
         End
         Begin VB.Menu chkMiniMapaNPC 
            Caption         =   "NPC"
            Shortcut        =   {F3}
         End
         Begin VB.Menu chkMiniMapaColores 
            Caption         =   "Colores del piso"
            Shortcut        =   {F4}
         End
         Begin VB.Menu chkMiniMapaAcciones 
            Caption         =   "Acciones"
            Shortcut        =   {F5}
         End
         Begin VB.Menu chkMiniMapaTriggers 
            Caption         =   "Triggers"
            Shortcut        =   {F8}
         End
         Begin VB.Menu chkMiniMapaPiso 
            Caption         =   "Piso"
            Shortcut        =   {F9}
         End
      End
      Begin VB.Menu ver_bloqueos 
         Caption         =   "Bloqueos"
         Shortcut        =   ^B
      End
      Begin VB.Menu ver_triggers 
         Caption         =   "Triggers"
         Shortcut        =   ^T
      End
      Begin VB.Menu ver_acciones 
         Caption         =   "Acciones"
         Shortcut        =   ^A
      End
      Begin VB.Menu ver_luces 
         Caption         =   "Tiles donde hay luces"
         Shortcut        =   ^L
      End
      Begin VB.Menu ver_graficos 
         Caption         =   "Tiles donde hay graficos"
      End
      Begin VB.Menu mnuVerCantidadObjetos 
         Caption         =   "Cantidad de objetos"
         Shortcut        =   ^O
      End
      Begin VB.Menu ver_particulas 
         Caption         =   "Partículas"
         Shortcut        =   ^{F7}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNumeroTilePiso 
         Caption         =   "Número de tile del Piso"
         Shortcut        =   ^{F9}
      End
      Begin VB.Menu mnuNumeroEfectoSonido 
         Caption         =   "Número de Efecto de Sonido"
      End
      Begin VB.Menu mnuTodoDia 
         Caption         =   "Todo de dia"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuTechosTransparentes 
         Caption         =   "Techos transparentes"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu ver_ZonadeNacCriatura 
         Caption         =   "Zona de Nacimiento de Criatura"
      End
      Begin VB.Menu ver_ZonadeNacCriaturas 
         Caption         =   "Zonas de Nacimiento de Criaturas"
      End
      Begin VB.Menu sdfsdfsdfsdf 
         Caption         =   "-"
      End
      Begin VB.Menu ver_grilla 
         Caption         =   "Grilla"
         Shortcut        =   ^G
      End
      Begin VB.Menu ver_area 
         Caption         =   "Area de juego"
         Shortcut        =   ^J
      End
      Begin VB.Menu ver_char 
         Caption         =   "Personaje (modo caminata)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnu_tools 
      Caption         =   "Recursos"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnu_paker 
         Caption         =   "Agregar/cambiar archivos de Recursos"
         Enabled         =   0   'False
      End
      Begin VB.Menu separadorasds 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ventana_indexar 
         Caption         =   "Configuración de Gráficos"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTilesets 
         Caption         =   "Configuración de Pisos"
      End
      Begin VB.Menu mnuConfigEfectos 
         Caption         =   "Configuración de Efectos"
      End
      Begin VB.Menu mnuConfigSonidos 
         Caption         =   "Configuración de Sonidos"
      End
      Begin VB.Menu mnuConfigPisadas 
         Caption         =   "Configuración de Sonidos de Pisadas"
      End
      Begin VB.Menu mnuConfigPersonajes 
         Caption         =   "Configuración de Personajes"
      End
      Begin VB.Menu separadorConfigurar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfigCriaturas 
         Caption         =   "Configuración de Criaturas"
      End
      Begin VB.Menu mnuConfigObjetos 
         Caption         =   "Configuración de Objetos"
      End
      Begin VB.Menu mnuConfigHechzos 
         Caption         =   "Configuración de Hechizos"
      End
      Begin VB.Menu mnuConfigEntidades 
         Caption         =   "Configuración de Entidades"
      End
      Begin VB.Menu mnuConfigNiveles 
         Caption         =   "Configuración de Niveles"
      End
      Begin VB.Menu mnuConfigFacciones 
         Caption         =   "Configuración de Facciones"
         Begin VB.Menu mnuConfigArmadaReal 
            Caption         =   "Armada Real"
         End
         Begin VB.Menu mnuConfigLegiónOscura 
            Caption         =   "Legión Oscura"
         End
      End
      Begin VB.Menu separadorConfigurar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfigRings 
         Caption         =   "Configuración de Rings"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuConfigDescansos 
         Caption         =   "Configuración de Descansos"
         Enabled         =   0   'False
      End
      Begin VB.Menu separadorConfigurar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfigAspectos 
         Caption         =   "Configuración de Aspectos"
      End
   End
   Begin VB.Menu mnuEquipo 
      Caption         =   "Equipo"
      Begin VB.Menu mnu_CDMObtenerNovedades 
         Caption         =   "Obtener novedades"
         Enabled         =   0   'False
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnu_CDMCompartirNovedades 
         Caption         =   "Compartir mis novedades"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_CDMPublicarServidor 
         Caption         =   "Publicar en el Servidor"
      End
   End
   Begin VB.Menu cdm_ayuda 
      Caption         =   "Ayuda"
      Begin VB.Menu cmd_acerca_de 
         Caption         =   "Acerca de..."
      End
      Begin VB.Menu cmd_manual 
         Caption         =   "Ingresar al manual"
      End
      Begin VB.Menu sepAyuda1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpcionesDelEditor 
         Caption         =   "Opciones del Editor"
      End
      Begin VB.Menu mnuReportarBug 
         Caption         =   "Reportar Error"
      End
   End
   Begin VB.Menu mnuCambiosPendientes 
      Caption         =   "Cliente"
      NegotiatePosition=   2  'Middle
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'http://go.microsoft.com/?linkid=9776485

Option Explicit

Private Enum eSolapasEditor
    Mapa = 0
    Herramientas = 1
    Tilesets = 2
    Montañas = 3
    Particulas = 4
    Triggers = 5
    objetos = 6
    Luces = 7
    Graficos = 8
    Acciones = 9
    Bloqueos = 10
    Npcs = 11
    Entidades = 12
End Enum

' Zona del Mundo seleccionada para trabajar
Private zonaActual As String

Public tX As Integer
Public tY As Integer

Public MouseX As Long
Public MouseY As Long

Public MouseBoton As Long
Public MouseShift As Long

Public clicX As Long
Public clicY As Long

Private ignorarMouseUp As Boolean 'Esto sirve cuando se utiliza el copiar/cortar.
Public focoEnElRender As Boolean
Private ultimaSolapaseleccionada As eSolapasEditor
Private comandoMontaniaActual As cComandoInsertarMotania
Public editandoTileSets As Boolean

' Tool Graficos temporal
Private capaSeleccionada As Byte
Private graficoSeleccionadoIndex As Integer

' Tool criaturas
Private criaturaSeleccion_Index As Integer
Private criaturaSeleccion_Zona As Byte

'Tool Luz, Seleccion temporal
Private luzSeleccionada As tLuzPropiedades

' EL usuario selecciona en el menu con que elementos desea trabajar
' En las herramientas de pegado, eliminado o creación de predefinidos
Private trabajandoConElementos As Tools

' El usuario esta viendo la barra de herramientas?. Solo tiene sentido cuando
' esta activada la opción "Ocultar barra de herramientas" del Editor
Private mostrandoBarra As Boolean

Sub ActChecks()

frmMain.ctriggers.value = DRAWTRIGGERS
frmMain.Bloqueos.value = DRAWBLOQUEOS
frmMain.ckbMostrarAcciones.value = dibujarAccionTile
frmMain.chkMostrarCantidadObjeto.value = dibujarCantidadObjetos
frmMain.chkPuntosDondeHayLuces.value = ME_Tools.mostrarTileDondeHayLuz
frmMain.chkMostrarNumeroTileSet.value = ME_Tools.mostrarTileNumber
frmMain.chkTransparentarTechos.value = ME_Tools.dibujarTechosTransparentes
frmMain.chkZonaNacimientoCriaturas.value = ME_Tools.dibujarZonaNacimientoCriaturas
frmMain.chkVerZonaDondeNaceCraitura.value = ME_Tools.dibujarZonaNacimientoCriatura
frmMain.chkVerEfectosdeSonido.value = ME_Tools.mostrarTileEfectoSonidoPasos
frmMain.chkVer_Graficos.value = ME_Tools.mostrarTileDondeHayGraficos

End Sub


Sub enable_agua_buttons()

If AbriendoMapa Then Exit Sub

mapinfo.agua_rect.left = minl(val(Me.lblAguaX1.caption), val(lblAguaX2.caption))
mapinfo.agua_rect.right = maxl(val(lblAguaX1.caption), val(lblAguaX2.caption))
mapinfo.agua_rect.top = minl(val(lblAguaY1.caption), val(lblAguaY2.caption))
mapinfo.agua_rect.bottom = maxl(val(lblAguaY1.caption), val(lblAguaY2.caption))
mapinfo.agua_profundidad = agua_profundidad.value

mapinfo.agua_tileset = Me_indexar_Pisos.obtenerIDTileSet(Me.lblTexturaSeleccionadaAguaTierra)
mapinfo.UsaAguatierra = chkAgua.value <> vbUnchecked

If Not prgRun Then Exit Sub

Call Engine_Landscape_Water.RemakeWaterTilenumbers(mapinfo.agua_rect.left, mapinfo.agua_rect.top, mapinfo.agua_rect.right, mapinfo.agua_rect.bottom)
'Call Engine_Landscape_Water.recalcular_colores_agua
'Call Engine_Landscape_Water.recalcular_opacidades_agua

setValoresAgua
End Sub

Public Sub setValoresAgua()
    
    ' Agua
    Me.lblAguaX1.caption = mapinfo.agua_rect.left
    Me.lblAguaX2.caption = mapinfo.agua_rect.right
    Me.lblAguaY1.caption = mapinfo.agua_rect.top
    Me.lblAguaY2.caption = mapinfo.agua_rect.bottom

    Me.lblTexturaSeleccionadaAguaTierra.caption = Engine_Tilesets.Tilesets(mapinfo.agua_tileset).nombre
    
    agua_profundidad.value = mapinfo.agua_profundidad
    agua_profundidad_lbl.caption = "Nivel del agua: " & mapinfo.agua_profundidad
        
    chkAgua.value = IIf(mapinfo.UsaAguatierra, vbChecked, vbUnchecked)
        
    If chkAgua.value = 0 Then 'No tiene agua
        Call activarMenuAguaTierra(False)
    Else
        Call activarMenuAguaTierra(True)
    End If
    
    ' Luz del mapa
    Me.chkElMapa.value = IIf(mapinfo.ColorPropio, vbChecked, vbUnchecked)
    LabelCol.BackColor = RGB(mapinfo.BaseColor.r, mapinfo.BaseColor.g, mapinfo.BaseColor.b)

End Sub

Private Sub activarMenuAguaTierra(activoMenuAguaTierra As Boolean)
    lblTexturaAguaTierra.Enabled = activoMenuAguaTierra
    lblTexturaSeleccionadaAguaTierra.Enabled = activoMenuAguaTierra
    lblTexturaAguaTierraArea.Enabled = activoMenuAguaTierra
    
    Dim i As Byte
    
    For i = 0 To 3
        lblTexturaAguaTierraAreaEjes(i).Enabled = activoMenuAguaTierra
    Next i
    
    agua_profundidad_lbl.Enabled = activoMenuAguaTierra
    agua_profundidad.Enabled = activoMenuAguaTierra
    cmdSeleccionarAreaAguaTierra.Enabled = activoMenuAguaTierra
    agua_profundidad_lbl.Enabled = activoMenuAguaTierra
    agua_profundidad.Enabled = activoMenuAguaTierra
    cmdSeleccionarAreaAguaTierra.Enabled = activoMenuAguaTierra
End Sub
Private Sub agua_profundidad_Change()
    enable_agua_buttons
End Sub

Private Sub agua_profundidad_Scroll()
    enable_agua_buttons
End Sub

Private Sub asd_Click()

End Sub

Private Sub Bloqueos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ver_bloqueos_Click
End Sub


'Private Sub cdm_enviar_todo_mnu_Click()
'CDM_EnviarTodo
'End Sub

Private Sub chkAgua_Click()
    enable_agua_buttons
End Sub

Private Sub chkAnimarDia_Click()
    timerAnimarDia.Enabled = chkAnimarDia.value = vbChecked
    
End Sub

Private Sub chkElMapa_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mapinfo.ColorPropio = Me.chkElMapa.value = vbChecked
    VBC2RGBC LabelCol.BackColor, mapinfo.BaseColor
End Sub

Private Sub chkForzarPisoCorrectoOff_Click()

    
    Me_Tools_TileSet.NoForzarInserccionCorrecta = (Me.chkForzarPisoCorrectoOff.value = vbChecked)
    
    If Me_Tools_TileSet.NoForzarInserccionCorrecta Then
        Call GUI_Alert("Con esta opción activada el Editor no va a chequear en donde estás insertando el piso. " & vbCrLf & "Lo podés poner en cualquier lugar y, por lo tanto, equivocarte, poner el piso mal y generar un error de encaje entre los suelos.", "Cuidado")
    End If
End Sub

Private Sub chkLuzSobrePersonaje_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Render_Radio_Luz = Not Render_Radio_Luz
End Sub

Private Sub chkMiniMapa_Click()
    chkMiniMapa.checked = Not chkMiniMapa.checked
    ME_MiniMap.miniMapVisible = chkMiniMapa.checked
End Sub

Private Sub chkMiniMapaAcciones_Click()
    chkMiniMapaAcciones.checked = Not chkMiniMapaAcciones.checked
    
    If chkMiniMapaAcciones.checked Then
        ME_MiniMap.miniMapaTipo = ME_MiniMap.miniMapaTipo Or eMiniMapaTipo.emmAcciones
    Else
        ME_MiniMap.miniMapaTipo = ME_MiniMap.miniMapaTipo And (Not eMiniMapaTipo.emmAcciones)
    End If
    
    miniMap_Redraw
End Sub

Private Sub chkMiniMapaBloqueos_Click()
    chkMiniMapaBloqueos.checked = Not chkMiniMapaBloqueos.checked
    
    If chkMiniMapaBloqueos.checked Then
        ME_MiniMap.miniMapaTipo = ME_MiniMap.miniMapaTipo Or eMiniMapaTipo.emmBloqueos
    Else
        ME_MiniMap.miniMapaTipo = ME_MiniMap.miniMapaTipo And (Not eMiniMapaTipo.emmBloqueos)
    End If
    
    miniMap_Redraw
End Sub

Private Sub chkMiniMapaColores_Click()
    chkMiniMapaColores.checked = Not chkMiniMapaColores.checked
    
    If chkMiniMapaColores.checked Then
        ME_MiniMap.miniMapaTipo = ME_MiniMap.miniMapaTipo Or eMiniMapaTipo.emmColores
    Else
        ME_MiniMap.miniMapaTipo = ME_MiniMap.miniMapaTipo And (Not eMiniMapaTipo.emmColores)
    End If
    
    miniMap_Redraw
End Sub

Private Sub chkMiniMapaLuces_Click()
    chkMiniMapaLuces.checked = Not chkMiniMapaLuces.checked
    
    If chkMiniMapaLuces.checked Then
        ME_MiniMap.miniMapaTipo = ME_MiniMap.miniMapaTipo Or eMiniMapaTipo.emmLuces
    Else
        ME_MiniMap.miniMapaTipo = ME_MiniMap.miniMapaTipo And (Not eMiniMapaTipo.emmLuces)
    End If
    
    miniMap_Redraw
End Sub


Private Sub chkMiniMapaNPC_Click()
    chkMiniMapaNPC.checked = Not chkMiniMapaNPC.checked
    
    If chkMiniMapaNPC.checked Then
        ME_MiniMap.miniMapaTipo = ME_MiniMap.miniMapaTipo Or eMiniMapaTipo.emmNPC
    Else
        ME_MiniMap.miniMapaTipo = ME_MiniMap.miniMapaTipo And (Not eMiniMapaTipo.emmNPC)
    End If
    
    miniMap_Redraw
End Sub

Private Sub chkMiniMapaPiso_Click()

    chkMiniMapaPiso.checked = Not chkMiniMapaPiso.checked
    
    If chkMiniMapaPiso.checked Then
        ME_MiniMap.miniMapaTipo = ME_MiniMap.miniMapaTipo Or eMiniMapaTipo.emmPiso
    Else
        ME_MiniMap.miniMapaTipo = ME_MiniMap.miniMapaTipo And (Not eMiniMapaTipo.emmAcciones)
    End If
    
    miniMap_Redraw

End Sub

Private Sub chkMiniMapaTriggers_Click()
    chkMiniMapaTriggers.checked = Not chkMiniMapaTriggers.checked
    
    If chkMiniMapaTriggers.checked Then
        ME_MiniMap.miniMapaTipo = ME_MiniMap.miniMapaTipo Or eMiniMapaTipo.emmTriggers
    Else
        ME_MiniMap.miniMapaTipo = ME_MiniMap.miniMapaTipo And (Not eMiniMapaTipo.emmTriggers)
    End If
    
    miniMap_Redraw

End Sub

Private Sub chkMostrarCantidadObjeto_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call mnuVerCantidadObjetos_Click
End Sub

Private Sub chkMostrarNumeroTileSet_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call mnuNumeroTilePiso_Click
End Sub

Private Sub chkPuntosDondeHayLuces_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ver_luces_Click
End Sub

Private Sub chkTransparentarTechos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     Call mnuTechosTransparentes_Click
End Sub

Private Sub chkUtilizarBrillo_Click()
    Call setEnabledHijos((Me.chkUtilizarBrillo.value = 1), Me.FraBrillo, Me)
    ActualizarEstadoHerramientaLuz
End Sub

Private Sub chkVer_Graficos_Click()
    Call ver_graficos_Click
    ActChecks
End Sub

Private Sub chkVerEfectosdeSonido_Click()
    Call mnuNumeroEfectoSonido_Click
End Sub

Private Sub chkVerZonaDondeNaceCraitura_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ver_ZonadeNacCriatura_Click
End Sub

Private Sub chkZonaNacimientoCriaturas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ver_ZonadeNacCriaturas_Click
End Sub

Private Sub ckbMostrarAcciones_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ver_acciones_Click
End Sub

Private Sub cmd_acerca_de_Click()
    frmAcercaDe.Show , Me
End Sub

Private Sub cmd_Click()
      GUI_Alert "Mientras más brillosa sea la luz, más presente estará el color de la luz. Una luz con brillo 1% apenas se notará la luz."
End Sub

Private Sub cmd_manual_Click()
    ShellExecute hwnd, "open", "http://www.tdsx.com.ar/manual", vbNullString, vbNullString, vbNormalFocus
End Sub



Private Sub cmdAyudaPredefinidos_Click()
    GUI_Alert "Para crear un predefinido seleccione un area del mapa y luego presione " & Chr$(255) & "CONTROL + P" & Chr$(255) & ". Luego deberá elegir un nombre para este preset.", "Información"
End Sub

Private Sub cmdBorrarEntidad_Click()
    Call Me_Tools_Entidades.click_BorrarEntidad
End Sub

Private Sub cmdBorrarEntidadMultiple_Click()
    Dim posicion As Byte
    Dim posx As Integer
    Dim posy As Integer
    Dim backupHerramienta As Integer
    
    If Me.lstEntidadesEnTile.listIndex > -1 Or Me.lstEntidadesEnTile.ListCount = 1 Then
        If Me.lstEntidadesEnTile.listIndex > -1 Then
            posicion = Me.lstEntidadesEnTile.listIndex + 1
        Else
            posicion = 1
        End If
                                
        Call obtenerPosicionDeEntidadSeleccionada(posx, posy)
                
        backupHerramienta = eHerramientasEntidades.insertar
                
        'Elimino
        Call Me_Tools_Entidades.seleccionarEntidadBorrado(posicion)

        Call click_BorrarEntidadEnPos
        
        Call modSeleccionArea.puntoArea(areaSeleccionada, posx, posy)
        click_tool vbLeftButton
    
        'Restauramos lo que teniamos antes de utilizar esta herramienta
        Me_Tools_Entidades.herramientaInternaEntidades = backupHerramienta
        Call actualizarEntidadAccionSeleccionada
        Call Me_Tools_Entidades.activarUltimaHerramienta
        
        
        Call actualizarListaEntidadesEnTile(posx, posy)

    Else
        MsgBox "Tenes que seleccionar cual de las entidades que hay en el tile queres eliminar"
    End If
End Sub

Public Sub actualizarListaEntidadesEnTile(tileX As Integer, tileY As Integer)
    Dim i As Integer
    Dim tempInt As Integer
    Dim TempStr As String
    
    frmMain.lblEntidadEnTile = "Entidades en (" & tileX & "," & tileY & ")"
    frmMain.lstEntidadesEnTile.Clear
                                
    'Recorro la lista enlazada agregando a la lista del editor
    i = 1
                  
    tempInt = EntidadesMap(tileX, tileY)
    
    Do While (tempInt > 0)
        
        TempStr = EntidadesIndexadas(Engine_Entidades.Entidades(tempInt).numeroIndexadoEntidad).nombre
                                    
        If Not Engine_Entidades.Entidades(tempInt).accion Is Nothing Then
            TempStr = TempStr & "(" & Engine_Entidades.Entidades(tempInt).accion.GetNombre & ")"
        End If
        frmMain.lstEntidadesEnTile.AddItem i & " - " & TempStr
        
        tempInt = Engine_Entidades.Entidades(tempInt).Next
        i = i + 1
    Loop
                            
End Sub

'*****************************************************************************/
'Luces
Private Sub cmdBorrarLuz_Click()
    Call Me_Tools_Luces.click_BorrarLuz
End Sub



Private Sub cmdCompilarShader_Click()
'PixelShaderBump = CreateShaderFromCode(txtShader.text)

'Call Me.lstPixelShader_Viejos.AddItem(HelperStrings.QuitarDobleEspacios(Replace$(txtShader.text, vbNewLine, "")))

PixelShaderCatalog(ePixelShaders.Agua).codigo = txtShader.text
Engine_PixelShaders.Engine_PixelShaders_EngineReiniciado


End Sub

Private Sub cmdConfigurarVertex_Click()
    Call frmVertexShader.Show(, Me)
End Sub

Private Sub cmdInsertarEntidad_Click()
    Call Me_Tools_Entidades.click_InsertarEntidad
End Sub

Private Sub obtenerPosicionDeEntidadSeleccionada(posx As Integer, posy As Integer)
    Dim texto As String

    texto = Me.lblEntidadEnTile.caption
    
    'La posicion se encuentra en (X,Y)
    posx = val(mid$(texto, InStr(1, texto, "(") + 1, InStr(1, texto, ",") - InStr(1, texto, "(")))
    posy = val(mid$(texto, InStr(1, texto, ",") + 1, InStr(1, texto, ")") - InStr(1, texto, ",")))
    
End Sub
Private Sub cmdInsertarEntidadMultipl_Click()
    Dim posicion As Byte
    Dim posx As Integer
    Dim posy As Integer
    Dim backupHerramienta As Integer
    
    Dim idAccion As Integer
    Dim idEntidad As Integer
    Dim accion As iAccionEditor
    
    If Me.lstEntidadesEnTile.listIndex >= 0 Then
        'Obtengo en que posicion de la lista lo voy a insertar
        posicion = Me.lstEntidadesEnTile.listIndex + 1
  
        ' Obtengo lo que voy a isnertar
        idAccion = Me.lstConBuscadorAcciones.obtenerIDValor
        
        If idAccion > 0 Then
            Set accion = ME_modAccionEditor.obtenerAccionID(idAccion)
        Else
            Set accion = Nothing
        End If
               
        idEntidad = Me.lstEntidades.obtenerIDValor
    
        'Obtengo en que tile lo voy a insertar
        obtenerPosicionDeEntidadSeleccionada posx, posy
        
        'Hago un backup de la herramienta seleccionada
        backupHerramienta = eHerramientasEntidades.insertar
        
        'Seteamos con lo que vamos a trbajar
        Call Me_Tools_Entidades.seleccionarEntidad(idEntidad, accion, posicion)

        'Activamos la herramienta
        Call click_InsertarEntidadEnPos
        
        'Seleccionamos el area donde vamos a trabajar
        Call modSeleccionArea.puntoArea(areaSeleccionada, posx, posy)
        
        'Ejecutamos la acción sobre esa area
        click_tool vbLeftButton
        
        'Restauramos lo que teniamos antes de utilizar esta herramienta
        Me_Tools_Entidades.herramientaInternaEntidades = backupHerramienta
        Call actualizarEntidadAccionSeleccionada
        Call Me_Tools_Entidades.activarUltimaHerramienta
    End If
End Sub

Private Sub cmdInsertarLuz_Click()
    Call Me_Tools_Luces.seleccionarLuz(luzSeleccionada)
    Call Me_Tools_Luces.click_InsertarLuz
End Sub
Private Sub chkAnimacionFuego_Click()
    ActualizarEstadoHerramientaLuz
End Sub

Private Sub chkLuzCuadrada_Click()
    ActualizarEstadoHerramientaLuz
End Sub

Private Sub chkPrendeEn_Click()
    If chkPrendeEn.value = vbChecked Then
        horaInicioLuz.Enabled = True
        horaFinLuz.Enabled = True
        
        UpdateScrollsHora
    Else
        luzSeleccionada.luzInicio = 0
        luzSeleccionada.luzFin = 0

        horaInicioLuz.Enabled = False
        horaFinLuz.Enabled = False
    End If
    
    Call Me_Tools_Luces.seleccionarLuz(luzSeleccionada)
End Sub


Private Sub cmdLastimarPersonajes_Click()
    FX_Hit_Create UserCharIndex, Me.scrSangre_Cantidad.value, 2000, mzRed
    Sangre_Crear UserCharIndex, Me.scrSangre_Cantidad.value, 3000, Me.scrlSangre_Altura.value
End Sub

Private Sub cmdMasOpciones_Click()
    load frmMisc

    frmMisc.Show , Me

    Call modPosicionarFormulario.posicionarAbajoCentro(Me, frmMisc)
End Sub

Private Sub cmdPonerAntorcha_Click()
    Dim luzBackup As tLuzPropiedades
    
    Dim comando As cComandoInsertarAntorcha
    Set comando = New cComandoInsertarAntorcha
    
    ' La borramos
    If CharList(UserCharIndex).luz > 0 Then
        luzBackup = luzSeleccionada
        luzBackup.LuzRadio = 0
        
        Call comando.crear(UserCharIndex, luzBackup)
        Call ME_Tools.ejecutarComando(comando)
    Else ' La agregamos
        Call comando.crear(UserCharIndex, luzSeleccionada)
        Call ME_Tools.ejecutarComando(comando)
    End If
End Sub

Private Sub cmdProbarClima_Click()
Dim seleccionado As Tipos_Clima
Dim puede As Boolean
If Me.cmbClimaActual.listIndex >= 0 Then
    seleccionado = ME_Climas.obtenerTipoClima(Me.cmbClimaActual.list(Me.cmbClimaActual.listIndex))
    
    puede = True
    
    If seleccionado = ClimaLluvia And Not mapinfo.puede_lluvia Then
        puede = False
    ElseIf seleccionado = ClimaNeblina And Not mapinfo.puede_neblina Then
        puede = False
    ElseIf seleccionado = ClimaNiebla And Not mapinfo.puede_niebla Then
        puede = False
    ElseIf seleccionado = ClimaNieve And Not mapinfo.puede_nieve Then
        puede = False
    ElseIf seleccionado = ClimaNublado And Not mapinfo.puede_nublado Then
        puede = False
    ElseIf seleccionado = ClimaTormenta_de_arena And Not mapinfo.puede_sandstorm Then
        puede = False
    End If
    
    If puede Then
        Call Cambiar_estado_climatico(seleccionado)
    Else
        GUI_Alert "Este mapa no admite el clima seleccionado. Ve al menú Mapa -> Propiedades y estabelce los climas que se permiten en este mapa.", "Error"
    End If
End If


End Sub

Private Sub cmdPropiedadesMapa_Click()


    'Para que no piensen que se tildo
    frmCargando.Show , Me
    DoEvents
    
    'Cargamos el formulario
    load frmEditorGenerico
    
    'Lo configuramos
    frmEditorGenerico.ITEM_SENALADOR = "NOMBRE"
    frmEditorGenerico.ITEM_TIPO = "mapa"
    frmEditorGenerico.ITEM_TIPO_PLURAL = "mapas"
    frmEditorGenerico.ITEM_VERSIONADO = "PROPIEDAD_MAPA"
    frmEditorGenerico.caption = "Propiedades del Mapa"
    
    'Iniciamos
    Call frmEditorGenerico.iniciar(ME_Mapas.mapDataConfig)
    
    'Cargamos el archivo de informacion
    frmEditorGenerico.showFile ME_Mapas.mapDataFile
    
    'Lo mostramos
    Call modPosicionarFormulario.posicionarAbajoCentro(Me, frmEditorGenerico)
    
    'Si estamos en un mapa con un número definido, seteamos ese como default
    If THIS_MAPA.numero > 0 Then
        Call frmEditorGenerico.seleccionar(THIS_MAPA.numero)
    Else
        MsgBox "El mapa que estas modificando no tiene ningún número de mapa. Si un mapa no tiene número no se le pueden establecer sus propiedades. Para establecer que número de mapa es andá al menú Mapa -> Definir número.", vbExclamation, Me.caption
    End If
    
    frmCargando.Hide
    'Mostramos el formulario
    frmEditorGenerico.Show vbModal, frmMain
    
    'Vemos si se modifo alguna propiedad que tengo que mostrar en vivo
    If THIS_MAPA.numero > 0 Then
    
        Call ME_Mapas.cargarInformacionDeMapa(THIS_MAPA.numero, mapinfo)
        
        THIS_MAPA.nombre = mapinfo.Name
        
        Call act_titulo
    End If
 
End Sub

Private Sub cmdSolapas_Click(Index As Integer)
    ME_Tools.deseleccionarTool
    Call activarUltimaHerramientaCorrespondienteASolapa(Index)
End Sub

Private Sub cmdSeleccionarAreaAguaTierra_Click()
    If Me.cmdSeleccionarAreaAguaTierra.caption = "Cancelar" Then
        'Restablezco botones
        Me.cmdSeleccionarAreaAguaTierra.caption = "Seleccionar"
        
        Call Me_Tools_TileSet.EsconderVentanaTilesets
        
        Me.tilesets_area_sel_agua.Enabled = False
        Me.tilesets_area_sel_agua.Visible = False
    Else
        'Nos aseguramos la seleccion de la herramienta
        Call select_tool(Tools.tool_tileset)

        Area_Tileset.arriba = mapinfo.agua_rect.top
        Area_Tileset.abajo = mapinfo.agua_rect.bottom
        Area_Tileset.derecha = mapinfo.agua_rect.right
        Area_Tileset.izquierda = mapinfo.agua_rect.left
        
        tileset_actual = mapinfo.agua_tileset
        'Muestro la ventana
        Call Me_Tools_TileSet.MostrarVentanaTilesets(tileset_actual, tileset_actual_virtual)
        'Juego con los botones
        Me.cmdSeleccionarAreaAguaTierra.caption = "Cancelar"
        Me.tilesets_area_sel_agua.Enabled = True
        Me.tilesets_area_sel_agua.Visible = True
    End If
End Sub

Private Sub actualizarEntidadAccionSeleccionada()
    Dim idEntidad As Integer
    
    Dim accion As iAccionEditor
    Dim idAccion As Integer
    
    idAccion = Me.lstConBuscadorAcciones.obtenerIDValor
    If idAccion > 0 Then
        Set accion = ME_modAccionEditor.obtenerAccionID(idAccion)
    Else
        Set accion = Nothing
    End If
        
    idEntidad = Me.lstEntidades.obtenerIDValor
    
    Call Me_Tools_Entidades.seleccionarEntidad(idEntidad, accion, 0)
End Sub

Private Sub cmdTriggers_AplicarATodo_Click()
    Dim respuesta As VbMsgBoxResult
    Dim Trigger As Long
    
    Trigger = ME_Tools_Triggers.calcular_trigger_lista(frmMain.lstTriggers)
    
    If Trigger > 0 Then
        respuesta = MsgBox("¡¡CUIDADO!! ¿Queres aplicarle el trigger seleccionado a todo el mapa?", vbYesNo + vbExclamation, "Aplicar trigger")
    Else
        respuesta = MsgBox("¡¡CUIDADO!! ¿Queres borrar todos los triggers que hay en el mapa?", vbYesNo + vbExclamation, "Borrar triggers")
    End If
    
    If respuesta = vbYes Then
        
        'Seteamos con lo que vamos a trbajar
        Call ME_Tools_Triggers.establecerTrigger(Trigger)
        
        'Activamos la herramienta
        If Trigger > 0 Then
            Call ME_Tools_Triggers.click_InsertarTrigger
            Call selectToolMultiple(tool_triggers, "Insertar Trigger en todo el mapa")
        Else
            Call ME_Tools_Triggers.click_BorrarTrigger
            Call selectToolMultiple(tool_triggers, "Borrar Triggers en todo el mapa")
        End If
        
        'Seleccionamos el area donde vamos a trabajar
        Call modSeleccionArea.puntoArea(ME_Tools.areaSeleccionada, X_MINIMO_USABLE, Y_MINIMO_USABLE)
        Call modSeleccionArea.actualizarArea(ME_Tools.areaSeleccionada, X_MAXIMO_USABLE, Y_MAXIMO_USABLE)
        
        'Ejecutamos la acción sobre esa area
        click_tool vbLeftButton
        
        Call activarUltimaHerramientaCorrespondienteASolapa(eSolapasEditor.Triggers)
        
        Call ME_Tools.seleccionarTool(Nothing, tool_triggers)
    End If
End Sub

Private Sub acomodarElementos()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    ' Barra de Herramientas
    Me.SSTab1.top = (Me.ScaleHeight - Me.SSTab1.Height) + 34
        
    ' Barra de Botones
    Me.frmBotonera.top = Me.SSTab1.top

    If modPantalla.mostrarBarraHerramientas Then
        ' Máximo tamaño que puede ocupar el render
        Me.SSTab1.Visible = True
        Me.frmBotonera.Visible = True
    Else
        Me.SSTab1.Visible = mostrandoBarra
        Me.frmBotonera.Visible = mostrandoBarra
    End If
   
    Me.renderer.Width = Me.ScaleWidth
    Me.renderer.Height = Me.ScaleHeight

  ' Maximo tamaño que puede ocupar cada tile
    modPantalla.PixelesPorTile.y = renderer.Height / modPantalla.TilesPantalla.y
    modPantalla.PixelesPorTile.x = renderer.Width / modPantalla.TilesPantalla.x

    Me.caption = "Tamaño de cada tile " & modPantalla.PixelesPorTile.x & " * " & modPantalla.PixelesPorTile.y
    Me.lblTamanoTile.caption = "Tamaño de cada tile " & modPantalla.PixelesPorTile.x & " * " & modPantalla.PixelesPorTile.y
    
End Sub

Private Sub cmdPegarPersonaje_Click()
    Call Char_Start_Anim(UserCharIndex)
End Sub
Public Sub Form_Resize()
    Call acomodarElementos
End Sub

Private Sub ListaConBuscadorPresets_Change(valor As String, ID As Integer)

    Me_Tools_Presets.idPresetSeleccionado = ID

    Call Me_Tools_Presets.click_insertarPreset
End Sub

Private Sub ListaConBuscadorPresets_DblClic()
    Dim nuevoNombre As String
    Dim idpreset As Integer
    
    If ListaConBuscadorPresets.obtenerIDValor > 0 Then
        idpreset = ListaConBuscadorPresets.obtenerIDValor
        nuevoNombre = InputBox("Ingrese un nuevo nombre para el elemento presefinido '" & ListaConBuscadorPresets.obtenerValor & "'.", "Cambiar nombre a elemento predefinido.")
        
        If Len(nuevoNombre) > 0 Then
            Call ME_presets.cambiarNombrePreset(idpreset, nuevoNombre)
            Call Me_Tools_Presets.actualizarPresetEnListaUltimosUsados(idpreset)
            Call ME_presets.cargarListaPresets
        End If
    End If
    
End Sub

Private Sub lstConBuscadorAcciones_Change(valor As String, ID As Integer)
    Call actualizarEntidadAccionSeleccionada
End Sub

Private Sub lstEntidades_Change(valor As String, ID As Integer)
    Call actualizarEntidadAccionSeleccionada
End Sub

Private Sub lstPixelShader_Viejos_Click()
    Me.txtShader = Replace$(Me.lstPixelShader_Viejos.list(Me.lstPixelShader_Viejos.listIndex), ";", ";" & vbNewLine)
End Sub

Private Sub lstPixelShader_Viejos_DblClick()
    Dim shader As String
    
    shader = Replace$(Me.lstPixelShader_Viejos.list(Me.lstPixelShader_Viejos.listIndex), ";", ";" & vbNewLine)
    
    PixelShaderBump = CreateShaderFromCode(shader)
End Sub

Private Sub lstUltimosGraficosUsados_Click()
    
    graficoSeleccionadoIndex = val(lstUltimosGraficosUsados.list(lstUltimosGraficosUsados.listIndex))
    
    If graficoSeleccionadoIndex > 0 Then
        capaSeleccionada = activarCapasDeGrafico(capaSeleccionada, graficoSeleccionadoIndex)
            
        If capaSeleccionada > 0 Then
            Call ME_Tools_Graficos.establecerInfoGrhCapa(graficoSeleccionadoIndex, capaSeleccionada)
        End If
    End If
    
End Sub

Private Sub lstUltimosPredefinidosUtilizados_Click()

    Me_Tools_Presets.idPresetSeleccionado = val(Me.lstUltimosPredefinidosUtilizados.list(Me.lstUltimosPredefinidosUtilizados.listIndex))

    Call Me_Tools_Presets.click_insertarPreset
End Sub


Private Sub lstZonaNacimientoCriaturas_ItemCheck(item As Integer)
    Dim loopElemento As Byte
    
    If item = 0 Then
        criaturaSeleccion_Zona = 0
        lstZonaNacimientoCriaturas.Selected(0) = True
        
        For loopElemento = 1 To lstZonaNacimientoCriaturas.ListCount - 1
            lstZonaNacimientoCriaturas.Selected(loopElemento) = False
        Next
    Else
        lstZonaNacimientoCriaturas.Selected(loopElemento) = 0
        criaturaSeleccion_Zona = Me_Tools_Npc.calcularZonaLista(Me.lstZonaNacimientoCriaturas)
    End If
    
    If criaturaSeleccion_Index > 0 Then
        Call Me_Tools_Npc.seleccionarIndexNPC(criaturaSeleccion_Index, criaturaSeleccion_Zona)
    End If
End Sub

Private Sub lstZonaNacimientoCriaturas_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim resultado As VbMsgBoxResult
    Dim idSeleccionado As Byte
    Dim nombre As String
    
    If KeyCode = vbKeyDelete Then
       
        If Me.lstZonaNacimientoCriaturas.listIndex = -1 Then Exit Sub
        
        nombre = Me.lstZonaNacimientoCriaturas.list(Me.lstZonaNacimientoCriaturas.listIndex)
        
        idSeleccionado = CByte(val(mid$(nombre, 1, InStr(1, nombre, " ", vbTextCompare)))) - 1
        
        resultado = MsgBox("¿Estás seguro que queres borrar la zona de nacimiento '" & mapinfo.ZonasNacCriaturas(idSeleccionado).nombre & "'?. Se le quitará esta zona de nacimiento a las criaturas que la tengan. Si queres modificar está zona de nacimiento, seleccioná el area y pulsa CONTROL + N y ponelé este nombre.", vbExclamation + vbYesNo, "Zona de Nacimiento de Criaturas")
    
        If resultado = vbYes Then
        
        End If
        
    End If
    
End Sub

Private Sub luces_color_Click()
    On Error GoTo BotonCancelar:
    ColorDialog.CancelError = True 'Se produce un error si se toca cancelar
    ColorDialog.flags = cdlCCRGBInit
    ColorDialog.Color = luces_color.BackColor
    ColorDialog.ShowColor
        
    luces_color.BackColor = ColorDialog.Color
    VBC2RGBC luces_color.BackColor, luzSeleccionada.LuzColor
    Call Me_Tools_Luces.seleccionarLuz(luzSeleccionada)
    Exit Sub
BotonCancelar:
        Err.Clear
End Sub

Private Sub UpdateScrollsHora()

    luzSeleccionada.luzInicio = horaInicioLuz.value
    luzSeleccionada.luzFin = horaFinLuz.value
    
    lblInicio00.caption = "Inicio: " & obtener_hora_fraccion(horaInicioLuz.value)
    lblFin00.caption = "Fin: " & obtener_hora_fraccion(horaFinLuz.value)
    
    Call Me_Tools_Luces.seleccionarLuz(luzSeleccionada)
End Sub

Private Sub horaFinLuz_Change()
    UpdateScrollsHora
End Sub

Private Sub horaFinLuz_Scroll()
    UpdateScrollsHora
End Sub

Private Sub horaInicioLuz_Change()
    UpdateScrollsHora
End Sub

Private Sub horaInicioLuz_Scroll()
    UpdateScrollsHora
End Sub

Private Sub forzar_dia_c_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call mnuTodoDia_Click
End Sub

Private Sub hora_scroll_Change()
    lblHora.caption = "Hora: " & obtener_hora_fraccion(hora_scroll.value) & "hs."
End Sub

Private Sub hora_scroll_Scroll()
    Call hora_scroll_Change
End Sub

Private Sub ActualizarEstadoHerramientaLuz()
    luzSeleccionada.LuzTipo = 0
    luzSeleccionada.LuzTipo = luzSeleccionada.LuzTipo Or IIf(Me.chkAnimacionFuego, TipoLuces.Luz_Fuego, 0)
    luzSeleccionada.LuzTipo = luzSeleccionada.LuzTipo Or IIf(Me.chkLuzCuadrada, TipoLuces.Luz_Cuadrada, 0)
    luzSeleccionada.LuzTipo = luzSeleccionada.LuzTipo Or IIf(Me.chkUtilizarBrillo, TipoLuces.Luz_Normal, 0)
    
    Call Me_Tools_Luces.seleccionarLuz(luzSeleccionada)
End Sub

Private Sub luces_radio_Change()

    luzSeleccionada.LuzRadio = luces_radio.value
    
    luces_radio_label.caption = "Radio: " & luzSeleccionada.LuzRadio
    
    Call Me_Tools_Luces.seleccionarLuz(luzSeleccionada)
End Sub

Private Sub luces_radio_Scroll()
    luces_radio_Change
End Sub



Private Sub optimizar_luces_Click()
    Pre_Render_Lights
End Sub
'*****************************************************************************/
' Objetos
Private Sub cmdBorrarObjeto_Click()
    Call Me_Tools_Objetos.click_BorrarOBJ
End Sub
'*****************************************************************************/
' Pre sets
Private Sub cmdEliminarPreset_Click()
    Dim resultado As VbMsgBoxResult
    Dim idpreset As Integer
    
    If Me.ListaConBuscadorPresets.obtenerIDValor > 0 Then
            
        idpreset = Me.ListaConBuscadorPresets.obtenerIDValor
        
        resultado = MsgBox("¿Estas seguro que queres eliminar de la base de datos el elemento predefinido '" & Me.ListaConBuscadorPresets.obtenerValor & "'?.", vbYesNo Or vbExclamation)
    
        If resultado = vbYes Then
            'Elimino de la base de datos
            Call Me_indexar_Predefinidos.eliminar(idpreset)
            ' Elimino de la lista de ultimos usando
            Call Me_Tools_Presets.elimimarDeListaUltimosUsados(idpreset)
            ' Recargo la lista de presets
            Call ME_presets.cargarListaPresets
        End If
    Else
        Call MsgBox("Debe seleccionar un preset para eliminar.", vbExclamation + vbOKOnly)
    End If
    
End Sub
'*****************************************************************************/
'Acciones
Private Sub cmdEliminarAccion_Click()

Dim idSeleccionado As Integer
Dim accion As iAccionEditor
Dim estaSeguro As VbMsgBoxResult

If Me.listTileAccionActuales.listIndex = -1 Then
    MsgBox "Debe seleccionar alguna acción que se este utilizando actualmente."
Else
    idSeleccionado = val(mid$(Me.listTileAccionActuales.list(Me.listTileAccionActuales.listIndex), 1, InStr(1, Me.listTileAccionActuales.list(Me.listTileAccionActuales.listIndex), ")") - 1))
    
    Set accion = ME_modAccionEditor.obtenerAccionID(idSeleccionado)
    
    If esAccionnUsada(accion) Then
      estaSeguro = MsgBox("La acción " & accion.GetNombre & " se esta utilizando actualmente en este mapa. Si la eliminas se borrará de los lugares donde esta insertada. ¿Estas seguro que deseas borrar esta acción?", vbYesNo + vbExclamation)
      
        If estaSeguro = vbYes Then
            Call ME_Tools_Acciones.seleccionarAccion(Nothing)
            Call ME_modAccionEditor.eliminarAccionMapa(accion)
            Call ME_modAccionEditor.eliminarAccion(idSeleccionado)
            Call Me.refrescarListaUsando
        End If
    End If
End If

End Sub

Private Sub cmdBorrarAccion_Click()
    Call ME_Tools_Acciones.click_InsertarBorrarAccion
End Sub

Private Sub cmdInsertarAccion_Click()
    If Me.listTileAccionActuales.listIndex = -1 Or Me.listTileAccionActuales.ListCount = 0 Then
        MsgBox "Debe seleccionar alguna acción que se este utilizando actualmente."
    Else
        Dim accion As iAccionEditor
        Dim idAccion As Integer
    
        idAccion = val(left$(Me.listTileAccionActuales.list(Me.listTileAccionActuales.listIndex), InStr(1, Me.listTileAccionActuales.list(Me.listTileAccionActuales.listIndex), ")") - 1))
        Set accion = ME_modAccionEditor.obtenerAccionID(idAccion)
    
        Call ME_Tools_Acciones.seleccionarAccion(accion)
        Call ME_Tools_Acciones.click_InsertarAccion
    End If
End Sub

'*********************************************************
' Bloqueos
Private Sub cmdBorrarBloqueo_Click()
    Call ME_Tools_Triggers.click_BorrarBloqueo
End Sub
Private Sub cmdInsertarBloqueo_Click()
    Call ME_Tools_Triggers.click_InsertarBloqueo
End Sub

Private Sub cmdInsertarDobleBloqueo_Click()
    Call ME_Tools_Triggers.click_InsertarDobleBloqueo
End Sub

'*********************************************************
' Graficos
Private Sub cmdInsertarGrafico_Click()

    If capaSeleccionada = 0 Then
        MsgBox "Debe seleccionar una capa donde se va a insertar el grafico."
        Exit Sub
    End If
    
    If graficoSeleccionadoIndex = 0 Then
        MsgBox "Debe seleccionar una capa donde se va a insertar el grafico."
        Exit Sub
    End If
    
    Call ME_Tools_Graficos.click_InsertarGrafico
End Sub

Private Sub cmdBorrarGrafico_Click()
    Dim loopCapa As Byte
    
    ' Activamos todas las capas
    For loopCapa = 1 To CANTIDAD_CAPAS
        Me.grh_capa(loopCapa).Enabled = True
    Next

    Call ME_Tools_Graficos.click_BorrarGrafico
End Sub
'*********************************************************
' Criaturas NPCS
Private Sub cmdInsertarNpc_Click()
    Call Me_Tools_Npc.click_InsertarNPC
End Sub

Private Sub cmdBorrarNpc_Click()
    Call Me_Tools_Npc.click_BorrarNPC
End Sub
'*********************************************************
'Triggers
Private Sub cmdInsertarTrigger_Click()
  Call ME_Tools_Triggers.click_InsertarTrigger
End Sub
Private Sub cmdBorrarTrigger_Click()
    Call ME_Tools_Triggers.click_BorrarTrigger
End Sub

Private Sub cmdResetearListaTriggers_Click()
    Dim loopC As Byte
    
    For loopC = 0 To frmMain.lstTriggers.ListCount - 1
        frmMain.lstTriggers.Selected(loopC) = False
    Next loopC
End Sub
'*********************************************************
Private Sub ListaConBuscadorGraficos_Change(valor As String, ID As Integer)
     
    graficoSeleccionadoIndex = ID
    
    capaSeleccionada = activarCapasDeGrafico(capaSeleccionada, graficoSeleccionadoIndex)
           
    If capaSeleccionada > 0 Then
        Call ME_Tools_Graficos.establecerInfoGrhCapa(graficoSeleccionadoIndex, capaSeleccionada)
    End If
End Sub

Private Function activarCapasDeGrafico(capaActual As Byte, ID As Integer) As Byte

    Dim loopCapa As Byte
    Dim capasAplica As Byte
    'Si la capa seleccionada es imcompatible, la desactivo
    ' Inhabilito o habilito las capas segun donde pueda ponerse este grafico
    'Si hay una sola capa habilitada, la habilito automaticamente
    
    For loopCapa = 1 To CANTIDAD_CAPAS
        If GrhData(ID).Capa(loopCapa) Then capasAplica = capasAplica + 1
    Next
    
    ' Si la que estaba seleccionada no es compatible, al desactivo
    If capaSeleccionada > 0 Then
        If GrhData(ID).Capa(capaSeleccionada) = False Then
            Me.grh_capa(capaSeleccionada).value = False
            activarCapasDeGrafico = 0
        Else
            'Si es compatible, la selecciono
            activarCapasDeGrafico = capaSeleccionada
        End If
    End If
    
    ' Activo todas las posibles
    For loopCapa = 1 To CANTIDAD_CAPAS
        Me.grh_capa(loopCapa).Enabled = GrhData(ID).Capa(loopCapa)
        ' Si solo se puede activar en una capa, la activo
        If capasAplica = 1 And GrhData(ID).Capa(loopCapa) Then
            Me.grh_capa(loopCapa).value = True
            activarCapasDeGrafico = loopCapa
        End If
    Next
    
End Function
Private Sub ListaConBuscadorGraficos_GotFocus()
    Dim i As Byte
    
    For i = 0 To frmMain.lstGraficosCopiados.ListCount - 1
        frmMain.lstGraficosCopiados.Selected(i) = False
    Next
End Sub

Private Sub ListaConBuscadorNpcs_Change(valor As String, ID As Integer)
    
    criaturaSeleccion_Index = ID
    
    'Si la zona es 0, igual es valida
    Call Me_Tools_Npc.seleccionarIndexNPC(criaturaSeleccion_Index, criaturaSeleccion_Zona)
End Sub

Private Sub ListaConBuscadorObjetos_Change(valor As String, ID As Integer)
    Call Me_Tools_Objetos.seleccionarIndexObjeto(ID)
End Sub

Private Sub listTileAccionActuales_dblClick()
    Dim accion As iAccionEditor
    Dim idAccion As Integer
    
    idAccion = val(left$(Me.listTileAccionActuales.list(Me.listTileAccionActuales.listIndex), InStr(1, Me.listTileAccionActuales.list(Me.listTileAccionActuales.listIndex), ")") - 1))
    Set accion = ME_modAccionEditor.obtenerAccionID(idAccion)
    
    Call frmEditorAccion.Cargar(accion)
    Call frmEditorAccion.Show(vbModal, Me)
    Call ME_modAccionEditor.refrescarListaUsando(Me.listTileAccionActuales)
End Sub

Private Sub listTipoAccionesDisponibles_Click()

   Unload frmModificarAccion
   
   Dim nuevaAccionPadre As New cAccionCompuestaEditor
   Dim nuevaAccionHijo As New cAccionTileEditor
   Dim nombreNuevaAccion As String
   
   Set nuevaAccionHijo = listaAccionTileEditor.item(Me.listTipoAccionesDisponibles.listIndex + 1).Clonar
   
   Call frmModificarAccion.Cargar(nuevaAccionHijo)
   
   If frmModificarAccion.edicion(Me) Then
        nombreNuevaAccion = InputBox("Ingrese un nombre descriptivo para esta accion.", "Editor de Acciones")
        If Len(nombreNuevaAccion) > 0 Then
            Call nuevaAccionPadre.iAccionEditor_crear(nombreNuevaAccion, "")
            Call nuevaAccionPadre.agregarHijo(nuevaAccionHijo)
            Call ME_modAccionEditor.agregarNuevaAccion(nuevaAccionPadre)
            Call refrescarListaUsando
        End If
   End If
   
End Sub

Public Sub refrescarListaDisponibles(lista As ListBox)
    Call ME_modAccionEditor.refrescarListaDisponibles(Me.listTipoAccionesDisponibles)
End Sub

Public Sub refrescarListaUsando()
     Call ME_modAccionEditor.refrescarListaUsando(Me.listTileAccionActuales)
     
    'Acciones
    Dim aux As iAccionEditor

    Call Me.lstConBuscadorAcciones.vaciar
    Call Me.lstConBuscadorAcciones.addString(0, "0 - Ninguna")

    For Each aux In ME_modAccionEditor.listaAccionTileEditorUsando
     Call Me.lstConBuscadorAcciones.addString(aux.getID, aux.getID & " - " & aux.GetNombre)
    Next
End Sub
'/*************************************************************************/

Private Sub cmdIniciarMusicaAmbiente_Click()


If Me.cmdIniciarMusicaAmbiente.caption = "Iniciar Música Ambiente" Then
    
    If mapinfo.Music > 0 Then
        Call Engine_Sonido_Extend.Sonido_Play_Ambiente(mapinfo.Music)
        Me.cmdIniciarMusicaAmbiente.caption = "Parar Música Ambiente"
    Else
        GUI_Alert "El mapa no tiene una música ambiental establecida. Ve al menú Mapa -> Propiedades y estabelce una.", "Error"
    End If
Else
    Call Engine_Sonido_Extend.Sonido_Stop_Ambiente(mapinfo.Music)
    
    Me.cmdIniciarMusicaAmbiente.caption = "Iniciar Música Ambiente"
End If

End Sub
Private Sub Commandasd_Click()
Pre_Render_Lights
End Sub

Private Sub Commanddf_Click()
'FX_Projectile_Create_pos UserCharIndex, MouseTileX, MouseTileY, 2, 0.05
Dim i As Integer

i = Engine_Entidades.Entidades_Crear_Indexada(UserPos.x, UserPos.y, Rnd * 255, EntidadesIndexadas(2))
Engine_Entidades.Entidades_SetDestino i, MouseTileX * 32, MouseTileY * 32

End Sub

Private Sub grh_capa_Click(Index As Integer)
    capaSeleccionada = Index
    Call ME_Tools_Graficos.establecerInfoGrhCapa(graficoSeleccionadoIndex, capaSeleccionada)
End Sub

Private Sub Command4_Click()
    GUI_Alert "Con el click " & Chr$(255) & "derecho" & Chr$(255) & " del mouse se obtiene el trigger de una posición copiarlo."
End Sub


'Private Sub Command6_Click()
'    frmMain.asd.Picture = frmMain.asd.Image
'    frmMain.asd.Refresh
'    SavePicture frmMain.asd.Picture, Clientpath & CurMap & "h.bmp"
'End Sub

Private Sub Command7_Click()

Dim y As Integer, x As Integer
Dim j As Integer

For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
    For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
        With hMapData(x, y)
            For j = 0 To 3
                .hs(j) = 0
            Next j
            .h = 0
        End With
    Next x
Next y
End Sub

Private Sub ctriggers_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ver_triggers_Click
ActChecks
End Sub


Public Sub activarUltimaHerramientaCorrespondienteASolapa(solapaSeleccionada As Integer)



    Me.cmdSolapas(ultimaSolapaseleccionada).FontBold = False
    Me.cmdSolapas(solapaSeleccionada).FontBold = True
    Me.SSTab1.Tab = solapaSeleccionada
    ultimaSolapaseleccionada = solapaSeleccionada

    
    If (ME_Tools.TOOL_SELECC And Tools.tool_copiar) Or (ME_Tools.TOOL_SELECC And ME_Tools.TOOL_SELECC) Then
      '  Call restablecerherrmaientas
        Exit Sub
    End If


    Select Case ultimaSolapaseleccionada
            
        Case eSolapasEditor.Graficos  'Graficos
                    
            Call select_tool(Tools.tool_grh)

            Call ME_Tools_Graficos.activarUltimaHerramienta
            
        Case eSolapasEditor.Herramientas 'Predefinidos
            
            Call Me_Tools_Presets.activarUltimaHerramientaPresets
        
        Case eSolapasEditor.Triggers 'Triggers
            
            Call select_tool(Tools.tool_triggers)
            
            Call ME_Tools_Triggers.activarUltimaHerramientaTriggers
            
        Case eSolapasEditor.Npcs 'Npcs
            
            Call select_tool(Tools.tool_npc)
                        
            Call Me_Tools_Npc.activarUltimaHerramientaNPC
        
        '/* Segunda linea /*
        
        Case eSolapasEditor.Montañas 'Montañas
            
            Call select_tool(Tools.tool_montaña)
               
        Case eSolapasEditor.Tilesets 'Tilesets
            
            select_tool Tools.tool_tileset
            
            TiempoBotonTileSetApretado = GetTickCount
            
            Call Me_Tools_TileSet.establecerAreaTileSet(tileset_actual, tileset_actual_virtual, Area_Tileset)
            
        Case eSolapasEditor.Bloqueos 'Bloqueos
            
            select_tool Tools.tool_bloqueo
            Call ME_Tools_Triggers.activarUltimaHerramientaBloqueo
        
        Case eSolapasEditor.Particulas 'Particulas
        
            ultimaSolapaseleccionada = eSolapasEditor.Particulas
            frmMain.SSTab1.Tab = eSolapasEditor.Particulas
            select_tool Tools.tool_particles
            
        '/* Tercera linea /*
        
        Case eSolapasEditor.Acciones 'Acciones
                        
            Call select_tool(Tools.tool_grh)
            Call ME_Tools_Acciones.activarUltimaHerramientaAcciones
        
        Case eSolapasEditor.Luces 'Luces
                        
            Call select_tool(Tools.tool_luces)
            Call Me_Tools_Luces.activarUltimaHerramientaLuces
            
        Case eSolapasEditor.objetos 'Objetos
                    
            Call select_tool(Tools.tool_obj)
            Call Me_Tools_Objetos.activarUltimaHerramientaObjeto
        
        Case eSolapasEditor.Entidades 'Entidades
            
            Call select_tool(Tools.tool_entidades)
            Call Me_Tools_Entidades.activarUltimaHerramienta
        
        Case eSolapasEditor.Mapa  'Mapa
        
            Call select_tool(Tools.tool_none)
            
    End Select

    
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

clicX = x
    clicY = y
End Sub


Private Sub hlpTilesets_Click()
GUI_Alert "Manteniendo presionada la tecla " & Chr$(255) & "G" & Chr$(255) & " se despliega el menú de pisos donde se puede elegir el piso a insertar en el mapa.", "Información sobre TileSets"
End Sub



Private Sub LabelCol_Click()
    On Error GoTo BotonCancelar:
    
    ColorDialog.flags = cdlCCRGBInit
    ColorDialog.Color = LabelCol.BackColor
    ColorDialog.ShowColor
    LabelCol.BackColor = ColorDialog.Color
    VBC2RGBC LabelCol.BackColor, mapinfo.BaseColor
    
    Exit Sub
BotonCancelar:
        Err.Clear
End Sub

Private Sub linesolid_Click()
    If chkMostrarLineas.value = vbChecked Then
    D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_WIREFRAME
    ''lColorMod = D3DTOP_MODULATE Or D3DTOP_MODULATE2X
    Else
    D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
    ''lColorMod = D3DTOP_MODULATE
    End If
End Sub


Private Sub lstGraficosCopiados_MouseUp(Button As Integer, _
                                        Shift As Integer, _
                                        x As Single, _
                                        y As Single)

    Call frmMain.ListaConBuscadorGraficos.deseleccionar
    
    Dim loopCapa      As Byte
    Dim infoCapa      As String
    Dim capas(1 To CANTIDAD_CAPAS) As ME_Tools_Graficos.tCapasPosicion

    For loopCapa = 0 To Me.lstGraficosCopiados.ListCount - 1

        If Me.lstGraficosCopiados.Selected(loopCapa) Then
        
            infoCapa = Me.lstGraficosCopiados.list(loopCapa)
            
            If InStr(1, infoCapa, "(") > 0 And InStr(1, infoCapa, ")") > 0 Then
                capas(loopCapa + 1).GrhIndex = val(mid$(infoCapa, InStr(1, infoCapa, "(") + 1, (InStr(1, infoCapa, ")") - InStr(1, infoCapa, "(")) - 1))
            End If

            capas(loopCapa + 1).seleccionado = True
        Else
            capas(loopCapa + 1).GrhIndex = 0
            capas(loopCapa + 1).seleccionado = False
        End If

    Next

    Call ME_Tools_Graficos.establecerInfoGrhPosicion(capas)
    
    
End Sub


Private Sub lstTriggers_Click()
lblDescTrigger.caption = "Descripción: " & obtenerDescripcion(lstTriggers.listIndex)
Call ME_Tools_Triggers.establecerTrigger(ME_Tools_Triggers.calcular_trigger_lista(frmMain.lstTriggers))
End Sub

Private Sub luz_luminosidad_Change()
    Call actualizarLuminosidadBrillo
End Sub

Private Sub actualizarLuminosidadBrillo()
    luz_luminosidad_lbl.caption = "Brillo: " & Round((255 - luz_luminosidad.value) / 2.55, 1) & "%"
    
    luzSeleccionada.LuzBrillo = luz_luminosidad.value
    Call Me_Tools_Luces.seleccionarLuz(luzSeleccionada)
End Sub

Private Sub luz_luminosidad_Scroll()
    Call actualizarLuminosidadBrillo
End Sub

Private Sub mnu_CDMCompartirNovedades_Click()
    frmCDMCommitearEditor.Show , Me
End Sub

Private Sub mnuEquipo_Click()
    
    #If Produccion = 2 Then
        
        Me.mnu_CDMCompartirNovedades.caption = "Compartir - No disponible en Pre-Produccion"
        Me.mnu_CDMCompartirNovedades.Enabled = False
        
        Me.mnu_CDMObtenerNovedades.caption = "Obtener Novedades - No disponible en Pre-Produccion"
        Me.mnu_CDMObtenerNovedades.Enabled = False
        
    #End If
End Sub

Private Sub mnu_CDMObtenerNovedades_Click()
     frmCDMUpdateEditor.Show , Me
End Sub

Private Sub mnu_CDMPublicarServidor_Click()
    frmCDMPublicarServidor.Show , Me
End Sub

Private Sub mnu_paker_Click()
    frmConfigurarRecursos.Show , Me
End Sub

Private Sub mnu_ventana_indexar_Click()
    load frmConfigurarGraficos
   
    frmConfigurarGraficos.Show , Me

    Call modPosicionarFormulario.posicionarAbajoDerecha(Me, frmConfigurarGraficos)
End Sub

Private Sub mnuabrir_Click()
If THIS_MAPA.editado Then
    Dim r As Integer
    r = MsgBox("Los cambios se perderán, deseas guardar el mapa antes de abrir otro?", vbExclamation + vbYesNoCancel)
    If r = vbCancel Then Exit Sub
    If r = vbYes Then
        If THIS_MAPA.Path <> "" Then
            GuardarMapa (THIS_MAPA.Path)
        Else
            GuardarMapaComo
        End If
    End If
End If
ShowAbrirMapa
act_titulo
End Sub

Private Sub mnuBorrarSeleccion_Click()
    'Eliminar
    Dim tool As Long
    
    tool = TOOL_SELECC
    
    ignorarMouseUp = True
    
    Call Me_Tools_Seleccion.eliminar(areaSeleccionada, trabajandoConElementos)
    
    Call ME_Tools.seleccionarTool(Nothing, tool)
    Call activarUltimaHerramientaCorrespondienteASolapa(CInt(ultimaSolapaseleccionada))
    
    
End Sub

Private Sub mnuCambiosPendientes_Click()
     frmCambiosPendentes.Show , Me
End Sub

Private Sub mnuConfigAspectos_Click()
    Call lanzarGenerico("pixel.json", "pixels.dat", "aspecto", "aspectos", "NOMBRE", "ASPECTO", "Configuración de Aspectos")
    
    ' Recargamos los pixels
    Call Me_indexar_Pixels.cargarDesdeIni
End Sub

Private Sub mnuConfigCriaturas_Click()
    Call lanzarGenerico("criatura.json", "npcs.dat", "criatura", "criaturas", "NAME", "CRIATURA", "Configuración de las Criaturas")

    'Recargamos los objetos
    Call ME_obj_npc.cargarInformacionNPCs
    Call ME_obj_npc.cargarListaNPC
End Sub

Private Sub mnuConfigEfectos_Click()
    load frmConfigurarEfectos
   
    frmConfigurarEfectos.Show , Me

    Call modPosicionarFormulario.posicionarAbajoCentro(Me, frmConfigurarEfectos)
End Sub

Private Sub mnuConfigEntidades_Click()
    load frmConfigurarEntidades
   
    frmConfigurarEntidades.Show , Me

    Call modPosicionarFormulario.posicionarAbajoCentro(Me, frmConfigurarEntidades)
End Sub

Private Sub mnuConfigHechzos_Click()
    Call lanzarGenerico("hechizo.json", "hechizos.dat", "hechizo", "hechizos", "NOMBRE", "HECHIZO", "Configuración de los hechizos")
    
    'Luego de que cierra...
    Call Me_Hechizos.cargarInformacionHechizos
End Sub
Private Sub lanzarGenerico(archivoJSON As String, archivoDatos As String, Tipo As String, tipoPlural As String, senalador As String, tipoVersionado As String, descripcion As String)
    'Cargamos el archivo de configuracion
    Dim objDataConfig As cFileJSON
    Dim objDataFile As cFileINI
    
    'Para que no piensen que se tildo
    frmCargando.Show , Me
    DoEvents
    
    Set objDataConfig = New cFileJSON
    objDataConfig.init DBPath & "\JSON\" & archivoJSON
    
    'Cargamos el archivo donde esta la info de los objetos
    Set objDataFile = New cFileINI
    objDataFile.load DBPath & "\" & archivoDatos, objDataConfig
         
    'Cargamos el formulario
    load frmEditorGenerico
    
    'Lo configuramos
    frmEditorGenerico.ITEM_SENALADOR = senalador
    frmEditorGenerico.ITEM_TIPO = Tipo
    frmEditorGenerico.ITEM_TIPO_PLURAL = tipoPlural
    frmEditorGenerico.ITEM_VERSIONADO = tipoVersionado
    frmEditorGenerico.caption = descripcion
            
    'Iniciamos
    Call frmEditorGenerico.iniciar(objDataConfig)
            
    'Cargamos el archivo de informacion
    frmEditorGenerico.showFile objDataFile
    
    'Cartel de cargando
    frmCargando.Hide
    
    'Lo mostramos
    Call modPosicionarFormulario.posicionarAbajoCentro(Me, frmEditorGenerico)
        
    'Mostramos el formulario
    Call frmEditorGenerico.Show(vbModal, Me)
End Sub

Private Sub mnuConfigArmadaReal_Click()
    Call lanzarGenerico("rango_faccion.json", "armada.dat", "rango", "rangos", "NOMBRE", "ARMADA_RANGO", "Configuración de los rangos de la Armada Real")
End Sub

Private Sub mnuConfigLegiónOscura_Click()
    Call lanzarGenerico("rango_faccion.json", "legion.dat", "rango", "rangos", "NOMBRE", "LEGION_RANGO", "Configuración de los rangos de la Legión Oscura")
End Sub

Private Sub mnuConfigNiveles_Click()
    Call lanzarGenerico("nivel.json", "niveles.dat", "nivel", "niveles", "", "NIVEL", "Configuración de los niveles")
End Sub

Private Sub mnuConfigObjetos_Click()

    Call lanzarGenerico("objeto.json", "objetos.dat", "objeto", "objetos", "NAME", "OBJETO", "Configuración de los objetos")

    'Cuando se cierra el formulario
    'Recargamos los objetos
    Call ME_obj_npc.cargarInformacionObjetos
    Call ME_obj_npc.cargarListaObjetos
End Sub

Private Sub mnuConfigPersonajes_Click()
    load frmConfigurarPersonajes
   
    frmConfigurarPersonajes.Show , Me

    Call modPosicionarFormulario.posicionarAbajoCentro(Me, frmConfigurarPersonajes)
End Sub

Private Sub mnuConfigPisadas_Click()
    Call lanzarGenerico("pisadas.json", "pisadas.ini", "pisada", "pisadas", "NOMBRE", "PISADA", "Configuración de los Sonidos que se escuchan al pisar")
    
    Call Me_indexar_EfectosPisadas.cargarInformacionEfectosPisadas
End Sub

Private Sub mnuConfigSonidos_Click()
    load frmConfigurarSonidos

    frmConfigurarSonidos.Show , Me

    Call modPosicionarFormulario.posicionarAbajoCentro(Me, frmConfigurarSonidos)
End Sub

Private Sub mnuCopiar_Click()
    ' Chequeo que haya seleccionado algo
    
    If trabajandoConElementos = tool_none Then
        Call GUI_Alert("Tenes que ir al menú 'Edición -> Trabajar con...' y  seleccionar el tipo de elementos que queres copiar.", "No hay nada para copiar.")
        Exit Sub
    End If
    
    'Copiar
    Call Me_Tools_Seleccion.cargarAlPortapeles(areaSeleccionada)
    Call Me_Tools_Seleccion.copiar(1, trabajandoConElementos)
    
    Call refreshMenuCopiarPortapapeles
    
    ignorarMouseUp = True
End Sub

Private Sub refreshMenuCopiarPortapapeles()
    Dim loopP As Byte
    
    For loopP = 1 To UBound(portapapeles)
        If portapapeles(loopP).vacio = False Then
            Me.mnuPortapapeles(loopP).caption = loopP & ": " & portapapeles(loopP).nombre
            Me.mnuPortapapeles(loopP).Enabled = True
        Else
            Me.mnuPortapapeles(loopP).caption = loopP & ": < Vacio >"
            Me.mnuPortapapeles(loopP).Enabled = False
        End If
    Next
End Sub
Private Sub mnuCopiarBordes_Click()
    CopiarBordesMapaActual
End Sub

Private Sub mnuCortar_Click()
    ' Chequeo que haya seleccionado algun tipo de elemento para usar esta herramienta
    If trabajandoConElementos = tool_none Then
        Call GUI_Alert("Tenes que ir al menú 'Edición -> Trabajar con...' y  seleccionar el tipo de elementos que queres cortar.", "No hay nada para cortar.")
        Exit Sub
    End If
    
    'Cortar
    Call Me_Tools_Seleccion.cortar(areaSeleccionada, trabajandoConElementos)
    ignorarMouseUp = True
End Sub

Private Sub mnuCrearPredefinido_Click()
    
    ' Chequeo que haya seleccionado algun tipo de elemento para usar esta herramienta
    If trabajandoConElementos = tool_none Then
        Call GUI_Alert("Tenes que ir al menú 'Edición -> Trabajar con...' y  seleccionar el tipo de elementos que queres utilizar para crear el predefinido.", "Predefinido")
        Exit Sub
    End If
    
    If Me_Tools_Presets.nuevoPreset(trabajandoConElementos) Then
        GUI_Alert "Predefinido creado correctamente.", "Crear predefinido"
    Else
        GUI_Alert "No se ha podido crear el predefinido. Por favor, intente más tarde o contacte a un Administrador.", "Crear predefinido"
    End If
End Sub

Private Sub mnuCrearZonaCriatura_Click()
    
    Dim nombre As String
    Dim crear As VbMsgBoxResult
    
    'Me fijo la cantidad actual de zonas
    If UBound(mapinfo.ZonasNacCriaturas) - LBound(mapinfo.ZonasNacCriaturas) + 1 = 5 Then
        Call GUI_Alert("El máximo de zonas de criaturas en un mapa es 5.", "Máximo de zonas alcanzado.")
        Exit Sub
    End If
    
    ' Le solicito el nombre
    nombre = InputBox("Ingrese el nombre identificador de la zona. Máximo 15 cáracteres", "Crear Zona de Nacimiento de Criaturas")
    
    If nombre = vbNullString Then Exit Sub
    
    ' Acortamos el nombre por las dudas
    nombre = mid$(nombre, 1, 15)
    
    ' Ya existe una zona con este nombre?
    If Engine_Map.zonaExiste(nombre) Then
        crear = MsgBox("Ya existe una zona con ese nombre. ¿Desea remplazarla?", vbExclamation + vbYesNo, "Crear Zona de nacimiento de Criaturas")
    Else
        crear = vbTrue
    End If
    
    If crear = vbNo Then Exit Sub
    
    ' Lo agrego
    Call Engine_Map.AgregarZona(areaSeleccionada, nombre)
    
    ' Actualizo
    Call Me_Tools_Npc.cargarZonasDeNacimiento(mapinfo.ZonasNacCriaturas)
    
    ' Se la agrego...
    Call GUI_Alert("Zona creada correctamente.", "Crear Zona de Nacimiento de Criaturas")
    
End Sub

Private Sub mnuDeshacer_Click()
    
    Call ME_modComandos.desHacerAnteriorComando
    
    'Actualizo la vista
    'Map_render_2array
    rm2a
    Cachear_Tiles = True
End Sub

Private Sub mnuEdicion_Click()
    
    ' Hacer y deshacer
    If ME_modComandos.hayComandosDesHacer Then
        Me.mnuDeshacer.caption = "Deshacer " & ME_modComandos.obtenerDescripcionComandoDeshacer
        Me.mnuDeshacer.Enabled = True
    Else
        Me.mnuDeshacer.caption = "Deshacer"
        Me.mnuDeshacer.Enabled = False
    End If
    
    If ME_modComandos.hayComandosReHacer Then
        Me.mnuRehacer.caption = "Rehacer " & ME_modComandos.obtenerDescripcionComandoReHacer
        Me.mnuRehacer.Enabled = True
    Else
        Me.mnuRehacer.caption = "Rehacer"
        Me.mnuRehacer.Enabled = False
    End If
    
    ' Insertar piso en todo el mapa
    If ME_Tools.isToolSeleccionada(tool_tileset) Then
        Me.mnuInsertarPisoEnMapa.Enabled = True
    Else
        Me.mnuInsertarPisoEnMapa.Enabled = False
    End If
    
    ' Insertar piso en todo el mapa
    If ME_Tools.isToolSeleccionada(tool_tileset) Then
        Me.mnuInsertarPisoEnMapaBloqueos.Enabled = True
    Else
        Me.mnuInsertarPisoEnMapaBloqueos.Enabled = False
    End If
    
End Sub

Private Sub mnuExportarAImagen_Click()
    Dim escala As Single

    ' Le solicito la escala
    escala = val(InputBox("Ingrese un número con decimales entre 1 y 10, este número será la escala con la cual se generará la imagen que representa al mapa. Mientras mayor sea el número más chica será la imagen. Ejemplo: Si el número es 5, el tamaño de la imagen sera la mitad del tamaño real del mapa.", "Escala"))
   
    ' ¿La escala es correcta?
    If escala < 1 Or escala > 10 Then
        Call GUI_Alert("La escala seleccionada (" & escala & ") no es válida. Tiene que ser un número decimal entre 1 y 10.", "Escala seleccionada")
        Exit Sub
    End If
    
    ' Iniciamos. Se guardara en formato PNG y al terminar se abrira la carpeta con el archivo seleccionado
    Call frmExportarAux.exportarMapa(1 / escala, True, True)
End Sub

Public Sub ocultarRender(textoMensaje As String)
    frmMain.renderer.Visible = False
    frmMain.lblTapaSolapas.Visible = True
    frmMain.lblTapaSolapas.caption = textoMensaje
    frmMain.Refresh
End Sub

Public Sub mostrarRender()
    frmMain.renderer.Visible = True
    frmMain.lblTapaSolapas.Visible = False
End Sub


Private Sub mnuExportarZonaDeTrabajo_Click()
    Dim escala As Integer
    
    ' Le solicito la escala
    escala = val(InputBox("Ingrese un número con decimales entre 1 y 10, este número será la escala con la cual se generará la imagen que representa al mapa. Mientras mayor sea el número más chica será la imagen. Ejemplo: Si el número es 5, el tamaño de la imagen sera la mitad del tamaño real del mapa.", "Escala"))
   
    ' ¿La escala es correcta?
    If escala < 1 Or escala > 10 Then
        Call GUI_Alert("La escala seleccionada (" & escala & ") no es válida. Tiene que ser un número decimal entre 1 y 10.", "Escala seleccionada")
        Exit Sub
    End If

    ' Exportamos
    Call frmExportarAux.exportarZona(1 / escala)
End Sub

Private Sub mnuGenerarInformes_Click()
    load frmInformeMapa
   
    frmInformeMapa.Show , Me

    Call modPosicionarFormulario.posicionarAbajoCentro(Me, frmInformeMapa)
End Sub


Private Sub mnuInsertarBloqueoArea_Click()
    Dim respuesta As VbMsgBoxResult
    Dim x1 As Integer
    Dim y1 As Integer
    
    Dim loopX As Integer
    Dim loopY As Integer
    
    respuesta = MsgBox("Se aplicará el bloqueo en el area donde tenes el mouse delimitada por bloqueos o por el fin del mapa. ¿Estás seguro?", vbQuestion + vbYesNo, "Aplicar bloqueo al área")

    If respuesta = vbYesNo Then Exit Sub
    
    'El punto inciial es dond esta el mouse
    x1 = MouseTileX
    y1 = MouseTileY
                        
    ' Bloqueamos a través de una funcion recursiva
    Call ME_Tools_Triggers.generarBloqueoExpansivo(x1, y1)
    
    'Ejecutamos la acción sobre esa area
    click_tool vbLeftButton
        
    ' Restablecemos
    Call modSeleccionArea.puntoArea(ME_Tools.areaSeleccionada, x1, y1)
    Call modSeleccionArea.actualizarArea(ME_Tools.areaSeleccionada, x1, y1)
    Call activarUltimaHerramientaCorrespondienteASolapa(eSolapasEditor.Bloqueos)
  
End Sub

Private Sub mnuInsertarPisoEnMapa_Click()
    Dim respuesta As VbMsgBoxResult
   
    respuesta = MsgBox("Se aplicará el piso seleccionado a toda el área libre de pisos en donde está el mouse. ¿Estás seguro?. Podés volver para atrás con CONTROL + Z.", vbQuestion + vbYesNo, "Aplicar piso al mapa")
    
    If respuesta = vbNo Then Exit Sub

    Call aplicarInsertarExpansivo(MouseTileX, MouseTileY, False)
End Sub

Private Sub aplicarInsertarExpansivo(ByVal x As Byte, ByVal y As Byte, considerarBloqueos As Boolean)
   
    ' Guardamos el estado de las herramientas
    Call backupearEstadoHerramientas
    
    ' Generamos
    If Me_Tools_TileSet.generarPisoExpansivo(x, y, considerarBloqueos) Then
    
        'Ejecutamos la acción sobre esa area
        click_tool vbLeftButton
        
        ' Restablecemos
        Call restablecerBackupHerramientas
    
        ' Suponemos que estaba con los tilesets
        Call activarUltimaHerramientaCorrespondienteASolapa(eSolapasEditor.Tilesets)
    Else
        Call GUI_Alert("No se puede poner el piso acá. Revisa de estar haciendo las cosas bien.")
    End If

End Sub

Private Sub mnuInsertarPisoEnMapaBloqueos_Click()
    Dim respuesta As VbMsgBoxResult
   
    respuesta = MsgBox("Se aplicará el piso seleccionado a toda el área libre de pisos y bloqueos en donde está el mouse. ¿Estás seguro?. Podés volver para atrás con CONTROL + Z.", vbQuestion + vbYesNo, "Aplicar piso al mapa")
    
    If respuesta = vbNo Then Exit Sub

    Call aplicarInsertarExpansivo(MouseTileX, MouseTileY, True)
End Sub

Private Sub mnuNewMap_Click()
    NuevoMapa
    act_titulo
End Sub

Private Sub mnuNumeroEfectoSonido_Click()

    frmMain.mnuNumeroEfectoSonido.checked = Not frmMain.mnuNumeroEfectoSonido.checked
    
    If frmMain.mnuNumeroEfectoSonido.checked = True Then
        mostrarTileEfectoSonidoPasos = vbChecked
    Else
        mostrarTileEfectoSonidoPasos = vbUnchecked
    End If
    
    ActChecks
    
End Sub

Private Sub mnuNumeroTilePiso_Click()
    
    frmMain.mnuNumeroTilePiso.checked = Not frmMain.mnuNumeroTilePiso.checked
    
    If frmMain.mnuNumeroTilePiso.checked = True Then
        mostrarTileNumber = vbChecked
    Else
        mostrarTileNumber = vbUnchecked
    End If
    
    ActChecks
End Sub

Private Sub mnuOpcionesDelEditor_Click()
    Call frmOpcionesDelEditor.Show(, Me)
End Sub

Private Sub mnuOpenPak_Click()
    GUI_Load New VWOpenPakMap
End Sub

Private Sub mnuPegar_Click()
    'Pegar
    Call Me_Tools_Seleccion.pegar(areaSeleccionada)
End Sub

Private Sub mnuPortapapeles_Click(Index As Integer)
    Call Me_Tools_Seleccion.copiar(Index, trabajandoConElementos)
End Sub

Private Sub mnuRehacer_Click()
    Call ME_modComandos.reHacerSiguienteComando
    'Actualizo la vista
    'Map_render_2array
    rm2a
    Cachear_Tiles = True
End Sub

Public Function GuardarMapaActual() As Boolean
    Debug.Print THIS_MAPA.Path
    
    Me.caption = Me.caption & " - Guardando Mapa..."
    
    'GUARDO EN UN ARCHIVO INDIVIDUAL
    If THIS_MAPA.numero = 0 Then
        MsgBox "Estas guardando un mapa del mundo. Tenes que decir que número de mapa es."
        mnuSetNum_Click
        Exit Function
    End If
        
    If THIS_MAPA.Path <> "" Then
        Call GuardarMapa(THIS_MAPA.Path)
    Else
        If GuardarMapaComo = False Then
            MsgBox "No se guardó el mapa !!", vbCritical + vbOKOnly
            Exit Function
        End If
    End If
    
    'GUARDO EN EL PAQUETE DE MAPAS DEL EDITOR
    If THIS_MAPA.numero > 0 And FileExist(THIS_MAPA.Path) Then
        'Dim localeID As Integer
        'If pakMapasME.Puedo_Editar(THIS_MAPA.numero, CDM_UserPrivs, CDM_UserID) Then
            pakMapasME.Parchear THIS_MAPA.numero, THIS_MAPA.Path, 0
            Call versionador.modificado("RECURSO_MAPA", THIS_MAPA.numero)
            'localeID = CDM_Commit(THIS_MAPA.Path, THIS_MAPA.numero, CDM_Upd_Mapas)
            'CDM_Enviar localeID
        'Else
        '    MsgBox "No tenes permiso para parchear el mapa numero " & THIS_MAPA.numero, vbOKOnly + vbExclamation
        'End If
        GuardarMapaActual = True
    End If
    
    act_titulo
End Function

Private Sub mnuReportarBug_Click()
    Call frmBug.crearEnBlanco
    Call frmBug.Show(, Me)
End Sub

Private Sub mnuSalirDelEditor_Click()
    Call salirDelEditor
End Sub

Private Sub mnuSave_Click()
    SpoofCheck.Enabled = False
    Call GuardarMapaActual
    SpoofCheck.Enabled = True
End Sub

Private Sub mnuSaveAndCompile_Click()
    'Guarda el mapa en el archivo y en el pakme
    Call mnuSave_Click

    CompilarMapa
    
    act_titulo
End Sub

Private Sub mnuSaveAs_Click()
    SpoofCheck.Enabled = False
    
    If GuardarMapaComo = False Then
        SpoofCheck.Enabled = True
        MsgBox "No se guardó el mapa !!", vbCritical + vbOKOnly
        Exit Sub
    End If
    
    SpoofCheck.Enabled = True
    act_titulo
End Sub


Private Sub mnuSetNum_Click()

Dim tmp_num As Integer
Dim resultado As String

resultado = InputBox("Ingrese el número del mapa", Me.caption, THIS_MAPA.numero)

If IsNumeric(resultado) Then
    tmp_num = val(resultado)
    
    THIS_MAPA.numero = tmp_num
    
    Call ME_Mapas.cargarInformacionDeMapa(THIS_MAPA.numero, mapinfo)
        
    THIS_MAPA.nombre = mapinfo.Name
        
    Call act_titulo
End If

'If pakMapasME.Puedo_Editar(tmp_num, CDM_UserPrivs, CDM_UserID) Then
    
'Else
'    MsgBox "No podés parchear el mapa " & tmp_num
'End If


End Sub

Private Sub nombre_mapa_Change()

End Sub

Private Sub mnuTechosTransparentes_Click()
    
    frmMain.mnuTechosTransparentes.checked = Not frmMain.mnuTechosTransparentes.checked
    
    If frmMain.mnuTechosTransparentes.checked = True Then
        dibujarTechosTransparentes = vbChecked
    Else
        dibujarTechosTransparentes = vbUnchecked
    End If
    
    bTecho = bTecho
    
    ActChecks
    
End Sub

Private Sub mnuTilesets_Click()
    select_tool Tools.tool_tileset
    load frmConfigurarPisos
    frmConfigurarPisos.Show , Me
End Sub

Private Sub mnuTranslados_Click()
select_tool Tools.tool_acciones
End Sub


Private Sub mnuTodoDia_Click()
    mnuTodoDia.checked = Not mnuTodoDia.checked
    Forzar_Dia = mnuTodoDia.checked
    
    forzar_dia_c.value = IIf(Forzar_Dia, vbChecked, vbUnchecked)
End Sub

Private Sub mnuTrabajarCon_Elemento_Click(Index As Integer)
    Me.mnuTrabajarCon_Elemento(Index).checked = Not Me.mnuTrabajarCon_Elemento(Index).checked

    Call actualizarTrabajarConElemento
End Sub

Private Sub mnuTrabajarCon_Nada_Click()
    Dim loopMenu As Byte
           
    For loopMenu = Me.mnuTrabajarCon_Elemento.LBound To Me.mnuTrabajarCon_Elemento.UBound
        Me.mnuTrabajarCon_Elemento.item(loopMenu).checked = False
    Next loopMenu

    Call actualizarTrabajarConElemento
End Sub

Private Sub mnuTrabajarCon_Todo_Click()
    Dim loopMenu As Byte
        
    For loopMenu = Me.mnuTrabajarCon_Elemento.LBound To Me.mnuTrabajarCon_Elemento.UBound
        Me.mnuTrabajarCon_Elemento.item(loopMenu).checked = True
    Next loopMenu
    
    Call actualizarTrabajarConElemento
    
End Sub

Private Sub actualizarTrabajarConElemento()

    trabajandoConElementos = 0
    
    If Me.mnuTrabajarCon_Elemento(0).checked Then
        trabajandoConElementos = trabajandoConElementos Or tool_acciones
    End If
    
    If Me.mnuTrabajarCon_Elemento(1).checked Then
        trabajandoConElementos = trabajandoConElementos Or tool_npc
    End If
    
    If Me.mnuTrabajarCon_Elemento(2).checked Then
        trabajandoConElementos = trabajandoConElementos Or tool_grh
    End If
    
    If Me.mnuTrabajarCon_Elemento(3).checked Then
        trabajandoConElementos = trabajandoConElementos Or tool_luces
    End If
    
    If Me.mnuTrabajarCon_Elemento(4).checked Then
        trabajandoConElementos = trabajandoConElementos Or tool_triggers
    End If
    
    If Me.mnuTrabajarCon_Elemento(5).checked Then
        trabajandoConElementos = trabajandoConElementos Or tool_obj
    End If
    
    If Me.mnuTrabajarCon_Elemento(6).checked Then
        trabajandoConElementos = trabajandoConElementos Or tool_particles
    End If

    If Me.mnuTrabajarCon_Elemento(7).checked Then
        trabajandoConElementos = trabajandoConElementos Or tool_tileset
    End If
End Sub
Private Sub mnuVerCantidadObjetos_Click()
    
    frmMain.mnuVerCantidadObjetos.checked = Not frmMain.mnuVerCantidadObjetos.checked
    
    If frmMain.mnuVerCantidadObjetos.checked = True Then
        dibujarCantidadObjetos = vbChecked
    Else
        dibujarCantidadObjetos = vbUnchecked
    End If
    
    ActChecks
End Sub

Private Sub mnuZonaTrabajo_Click(Index As Integer)
    Call establecerZonaDeTrabajo(Index)
    'Mostramos el mundo para que eliga el mapa a abrir
    Call mnuOpenPak_Click
End Sub

Private Sub establecerZonaDeTrabajo(Zona As Integer)

    Dim loopElemento As Byte
    
    'Cargamos la info de la zona
    Call ME_Mundo.CargarArrayMapas(ME_Mundo.zonas(Zona).archivo)
        
    'Actualizamos los menes
    For loopElemento = Me.mnuZonaTrabajo.LBound To Me.mnuZonaTrabajo.UBound
        Me.mnuZonaTrabajo.item(loopElemento).checked = False
    Next
    
    zonaActual = ME_Mundo.zonas(Zona).nombre
    
    'Marco como chequeado el nuevo
    Me.mnuZonaTrabajo.item(Zona).checked = True
End Sub
Private Sub obj_cantidad_Change()

Dim cantidad As Integer

cantidad = val(obj_cantidad.text)

If cantidad <= 0 Then cantidad = 1
If cantidad > 10000 Then cantidad = 10000
obj_cantidad.text = cantidad

Call Me_Tools_Objetos.seleccionarCantidadObjeto(cantidad)

End Sub

Private Sub cmdInsertarObjeto_Click()
    Call obj_cantidad_Change
    Call Me_Tools_Objetos.click_InsertarOBJ
End Sub

Private Sub opt_montaña_Click(Index As Integer)
    mt_select = Index
    renderer.SetFocus
End Sub

Private Sub cmdInsertarPreset_Click()
    Call Me_Tools_Presets.click_insertarPreset
End Sub


Private Sub puede_neblina_Click()

End Sub

Private Sub radio_montaña_Change()
radio_montaña_lbl.caption = "Radio: " & radio_montaña.value
radio_montana = radio_montaña.value

End Sub

Private Sub renderer_Click()

'puede_mover = False

'Me.visible = True 'Esto hace que el editor se centre todo el tiempo en pantalla, lo que es molesto, por ejemplo si se esta con un excel atras

Call ConvertCPtoTP(MouseX, MouseY, modPantalla.PixelesPorTile.x, modPantalla.PixelesPorTile.y, tX, tY)

End Sub

Public Function ABRIR_Mapa(MapSelected As Integer, Optional ByVal crearSinoExiste As Boolean = True) As Boolean

    Dim existe As Boolean
    
    Call ME_FIFO.prepararWorkEspace
        
    existe = SwitchMap(MapSelected)
    
    If Not existe And crearSinoExiste Then
        Call LIMPIAR_MAPA
        GUI_Alert "El mapa no existía, se creó de 0."
        existe = True
    End If
    
    
    If existe Then
        'Luego de cargar el mapa refrescamos la lista de acciones usadas en este mapa
        Call ME_modAccionEditor.refrescarListaUsando(frmMain.listTileAccionActuales)
        
        'Cargamos la info del dat
        Call ME_Mapas.cargarInformacionDeMapa(MapSelected, mapinfo)

        'Seteamos algunos valores
        With THIS_MAPA
            .numero = MapSelected
            .Path = app.Path & "\Datos\tmpmap.cache"
            .nombre = mapinfo.Name
        End With
    
        Call frmMain.setValoresAgua
        Call frmMain.act_titulo
        
        Call miniMap_Redraw
        
        ABRIR_Mapa = True
    Else
        ABRIR_Mapa = False
    End If

End Function

Private Sub renderer_DblClick()
'Call ConvertCPtoTP(MouseX, MouseY, MouseTileX, MouseTileY)
'If TipoEditorParticulas Then
'    If emisor_editado > 0 And emisor_editado < Engine_Particles.emisores_particulas_count Then
'        Engine_Particles.Particle_Group_Set_TPPos emisor_testeo, MouseX - offset_map.x, MouseY - offset_map.y
'    End If
'End If
End Sub

Private Sub renderer_GotFocus()
    focoEnElRender = True
End Sub

Private Sub renderer_KeyDown(KeyCode As Integer, Shift As Integer)
If GUI_KeyDown(KeyCode, Shift) = False Then
    
    If KeyCode = vbKeyDown Or KeyCode = vbKeyS Then
        UserDirection = SOUTH
        Exit Sub
    ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyW Then
        UserDirection = NORTH
        Exit Sub
    ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyA Then
        UserDirection = WEST
        Exit Sub
    ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyD Then
        UserDirection = EAST
        Exit Sub
    ElseIf KeyCode = vbKeyInsert Then
        UserPos.x = maxl(UserPos.x - 10, SV_Constantes.X_MINIMO_JUGABLE)
        Exit Sub
    ElseIf KeyCode = vbKeyHome Then
        UserPos.y = maxl(UserPos.y + -10, SV_Constantes.Y_MINIMO_JUGABLE)
        Exit Sub
    ElseIf KeyCode = vbKeyEnd Then
        UserPos.y = minl(UserPos.y + 10, SV_Constantes.Y_MAXIMO_JUGABLE)
        Exit Sub
    ElseIf KeyCode = vbKeyPageUp Then
        UserPos.x = minl(UserPos.x + 10, SV_Constantes.X_MAXIMO_JUGABLE)
        Exit Sub
    End If
    
    Dim solapaSeleccionada As Integer
    
    If GetKeyState(vbKeyControl) < 0 Then Exit Sub
    
    
    Select Case KeyCode
    
        Case vbKeySpace
            
            If Not modPantalla.mostrarBarraHerramientas Then
                
                If Me.frmBotonera.Visible Then
                    Me.frmBotonera.Visible = False
                    Me.SSTab1.Visible = False
                    mostrandoBarra = False
                Else
                    Me.frmBotonera.Visible = True
                    Me.SSTab1.Visible = True
                    mostrandoBarra = True
                End If
            End If
            Exit Sub
    
        Case vbKeyG
            If mostrandoVWindows Then
                Call Me_Tools_TileSet.EsconderVentanaTilesets
            Else
                Call Me_Tools_TileSet.MostrarVentanaTilesets(tileset_actual, tileset_actual_virtual)
            End If
    End Select
    
    
    ME_Tools.deseleccionarTool
 
 
    Select Case KeyCode
                
        Case vbKeyF11
            #If medir Then
                ME_Render.mostrarTiempos = Not ME_Render.mostrarTiempos
                Exit Sub
            #End If
            
        Case vbKeyMultiply
        
            Engine.Engine_Toggle_fps_limit
            Exit Sub
            
        Case vbKeyEscape
            ignorarMouseUp = False
            Exit Sub
                       
        '/* Primera linea /*
        
        Case vbKeyR 'Graficos
            
            ME_Tools.deseleccionarTool
            solapaSeleccionada = eSolapasEditor.Graficos
            
        Case vbKeyT 'Predefinidos

           solapaSeleccionada = eSolapasEditor.Herramientas
        
        Case vbKeyY 'Triggers
            solapaSeleccionada = eSolapasEditor.Triggers
            
        Case vbKeyU 'Npcs
            solapaSeleccionada = eSolapasEditor.Npcs
        
        '/* Segunda linea /*
        
        Case vbKeyF 'Montañas
            
            solapaSeleccionada = eSolapasEditor.Montañas
               
        Case vbKeyG 'Tilesets
            
            solapaSeleccionada = eSolapasEditor.Tilesets
            
        Case vbKeyH 'Bloqueos
            
            solapaSeleccionada = eSolapasEditor.Bloqueos
        
        Case vbKeyV 'Particulas
        
            solapaSeleccionada = eSolapasEditor.Particulas

        '/* Tercera linea /*
        
        Case vbKeyC 'Acciones
            
            solapaSeleccionada = eSolapasEditor.Acciones
        
        Case vbKeyJ 'Luces
            
            solapaSeleccionada = eSolapasEditor.Luces
            
        Case vbKeyB 'Objetos
        
            solapaSeleccionada = eSolapasEditor.objetos
            
        Case vbKeyN 'Mapa
        
            solapaSeleccionada = eSolapasEditor.Mapa
        
        '/* Cuarta linea /*
        Case vbKeyE 'Entidades
        
            solapaSeleccionada = eSolapasEditor.Entidades
        

        
    End Select
    
    Call activarUltimaHerramientaCorrespondienteASolapa(solapaSeleccionada)
    
End If
End Sub

Private Sub renderer_KeyPress(KeyAscii As Integer)
    GUI_Keypress KeyAscii
End Sub

Private Sub renderer_KeyUp(KeyCode As Integer, Shift As Integer)
If GUI_KeyUp(KeyCode, Shift) = False Then

    'If KeyCode = vbKeyG And GetAsyncKeyState(vbKeyControl) = 0 Then
    '    MOSTRAR_TILESET = False
    '    EsconderVentanaTilesets
    'End If
    
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or _
        KeyCode = vbKeyS Or KeyCode = vbKeyW Or KeyCode = vbKeyD Or KeyCode = vbKeyA Then
        UserDirection = 0
    End If
End If
End Sub

Private Sub renderer_LostFocus()
    focoEnElRender = False
End Sub

Private Sub renderer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    MouseBoton = Button
    MouseShift = Shift
    
    Call ConvertCPtoTP(x, y, modPantalla.PixelesPorTile.x, modPantalla.PixelesPorTile.y, MouseTileX, MouseTileY)
    
    clickpos.x = MouseTileX
    clickpos.y = MouseTileY
    inicial_click_tile.x = MouseTileX
    inicial_click_tile.y = MouseTileY
    inicial_click.x = x
    inicial_click.y = y
    clickposp.x = x
    clickposp.y = y
    
    Dim tX As Long
    Dim tY As Long
    
    tX = 32 * x / modPantalla.PixelesPorTile.x
    tY = 32 * y / modPantalla.PixelesPorTile.y
    
    If GUI_MouseDown(Button, Shift, tX, tY) = False Then
       
        If editando_montaña Then
            ME_Tools.calcular_montaña clickposp.y - y
        End If
        
        If TOOL_SELECC = Tools.tool_montaña And editando_montaña = False Then
            editando_montaña = True
            Set comandoMontaniaActual = New cComandoInsertarMotania
            Call comandoMontaniaActual.crear(MouseTileX, MouseTileY, ME_Tools.radio_montana)
        End If
    
        'Seleccionar parte de la pantalla
        If Not editando_montaña Then
            If Button = vbLeftButton Then
                Call modSeleccionArea.puntoArea(areaSeleccionada, MouseTileX, MouseTileY)
                TOOL_SELECC = TOOL_SELECC Or Tools.tool_seleccion
            End If
        End If
    End If

End Sub

Private Sub renderer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    renderer.SetFocus

    If (MouseX <> x Or MouseY <> y) And Not focoEnElRender Then
        Dim frm As Form
        For Each frm In Forms
            If frm.Visible = True And Not frm Is Me Then
                Exit Sub
            End If
        Next
        renderer.SetFocus
    End If
    
    MouseX = x
    MouseY = y
    
    Call ConvertCPtoTP(x, y, modPantalla.PixelesPorTile.x, modPantalla.PixelesPorTile.y, MouseTileX, MouseTileY)
    
    Dim tX As Long
    Dim tY As Long
    
    tX = 32 * x / modPantalla.PixelesPorTile.x
    tY = 32 * y / modPantalla.PixelesPorTile.y
    
    If GUI_MouseMove(Button, Shift, tX, tY) = False Then
        
        If y <= Y_MAXIMO_VISIBLE Then console_alpha = True Else console_alpha = False
        
        TOOL_MOUSEOVER = (y - 28) / 12
        
        If Button = 0 Then
            clickpos.x = MouseTileX
            clickpos.y = MouseTileY
            clickposp.x = x
            clickposp.y = y
        Else
            final_click_tile.x = MouseTileX
            final_click_tile.y = MouseTileY
            final_click.x = x
            final_click.y = y
            
            If editando_montaña Then
                ME_Tools.calcular_montaña clickposp.y - y
            End If
        End If

        'Seleccionando area de la pantalla?
        If Button = vbLeftButton And (TOOL_SELECC And Tools.tool_seleccion) Then
          Call modSeleccionArea.actualizarArea(areaSeleccionada, MouseTileX, MouseTileY)
        Else
          Call modSeleccionArea.puntoArea(areaSeleccionada, MouseTileX, MouseTileY)
        End If
        
    End If
    
End Sub

Private Sub renderer_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    clicX = x
    clicY = y
    
    Call ConvertCPtoTP(x, y, modPantalla.PixelesPorTile.x, modPantalla.PixelesPorTile.y, MouseTileX, MouseTileY)
    
    Dim tX As Long
    Dim tY As Long
    
    tX = 32 * x / modPantalla.PixelesPorTile.x
    tY = 32 * y / modPantalla.PixelesPorTile.y
    
    If GUI_MouseUp(Button, Shift, tX, tY) = False Then
        If Not ignorarMouseUp Then
            click_tool Button
        Else
            ignorarMouseUp = False
        End If
        
        If TOOL_SELECC = Tools.tool_montaña And editando_montaña = True Then
            'Backup_HM
            If Not comandoMontaniaActual Is Nothing Then
                comandoMontaniaActual.leerNuevas
                
                Call ME_Tools.ejecutarComando(comandoMontaniaActual)
                
                Set comandoMontaniaActual = Nothing
            End If
            'Compute_Mountain
        End If
        
        'Seleccionar parte de la pantalla
        If TOOL_SELECC And Tools.tool_seleccion And Button = vbLeftButton Then
            TOOL_SELECC = (TOOL_SELECC Xor Tools.tool_seleccion)
            Call modSeleccionArea.reiniciarArea(areaSeleccionada)
        End If
        
    End If
    
    editando_montaña = False
    
    Cachear_Tiles = True
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseX = x
    MouseY = y
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '

Private Sub SpoofCheck_Timer()
Dim MapE As Mapa
Dim viejoCaption As String

MapE = THIS_MAPA

viejoCaption = Me.caption
Me.caption = Me.caption & " - Guardando copia de seguridad..."
'Lo gaurdo con otro nombre por si falla el algortimo de guardado
If ME_FIFO.Guardar_Mapa_ME(app.Path & "\Datos\tmpmap.cache~1") Then
    'Si esta todo ok, elimino la copia temporal
    If FileExist(app.Path & "\Datos\tmpmap.cache") Then
        Kill app.Path & "\Datos\tmpmap.cache"
    End If
    
    'Renombro el archivo
    Name (app.Path & "\Datos\tmpmap.cache~1") As (app.Path & "\Datos\tmpmap.cache")
    
    THIS_MAPA = MapE
End If

Me.caption = viejoCaption
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
SetEstadoFlagsMapaEnControles

On Error Resume Next
renderer.SetFocus
End Sub


Public Sub SetEstadoFlagsMapaEnControles()
    LabelCol.BackColor = RGB(mapinfo.BaseColor.r, mapinfo.BaseColor.g, mapinfo.BaseColor.b)
    Me.chkElMapa.value = IIf(mapinfo.ColorPropio, vbChecked, 0)
End Sub

Private Sub tilesets_area_sel_agua_Click()

'Oculto la ventana de tilesets
Call Me_Tools_TileSet.EsconderVentanaTilesets
    
'Desactivo el boton actual y activo el que me permite seleccionar
Me.tilesets_area_sel_agua.Enabled = False
Me.tilesets_area_sel_agua.Visible = False
Me.cmdSeleccionarAreaAguaTierra.Enabled = True
Me.cmdSeleccionarAreaAguaTierra.caption = "Seleccionar"

'Actualizo el area seleccionada
lblAguaX1.caption = Area_Tileset.izquierda
lblAguaX2.caption = Area_Tileset.derecha
lblAguaY1.caption = Area_Tileset.arriba
lblAguaY2.caption = Area_Tileset.abajo

'El tilesets y su nombre
Me.lblTexturaSeleccionadaAguaTierra.caption = Engine_Tilesets.Tilesets(tileset_actual).nombre

enable_agua_buttons


End Sub

Private Sub toll_mnu_sol_Click()
select_tool Tools.tool_sol
End Sub

Private Sub tool_acciones_Click()
frmMain.SSTab1.Tab = eSolapasEditor.Acciones
End Sub

Private Sub tool_luces_Click()
select_tool Tools.tool_luces
End Sub

Private Sub tool_montania_Click()
select_tool Tools.tool_montaña
End Sub

Private Sub tool_ninguna_Click()
select_tool tool_none
End Sub

Sub select_tool(ByVal tool As Tools)
TOOL_SELECC = 0
TOOL_SELECC = tool
selec_TOOL
End Sub

Private Sub tool_particulas_Click()
select_tool tool_particles
End Sub

Private Sub tool_tilesets_Click()
select_tool tool_tileset
End Sub

Private Sub tool_triggers_Click()
    frmMain.SSTab1.Tab = eSolapasEditor.Triggers
End Sub

Private Sub timerAnimarDia_Timer()
    If hora_scroll.value + 1 = 97 Then
        hora_scroll.value = 1
    Else
        hora_scroll.value = hora_scroll.value + 1
    End If
    
End Sub

Private Sub ver_acciones_Click()

frmMain.ver_acciones.checked = Not frmMain.ver_acciones.checked

If frmMain.ver_acciones.checked = True Then
    dibujarAccionTile = vbChecked
Else
    dibujarAccionTile = vbUnchecked
End If

ActChecks
End Sub

Private Sub ver_area_Click()
ver_area.checked = Not ver_area.checked
DRAWCLIENTAREA = IIf(ver_area.checked, 1, 0)
End Sub

Private Sub ver_bloqueos_Click()

ver_bloqueos.checked = Not ver_bloqueos.checked

If ver_bloqueos.checked = True Then
    DRAWBLOQUEOS = vbChecked
Else
    If TOOL_SELECC And Tools.tool_bloqueo Then
        DRAWBLOQUEOS = vbGrayed
    Else
        DRAWBLOQUEOS = vbUnchecked
    End If
End If

ActChecks

End Sub

Private Sub ver_char_Click()
ToggleWalkMode
End Sub

Private Sub ver_graficos_Click()
    Me.ver_graficos.checked = Not Me.ver_graficos.checked
    
    If Me.ver_graficos.checked = True Then
        ME_Tools.mostrarTileDondeHayGraficos = vbChecked
    Else
        ME_Tools.mostrarTileDondeHayGraficos = vbUnchecked
    End If
    
    ActChecks
End Sub

Private Sub ver_grilla_Click()
ver_grilla.checked = Not ver_grilla.checked
DRAWGRILLA = IIf(ver_grilla.checked, 1, 0)
End Sub

Private Sub ver_luces_Click()
    ver_luces.checked = Not ver_luces.checked
    
    If ver_luces.checked = True Then
        ME_Tools.mostrarTileDondeHayLuz = vbChecked
    Else
        ME_Tools.mostrarTileDondeHayLuz = vbUnchecked
    End If
    
    ActChecks
End Sub

Private Sub ver_particulas_Click()
    Dim DUMMY As Byte
End Sub

Private Sub ver_triggers_Click()
ver_triggers.checked = Not ver_triggers.checked

If ver_triggers.checked = True Then
    DRAWTRIGGERS = vbChecked
Else
    If TOOL_SELECC And Tools.tool_triggers Then
        DRAWTRIGGERS = vbGrayed
    Else
        DRAWTRIGGERS = vbUnchecked
    End If
End If

ActChecks
End Sub


Public Function GuardarMapaComo() As Boolean
    Dim tmp_path As String
    On Error GoTo Cancelar:
    
    ColorDialog.filter = "Mapa TDS (*.tmap)|*.tmap"
    ColorDialog.flags = cdlOFNHideReadOnly
    ColorDialog.InitDir = DatosPath & "Mapas\Raw"
    ColorDialog.FileName = THIS_MAPA.numero & ".tmap"
    ColorDialog.DefaultExt = "tmap"
    ColorDialog.DialogTitle = "Guardar como..."
    Me.ColorDialog.ShowSave
    tmp_path = ColorDialog.FileName
    If FileExist(tmp_path, vbNormal) = True Then
        If MsgBox("¿Desea sobrescribir " & tmp_path & "?", vbExclamation + vbYesNo) = vbNo Then
            GuardarMapaComo = False
            Exit Function
        Else
            Kill tmp_path
        End If
    End If
    THIS_MAPA.Path = tmp_path
    GuardarMapaComo = GuardarMapa(tmp_path)
    
    GuardarMapaComo = True
    
    Exit Function
Cancelar:
    GuardarMapaComo = False
End Function

Public Sub ShowAbrirMapa()
Dim tmp_path As String

On Error GoTo BotonCancelar

ColorDialog.CancelError = True 'Si la persona toca cancelar, se genera un error
ColorDialog.filter = "Mapa TDS (*.tmap)|*.tmap|am(*.am)|*.am"
ColorDialog.flags = cdlOFNHideReadOnly
ColorDialog.InitDir = DatosPath & "Mapas\Raw"
ColorDialog.FileName = "SinTitulo.tmap"
ColorDialog.DefaultExt = "tmap"
ColorDialog.DialogTitle = "Abrir mapa"
Me.ColorDialog.ShowOpen
tmp_path = ColorDialog.FileName

AbrirMapa tmp_path

SetEstadoFlagsMapaEnControles

Exit Sub
BotonCancelar:
    Err.Clear
    Exit Sub
End Sub

Public Sub act_titulo()
Me.caption = "Tierras del Sur - Editor del Mundo - " & Trim$(THIS_MAPA.nombre) & " - (" & THIS_MAPA.numero & ")" & IIf(THIS_MAPA.editado, " *(Copia de seguridad)*", "") & " - " & CDM.cerebro.Usuario.nombre
If THIS_MAPA.numero = 0 Then
    mnuSaveAndCompile.Enabled = False
Else
    mnuSaveAndCompile.Enabled = True
End If
End Sub

Private Sub actualizarMenus()
    Dim Enabled As Boolean
    
    Enabled = cerebro.Usuario.tienePermisos("RECURSOS", ePermisosCDM.lectura)
    
    Me.mnu_paker.Visible = Enabled
    Me.mnu_paker.Enabled = Enabled
    Me.separadorasds.Visible = Enabled
    Me.separadorasds.Enabled = Enabled
        
    Enabled = cerebro.Usuario.tienePermisos("CONFIG.GRAFICOS", ePermisosCDM.lectura)

    Me.mnu_ventana_indexar.Visible = Enabled
    Me.mnu_ventana_indexar.Enabled = Enabled
    
    Enabled = cerebro.Usuario.tienePermisos("CONFIG.PISOS", ePermisosCDM.lectura)
    
    Me.mnuTilesets.Visible = Enabled
    Me.mnuTilesets.Enabled = Enabled

    Enabled = cerebro.Usuario.tienePermisos("EDITOR.CDM", ePermisosCDM.escritura)
    
    Me.mnu_CDMCompartirNovedades.Visible = Enabled
    Me.mnu_CDMCompartirNovedades.Enabled = Enabled
    
    Enabled = cerebro.Usuario.tienePermisos("EDITOR.CDM", ePermisosCDM.lectura)
    
    Me.mnu_CDMObtenerNovedades.Visible = Enabled
    Me.mnu_CDMObtenerNovedades.Enabled = Enabled
    
    ' ¿Hace falta mostrar el menu?
    Enabled = (Me.mnu_CDMCompartirNovedades.Visible Or Me.mnu_CDMObtenerNovedades.Visible)

    Me.mnuEquipo.Visible = Enabled
    Me.mnuEquipo.Enabled = Enabled
    
End Sub
Private Sub Form_Load()

' Climas disponibles
Call ME_Climas.cargarClimasDisponiblesEnCombo(Me.cmbClimaActual)

Call Me.cargarPreferenciasUsuario

Call Me.activarUltimaHerramientaCorrespondienteASolapa(eSolapasEditor.Mapa)

Call actualizarMenus

luzSeleccionada.LuzRadio = 3
Call Me_Tools_Luces.seleccionarLuz(luzSeleccionada)

'Cargamos las zonas entre las cuales se puede cambiar
Dim loopZona As Byte

If ME_Mundo.CantidadZonasCargadas > 0 Then
    For loopZona = 0 To UBound(ME_Mundo.zonas)
        If loopZona > 0 Then load frmMain.mnuZonaTrabajo(loopZona)
        frmMain.mnuZonaTrabajo.item(loopZona).caption = ME_Mundo.zonas(loopZona).nombre
        frmMain.mnuZonaTrabajo.item(loopZona).Enabled = True
    Next
    
    '¿Tiene una zona predefinida valida?
    If ME_Mundo.obtenerIDZona(zonaActual) >= 0 Then
        Call establecerZonaDeTrabajo(ME_Mundo.obtenerIDZona(zonaActual))
    Else
        Call establecerZonaDeTrabajo(0)
    End If
End If
IniciarRuedaMouse renderer.hwnd

txtShader.text = Engine.ShaderSombra

'Activamos el timer que guarda el mapa cada determinado tiempo en la copia de seguridad
SpoofCheck.Enabled = True

frmMain.editandoTileSets = True



End Sub

Private Sub Form_Terminate()
prgRun = False
End Sub

Public Function salirDelEditor(Optional ByVal sinPreguntar As Boolean = False) As Boolean
    Dim respuesta As VbMsgBoxResult
    
    If sinPreguntar Then
        Call guardarPreferenciasUsuario
        Call cerrarEditor
        Exit Function
    End If
    
    respuesta = MsgBox("¿Estás seguro de que queres salir del editor?", vbYesNo + vbExclamation, Me.caption)
    
    If respuesta = vbYes Then
        Call guardarPreferenciasUsuario
        Call cerrarEditor
        salirDelEditor = True
    Else
        salirDelEditor = False
    End If
End Function


Private Sub Form_Unload(Cancel As Integer)
    Cancel = IIf(salirDelEditor, 0, 1)
End Sub

Public Sub cerrarEditor()
    prgRun = False
End Sub


Private Sub guardarUltimaZonaUtilizada()
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_bloqueos", IIf(frmMain.ver_bloqueos.checked, "SI", "NO"))
End Sub
Public Sub guardarPreferenciasUsuario()

    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ZONA", zonaActual)
    'Esto deberia ser sobre las variables y no los menu...
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_bloqueos", IIf(frmMain.ver_bloqueos.checked, "SI", "NO"))
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_triggers", IIf(frmMain.ver_triggers.checked, "SI", "NO"))
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_acciones", IIf(frmMain.ver_acciones.checked, "SI", "NO"))
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_luces", IIf(frmMain.ver_luces.checked, "SI", "NO"))
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_CantidadObjetos", IIf(frmMain.mnuVerCantidadObjetos.checked, "SI", "NO"))
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_particulas", IIf(frmMain.ver_particulas.checked, "SI", "NO"))
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_NumeroTilePiso", IIf(frmMain.mnuNumeroTilePiso.checked, "SI", "NO"))
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_TodoDia", IIf(frmMain.mnuTodoDia.checked, "SI", "NO"))
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_TechosTransparentes", IIf(frmMain.mnuTechosTransparentes.checked, "SI", "NO"))
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_ZonaNacimientoCriaturas", IIf(frmMain.ver_ZonadeNacCriaturas.checked, "SI", "NO"))
     Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_ZonaNacimientoCriatura", IIf(frmMain.ver_ZonadeNacCriatura.checked, "SI", "NO"))
     
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_grilla", IIf(frmMain.ver_grilla.checked, "SI", "NO"))
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_area", IIf(frmMain.ver_area.checked, "SI", "NO"))
    
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_MiniMapa", IIf(frmMain.chkMiniMapa.checked, "SI", "NO"))
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_MiniMapaBloqueos", IIf(frmMain.chkMiniMapaBloqueos.checked, "SI", "NO"))
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_MiniMapaLuces", IIf(frmMain.chkMiniMapaLuces.checked, "SI", "NO"))
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_MiniMapaNPC", IIf(frmMain.chkMiniMapaNPC.checked, "SI", "NO"))
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_MiniMapaColores", IIf(frmMain.chkMiniMapaColores.checked, "SI", "NO"))
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_MiniMapaAcciones", IIf(frmMain.chkMiniMapaAcciones.checked, "SI", "NO"))
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ver_MiniMapaTriggers", IIf(frmMain.chkMiniMapaTriggers.checked, "SI", "NO"))
End Sub

Public Sub cargarPreferenciasUsuario()
    zonaActual = ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ZONA")
    'Esto deberia ser sobre las variables y no los menu...
    If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ver_bloqueos") = "SI" Then ver_bloqueos_Click
    If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ver_triggers") = "SI" Then ver_triggers_Click
    If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ver_acciones") = "SI" Then ver_acciones_Click
    If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ver_luces") = "SI" Then ver_luces_Click
    If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ver_CantidadObjetos") = "SI" Then mnuVerCantidadObjetos_Click
    If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ver_particulas") = "SI" Then ver_particulas_Click
    If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ver_NumeroTilePiso") = "SI" Then mnuNumeroTilePiso_Click
    If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ver_TodoDia") = "SI" Then mnuTodoDia_Click
    If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ver_TechosTransparentes") = "SI" Then mnuTechosTransparentes_Click
    If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ver_ZonaNacimientoCriaturas") = "SI" Then ver_ZonadeNacCriaturas_Click
    
    If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ver_grilla") = "SI" Then ver_grilla_Click
    If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ver_area") = "SI" Then ver_area_Click
    If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ver_char") = "SI" Then ver_char_Click
    
    If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ver_MiniMapa") = "SI" Then chkMiniMapa_Click
    If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ver_MiniMapaBloqueos") = "SI" Then chkMiniMapaBloqueos_Click
    If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ver_MiniMapaLuces") = "SI" Then chkMiniMapaLuces_Click
    If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ver_MiniMapaNPC") = "SI" Then chkMiniMapaNPC_Click
    If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ver_MiniMapaColores") = "SI" Then chkMiniMapaColores_Click
    If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ver_MiniMapaAcciones") = "SI" Then chkMiniMapaAcciones_Click
    If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ver_MiniMapaTriggers") = "SI" Then chkMiniMapaTriggers_Click
End Sub

Private Sub ver_ZonadeNacCriatura_Click()
    'DIBUJA UN TEXTO QUE INDICA EN QUE ZONAS PUEDE NACER UNA CRIATURA
    frmMain.ver_ZonadeNacCriatura.checked = Not frmMain.ver_ZonadeNacCriatura.checked
    
    If frmMain.ver_ZonadeNacCriatura.checked = True Then
        ME_Tools.dibujarZonaNacimientoCriatura = vbChecked
    Else
        ME_Tools.dibujarZonaNacimientoCriatura = vbUnchecked
    End If
    
    ActChecks
    
End Sub



Private Sub ver_ZonadeNacCriaturas_Click()
    'MARCA LAS ZONAS
    frmMain.ver_ZonadeNacCriaturas.checked = Not frmMain.ver_ZonadeNacCriaturas.checked
    
    If frmMain.ver_ZonadeNacCriaturas.checked = True Then
        ME_Tools.dibujarZonaNacimientoCriaturas = vbChecked
    Else
        ME_Tools.dibujarZonaNacimientoCriaturas = vbUnchecked
    End If
    
    ActChecks
End Sub
