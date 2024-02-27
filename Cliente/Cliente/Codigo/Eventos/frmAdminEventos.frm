VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAdminEventos 
   Caption         =   "Administrador de eventos"
   ClientHeight    =   10860
   ClientLeft      =   -60
   ClientTop       =   105
   ClientWidth     =   17595
   LinkTopic       =   "Form1"
   ScaleHeight     =   724
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1173
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmMensaje 
      Height          =   1575
      Left            =   3840
      TabIndex        =   125
      Top             =   2880
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton cmdAceptarMensaje 
         Caption         =   "Aceptar"
         Height          =   360
         Left            =   3360
         TabIndex        =   126
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblMensaje 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mensaje de aviso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   120
         TabIndex        =   127
         Top             =   240
         Width           =   4125
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame frmInfoEvento 
      Caption         =   "Información del Evento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   13200
      TabIndex        =   46
      Top             =   9480
      Visible         =   0   'False
      Width           =   12495
      Begin VB.Frame frmInscribirParticipantes 
         Height          =   4815
         Left            =   3600
         TabIndex        =   118
         Top             =   1680
         Visible         =   0   'False
         Width           =   6855
         Begin VB.CommandButton cmdVerificarInscripcionManual 
            Caption         =   "Validar Personajes"
            Height          =   360
            Left            =   3000
            TabIndex        =   123
            Top             =   4320
            Width           =   1815
         End
         Begin VB.CommandButton cmdVolverFromInscripcion 
            Caption         =   "Volver"
            Height          =   360
            Left            =   4920
            TabIndex        =   122
            Top             =   4320
            Width           =   1815
         End
         Begin VB.CommandButton cmdInscribir 
            Caption         =   "Inscribir"
            Height          =   360
            Left            =   1080
            TabIndex        =   121
            Top             =   4320
            Width           =   1810
         End
         Begin VB.TextBox txtEquiposManual 
            Appearance      =   0  'Flat
            Height          =   3495
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   119
            Text            =   "frmAdminEventos.frx":0000
            Top             =   720
            Width           =   6615
         End
         Begin VB.Label lblForTextEquipos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ingresá el nombre de cada personaje separado por una coma "","" y a cada equipo por un enter (salto de linea)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   120
            TabIndex        =   120
            Top             =   240
            Width           =   6660
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton cmdPublicar 
         Caption         =   "Publicar"
         Height          =   360
         Left            =   8760
         TabIndex        =   117
         Top             =   7080
         Width           =   1575
      End
      Begin VB.CommandButton cmdInscribirManual 
         Caption         =   "Inscribir Participantes"
         Height          =   360
         Left            =   6240
         TabIndex        =   116
         Top             =   7080
         Width           =   2415
      End
      Begin VB.CommandButton cmdCancelarEvento 
         BackColor       =   &H000000FF&
         Caption         =   "Cancelar este evento"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8640
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdCerrarInfoEvento 
         Caption         =   "Volver"
         Height          =   360
         Left            =   11280
         TabIndex        =   47
         Top             =   7080
         Width           =   990
      End
      Begin RichTextLib.RichTextBox rtbInfoEvento 
         Height          =   7335
         Left            =   120
         TabIndex        =   113
         Top             =   240
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   12938
         _Version        =   393217
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmAdminEventos.frx":0039
      End
   End
   Begin VB.Frame frmOtros 
      Caption         =   "Otras configuraciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   11040
      TabIndex        =   96
      Top             =   9120
      Visible         =   0   'False
      Width           =   5415
      Begin VB.ComboBox cmbIdentificarEquipos 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblIdentificarEquipos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Identificar a los equipos con:"
         Height          =   195
         Left            =   120
         TabIndex        =   98
         Top             =   420
         Width           =   2025
      End
   End
   Begin VB.Frame FraLimiteDe 
      Caption         =   "Limite de Objetos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   18480
      TabIndex        =   91
      Top             =   4560
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox txtCantidadObjetos 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         TabIndex        =   95
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox chkNoPermitirOroBilletera 
         Appearance      =   0  'Flat
         Caption         =   "No permitir oro en la billetera. (deben tener 0 monedas en la billetera)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   3960
         Width           =   5175
      End
      Begin VB.CheckBox chkTemplate 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4320
         TabIndex        =   93
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox chkNoPermitirItems 
         Appearance      =   0  'Flat
         Caption         =   "No permitir otros objetos (lista en blanco = personajes desnudos)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   3720
         Width           =   4935
      End
      Begin TDS_1.GridTextConAutoCompletar GridTextListaObjetos 
         Height          =   3375
         Left            =   120
         TabIndex        =   149
         Top             =   360
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5953
      End
   End
   Begin VB.Frame frmRestriccionesPersonaje 
      Caption         =   "Restricciones del Personaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   5760
      TabIndex        =   84
      Top             =   7920
      Visible         =   0   'False
      Width           =   5415
      Begin TDS_1.UpDownText sliderArmadaRango 
         Height          =   300
         Left            =   2160
         TabIndex        =   132
         Top             =   1630
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   529
         MaxValue        =   0
         MinValue        =   0
         Value           =   0
         Enabled         =   -1  'True
         Blanqueado      =   0   'False
      End
      Begin TDS_1.UpDownText sliderLegionRango 
         Height          =   300
         Left            =   2160
         TabIndex        =   131
         Top             =   1150
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   529
         MaxValue        =   0
         MinValue        =   0
         Value           =   0
         Enabled         =   -1  'True
         Blanqueado      =   0   'False
      End
      Begin VB.CheckBox chkCiudadanos 
         Appearance      =   0  'Flat
         Caption         =   "Ciudadanos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkCriminales 
         Appearance      =   0  'Flat
         Caption         =   "Criminales"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   87
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox chkLegionarios 
         Appearance      =   0  'Flat
         Caption         =   "Legionarios. Rango >="
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CheckBox chkArmadas 
         Appearance      =   0  'Flat
         Caption         =   "Armadas.     Rango >="
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblAlineacionListaAlerta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Si se tilda alguna de estas opciónes se activa la restricción a la alineación del personaje."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   90
         Top             =   3840
         Width           =   5175
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAlineacionLista 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Solo pueden ingresar:"
         Height          =   195
         Left            =   120
         TabIndex        =   89
         Top             =   315
         Width           =   1545
      End
   End
   Begin VB.Frame frmRingsDescansos 
      Height          =   4335
      Left            =   240
      TabIndex        =   53
      Top             =   7800
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Frame frmDescansos 
         Caption         =   "Descansos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   5175
         Begin VB.CheckBox chkDescansoBoveda 
            Appearance      =   0  'Flat
            Caption         =   "Con Bóveda (pierde sentido limitar objetos)"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.Frame frmRings 
         Caption         =   "Rings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   54
         Top             =   2280
         Width           =   5175
         Begin VB.CheckBox chkRingAcuatico 
            Appearance      =   0  'Flat
            Caption         =   "Acuatico (¡¡obligar a llevar barca!!)"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   550
            Width           =   3495
         End
         Begin VB.CheckBox chkPlantado 
            Appearance      =   0  'Flat
            Caption         =   "De plantes (solo para 1vs1)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   3375
         End
      End
   End
   Begin VB.ComboBox cmbRestricciones 
      Height          =   315
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   120
      Width           =   5415
   End
   Begin VB.Frame frmOtrosEventos 
      Caption         =   "Eventos actuales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   7320
      TabIndex        =   27
      Top             =   4920
      Width           =   5415
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   1080
         TabIndex        =   115
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton cmdVerInfo 
         Caption         =   "Ver Info"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3240
         TabIndex        =   30
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton cmdVerEventos 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   4320
         TabIndex        =   29
         Top             =   240
         Width           =   975
      End
      Begin VB.ListBox lstEstadoEventos 
         Appearance      =   0  'Flat
         Height          =   1590
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label lblCantidadEventos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   120
         TabIndex        =   124
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame lblConfigGeneral 
      Caption         =   "Configuración General"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin TDS_1.UpDownText txtCantidadEquiposTorneoMax 
         Height          =   330
         Left            =   3120
         TabIndex        =   146
         Top             =   2565
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   582
         MaxValue        =   0
         MinValue        =   0
         Value           =   0
         Enabled         =   -1  'True
         Blanqueado      =   0   'False
      End
      Begin TDS_1.UpDownText txtCantidadEquiposTorneoMin 
         Height          =   330
         Left            =   1200
         TabIndex        =   145
         Top             =   2565
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   582
         MaxValue        =   0
         MinValue        =   0
         Value           =   0
         Enabled         =   -1  'True
         Blanqueado      =   0   'False
      End
      Begin TDS_1.UpDownText txtCantidadParticipantesTorneo 
         Height          =   330
         Left            =   2700
         TabIndex        =   144
         Top             =   2955
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         MaxValue        =   0
         MinValue        =   0
         Value           =   0
         Enabled         =   -1  'True
         Blanqueado      =   0   'False
      End
      Begin TDS_1.UpDownText txtMaxNivel 
         Height          =   330
         Left            =   2520
         TabIndex        =   143
         Top             =   3360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         MaxValue        =   0
         MinValue        =   0
         Value           =   0
         Enabled         =   -1  'True
         Blanqueado      =   0   'False
      End
      Begin TDS_1.UpDownText txtMinNivel 
         Height          =   330
         Left            =   1680
         TabIndex        =   142
         Top             =   3345
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         MaxValue        =   0
         MinValue        =   0
         Value           =   0
         Enabled         =   -1  'True
         Blanqueado      =   0   'False
      End
      Begin VB.Frame frmTiempos 
         Caption         =   "Tiempos (en minutos)"
         Height          =   1455
         Left            =   4080
         TabIndex        =   43
         Top             =   1440
         Width           =   2895
         Begin TDS_1.UpDownText sldMinutosToleranciaDeslogueo 
            Height          =   330
            Left            =   2120
            TabIndex        =   130
            Top             =   1020
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   582
            MaxValue        =   0
            MinValue        =   0
            Value           =   0
            Enabled         =   -1  'True
            Blanqueado      =   0   'False
         End
         Begin TDS_1.UpDownText txtTiempoInscripcion 
            Height          =   330
            Left            =   2120
            TabIndex        =   129
            Top             =   645
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   582
            MaxValue        =   0
            MinValue        =   0
            Value           =   0
            Enabled         =   -1  'True
            Blanqueado      =   0   'False
         End
         Begin TDS_1.UpDownText txtTiempoAviso 
            Height          =   330
            Left            =   2120
            TabIndex        =   128
            Top             =   285
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   582
            MaxValue        =   0
            MinValue        =   0
            Value           =   0
            Enabled         =   -1  'True
            Blanqueado      =   0   'False
         End
         Begin VB.Label lblTiempoEsperaDeslogueo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tolerancia si se desloguea:"
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label lblTiempoAviso 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Durante el cual se anuncia:"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            ToolTipText     =   "Cuanto tiempo se va a estar avisando a los jugadores que se va a hacer este evento."
            Top             =   360
            Width           =   1950
         End
         Begin VB.Label lblTiempoInscripcion 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Que dura la inscripción:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   720
            Width           =   1665
         End
      End
      Begin VB.CommandButton cmdDeathCrear 
         Caption         =   "Crear Evento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   7200
         Width           =   6855
      End
      Begin VB.ComboBox ComEstrellasEvento 
         Height          =   315
         ItemData        =   "frmAdminEventos.frx":00BB
         Left            =   5040
         List            =   "frmAdminEventos.frx":00BD
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Frame FraApuestas 
         Caption         =   "Apuestas"
         Height          =   1215
         Left            =   4080
         TabIndex        =   21
         Top             =   120
         Width           =   2895
         Begin TDS_1.UpDownText txtTiempoApuestas 
            Height          =   330
            Left            =   1680
            TabIndex        =   140
            Top             =   840
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   582
            MaxValue        =   0
            MinValue        =   0
            Value           =   0
            Enabled         =   -1  'True
            Blanqueado      =   0   'False
         End
         Begin VB.CheckBox chkApuestasActivadas 
            Appearance      =   0  'Flat
            Caption         =   "Activadas"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   230
            Width           =   1335
         End
         Begin VB.TextBox txtApuestasPozoInicial 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   22
            Text            =   "0"
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblTiempoApuestas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Minutos para apostar:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   104
            Top             =   920
            Width           =   1530
         End
         Begin VB.Label lblPozoInicial 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pozo inicial:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   530
            Width           =   840
         End
      End
      Begin VB.OptionButton OptEvento 
         Caption         =   "Con evento"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   3720
         UseMaskColor    =   -1  'True
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptEvento 
         Caption         =   "Sin evento"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   13
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox txtPrecioInscripcionTorneo 
         Appearance      =   0  'Flat
         Height          =   280
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   6
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtDescripcionTorneo 
         Appearance      =   0  'Flat
         Height          =   1455
         Left            =   1200
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtNombreTorneo 
         Appearance      =   0  'Flat
         Height          =   280
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
      Begin VB.Frame frameSubEvento 
         Height          =   3375
         Left            =   120
         TabIndex        =   15
         Top             =   3720
         Width           =   6915
         Begin TDS_1.UpDownText txtAlMejorDe 
            Height          =   330
            Left            =   1800
            TabIndex        =   141
            Top             =   300
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   582
            MaxValue        =   0
            MinValue        =   0
            Value           =   0
            Enabled         =   -1  'True
            Blanqueado      =   0   'False
         End
         Begin VB.Frame frmSinEvento 
            BorderStyle     =   0  'None
            Height          =   3015
            Left            =   360
            TabIndex        =   110
            Top             =   2280
            Width           =   6765
            Begin VB.OptionButton optSinSubEvento 
               Caption         =   "Sumonear inmediatamente a los descansos"
               Height          =   495
               Index           =   0
               Left            =   120
               TabIndex        =   112
               Top             =   0
               Value           =   -1  'True
               Width           =   3375
            End
            Begin VB.OptionButton optSinSubEvento 
               Caption         =   "No sumonear. Generar lista de inscriptos."
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   111
               Top             =   480
               Width           =   3255
            End
         End
         Begin VB.Frame frmCaenItems 
            Height          =   615
            Left            =   3120
            TabIndex        =   106
            Top             =   1440
            Width           =   3615
            Begin VB.ComboBox cmbCaenItemsTipo 
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   108
               Top             =   240
               Width           =   3255
            End
            Begin VB.CheckBox chkCaenItem 
               Appearance      =   0  'Flat
               Caption         =   "Por los Objetos"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   120
               TabIndex        =   107
               Top             =   0
               Width           =   1455
            End
         End
         Begin VB.Frame frmGanadorQuedaCancha 
            Height          =   1095
            Left            =   3120
            TabIndex        =   99
            Top             =   240
            Width           =   3615
            Begin TDS_1.UpDownText sliderDebeEsperar 
               Height          =   325
               Left            =   2280
               TabIndex        =   148
               Top             =   680
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   582
               MaxValue        =   0
               MinValue        =   0
               Value           =   0
               Enabled         =   -1  'True
               Blanqueado      =   0   'False
            End
            Begin TDS_1.UpDownText sliderGanadorCancha 
               Height          =   325
               Left            =   2800
               TabIndex        =   147
               Top             =   240
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   582
               MaxValue        =   0
               MinValue        =   0
               Value           =   0
               Enabled         =   -1  'True
               Blanqueado      =   0   'False
            End
            Begin VB.CheckBox chkGanadorQuedaEnCancha 
               Appearance      =   0  'Flat
               Caption         =   "Ganador queda en cancha"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   100
               Top             =   0
               Width           =   2295
            End
            Begin VB.Label lblSliderDebeEsperar 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Equipo que pierde debe esperar para volver a jugar"
               Enabled         =   0   'False
               Height          =   435
               Left            =   120
               TabIndex        =   102
               Top             =   600
               Width           =   2535
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblForsliderGanadorCancha 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hasta ganar la cantidad de eventos:"
               Enabled         =   0   'False
               Height          =   195
               Left            =   120
               TabIndex        =   101
               Top             =   320
               Width           =   2580
            End
         End
         Begin VB.CheckBox chkConVuelta 
            Appearance      =   0  'Flat
            Caption         =   "Con vuelta"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   840
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   20
            Top             =   1800
            Width           =   1095
         End
         Begin VB.CheckBox chkClasificacionCompleta 
            Appearance      =   0  'Flat
            Caption         =   "Clasificacion Completa"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   840
            TabIndex        =   19
            Top             =   1320
            Width           =   2055
         End
         Begin VB.OptionButton OptTipoSubEvento 
            Caption         =   "DeathMatch"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OptTipoSubEvento 
            Caption         =   "PlayOff"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   1080
            Width           =   1575
         End
         Begin VB.OptionButton OptTipoSubEvento 
            Caption         =   "Liga"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   16
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label lblCombatesAlMejorDe 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Combates al mejor de:"
            Height          =   195
            Left            =   120
            TabIndex        =   103
            Top             =   360
            Width           =   1560
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   3960
         X2              =   3960
         Y1              =   3840
         Y2              =   135
      End
      Begin VB.Label lblCantidadEstrellas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Importancia:"
         Height          =   195
         Left            =   4080
         TabIndex        =   25
         Top             =   3050
         Width           =   870
      End
      Begin VB.Label lblNivelInclusive 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "inclusive"
         Height          =   195
         Left            =   3240
         TabIndex        =   12
         Top             =   3405
         Width           =   615
      End
      Begin VB.Label lblNivelA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         Height          =   195
         Left            =   2355
         TabIndex        =   11
         Top             =   3405
         Width           =   90
      End
      Begin VB.Label lblNivel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Personajes de nivel"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   3405
         Width           =   1380
      End
      Begin VB.Label lblCantidadIntegrantes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de integrantes por equipo:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   3000
         Width           =   2520
      End
      Begin VB.Label lblCupoMaximo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cupo Maximo:"
         Height          =   195
         Left            =   2040
         TabIndex        =   8
         Top             =   2640
         Width           =   1005
      End
      Begin VB.Label lblEquiposMin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cupo Mínimo:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   990
      End
      Begin VB.Label lblPrecioInscripcion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Precio Inscripción por persona:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   2205
         Width           =   2190
      End
      Begin VB.Label lblDescripcion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   840
      End
      Begin VB.Label lblNombreEvento 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame FraPremios 
      Caption         =   "Premios (por equipo)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   7320
      TabIndex        =   48
      Top             =   480
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CheckBox chkTipoPremio 
         Appearance      =   0  'Flat
         Caption         =   "El premio está expresado en porcentajes (%). El oro a entregar será el porcentaje en base a lo recaudado."
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   3840
         Width           =   5175
      End
      Begin VB.TextBox txtOroPremio 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   480
         MaxLength       =   11
         TabIndex        =   49
         Text            =   "0"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblPuestoX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1º"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   52
         Top             =   520
         Width           =   150
      End
      Begin VB.Label lblPremiosOro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Oro por equipo"
         Height          =   195
         Left            =   360
         TabIndex        =   51
         Top             =   240
         Width           =   1050
      End
   End
   Begin VB.Frame FraClasesPermitidas 
      Caption         =   "Clases y Razas permitidas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4200
      Left            =   12840
      TabIndex        =   59
      Top             =   360
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CheckBox chkRazaX 
         Appearance      =   0  'Flat
         Caption         =   "Enano"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   81
         Top             =   720
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkRazaX 
         Appearance      =   0  'Flat
         Caption         =   "Humano"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   80
         Top             =   480
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkRazaX 
         Appearance      =   0  'Flat
         Caption         =   "Gnomo"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   3240
         TabIndex        =   79
         Top             =   1440
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkRazaX 
         Appearance      =   0  'Flat
         Caption         =   "Elfo Oscuro"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   3240
         TabIndex        =   78
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkRazaX 
         Appearance      =   0  'Flat
         Caption         =   "Elfo"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   77
         Top             =   960
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkRazaX 
         Appearance      =   0  'Flat
         Caption         =   "Todas"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   76
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkClaseX 
         Appearance      =   0  'Flat
         Caption         =   "Mago"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   75
         Top             =   480
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkClaseX 
         Appearance      =   0  'Flat
         Caption         =   "Todas"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkClaseX 
         Appearance      =   0  'Flat
         Caption         =   "Clerigo"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   73
         Top             =   480
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkClaseX 
         Appearance      =   0  'Flat
         Caption         =   "Guerrero"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   72
         Top             =   720
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkClaseX 
         Appearance      =   0  'Flat
         Caption         =   "Asesino"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1320
         TabIndex        =   71
         Top             =   720
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkClaseX 
         Appearance      =   0  'Flat
         Caption         =   "Ladrón"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   70
         Top             =   960
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkClaseX 
         Appearance      =   0  'Flat
         Caption         =   "Bardo"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   1320
         TabIndex        =   69
         Top             =   960
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkClaseX 
         Appearance      =   0  'Flat
         Caption         =   "Druida"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   68
         Top             =   1200
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkClaseX 
         Appearance      =   0  'Flat
         Caption         =   "Paladin"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   1320
         TabIndex        =   67
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkClaseX 
         Appearance      =   0  'Flat
         Caption         =   "Cazador"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   66
         Top             =   1440
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkClaseX 
         Appearance      =   0  'Flat
         Caption         =   "Pescador"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   1320
         TabIndex        =   65
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkClaseX 
         Appearance      =   0  'Flat
         Caption         =   "Herrero"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   64
         Top             =   1680
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkClaseX 
         Appearance      =   0  'Flat
         Caption         =   "Leñador"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   1320
         TabIndex        =   63
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkClaseX 
         Appearance      =   0  'Flat
         Caption         =   "Minero"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   62
         Top             =   2160
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkClaseX 
         Appearance      =   0  'Flat
         Caption         =   "Carpintero"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   1320
         TabIndex        =   61
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkClaseX 
         Appearance      =   0  'Flat
         Caption         =   "Pirata"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   60
         Top             =   1920
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         DrawMode        =   9  'Not Mask Pen
         X1              =   2760
         X2              =   2760
         Y1              =   2520
         Y2              =   120
      End
   End
   Begin VB.Frame frmHechizosPermitidos 
      Caption         =   "Hechizos permitidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   16800
      TabIndex        =   82
      Top             =   9120
      Width           =   5415
      Begin VB.CheckBox chkHechizo 
         Appearance      =   0  'Flat
         Caption         =   "Todas"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   83
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame frmRestriccionesGenerales 
      Caption         =   "Restriccines de Equipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   13080
      TabIndex        =   33
      Top             =   5160
      Width           =   5415
      Begin TDS_1.UpDownText txtAlMenosTrabajadoras 
         Height          =   330
         Left            =   2640
         TabIndex        =   139
         Top             =   2350
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         MaxValue        =   0
         MinValue        =   0
         Value           =   0
         Enabled         =   -1  'True
         Blanqueado      =   0   'False
      End
      Begin TDS_1.UpDownText txtMaxSumatoriaNiveles 
         Height          =   330
         Left            =   2280
         TabIndex        =   138
         Top             =   3330
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         MaxValue        =   0
         MinValue        =   0
         Value           =   0
         Enabled         =   -1  'True
         Blanqueado      =   0   'False
      End
      Begin TDS_1.UpDownText txtAlMenosMagicas 
         Height          =   330
         Left            =   2640
         TabIndex        =   137
         Top             =   1995
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         MaxValue        =   0
         MinValue        =   0
         Value           =   0
         Enabled         =   -1  'True
         Blanqueado      =   0   'False
      End
      Begin TDS_1.UpDownText txtAlMenosSemiMagicas 
         Height          =   330
         Left            =   2640
         TabIndex        =   136
         Top             =   1650
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         MaxValue        =   0
         MinValue        =   0
         Value           =   0
         Enabled         =   -1  'True
         Blanqueado      =   0   'False
      End
      Begin TDS_1.UpDownText txtAlMenosNoMagicas 
         Height          =   330
         Left            =   2640
         TabIndex        =   135
         Top             =   1290
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         MaxValue        =   0
         MinValue        =   0
         Value           =   0
         Enabled         =   -1  'True
         Blanqueado      =   0   'False
      End
      Begin TDS_1.UpDownText txtMaxRepeRaza 
         Height          =   330
         Left            =   2160
         TabIndex        =   134
         Top             =   840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         MaxValue        =   0
         MinValue        =   0
         Value           =   0
         Enabled         =   -1  'True
         Blanqueado      =   0   'False
      End
      Begin TDS_1.UpDownText txtCantidadMaximaRepeClase 
         Height          =   330
         Left            =   2160
         TabIndex        =   133
         Top             =   450
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         MaxValue        =   0
         MinValue        =   0
         Value           =   0
         Enabled         =   -1  'True
         Blanqueado      =   0   'False
      End
      Begin VB.CheckBox chkAlMenosTrabajadores 
         Appearance      =   0  'Flat
         Caption         =   "Clases trabajadoras al menos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   2400
         Width           =   2415
      End
      Begin VB.CheckBox chkNoRepetirRaza 
         Appearance      =   0  'Flat
         Caption         =   "No repetir raza mas de"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   870
         Width           =   1935
      End
      Begin VB.CheckBox chkNoRepetirClase 
         Appearance      =   0  'Flat
         Caption         =   "No repetir clase más de"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   520
         Width           =   2025
      End
      Begin VB.CheckBox chkSoloPremium 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "(premium)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   40
         Top             =   2880
         Width           =   975
      End
      Begin VB.CheckBox chkSoloPJEnCuentas 
         Appearance      =   0  'Flat
         Caption         =   "Solo personajes en cuentas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2880
         Value           =   1  'Checked
         Width           =   2280
      End
      Begin VB.CheckBox chkNoRepetirClan 
         Appearance      =   0  'Flat
         Caption         =   "Torneo de Clanes. No se puede repetir clan."
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   3735
      End
      Begin VB.CheckBox chkAlMenosNoMagicas 
         Appearance      =   0  'Flat
         Caption         =   "Clases no mágicas al menos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1350
         Width           =   2355
      End
      Begin VB.CheckBox chkAlMenosMagicas 
         Appearance      =   0  'Flat
         Caption         =   "Clases mágicas al menos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   2040
         Width           =   2115
      End
      Begin VB.CheckBox chkAlMenosSemiMagicas 
         Appearance      =   0  'Flat
         Caption         =   "Clases semi mágicas al menos:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1680
         Width           =   2505
      End
      Begin VB.CheckBox chkMaxSumatoriaNiveles 
         Appearance      =   0  'Flat
         Caption         =   "Máxima suma de niveles"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   3360
         Width           =   2025
      End
   End
End
Attribute VB_Name = "frmAdminEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private hechizos() As String
Private objeto() As String

Private reglaHechizos() As Byte

Private Const CANTIDAD_HECHIZOS = 41

'Hechizos
Private Enum eHechizos
    Resucitar = 11
    Provocar_Hambre = 12
    Terrible_Hambre = 13
    
    Invisibilidad = 14
        
    Llamado_naturaleza = 16
    Invocar_Zombies = 17
    Torpeza = 19
    
    Debilidad = 21
    
    Invocar_elemetanl_fuego = 26
    Invocoar_elemental_agua = 27
    Invocoar_elemental_tierra = 28
    Implorar_ayuda = 29
    
    Estupidez = 31
    Ayuda_espiritu_indomable = 33
    Mimetismo = 35
    
    Invocar_Mascotas = 39
End Enum

Private eventoActual As tConfigEvento
Private eventoActualEstado As eEstadoEvento

Private Sub Cerrar_Click()
    Unload Me
End Sub

Private Sub chkApuestasActivadas_Click()
    Dim habilitado As Boolean
    
    If Me.chkApuestasActivadas.value = 1 Then
        habilitado = True
    Else
        habilitado = False
    End If
    
    Me.lblPozoInicial.Enabled = habilitado
    Me.txtApuestasPozoInicial.Enabled = habilitado
    
    Me.txtTiempoApuestas.Enabled = habilitado
    Me.lblTiempoApuestas.Enabled = habilitado

End Sub

Private Sub chkCaenItem_Click()
    
    If Me.chkCaenItem.value = 1 Then
        Me.cmbCaenItemsTipo.Enabled = True
    Else
        Me.cmbCaenItemsTipo.Enabled = False
    End If

End Sub

Private Sub chkClaseX_Click(Index As Integer)
Dim loopCheck As Byte

If Index = 0 Then
    For loopCheck = 1 To Me.chkClaseX.UBound
        Me.chkClaseX(loopCheck).value = Me.chkClaseX(0).value
    Next
End If

End Sub

Private Sub chkGanadorQuedaEnCancha_Click()

    Dim habilitado As Boolean
    
    If Me.chkGanadorQuedaEnCancha.value = 1 Then
        habilitado = True
    Else
        habilitado = False
    End If
    
    Me.lblForsliderGanadorCancha.Enabled = habilitado
    Me.lblSliderDebeEsperar.Enabled = habilitado
    
    Me.sliderGanadorCancha.Enabled = habilitado
    Me.sliderDebeEsperar.Enabled = habilitado
    
End Sub

Private Sub chkHechizo_Click(Index As Integer)

Dim loopCheck As Byte

If Index = 0 Then
    For loopCheck = 1 To Me.chkHechizo.UBound
        Me.chkHechizo(loopCheck).value = Me.chkHechizo(0).value
    Next
End If

End Sub

Private Sub chkRazaX_Click(Index As Integer)
    Dim loopCheck As Byte
    
    If Index = 0 Then
        For loopCheck = 1 To Me.chkRazaX.UBound
            Me.chkRazaX(loopCheck).value = Me.chkRazaX(0).value
        Next
    End If
End Sub

Private Sub cmbRestricciones_Click()
        
    Dim pantalla As Frame
    
    ' Ocultamos todos los frames
    Call ocultarRestricciones
    
    Select Case Me.cmbRestricciones.ListIndex
        Case 0 'Premios
            
            Set pantalla = Me.FraPremios
        Case 1 'Clases Permitidas
            Set pantalla = Me.FraClasesPermitidas
        Case 2 'Hechizos Permitidos
            Set pantalla = Me.frmHechizosPermitidos
        Case 3 'Limite de items
            Set pantalla = Me.FraLimiteDe
        Case 4 'Restricciones de equipo
            Set pantalla = Me.frmRestriccionesGenerales
        Case 5 'Restricciones de los personajes
            Set pantalla = Me.frmRestriccionesPersonaje
        Case 6 ' Otras
            Set pantalla = Me.frmOtros
        Case 7 ' Rings y descansos
           Set pantalla = Me.frmRingsDescansos

    End Select
    
    pantalla.visible = True
    pantalla.top = 38
    pantalla.left = 487
    
End Sub

Private Sub cmdAceptarMensaje_Click()
    Me.frmMensaje.visible = False
End Sub

Private Sub cmdCancelarEvento_Click()
    Me.cmdCancelarEvento.Enabled = False
    
    Call sSendData(Paquetes.ComandosSemi, SemiDios2.CancelarEvento, eventoActual.Nombre)
    Call sSendData(Paquetes.ComandosSemi, SemiDios1.ObtenerEventos)
        
    Me.frmInfoEvento.visible = False
End Sub

Private Sub cmdCerrarInfoEvento_Click()
    Me.frmInfoEvento.visible = False
End Sub

Private Function validarFormulario() As Boolean
    Dim minCantidadEquipos As Byte
    Dim maxCantidadEquipos As Byte
    Dim nombreEvento As String
    Dim cantidadIntegrantesPorEquipo As Byte
    Dim totalPremio As Long
    Dim loopPremio As Integer
    Dim resultado As VbMsgBoxResult
    Dim resultadoChequeo As Boolean
    
    nombreEvento = Trim$(Me.txtNombreTorneo)
    minCantidadEquipos = Me.txtCantidadEquiposTorneoMin.value
    maxCantidadEquipos = Me.txtCantidadEquiposTorneoMax.value
    cantidadIntegrantesPorEquipo = Me.txtCantidadParticipantesTorneo.value
    
    ' Nombre del Evento
    If Len(nombreEvento) = 0 Then
        Call procesarMensaje("Te olvidaste de ponerle el nombre al evento.")
        Exit Function
    End If
    
    If InStr(1, nombreEvento, "-") > 0 Then
        Call procesarMensaje("El nombre del evento no puede tener guiones (-).")
        Exit Function
    End If
    
    'Chequeo la cantidad minima y maxima de equipos
    If minCantidadEquipos < 1 Then
        Call procesarMensaje("La cantidad maxima de equipos debe ser mayor o igual a 1.")
        Exit Function
    End If
    
    If maxCantidadEquipos < 1 Then
        Call procesarMensaje("La cantidad maxima de equipos debe ser mayor o igual a 1.")
        Exit Function
    End If
    
    If maxCantidadEquipos < minCantidadEquipos Then
        Call procesarMensaje("La cantidad maxima de equipos participantes admitidos es menor a la cantidad minima!.")
        Exit Function
    End If
    
    If maxCantidadEquipos > 255 Then
        Call procesarMensaje("¡Demasiados equipos! El máximo es 255.")
        Exit Function
    End If
    
    ' Cantidad de Integrantes por equipo
    If cantidadIntegrantesPorEquipo = 0 Then
        Call procesarMensaje("¿Cuántos integrantes por equipo?")
        Exit Function
    End If
    
    If cantidadIntegrantesPorEquipo > 255 Then
        Call procesarMensaje("Demasiados integrantes por equipo. El máximo es 255.")
        Exit Function
    End If
    
    ' Precios inscripcion
    ' Importancia del evento
    ' Tiempo durante el cual se anuncia
    ' Tiempo que dura la inscripcion
    ' Tiempo tolerancia un usuario cierra.

    ' Premio
    totalPremio = 0
    For loopPremio = Me.txtOroPremio.count - 1 To 0 Step -1
        
        If val(Me.txtOroPremio(loopPremio)) = 0 Then
            
            If totalPremio > 0 And loopPremio > 0 Then
                resultado = MsgBox("El " & (loopPremio + 1) & "º lugar no tiene premio, ¿esta bien?. Porque el puesto número " & loopPremio + 2 & " si entrega.", vbQuestion + vbYesNo)
            
                If resultado = vbNo Then Exit Function
            End If
        
        Else
            totalPremio = totalPremio + CLng(Me.txtOroPremio(loopPremio))
        End If
    Next loopPremio

    ' ¿No entrega nada?
    If totalPremio = 0 Then
        resultado = MsgBox("El evento no entrega premios, ¿esta bien?.", vbQuestion + vbYesNo)
        If resultado = vbNo Then Exit Function
    End If

    If Me.chkTipoPremio.value = vbChecked Then ' ¿Porcentaje?
        If totalPremio > 500 Then
            Call procesarMensaje("Seleccionaste que los premios están entregados en función de la recaudacion. Entregar un 500% por sobre lo recaudado, es demasiado.")
            Exit Function
        End If
    End If

    ' -- Restricciones que afectan al personaje --
    If Not (Me.txtMinNivel.value > 0 And Me.txtMinNivel.value <= STAT_MAXELV) Then
        Call procesarMensaje("¿Nivel mínimo para entrar?.")
        Exit Function
    End If
    
    If Not (Me.txtMaxNivel.value > 0 And Me.txtMaxNivel.value <= STAT_MAXELV) Then
        Call procesarMensaje("¿Nivel máximo para entrar?.")
        Exit Function
    End If
  
    If Not (Me.txtMaxNivel.value >= Me.txtMinNivel.value) Then
        Call procesarMensaje("El nivel mínimo es mayor al nivel maximo")
        Exit Function
    End If

    ' Tipo Cuenta
    ' Clases Permitidas
    ' Razas Permitidas
    ' Alineacion
        
    ' --- Apuestas
    
    
    ' --- Restricciones al equipo
    If Me.chkAlMenosMagicas.value = vbChecked And Me.txtAlMenosMagicas.value < 1 Then
        Call procesarMensaje("Si hay un mínimo de cantidad de personajes de clases MAGICAS no tiene sentido que sea menor a 1!.")
        Exit Function
    End If
           
    If Me.chkAlMenosNoMagicas.value = vbChecked And Me.txtAlMenosNoMagicas.value < 1 Then
        Call procesarMensaje("Si hay un mínimo de cantidad de personajes de clases NO MAGICAS no tiene sentido que sea menor a 1!.")
        Exit Function
    End If
    
    If Me.chkAlMenosSemiMagicas.value = vbChecked And Me.txtAlMenosSemiMagicas.value < 1 Then
        Call procesarMensaje("Si hay un mínimo de cantidad de personajes de clases SEMI MAGICAS no tiene sentido que sea menor a 1!.")
        Exit Function
    End If

    ' --- Identificacion del equipo
    If Me.cmbIdentificarEquipos.ListIndex = -1 Then
        Call procesarMensaje("Tenes que seleccionar la manera con la cual se identificiará a los equipos.")
        Exit Function
    End If
        
    ' El evento en si
    If Me.OptEvento(0).value = True Then
        ' Automatico
        resultadoChequeo = validarTorneoAutomatico
    Else
        ' Manual
        resultadoChequeo = validarTorneoManual
    End If
    
    If Not resultadoChequeo Then Exit Function
            
    validarFormulario = True
End Function


Private Function validarListaPersonajes(ByVal equiposMax As Byte, ByVal integrantesPorEquipo As Byte) As Boolean

' Ingreso manual de participantes
Dim infoManual As String
Dim loopEquipo As Byte
Dim loopIntegrante As Byte
Dim infoEquipos() As String
Dim infoEquipo() As String
Dim respuesta As VbMsgBoxResult

'Tengo que obtener una lista con los equipos y saber, quienes estan offline y cuantos faltan
infoManual = Trim$(Me.txtEquiposManual.text)

Call quitarEnters(infoManual)
Call quitarEntersAdelante(infoManual)
Call quitarDobleEnters(infoManual)
    
If Len(infoManual) = 0 Then
    Call procesarMensaje("Tenés que ingresar al menos un equipo con el formato: Nombre personaje 1, nombre personaje2, etc.")
    Exit Function
End If

Me.txtEquiposManual.text = infoManual
        
'Chequeamos que la lista este ok
infoEquipos = Split(infoManual, vbCrLf)
        
'Chequeamos la cantidad de equipos
If (UBound(infoEquipos) + 1 > equiposMax) Then
    Call procesarMensaje("La cantidad de equipos que ingresaste (" & (UBound(infoEquipos) + 1) & ") es mayor a la cantidad máxima de participantes (" & equiposMax & ") que puede tener el evento.")
    Exit Function
End If

'Chequeamos la cantidad de equipos
If (UBound(infoEquipos) + 1 < equiposMax) Then
    respuesta = MsgBox("La cantidad de equipos ingresada (" & UBound(infoEquipos) + 1 & ") es menor a la cantidad máxima de participantes (" & equiposMax & "). Se abrirán las inscripciones para los integrantes faltantes. ¿Está bien esto?", vbYesNo + vbQuestion)
    
    If respuesta = vbNo Then Exit Function
End If

'Chequeamos la cantidad de integrantes
For loopEquipo = LBound(infoEquipos) To UBound(infoEquipos)
    infoEquipo = Split(infoEquipos(loopEquipo), ",")
    
    ' Cuando se pueda poner una cantidad minima/maxima de integrantes por equipo
    ' se debe chequear que la cantidad este entre ambos (si el split no devuelve nada, sera-1)
    If UBound(infoEquipo) + 1 = integrantesPorEquipo Then
        ' Chequeo los equipos
        For loopIntegrante = LBound(infoEquipo) To UBound(infoEquipo)
            If Trim$(infoEquipo(loopIntegrante)) = "" Then
                Call procesarMensaje("Falta el integrante " & loopIntegrante + 1 & " del equipo " & loopEquipo + 1 & ".")
                Exit Function
            End If
        Next loopIntegrante
    Else
        Call procesarMensaje("La cantidad de integrantes del equipo " & loopEquipo + 1 & " (" & (UBound(infoEquipo) + 1) & ") no corresponde con la cantidad de integrantes (" & integrantesPorEquipo & ") configuradas en el evento.")
        Exit Function
    End If
Next
 
validarListaPersonajes = True

End Function
Private Function validarTorneoManual() As Boolean

    validarTorneoManual = True
    
End Function

Private Function validarTorneoAutomatico() As Boolean
    Dim auxCoeficiente As Double

    validarTorneoAutomatico = False
    
    If Me.txtCantidadEquiposTorneoMin.value <= 1 Then
        Call procesarMensaje("Sí es un evento automatico tiene que haber al menos dos participantes.")
        Exit Function
    End If
    
    If (Me.txtAlMejorDe.value Mod 2) = 0 Then
        Call procesarMensaje("Mal puesta AL MEJOR DE. No se puede una cantidad par sino puede haber empate.")
        Exit Function
    End If
    
    ' ¿PlayOff?
    If Me.OptTipoSubEvento(1).value = True Then
        Me.txtCantidadEquiposTorneoMin.value = Me.txtCantidadEquiposTorneoMax.value
        
         auxCoeficiente = Log(Me.txtCantidadEquiposTorneoMax.value) / Log(2)
         auxCoeficiente = auxCoeficiente - Int(auxCoeficiente)
         
        If auxCoeficiente > 0 Then
            Call procesarMensaje("Esta mal la cantidad de equipos participantes. Debe ser potencia de 2. Esto es 2, 4, 8, 16, 32, 64, 128, 256.")
            Exit Function
        End If
    End If
        
    validarTorneoAutomatico = True

End Function

Private Function formularioAEstructura(configEvento As tConfigEvento) As Boolean
    Dim loopC As Integer
    Dim cantidad As Integer
    Dim TempByte As Byte
    Dim tempBool As Boolean
    
    ' Generales
    configEvento.Nombre = Me.txtNombreTorneo.text
    configEvento.descripcion = Me.txtDescripcionTorneo.text
    
    configEvento.cantEquiposMinimo = Me.txtCantidadEquiposTorneoMin.value
    configEvento.cantEquiposMaxima = Me.txtCantidadEquiposTorneoMax.value
    
    If Me.txtPrecioInscripcionTorneo.text = "" Then
        configEvento.costoInscripcion = 0
    Else
        configEvento.costoInscripcion = CLng(Me.txtPrecioInscripcionTorneo.text)
    End If
    
    configEvento.cantidadIntegrantesEquipo = Me.txtCantidadParticipantesTorneo.value
    
    ' Tiempos
    configEvento.tiempoAnuncio = Me.txtTiempoAviso.value
    configEvento.tiempoInscripcion = Me.txtTiempoInscripcion.value
    configEvento.tiempoTolerancia = Me.sldMinutosToleranciaDeslogueo.value
    
    ' Apuestas
    configEvento.apuestas.activadas = (Me.chkApuestasActivadas.value = 1)
    
    If configEvento.apuestas.activadas Then
        configEvento.apuestas.pozoInicial = val(Me.txtApuestasPozoInicial.text)
        configEvento.apuestas.tiempoAbiertas = Me.txtTiempoApuestas.value
    Else
        configEvento.apuestas.pozoInicial = 0
        configEvento.apuestas.tiempoAbiertas = 0
    End If

    ' Importancia del evento
    configEvento.importanciaEvento = ComEstrellasEvento.itemData(ComEstrellasEvento.ListIndex)

    ' Premios
    ' --- Cuento la cantidad de premios validos
    
    If Me.chkTipoPremio.value = vbUnchecked Then
        configEvento.premio.tipo = monedasDeOro
    Else
        configEvento.premio.tipo = porcentajeSobreAcumulado
    End If
    
    If Me.txtOroPremio.count > 1 Then
        cantidad = Me.txtOroPremio.count - 1
        ReDim configEvento.premio.valores(1 To cantidad)
               
        For loopC = 1 To cantidad
            If Not Me.txtOroPremio(loopC - 1).text = "" Then
                configEvento.premio.valores(loopC) = CLng(Me.txtOroPremio(loopC - 1).text)
            Else
                configEvento.premio.valores(loopC) = 0
            End If
        Next
    Else
        ReDim configEvento.premio.valores(1 To 1)
        configEvento.premio.valores(1) = 0
    End If
    
    ' Ring
    TempByte = 4 ' La caracteristica inicial es que sea un ring de torneo
    If Me.chkPlantado.value = vbChecked Then TempByte = TempByte Or 8
    If Me.chkRingAcuatico.value = vbChecked Then TempByte = TempByte Or 16
    configEvento.tipoRing = TempByte
    
    ' Descanso
    TempByte = 1 ' La caracteristica inicial es que sea un descnso de torneo
    If Me.chkDescansoBoveda.value = vbChecked Then TempByte = TempByte Or 4
    configEvento.tipoDescanso = TempByte

    ' Identificacion del equipo
    configEvento.comoIdentificarEquipo = Me.cmbIdentificarEquipos.itemData(Me.cmbIdentificarEquipos.ListIndex)

    ' Restricciones para el Personaje
    ' - Nivel
    If Me.txtMinNivel.value = 1 And Me.txtMaxNivel.value = STAT_MAXELV Then
        configEvento.restriccionesPersonaje.Nivel.activada = False
    Else
        configEvento.restriccionesPersonaje.Nivel.activada = True
        configEvento.restriccionesPersonaje.Nivel.minimo = Me.txtMinNivel.value
        configEvento.restriccionesPersonaje.Nivel.maximo = Me.txtMaxNivel.value
    End If

    ' - Alineacion
    If Me.chkCiudadanos.value = 1 Then
        configEvento.restriccionesPersonaje.Alineacion.activada = True
        configEvento.restriccionesPersonaje.Alineacion.ciudadano = True
    End If
    
    If Me.chkCriminales.value = 1 Then
        configEvento.restriccionesPersonaje.Alineacion.activada = True
        configEvento.restriccionesPersonaje.Alineacion.criminal = True
    End If
    
    If Me.chkLegionarios.value = 1 Then
        configEvento.restriccionesPersonaje.Alineacion.activada = True
        configEvento.restriccionesPersonaje.Alineacion.caos.activada = True
        configEvento.restriccionesPersonaje.Alineacion.caos.cantidad = Me.sliderLegionRango.value
    End If
    
    If Me.chkArmadas.value = 1 Then
        configEvento.restriccionesPersonaje.Alineacion.activada = True
        configEvento.restriccionesPersonaje.Alineacion.armada.activada = True
        configEvento.restriccionesPersonaje.Alineacion.armada.cantidad = Me.sliderArmadaRango.value
    End If
    
    ' - Clases Permitidas
    tempBool = True ' todas
    For loopC = 1 To 15
        configEvento.restriccionesPersonaje.Clase.clasesPermitidas(loopC) = IIf(Me.chkClaseX(loopC).value = vbChecked, 1, 0)
        tempBool = (tempBool And configEvento.restriccionesPersonaje.Clase.clasesPermitidas(loopC))
    Next
    ' Si estan todas activadas (tempbool = true) no tiene sentido esta restriccion
    configEvento.restriccionesPersonaje.Clase.activada = Not tempBool
    
    ' - Razas Permitidas
    tempBool = True
    For loopC = 1 To 5
        configEvento.restriccionesPersonaje.Raza.razasPermitidas(loopC) = IIf(Me.chkRazaX(loopC).value = vbChecked, 1, 0)
        tempBool = (tempBool And configEvento.restriccionesPersonaje.Raza.razasPermitidas(loopC))
    Next
    
    ' Si estan todas activadas (tempbool = true) no tiene sentido esta restriccion
    configEvento.restriccionesPersonaje.Raza.activada = Not tempBool
    
    ' - Personajes en Cuentas
    configEvento.restriccionesPersonaje.tipoCuenta = ninguna
    
    If Me.chkSoloPJEnCuentas.value = vbChecked Then
        If Me.chkSoloPremium.value = vbChecked Then
            configEvento.restriccionesPersonaje.tipoCuenta = eCuenta.premium
        Else
            configEvento.restriccionesPersonaje.tipoCuenta = eCuenta.todas
        End If
    End If

    ' - Iventario
    Call limiteDeItemsAConfig(configEvento.restriccionesPersonaje.inventario)
    
    ' Restricciones del equipo
    
    ' - ¿Torneo de Clanes?
    configEvento.restriccionesEquipo.repeticionClan.activada = (Me.chkNoRepetirClan.value = vbChecked)
    configEvento.restriccionesEquipo.repeticionClan.cantidad = 0
    
    ' - Repetir clase.
    configEvento.restriccionesEquipo.repeticionClase.activada = (Me.chkNoRepetirClase.value = vbChecked)
    configEvento.restriccionesEquipo.repeticionClase.cantidad = Me.txtCantidadMaximaRepeClase.value
    
    ' - Repetir raza.
    configEvento.restriccionesEquipo.repeticionRaza.activada = (Me.chkNoRepetirRaza.value = vbChecked)
    configEvento.restriccionesEquipo.repeticionRaza.cantidad = Me.txtMaxRepeRaza.value
    
    ' - Sumatoria de Niveles
    configEvento.restriccionesEquipo.limiteSumaDeNivel.activada = (Me.chkMaxSumatoriaNiveles.value = vbChecked)
    configEvento.restriccionesEquipo.limiteSumaDeNivel.cantidad = txtMaxSumatoriaNiveles.value
    
    ' - Limite (mininimo)
    ' -- Clases No magicas
    If Me.chkAlMenosNoMagicas.value = vbChecked Or Me.chkAlMenosSemiMagicas.value = vbChecked Or Me.chkAlMenosMagicas.value = vbChecked Then
        configEvento.restriccionesEquipo.grupoClases.activada = True
        
        ' -- Clases magicas
        If Me.chkAlMenosNoMagicas.value = vbChecked Then
            configEvento.restriccionesEquipo.grupoClases.noMagicas = txtAlMenosNoMagicas.value
        Else
            configEvento.restriccionesEquipo.grupoClases.noMagicas = 0
        End If
        
        ' -- Clase semi magicas
        If Me.chkAlMenosSemiMagicas.value = vbChecked Then
            configEvento.restriccionesEquipo.grupoClases.semiMagicas = txtAlMenosSemiMagicas.value
        Else
            configEvento.restriccionesEquipo.grupoClases.semiMagicas = 0
        End If
    
        ' -- Clases no  magicas
        If Me.chkAlMenosMagicas.value = vbChecked Then
            configEvento.restriccionesEquipo.grupoClases.magicas = txtAlMenosMagicas.value
        Else
            configEvento.restriccionesEquipo.grupoClases.magicas = 0
        End If
        
        ' -- Clases Trabajadoras
        If Me.chkAlMenosTrabajadores.value = vbChecked Then
            configEvento.restriccionesEquipo.grupoClases.trabajadoras = chkAlMenosTrabajadores.value
        Else
            configEvento.restriccionesEquipo.grupoClases.trabajadoras = 0
        End If

    Else

    End If
    
    ' Reglas del evento
    For loopC = 1 To CANTIDAD_HECHIZOS
        configEvento.reglas.hechizos(loopC) = (Me.chkHechizo(loopC).value = vbChecked)
    Next loopC
        
    ' Evento
    If Me.OptEvento(0).value = True Then ' Evento Automatico
    
        configEvento.automatico = True
        configEvento.configAutomatico.maxsRounds = Me.txtAlMejorDe.value
        
        ' El evento se repite una y otra vez hasta que alguien lo gane N veces?
        configEvento.configAutomatico.configCircular.activado = (Me.chkGanadorQuedaEnCancha.value = vbChecked)
 
        If configEvento.configAutomatico.configCircular.activado Then
            configEvento.configAutomatico.configCircular.cantidadAGanar = Me.sliderGanadorCancha.value
            configEvento.configAutomatico.configCircular.eventosExcluido = Me.sliderDebeEsperar.value
        End If
        
        If Me.OptTipoSubEvento(0).value = True Then
        
            configEvento.configAutomatico.tipo = eEventoTipoAutomatico.deathmatch
            
        ElseIf Me.OptTipoSubEvento(2).value = True Then
            
            configEvento.configAutomatico.tipo = eEventoTipoAutomatico.liga
            configEvento.configAutomatico.ligaConfig.conVuelta = (Me.chkConVuelta.value = vbChecked)
                        
        ElseIf Me.OptTipoSubEvento(1).value = True Then
            
            configEvento.configAutomatico.tipo = eEventoTipoAutomatico.playoff
            configEvento.configAutomatico.playOffConfig.clasificacionCompleta = (Me.chkClasificacionCompleta.value = vbChecked)
            
        End If
        
    ElseIf Me.OptEvento(1).value = True Then ' Evento manual. Solo se utiliza el sistema de inscripcion.

        configEvento.automatico = False
        configEvento.configManual.transportarInmediato = Me.optSinSubEvento(0).value
        
    End If
    
    formularioAEstructura = True
End Function

Private Sub limiteDeItemsAConfig(objetos As tEventoRestriccionObjetos)

Dim cantidadObjetos As Integer
Dim loopC As Byte
Dim loopReal As Byte

objetos.BilleteraVacia = (Me.chkNoPermitirOroBilletera.value = vbChecked)
objetos.restringir = (Me.chkNoPermitirItems.value = vbChecked)

' ¿Objetos?
If GridTextListaObjetos.obtenerCantidadCampos = 0 Then
    cantidadObjetos = 0
Else
    For loopC = 0 To Me.GridTextListaObjetos.obtenerCantidadCampos - 1
        If Me.GridTextListaObjetos.obtenerID(loopC) > 0 Then
            cantidadObjetos = cantidadObjetos + 1
        End If
    Next
End If

' ¿Esta activada?
If objetos.BilleteraVacia = False And objetos.restringir = False And cantidadObjetos = 0 Then
    objetos.activada = False
    Exit Sub
End If

' Marcamos como activada la restriccion
objetos.activada = True

If cantidadObjetos = 0 Then
    ReDim objetos.objetos(1 To 1) As tEventoObjetoRestringido
    objetos.objetos(1).id = 0
    Exit Sub
End If

' Cargamos los objetos limitados
ReDim objetos.objetos(1 To cantidadObjetos) As tEventoObjetoRestringido

loopReal = 1
For loopC = 0 To Me.GridTextListaObjetos.obtenerCantidadCampos - 1

    If Me.GridTextListaObjetos.obtenerID(loopC) > 0 Then

        objetos.objetos(loopReal).id = CInt(Me.GridTextListaObjetos.obtenerID(loopC))
        objetos.objetos(loopReal).cantidad = CInt(val(Me.GridTextListaObjetos.getValorDinamico("txtCantidad", loopC)))
        
        If Me.GridTextListaObjetos.getValorDinamico("chkMinimo", loopC) = 1 Then
            objetos.objetos(loopReal).tipo = eRangoLimite.minimo
        Else
            objetos.objetos(loopReal).tipo = eRangoLimite.maximo
        End If

        loopReal = loopReal + 1
    End If
        
Next

End Sub
Private Sub enviarConfiguracionDeEvento(configEvento As tConfigEvento)

''Creo el Evento
Dim infoEvento As String
Dim parametros() As String

Dim tempBool As Boolean
Dim tempString As String
Dim TempByte As Byte
Dim loopC As Integer

' Version del formulario
infoEvento = ByteToString(1)

If configEvento.automatico = True Then
    
    ' Evento automatico
    infoEvento = infoEvento & "¦¦" & ByteToString(0)
        
    ' Combates al mejor de
    infoEvento = infoEvento & ByteToString(configEvento.configAutomatico.maxsRounds)
    
    ' ¿Por los items?
    If configEvento.configAutomatico.objetosEnJuego.activado Then
        infoEvento = infoEvento & ByteToString(configEvento.configAutomatico.objetosEnJuego.cuando)
    Else
        infoEvento = infoEvento & ByteToString(eEventoCaenItems.nunca)
    End If
    
    ' ¿Evento circular?
    If configEvento.configAutomatico.configCircular.activado Then
        infoEvento = infoEvento & ByteToString(1) & ByteToString(configEvento.configAutomatico.configCircular.cantidadAGanar) & ByteToString(configEvento.configAutomatico.configCircular.eventosExcluido)
    Else
        infoEvento = infoEvento & ByteToString(0)
    End If
            
    ' Tipo de organizacion del evento y configuracion de este
    If configEvento.configAutomatico.tipo = eEventoTipoAutomatico.deathmatch Then
        infoEvento = infoEvento & ByteToString(eEventoTipoAutomatico.deathmatch)
    ElseIf configEvento.configAutomatico.tipo = eEventoTipoAutomatico.liga Then
        infoEvento = infoEvento & ByteToString(eEventoTipoAutomatico.liga) & configEvento.configAutomatico.ligaConfig.conVuelta & ";"
    ElseIf configEvento.configAutomatico.tipo = eEventoTipoAutomatico.playoff Then
        infoEvento = infoEvento & ByteToString(eEventoTipoAutomatico.playoff) & configEvento.configAutomatico.playOffConfig.clasificacionCompleta & ";"
    End If
        
Else
    ' Solo inscripcion
    infoEvento = infoEvento & "¦¦" & ByteToString(1) & ByteToString(IIf(configEvento.configManual.transportarInmediato, 1, 0))
End If

''******************************************************************************
'' CONFIGURACION GENERAL
''******************************************************************************
infoEvento = infoEvento & "¦¦" & configEvento.Nombre
infoEvento = infoEvento & "¦¦" & ByteToString(configEvento.importanciaEvento) & configEvento.descripcion
infoEvento = infoEvento & "¦¦" & configEvento.cantEquiposMinimo & "¦¦" & configEvento.cantEquiposMaxima
infoEvento = infoEvento & "¦¦" & configEvento.cantidadIntegrantesEquipo
infoEvento = infoEvento & "¦¦" & configEvento.costoInscripcion
infoEvento = infoEvento & "¦¦" & configEvento.tiempoAnuncio & "¦¦" & configEvento.tiempoInscripcion & "¦¦" & configEvento.tiempoTolerancia

'Pagos
infoEvento = infoEvento & "¦¦" & ByteToString(configEvento.premio.tipo) & ByteToString(UBound(configEvento.premio.valores))

For loopC = 1 To UBound(configEvento.premio.valores)
    infoEvento = infoEvento & LongToString(configEvento.premio.valores(loopC))
Next

'
infoEvento = infoEvento & "¦¦" & Chr$(configEvento.tipoRing)
infoEvento = infoEvento & "¦¦" & Chr$(configEvento.tipoDescanso)
infoEvento = infoEvento & "¦¦" & Chr$(configEvento.comoIdentificarEquipo)

''******************************************************************************
'' Condiciones variables
''******************************************************************************

'Apuestas
If configEvento.apuestas.activadas = True Then
   infoEvento = infoEvento & "¦¦" & eEventoCondicion.apuestasActivadas & ";" & LongToString(configEvento.apuestas.pozoInicial) & ByteToString(configEvento.apuestas.tiempoAbiertas)
End If

' Restricciones del Equipo

' - Clan
If configEvento.restriccionesEquipo.repeticionClan.activada Then
    infoEvento = infoEvento & "¦¦" & eEventoCondicion.clanRepetir
End If

' - Clase
If configEvento.restriccionesEquipo.repeticionClase.activada Then
    infoEvento = infoEvento & "¦¦" & eEventoCondicion.claseRepetir & ";" & ByteToString(configEvento.restriccionesEquipo.repeticionClase.cantidad)
End If

' - Raza
If configEvento.restriccionesEquipo.repeticionRaza.activada Then
    infoEvento = infoEvento & "¦¦" & eEventoCondicion.razaRepetir & ";" & ByteToString(configEvento.restriccionesEquipo.repeticionRaza.cantidad)
End If

' - Sumatoria de niveles
If configEvento.restriccionesEquipo.limiteSumaDeNivel.activada Then
    infoEvento = infoEvento & "¦¦" & eEventoCondicion.nivelesSumatoria & ";" & ITS(configEvento.restriccionesEquipo.limiteSumaDeNivel.cantidad)
End If

' - Minima cantidad de personajes por grupo de clase
If configEvento.restriccionesEquipo.grupoClases.activada Then

    infoEvento = infoEvento & "¦¦" & eEventoCondicion.clasesGrupo & ";"
    
    With configEvento.restriccionesEquipo.grupoClases
    infoEvento = infoEvento & ByteToString(.magicas) & ByteToString(.semiMagicas) & ByteToString(.noMagicas) & ByteToString(.trabajadoras)
    End With

End If
    
' Restricciones de Personajes

' - Restriccion de Nivel
If configEvento.restriccionesPersonaje.Nivel.activada Then
    infoEvento = infoEvento & "¦¦" & eEventoCondicion.nivelMinMax & ";" & ByteToString(configEvento.restriccionesPersonaje.Nivel.minimo) & ByteToString(configEvento.restriccionesPersonaje.Nivel.maximo)
End If

' - Personajes en Cuentas. Por defecto no hay restriccion.
If Not configEvento.restriccionesPersonaje.tipoCuenta = ninguna Then
    infoEvento = infoEvento & "¦¦" & eEventoCondicion.personajesCuenta & ";" & Chr$(configEvento.restriccionesPersonaje.tipoCuenta)
End If

' - Clases
If configEvento.restriccionesPersonaje.Clase.activada Then
    tempString = ""
    For loopC = 1 To 15
        If configEvento.restriccionesPersonaje.Clase.clasesPermitidas(loopC) = False Then
            tempString = tempString & Chr$(loopC)
        End If
    Next

    infoEvento = infoEvento & "¦¦" & eEventoCondicion.clasesPermitidas & ";" & tempString
End If

' - Razas. Por defecto todas pueden entrar.
If configEvento.restriccionesPersonaje.Raza.activada Then
    tempString = ""
    For loopC = 1 To 5
        If configEvento.restriccionesPersonaje.Raza.razasPermitidas(loopC) = False Then
           tempString = tempString & Chr$(loopC)
        End If
    Next

    infoEvento = infoEvento & "¦¦" & eEventoCondicion.razasPermitidas & ";" & tempString
End If

' - Alineaciones permitidas.
If configEvento.restriccionesPersonaje.Alineacion.activada Then
    
    TempByte = 0
    If configEvento.restriccionesPersonaje.Alineacion.ciudadano Then TempByte = (TempByte Or eEventoPersonajesAlineacion.Ciudadanos)

    If configEvento.restriccionesPersonaje.Alineacion.criminal Then TempByte = (TempByte Or eEventoPersonajesAlineacion.criminales)

    If configEvento.restriccionesPersonaje.Alineacion.caos.activada = True Then TempByte = (TempByte Or eEventoPersonajesAlineacion.Legionarios)
    
    If configEvento.restriccionesPersonaje.Alineacion.armada.activada = True Then TempByte = (TempByte Or eEventoPersonajesAlineacion.Armadas)

    infoEvento = infoEvento & "¦¦" & eEventoCondicion.alineacionesPermitidas & ";" & ByteToString(TempByte) & ByteToString(configEvento.restriccionesPersonaje.Alineacion.armada.cantidad) & ByteToString(configEvento.restriccionesPersonaje.Alineacion.caos.cantidad)
End If

' - Objetos
If configEvento.restriccionesPersonaje.inventario.activada Then

    infoEvento = infoEvento & "¦¦" & eEventoCondicion.objetosPermitidos & ";" & ByteToString(IIf(configEvento.restriccionesPersonaje.inventario.restringir, 1, 0)) & ByteToString(IIf(configEvento.restriccionesPersonaje.inventario.BilleteraVacia, 1, 0))

    For loopC = LBound(configEvento.restriccionesPersonaje.inventario.objetos) To UBound(configEvento.restriccionesPersonaje.inventario.objetos)
    
        If configEvento.restriccionesPersonaje.inventario.objetos(loopC).id > 0 Then
    
            With configEvento.restriccionesPersonaje.inventario.objetos(loopC)
                infoEvento = infoEvento & ";" & ITS(.id) & ITS(.cantidad) & ByteToString(.tipo)
            End With
        End If
    Next
       
End If

' Reglas durante el juego
' - Hechizos
infoEvento = infoEvento & "¦¦" & eEventoCondicion.hechizosPermitidos & ";"

For loopC = 1 To UBound(configEvento.reglas.hechizos)
    If Not configEvento.reglas.hechizos(loopC) Then
        infoEvento = infoEvento & ByteToString(loopC)
    End If
Next loopC

Call sSendData(Paquetes.ComandosSemi, SemiDios2.CrearEvento, infoEvento, True)
Call sSendData(Paquetes.ComandosSemi, SemiDios1.ObtenerEventos, "", True)

End Sub

Private Sub cmdDeathCrear_Click()

' Creamos la confirmacion
Dim configEvento As tConfigEvento

' Revisa configuracion
If Not validarFormulario() Then Exit Sub

' La cargamos
If Not formularioAEstructura(configEvento) Then Exit Sub

' La enviamos
Call enviarConfiguracionDeEvento(configEvento)

End Sub

Private Sub quitarEnters(cadena As String)
    If InStrRev(cadena, vbCrLf) = Len(cadena) - 1 Then
        cadena = mid$(cadena, 1, Len(cadena) - 2)
        Call quitarEnters(cadena)
    End If
End Sub

Private Sub quitarEntersAdelante(cadena As String)
    If InStr(1, cadena, vbCrLf) = 1 Then
        cadena = mid(cadena, Len(vbCrLf) + 1)
        Call quitarEntersAdelante(cadena)
    End If
End Sub

Private Sub quitarDobleEnters(cadena As String)
    If InStr(1, cadena, vbCrLf & vbCrLf) > 0 Then
        cadena = Replace$(cadena, vbCrLf & vbCrLf, vbCrLf)
        Call quitarDobleEnters(cadena)
    End If
End Sub

Private Sub cmdInscribir_Click()
 If Not validarListaPersonajes(eventoActual.cantEquiposMaxima, eventoActual.cantidadIntegrantesEquipo) Then Exit Sub
 
 Call inscribirParticipantes(False)
End Sub

Private Sub inscribirParticipantes(testeo As Boolean)
    Dim TempStr As String

    TempStr = ByteToString(IIf(testeo, 1, 0)) & ByteToString(Len(eventoActual.Nombre)) & eventoActual.Nombre & Replace$(Me.txtEquiposManual.text, vbCrLf, "-")

    Call sSendData(Paquetes.ComandosSemi, SemiDios2.inscribirEvento, TempStr)
    Call sSendData(Paquetes.ComandosSemi, SemiDios2.ObtenerInfoEvento, eventoActual.Nombre)
End Sub

Private Sub cmdInscribirManual_Click()
    Me.frmInscribirParticipantes.visible = True
    Me.cmdInscribirManual.Enabled = False
    Me.cmdPublicar.Enabled = False
    Me.cmdCerrarInfoEvento.Enabled = False
End Sub

Private Sub cmdMensaje_Click()
    Me.frmMensaje.visible = False
End Sub

Private Sub cmdPublicar_Click()
    Dim respuesta As VbMsgBoxResult
    
    respuesta = MsgBox("¿Estás seguro que queres publicar el evento?", vbQuestion + vbYesNo, eventoActual.Nombre)
    
    If respuesta = vbNo Then Exit Sub
    
    Me.cmdPublicar.Enabled = False
    Call sSendData(Paquetes.ComandosSemi, SemiDios2.publicarEvento, eventoActual.Nombre)
    Call sSendData(Paquetes.ComandosSemi, SemiDios2.ObtenerInfoEvento, eventoActual.Nombre)
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVerEventos_Click()
   Me.cmdVerEventos.Enabled = False
   Me.cmdVerInfo.Enabled = False
   Me.lstEstadoEventos.ListIndex = -1
   Call sSendData(Paquetes.ComandosSemi, SemiDios1.ObtenerEventos, "")
End Sub
Private Sub ocultarRestricciones()
    Me.frmHechizosPermitidos.visible = False
    Me.frmRestriccionesGenerales.visible = False
    Me.FraPremios.visible = False
    Me.FraLimiteDe.visible = False
    Me.FraClasesPermitidas.visible = False
    Me.frmRestriccionesPersonaje.visible = False
    Me.frmRingsDescansos.visible = False
    Me.frmOtros.visible = False
End Sub

Private Sub cmdVerificarInscripcionManual_Click()
 If Not validarListaPersonajes(eventoActual.cantEquiposMaxima, eventoActual.cantidadIntegrantesEquipo) Then Exit Sub
  
 Call inscribirParticipantes(True)
End Sub

Private Sub cmdVerInfo_Click()
    Dim nombreEvento As String
    Dim finNombre As Byte
    
    If Me.lstEstadoEventos.ListIndex >= 0 Then
        
        nombreEvento = Me.lstEstadoEventos.list(Me.lstEstadoEventos.ListIndex)
        finNombre = InStr(1, nombreEvento, "-")
        
        If finNombre > 1 Then
            nombreEvento = mid$(nombreEvento, 1, finNombre - 1)
            
            Call sSendData(Paquetes.ComandosSemi, SemiDios2.ObtenerInfoEvento, nombreEvento)

            Exit Sub
        End If
        
    End If

    Me.cmdVerInfo.Enabled = False
    Call procesarMensaje("Debe seleccionar un evento para obtener la info.")
End Sub

Public Sub parsearInfoEvento(datos As String)
    Dim longitudNombre As String
    
    eventoActualEstado = StringToByte(datos, 1)
    eventoActual.cantidadIntegrantesEquipo = StringToByte(datos, 2)
    eventoActual.cantEquiposMaxima = StringToByte(datos, 3)
    
    longitudNombre = StringToByte(datos, 4)
    eventoActual.Nombre = mid$(datos, 5, longitudNombre)
    mostrarInfoEvento (mid$(datos, 6 + longitudNombre - 1))
End Sub
Private Sub mostrarInfoEvento(info As String)

    If eventoActualEstado = esperandoConfirmacionInicio Then
        Me.cmdPublicar.visible = True
        Me.cmdInscribirManual.visible = True
        
        Me.cmdPublicar.Enabled = True
        Me.cmdInscribirManual.Enabled = True
        
    Else
        Me.cmdPublicar.visible = False
        Me.cmdInscribirManual.visible = False
        Me.cmdPublicar.Enabled = False
        Me.cmdInscribirManual.Enabled = False
    End If

    Me.cmdCancelarEvento.Enabled = True
    
    Me.frmInfoEvento.visible = True
    Me.frmInfoEvento.top = 0
    Me.frmInfoEvento.left = 120
    
    Me.rtbInfoEvento.text = info


    Me.frmInfoEvento.Caption = "Información del evento " & eventoActual.Nombre & "."
End Sub
Private Sub cmdVolverFromInscripcion_Click()
    Me.cmdInscribirManual.Enabled = True
    Me.cmdPublicar.Enabled = True
    Me.cmdCerrarInfoEvento.Enabled = True
    Me.frmInscribirParticipantes.visible = False
End Sub

Private Sub Form_Load()
    

Dim loopHechizo As Byte

' Importancia del evento
Call Me.ComEstrellasEvento.AddItem("1 - Sin importancia")
Me.ComEstrellasEvento.itemData(ComEstrellasEvento.NewIndex) = 1
Call Me.ComEstrellasEvento.AddItem("2 - Evento Reducido")
Me.ComEstrellasEvento.itemData(ComEstrellasEvento.NewIndex) = 2
Call Me.ComEstrellasEvento.AddItem("3 - Evento Normal")
Me.ComEstrellasEvento.itemData(ComEstrellasEvento.NewIndex) = 3
Call Me.ComEstrellasEvento.AddItem("4 - Gran Evento")
Me.ComEstrellasEvento.itemData(ComEstrellasEvento.NewIndex) = 4
Call Me.ComEstrellasEvento.AddItem("5 - Evento Historico")
Me.ComEstrellasEvento.itemData(ComEstrellasEvento.NewIndex) = 5

Me.ComEstrellasEvento.ListIndex = 2 ' Importancia por defecto

' Restricciones
Call Me.cmbRestricciones.AddItem("Premios", 0)
Call Me.cmbRestricciones.AddItem("Clases y Razas Permitidas.", 1)
Call Me.cmbRestricciones.AddItem("Hechizos permitidos.", 2)
Call Me.cmbRestricciones.AddItem("Limite de objetos y oro.", 3)
Call Me.cmbRestricciones.AddItem("Restricciones de Equipo.", 4)
Call Me.cmbRestricciones.AddItem("Restricciones del Personaje.", 5)
Call Me.cmbRestricciones.AddItem("Rings y Descansos.", 6)
Call Me.cmbRestricciones.AddItem("Otras configuraciones.", 6)

' Tipo de caida de Objetos
Call Me.cmbCaenItemsTipo.AddItem("Al finalizar el evento")
Me.cmbCaenItemsTipo.itemData(Me.cmbCaenItemsTipo.NewIndex) = eEventoCaenItems.alFinalizarEvento

Call Me.cmbCaenItemsTipo.AddItem("Al finalizar el combate")
Me.cmbCaenItemsTipo.itemData(Me.cmbCaenItemsTipo.NewIndex) = eEventoCaenItems.alFinalizarCombate

' Identificacion de los equipos
Call Me.cmbIdentificarEquipos.AddItem("Nombre de los personajes")
Me.cmbIdentificarEquipos.itemData(Me.cmbIdentificarEquipos.NewIndex) = eEventoIdentificacionEquipo.identificaPersonajes
Me.cmbIdentificarEquipos.ListIndex = Me.cmbIdentificarEquipos.NewIndex ' Selección por defecto

Call Me.cmbIdentificarEquipos.AddItem("Nombre del Clan")
Me.cmbIdentificarEquipos.itemData(Me.cmbIdentificarEquipos.NewIndex) = eEventoIdentificacionEquipo.identificaClan

Call Me.cmbIdentificarEquipos.AddItem("Nombre de la Faccion")
Me.cmbIdentificarEquipos.itemData(Me.cmbIdentificarEquipos.NewIndex) = eEventoIdentificacionEquipo.identificaFaccion

'
Me.sliderGanadorCancha.MinValue = 2
Me.sliderGanadorCancha.MaxValue = 20
Me.sliderGanadorCancha.value = 2

Me.sliderDebeEsperar.MinValue = 1
Me.sliderDebeEsperar.MaxValue = 10
Me.sliderDebeEsperar.value = 1

Me.txtMinNivel.MinValue = 1
Me.txtMinNivel.MaxValue = STAT_MAXELV
Me.txtMinNivel.blanqueado = True

Me.txtMaxNivel.MinValue = 1
Me.txtMaxNivel.MaxValue = STAT_MAXELV
Me.txtMaxNivel.blanqueado = True

Me.txtCantidadParticipantesTorneo.MinValue = 1
Me.txtCantidadParticipantesTorneo.MaxValue = 255
Me.txtCantidadParticipantesTorneo.blanqueado = True

' Cantidad de equipos minima y maxima
Me.txtCantidadEquiposTorneoMin.MinValue = 1
Me.txtCantidadEquiposTorneoMin.MaxValue = 255
Me.txtCantidadEquiposTorneoMin.blanqueado = True

Me.txtCantidadEquiposTorneoMax.MinValue = 1
Me.txtCantidadEquiposTorneoMax.MaxValue = 255
Me.txtCantidadEquiposTorneoMax.blanqueado = True

' Combates al mejor de ...
Me.txtAlMejorDe.MinValue = 1
Me.txtAlMejorDe.MaxValue = 255
Me.txtAlMejorDe.value = 3

' Restricciones de Raza
Me.txtMaxRepeRaza.MaxValue = 255
Me.txtMaxRepeRaza.MinValue = 0
Me.txtMaxRepeRaza.blanqueado = True

' Restriccion de Clase
Me.txtCantidadMaximaRepeClase.MaxValue = 255
Me.txtCantidadMaximaRepeClase.MinValue = 0
Me.txtCantidadMaximaRepeClase.blanqueado = True

' Restricciones de Grupo de Clases
Me.txtAlMenosNoMagicas.MaxValue = 255
Me.txtAlMenosNoMagicas.MinValue = 1
Me.txtAlMenosNoMagicas.value = 1

Me.txtAlMenosSemiMagicas.MaxValue = 255
Me.txtAlMenosSemiMagicas.MinValue = 1
Me.txtAlMenosSemiMagicas.value = 1

Me.txtAlMenosMagicas.MaxValue = 255
Me.txtAlMenosMagicas.MinValue = 1
Me.txtAlMenosMagicas.value = 1

Me.txtAlMenosTrabajadoras.MaxValue = 255
Me.txtAlMenosTrabajadoras.MinValue = 1
Me.txtAlMenosTrabajadoras.value = 1

' Sumatoria de niveles
Me.txtMaxSumatoriaNiveles.MaxValue = 32000
Me.txtMaxSumatoriaNiveles.MinValue = 1
Me.txtMaxSumatoriaNiveles.blanqueado = 0

' Miscelanios
Me.sldMinutosToleranciaDeslogueo.MaxValue = 20
Me.sldMinutosToleranciaDeslogueo.MinValue = 0
Me.sldMinutosToleranciaDeslogueo.value = 3

' Tiempos
Me.txtTiempoAviso.MaxValue = 3200
Me.txtTiempoAviso.MinValue = 0
Me.txtTiempoAviso.value = 10

Me.txtTiempoInscripcion.MaxValue = 3200
Me.txtTiempoInscripcion.MinValue = 0
Me.txtTiempoInscripcion.value = 10

Me.txtTiempoApuestas.MaxValue = 60
Me.txtTiempoApuestas.MinValue = 5
Me.txtTiempoApuestas.value = 5

Me.sldMinutosToleranciaDeslogueo.MaxValue = 30
Me.sldMinutosToleranciaDeslogueo.MinValue = 0
Me.sldMinutosToleranciaDeslogueo.value = 3

' Alineaciones
Me.sliderArmadaRango.MaxValue = 5
Me.sliderArmadaRango.MinValue = 0
Me.sliderArmadaRango.value = 0

Me.sliderLegionRango.MaxValue = 5
Me.sliderLegionRango.MinValue = 0
Me.sliderLegionRango.value = 0

' Eventos uatomatico por default
Me.frmSinEvento.visible = False

' Cargo los nombres e ids de los objetos
Call cargarObjetos

' Cargo los nombres e ids de los hechizos
Call cargarHechizos

' Cargo las reglas template sobre los hechizos permitos
Call cargarReglasbasicasHechizos(reglaHechizos)

' Genero el conjunto de textbox que van a permitir seleccionar los hechizos permitidos
Call generarListaHechizos

' Genero la grilla de limite de objetos
Call generarGrillaObjetos

    
' Selecciono la Restriccion inicial
Me.cmbRestricciones.ListIndex = 0
End Sub

Private Sub generarListaHechizos()

Dim loopHechizo As Integer

For loopHechizo = 1 To UBound(hechizos())
    Call Load(Me.chkHechizo(loopHechizo))

    With Me.chkHechizo(loopHechizo)

        .top = 240 + (loopHechizo \ 4) * 260
        .width = 1290
        .left = 120 + (loopHechizo Mod 4) * 1300
        .visible = True
        .Caption = mid(hechizos(loopHechizo), 1, 11)
        .ToolTipText = hechizos(loopHechizo)

        If reglaHechizos(loopHechizo) = 1 Then
            .value = 0
        Else
            .value = 1
        End If
    End With
Next

End Sub
Private Sub cargarReglasbasicasHechizos(ByRef reglas() As Byte)

    ReDim reglas(1 To 41) As Byte  'TODO. Cambiar el 41 por una variable de cantidad max de hechizos
    
    reglas(eHechizos.Ayuda_espiritu_indomable) = 1
    reglas(eHechizos.Debilidad) = 1
    reglas(eHechizos.Estupidez) = 1
    reglas(eHechizos.Implorar_ayuda) = 1
    reglas(eHechizos.Invisibilidad) = 1
    reglas(eHechizos.Invocar_elemetanl_fuego) = 1
    reglas(eHechizos.Invocar_Mascotas) = 1
    reglas(eHechizos.Invocar_Zombies) = 1
    reglas(eHechizos.Invocoar_elemental_agua) = 1
    reglas(eHechizos.Invocoar_elemental_tierra) = 1
    reglas(eHechizos.Llamado_naturaleza) = 1
    reglas(eHechizos.Provocar_Hambre) = 1
    reglas(eHechizos.Resucitar) = 1
    reglas(eHechizos.Terrible_Hambre) = 1
    reglas(eHechizos.Torpeza) = 1
    reglas(eHechizos.Mimetismo) = 1

End Sub

Private Sub cargarObjetos()
ReDim objeto(1 To 216) As String

objeto(1) = "1- Manzana Roja"
objeto(2) = "2- Espada larga"
objeto(3) = "3- Hacha orca"
objeto(4) = "9- Horquilla"
objeto(5) = "10- Cofre abierto"
objeto(6) = "11- Cofre cerrado"
objeto(7) = "15- Daga"
objeto(8) = "19- Espada dos manos"
objeto(9) = "21- Porcion de tarta"
objeto(10) = "22- Frutas del bosque"
objeto(11) = "23- Pan de trigo"
objeto(12) = "24- Pan de maiz"
objeto(13) = "25- Pastel"
objeto(14) = "26- Pollo"
objeto(15) = "27- Chuleta"
objeto(16) = "28- Queso de cabra"
objeto(17) = "29- Sandia"
objeto(18) = "30- Armadura de cuero"
objeto(19) = "31- Vestimentas comunes (H/E/EO-H/M)"
objeto(20) = "32- Vestimentas comunes (H/E/EO-H/M)"
objeto(21) = "35- Vestimentas comunes (H/E/EO-H/M)"
objeto(22) = "36- Poción amarilla"
objeto(23) = "37- Poción azul"
objeto(24) = "38- Poción roja"
objeto(25) = "39- Poción verde"
objeto(26) = "40- Atlas Argentum"
objeto(27) = "41- Libro antiguo"
objeto(28) = "42- Vino"
objeto(29) = "43- Botella de agua"
objeto(30) = "45- Ropa de clan (H/E/EO-H/M)"
objeto(31) = "58- Leña"
objeto(32) = "64- Manzana Roja"
objeto(33) = "80- Ropa de clan (H/E/EO-H/M)"
objeto(34) = "81- Sandia"
objeto(35) = "123- Espada vikinga"
objeto(36) = "124- Katana"
objeto(37) = "125- Sable"
objeto(38) = "126- Hacha larga de guerra"
objeto(39) = "127- Hacha de leñador"
objeto(40) = "129- Hacha de guerra dos filos"
objeto(41) = "130- Escudo de hierro"
objeto(42) = "131- Casco de hierro completo"
objeto(43) = "132- Casco de hierro "
objeto(44) = "133- Escudo Imperial"
objeto(45) = "135- Túnica de la cruz roja (H/E/EO-H/M)"
objeto(46) = "136- Ramitas"
objeto(47) = "138- Caña de pescar"
objeto(48) = "139- Merluza"
objeto(49) = "155- Cama"
objeto(50) = "156- Cama"
objeto(51) = "157- Copa de plata"
objeto(52) = "158- Banana"
objeto(53) = "159- Hacha de bárbaro"
objeto(54) = "160- Cerveza"
objeto(55) = "161- Jugo de frutas"
objeto(56) = "162- Silla"
objeto(57) = "163- Cuchara"
objeto(58) = "164- Espada corta "
objeto(59) = "165- Daga +1"
objeto(60) = "166- Poción violeta"
objeto(61) = "167- Mueble rustico"
objeto(62) = "168- Silla"
objeto(63) = "170- Vestimenta de mujer (H/E/EO-M)"
objeto(64) = "187- Piquete de minero"
objeto(65) = "192- Hierro"
objeto(66) = "193- Plata"
objeto(67) = "194- Oro"
objeto(68) = "195- Armadura de placas completa"
objeto(69) = "196- Túnica de mago (H/E/EO-H/M)"
objeto(70) = "198- Serrucho"
objeto(71) = "236- Ropa de pordiosero"
objeto(72) = "237- Tunica de Roja"
objeto(73) = "238- Túnica Azul (H/E/EO-H/M)"
objeto(74) = "239- Túnica Roja (H/E/EO-H/M)"
objeto(75) = "240- Ropa común (E/G-H)"
objeto(76) = "243- Armadura de placas completa (E/G-H/M)"
objeto(77) = "356- Armadura De las sombras"
objeto(78) = "357- Túnica de druida (H/E/EO-H/M)"
objeto(79) = "359- Cota de mallas"
objeto(80) = "360- Armadura de cazador"
objeto(81) = "362- Tambor"
objeto(82) = "365- Daga + 2"
objeto(83) = "366- Daga + 3"
objeto(84) = "367- Daga + 4"
objeto(85) = "371- Vestimentas de noble"
objeto(86) = "381- Túnica de monje (H/E/EO-H/M)"
objeto(87) = "382- Vestido azul (H/E/EO-M)"
objeto(88) = "386- Lingote de hierro"
objeto(89) = "387- Lingote de plata"
objeto(90) = "388- Lingote de oro"
objeto(91) = "389- Martillo de Herrero"
objeto(92) = "390- Armadura de placas completa + 1 "
objeto(93) = "391- Armadura de placas completa + 2"
objeto(94) = "392- Armadura de gala (E/G-H/M)"
objeto(95) = "393- Armadura de placas completa + 2 (E/G-H/M)"
objeto(96) = "398- Garrote"
objeto(97) = "399- Cimitarra"
objeto(98) = "400- Vara de mago"
objeto(99) = "401- Martillo de guerra"
objeto(100) = "402- Espada mata dragones"
objeto(101) = "403- Espada de plata"
objeto(102) = "404- Escudo de tortuga"
objeto(103) = "405- Casco de plata"
objeto(104) = "414- Piel de lobo"
objeto(105) = "415- Piel de oso pardo"
objeto(106) = "416- Piel de oso polar"
objeto(107) = "460- Daga (Newbies)"
objeto(108) = "461- Poción roja (Newbies)"
objeto(109) = "462- Poción azul (Newbies)"
objeto(110) = "469- Laúd"
objeto(111) = "474- Barca"
objeto(112) = "475- Galera"
objeto(113) = "476- Galeon"
objeto(114) = "478- Arco simple"
objeto(115) = "479- Arco compuesto"
objeto(116) = "480- Flecha"
objeto(117) = "483- Armadura de Dragón"
objeto(118) = "484- Armadura de herrero"
objeto(119) = "485- Armadura Legendaria"
objeto(120) = "486- Ropa de minero (E/G-H/M)"
objeto(121) = "487- Dama de las Tinieblas"
objeto(122) = "488- White Lady"
objeto(123) = "489- Armadura de placas azules"
objeto(124) = "490- Ropa de minero (H/E/EO-H/M)"
objeto(125) = "493- Armadura de Placas Roja (Mujer)"
objeto(126) = "495- Armadura Escarlata"
objeto(127) = "496- Armadura de la Ciénaga"
objeto(128) = "497- Armadura de placas de gala"
objeto(129) = "499- Cota de Mallas (H/E/EO-M)"
objeto(130) = "500- Armadura Bruñida (E/G-H/M)"
objeto(131) = "501- Ropa de Campesino (H/E/EO-H/M)"
objeto(132) = "502- Ropa de Campesino (E/G-H/M)"
objeto(133) = "503- Trampa Visual (H/E/EO-H/M)"
objeto(134) = "504- Ropa Común (H/E/EO-M)"
objeto(135) = "505- Ropa Común (H/E/EO-M)"
objeto(136) = "506- Ropa Común (H/E/EO-M)"
objeto(137) = "507- Ropa Común Ovispo (E/G-H)"
objeto(138) = "508- Ropa Estuaria (H/E/EO-M)"
objeto(139) = "510- Vestido Indulgente (H/E/EO-M)"
objeto(140) = "511- Vestido De Novia Sensual (H/E-M)"
objeto(141) = "512- Vestido Calipso (E/G-M)"
objeto(142) = "513- Vestido de Bruja (H/E-M)"
objeto(143) = "514- Vestido de Bruja (EO-M)"
objeto(144) = "515- Ropa de Carpintero (H/E/EO-H/M)"
objeto(145) = "519- Túnica Legendaria (H/E/EO-M/H)"
objeto(146) = "524- Túnica Roja(G-H/M)"
objeto(147) = "525- Túnica Roja combinada(G-H)"
objeto(148) = "526- Túnica Roja combinada(G-M)"
objeto(149) = "527- Ropa común (E/G-M)"
objeto(150) = "529- Armadura de cuero (H/E/EO-M)"
objeto(151) = "533- Botella agua"
objeto(152) = "534- Botella vacia"
objeto(153) = "540- Flauta"
objeto(154) = "543- Red de pesca"
objeto(155) = "544- Pejerrey"
objeto(156) = "545- Pez espada"
objeto(157) = "546- Salmón"
objeto(158) = "550- Vestido De Novia Sensual (EO-M)"
objeto(159) = "551- Flecha envenenada"
objeto(160) = "552- Flecha +2"
objeto(161) = "553- Flecha incendiaria"
objeto(162) = "559- Daga de Plata"
objeto(163) = "612- Armadura de cazador (E/G-H/M)"
objeto(164) = "613- Armadura de cuero (E/G-H/M)"
objeto(165) = "614- Túnica de druida (E/G-H/M)"
objeto(166) = "615- Túnica Legendaria (E/G-H/M)"
objeto(167) = "616- Cota de mallas (E/G-H/M)"
objeto(168) = "617- Túnica de monje (E/G-H/M)"
objeto(169) = "619- Vestimentas de noble (E/G-H/M)"
objeto(170) = "620- Túnica gris (E/G-H/M)"
objeto(171) = "621- Sombrero de Aprendiz"
objeto(172) = "622- Sombrero de Mago"
objeto(173) = "623- Báculo Engarzada"
objeto(174) = "624- Bastón nudoso"
objeto(175) = "625- Vara de fresno"
objeto(176) = "626- Arco simple reforzado"
objeto(177) = "627- Arco compuesto reforzado"
objeto(178) = "628- Arco de Cazador"
objeto(179) = "629- Poción de energía"
objeto(180) = "630- Hacha dorada"
objeto(181) = "632- Túnica de dioses"
objeto(182) = "633- Piel de lobo invernal"
objeto(183) = "642- Leña de Tejo"
objeto(184) = "643- Laúd mágico"
objeto(185) = "644- Anillo de resistencia mágica"
objeto(186) = "645- Anillo de protección mágica"
objeto(187) = "646- Anillo de Disolución Mágica"
objeto(188) = "647- Poción Negra"
objeto(189) = "648- Anillo mágico"
objeto(190) = "649- Anillo de los Dioses"
objeto(191) = "650- Pocion amarilla (Newbies)"
objeto(192) = "651- Pocion verde (Newbies)"
objeto(193) = "657- Espada Especial"
objeto(194) = "659- Laúd Especial"
objeto(195) = "661- Armadura del Gran Cazador"
objeto(196) = "662- Armadura del Gran Cazador (E/G-H/M)"
objeto(197) = "664- Vestimenta de mujer (Newbies)"
objeto(198) = "665- Equipo invernal (H-E-EO/M)"
objeto(199) = "666- Equipo invernal (H-E-EO/H)"
objeto(200) = "667- Equipo invernal (E/G-H/M)"
objeto(201) = "668- Túnica de dioses (E/G-H/M)"
objeto(202) = "670- Vestido Negro (E/G-M)"
objeto(203) = "671- Armadura de cazador (H/E/EO-M)"
objeto(204) = "672- Ropa de Carpintero (E/G-H/M)"
objeto(205) = "682- Poción violeta (Newbies)"
objeto(206) = "684- Fragmento de cristal"
objeto(207) = "685- Piquete de oro"
objeto(208) = "688- Armadura Klox"
objeto(209) = "689- Armadura de placas de gala (E/G-H/M)"
objeto(210) = "730- Mapa"
objeto(211) = "731- Mapa (Newbies)"
objeto(212) = "732- Caballito de mar"
objeto(213) = "733- Armadura De las sombras"
objeto(214) = "828- Flecha de tejo"
objeto(215) = "830- Espada de plata"
objeto(216) = "831- Espada de plata"

End Sub

Private Sub cargarHechizos()

ReDim hechizos(1 To CANTIDAD_HECHIZOS) As String

hechizos(1) = "Curar veneno"
hechizos(2) = "Dardo magico"
hechizos(3) = "Curar heridas leves"
hechizos(4) = "Envenenar"
hechizos(5) = "Curar heridas graves"
hechizos(6) = "Flecha magica"
hechizos(7) = "Flecha electrica"
hechizos(8) = "Misil magico"
hechizos(9) = "Paralizar"
hechizos(10) = "Remover paralisis"
hechizos(11) = "Resucitar"
hechizos(12) = "Provocar hambre"
hechizos(13) = "Terrible hambre de Igôr"
hechizos(14) = "Invisibilidad"
hechizos(15) = "Tormenta de fuego"
hechizos(16) = "Llamado a la naturaleza"
hechizos(17) = "Invokar Zombies"
hechizos(18) = "Celeridad"
hechizos(19) = "Torpeza"
hechizos(20) = "Fuerza"
hechizos(21) = "Debilidad"
hechizos(22) = "Llamado a Uhkrul"
hechizos(23) = "Descarga Eléctrica"
hechizos(24) = "Inmovilizar"
hechizos(25) = "Apocalípsis"
hechizos(26) = "Invocar elemental de fuego"
hechizos(27) = "Invocar elemental de agua"
hechizos(28) = "Invocar elemental de tierra"
hechizos(29) = "Implorar ayuda"
hechizos(30) = "Ceguera"
hechizos(31) = "Estupidez"
hechizos(32) = "Ira de dioS"
hechizos(33) = "Ayuda del Espiritu Indomable"
hechizos(34) = "Remover Estupidez"
hechizos(35) = "Mimetismo"
hechizos(36) = "Detectar invisibilidad"
hechizos(37) = "Curación celestial"
hechizos(38) = "Habilidad Ilimitada"
hechizos(39) = "Invocar mascotas"
hechizos(40) = "Descarga Divina"
hechizos(41) = "Tormenta Divina"

End Sub

Private Sub lstEstadoEventos_Click()
    Me.lstEstadoEventos.ToolTipText = Me.lstEstadoEventos.list(Me.lstEstadoEventos.ListIndex)
    Me.cmdCancelarEvento.Enabled = True
    Me.cmdVerInfo.Enabled = True
End Sub

Private Sub formatearTextNumero(text As Textbox)
    Dim texto As String
    
    If Len(text.text) = 0 Then Exit Sub
    
    texto = Replace$(text, ".", "")
    
    text = FormatNumber(val(texto), 0, vbTrue, vbFalse, vbTrue)
    text.SelStart = Len(text.text)
End Sub

Private Sub OptEvento_Click(Index As Integer)
    If Index = 0 Then
        Me.frmSinEvento.visible = False
    ElseIf Index = 1 Then
        Me.frmSinEvento.visible = True
        Me.txtCantidadEquiposTorneoMin.Enabled = True
    End If
End Sub

Private Sub OptTipoSubEvento_Click(Index As Integer)
    If Index = 1 Then
        Me.txtCantidadEquiposTorneoMin.Enabled = False
        Me.txtCantidadEquiposTorneoMin.value = Me.txtCantidadEquiposTorneoMax.value
    Else
        Me.txtCantidadEquiposTorneoMin.Enabled = True
    End If
End Sub

Private Sub txtApuestasPozoInicial_Change()
    Call formatearTextNumero(Me.txtApuestasPozoInicial)
End Sub

Private Sub generarGrillaObjetos()
Dim loopItem As Integer

'Cargamos los sonidos
Call Me.GridTextListaObjetos.setNombreCampos("")
    
' Le agrego un control más que es un CheckBox.
Call Me.GridTextListaObjetos.agregarControlDinamico(Me.txtCantidadObjetos, "txtCantidad", "Cantidad")
Call Me.GridTextListaObjetos.agregarControlDinamico(Me.chkTemplate, "chkMinimo", "Mínimo")

Call Me.GridTextListaObjetos.setDescripcion(0, "Objeto")

Call Me.GridTextListaObjetos.iniciar

' Cargamos todos los objetos
For loopItem = 1 To UBound(objeto)
    Call Me.GridTextListaObjetos.addString(val(objeto(loopItem)), objeto(loopItem))
Next

Call Me.GridTextListaObjetos.seleccionarID(0, 0)

End Sub

Public Sub parsearInfoEventos(datos As String)
Dim infoEventos() As String
Dim loopEvento As Byte
Dim cantidad As Byte
Dim nombreEvento As String

Me.lstEstadoEventos.Clear

infoEventos = Split(datos, "||")

If Len(datos) > 0 Then

    For loopEvento = 0 To UBound(infoEventos) - 1
        nombreEvento = mid$(infoEventos(loopEvento), 1, InStr(1, infoEventos(loopEvento), "-"))
    
        Call Me.lstEstadoEventos.AddItem(infoEventos(loopEvento))
    Next loopEvento
    
    cantidad = UBound(infoEventos)
Else
    cantidad = 0
End If

Me.lblCantidadEventos = "Cantidad: " & cantidad
Me.cmdVerEventos.Enabled = True
End Sub

Public Sub procesarMensaje(mensaje As String)

    Me.lblMensaje.Caption = mensaje
    Me.frmMensaje.visible = True

End Sub


'Private Sub txtCantidadEquiposTorneoMax_Change(valor As Double)

    ' PlayOff
'    If Me.OptTipoSubEvento(1).value = True Then
'        Me.txtCantidadEquiposTorneoMin.value = valor
'    End If
'End Sub

Private Sub txtOroPremio_Change(Index As Integer)
    Dim valor As Byte
    
    Call formatearTextNumero(Me.txtOroPremio(Index))
    
    If val(Me.txtOroPremio(Index)) > 0 And Index <> 8 Then
        
        valor = Index + 1
        If valor > Me.txtOroPremio.UBound Then
            
            Call Load(Me.txtOroPremio(valor))
            Call Load(Me.lblPuestoX(valor))
            
            With Me.txtOroPremio(valor)
                .Height = 285
                .top = 480 + valor * 300
                .visible = True
                .text = "0"
            End With
            
            With Me.lblPuestoX(valor)
                .Caption = valor + 1 & "º"
                .top = 520 + valor * 300
                .visible = True
            End With
        End If

    End If
End Sub

Private Sub txtPrecioInscripcionTorneo_Change()
    Call formatearTextNumero(Me.txtPrecioInscripcionTorneo)
End Sub
