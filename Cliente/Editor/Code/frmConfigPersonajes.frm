VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfigurarPersonajes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar partes de los personajes"
   ClientHeight    =   5085
   ClientLeft      =   2730
   ClientTop       =   6645
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfigPersonajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   8535
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdActualizarNombres 
      Caption         =   "Actualizar nombres"
      Height          =   360
      Left            =   2760
      TabIndex        =   104
      Top             =   5520
      Width           =   3015
   End
   Begin VB.OptionButton cmdOptionManoHabil 
      Appearance      =   0  'Flat
      Caption         =   "Zurdo"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   1200
      TabIndex        =   101
      Top             =   4850
      Width           =   855
   End
   Begin VB.OptionButton cmdOptionManoHabil 
      Appearance      =   0  'Flat
      Caption         =   "Diestro"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   100
      Top             =   4850
      Width           =   975
   End
   Begin EditorTDS.ListaConBuscador lstlGrhDisponibles 
      Height          =   2055
      Left            =   8640
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3625
   End
   Begin TabDlg.SSTab tabOpciones 
      Height          =   4815
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      WordWrap        =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Cuerpos"
      TabPicture(0)   =   "frmConfigPersonajes.frx":1CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tree_Cuerpo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmPropiedades_Cuerpos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdEliminar_Cuerpos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdNuevo_Cuerpos"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdCancelar(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdAceptar(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Cabezas"
      TabPicture(1)   =   "frmConfigPersonajes.frx":1CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdCancelar(1)"
      Tab(1).Control(1)=   "cmdAceptar(1)"
      Tab(1).Control(2)=   "cmdEliminar_Cabezas"
      Tab(1).Control(3)=   "cmdNuevo_Cabezas"
      Tab(1).Control(4)=   "frmPropiedades_Cabezas"
      Tab(1).Control(5)=   "tree_Cabezas"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Cascos/Sombreros"
      TabPicture(2)   =   "frmConfigPersonajes.frx":1D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdCancelar(2)"
      Tab(2).Control(1)=   "cmdAceptar(2)"
      Tab(2).Control(2)=   "cmdEliminar_Cascos"
      Tab(2).Control(3)=   "cmdNuevo_Cascos"
      Tab(2).Control(4)=   "frmPropiedades_Cascos"
      Tab(2).Control(5)=   "tree_Cascos"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Escudos"
      TabPicture(3)   =   "frmConfigPersonajes.frx":1D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdCancelar(3)"
      Tab(3).Control(1)=   "cmdAceptar(3)"
      Tab(3).Control(2)=   "cmdEliminar_Escudos"
      Tab(3).Control(3)=   "cmdNuevo_Escudos"
      Tab(3).Control(4)=   "frmPropiedades_Escudos"
      Tab(3).Control(5)=   "tree_Escudos"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Armas"
      TabPicture(4)   =   "frmConfigPersonajes.frx":1D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdCancelar(4)"
      Tab(4).Control(1)=   "cmdAceptar(4)"
      Tab(4).Control(2)=   "cmdEliminar_Armas"
      Tab(4).Control(3)=   "cmdNuevo_Armas"
      Tab(4).Control(4)=   "frmPropiedades_Armas"
      Tab(4).Control(5)=   "tree_Armas"
      Tab(4).ControlCount=   6
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   360
         Index           =   4
         Left            =   -71280
         TabIndex        =   99
         Top             =   4400
         Width           =   2175
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   360
         Index           =   4
         Left            =   -69000
         TabIndex        =   98
         Top             =   4400
         Width           =   2415
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   360
         Index           =   3
         Left            =   -71280
         TabIndex        =   97
         Top             =   4400
         Width           =   2175
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   360
         Index           =   3
         Left            =   -69000
         TabIndex        =   96
         Top             =   4400
         Width           =   2415
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   360
         Index           =   2
         Left            =   -71280
         TabIndex        =   95
         Top             =   4400
         Width           =   2175
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   360
         Index           =   2
         Left            =   -69000
         TabIndex        =   94
         Top             =   4400
         Width           =   2415
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   360
         Index           =   1
         Left            =   -71280
         TabIndex        =   93
         Top             =   4400
         Width           =   2175
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   360
         Index           =   1
         Left            =   -69000
         TabIndex        =   92
         Top             =   4400
         Width           =   2415
      End
      Begin VB.CommandButton cmdEliminar_Armas 
         Caption         =   "Eliminar"
         Height          =   360
         Left            =   -73080
         TabIndex        =   81
         Top             =   4400
         Width           =   1575
      End
      Begin VB.CommandButton cmdNuevo_Armas 
         Caption         =   "Nuevo"
         Height          =   360
         Left            =   -74880
         TabIndex        =   80
         Top             =   4400
         Width           =   1575
      End
      Begin VB.CommandButton cmdEliminar_Escudos 
         Caption         =   "Eliminar"
         Height          =   360
         Left            =   -73080
         TabIndex        =   79
         Top             =   4400
         Width           =   1575
      End
      Begin VB.CommandButton cmdNuevo_Escudos 
         Caption         =   "Nuevo"
         Height          =   360
         Left            =   -74880
         TabIndex        =   78
         Top             =   4400
         Width           =   1575
      End
      Begin VB.CommandButton cmdEliminar_Cascos 
         Caption         =   "Eliminar"
         Height          =   360
         Left            =   -73080
         TabIndex        =   77
         Top             =   4400
         Width           =   1575
      End
      Begin VB.CommandButton cmdNuevo_Cascos 
         Caption         =   "Nuevo"
         Height          =   360
         Left            =   -74880
         TabIndex        =   76
         Top             =   4400
         Width           =   1575
      End
      Begin VB.CommandButton cmdEliminar_Cabezas 
         Caption         =   "Eliminar"
         Height          =   360
         Left            =   -73080
         TabIndex        =   75
         Top             =   4400
         Width           =   1575
      End
      Begin VB.CommandButton cmdNuevo_Cabezas 
         Caption         =   "Nuevo"
         Height          =   360
         Left            =   -74880
         TabIndex        =   74
         Top             =   4400
         Width           =   1575
      End
      Begin VB.Frame frmPropiedades_Escudos 
         Caption         =   "Propiedades"
         Height          =   3855
         Left            =   -71280
         TabIndex        =   55
         Top             =   480
         Width           =   4695
         Begin VB.CommandButton cmdRestablecer_Escudos 
            Caption         =   "Restablecer"
            Height          =   360
            Left            =   2280
            TabIndex        =   89
            Top             =   3360
            Width           =   2295
         End
         Begin VB.CommandButton cmdAplicar_Escudos 
            Caption         =   "Aplicar"
            Height          =   360
            Left            =   120
            TabIndex        =   88
            ToolTipText     =   "Guarda los cambios realizados"
            Top             =   3360
            Width           =   2055
         End
         Begin VB.TextBox txtEscudo 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   66
            Top             =   1320
            Width           =   3495
         End
         Begin VB.TextBox txtEscudo 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   65
            Top             =   1680
            Width           =   3495
         End
         Begin VB.TextBox txtEscudo 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   64
            Top             =   2040
            Width           =   3495
         End
         Begin VB.TextBox txtEscudo 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   63
            Top             =   960
            Width           =   3495
         End
         Begin VB.TextBox txtEscudoNombre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            TabIndex        =   62
            Top             =   570
            Width           =   3495
         End
         Begin VB.Label lblEscudoNumeroResultado 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            Height          =   195
            Left            =   1080
            TabIndex        =   67
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblEscudoNumero 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Numero:"
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblEscudoNombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblEscudoNorte 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Norte"
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   960
            Width           =   405
         End
         Begin VB.Label lblEscudoEste 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Este"
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   1320
            Width           =   315
         End
         Begin VB.Label lblEscudoSur 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sur"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label lblEscudoOeste 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Oeste"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   2040
            Width           =   435
         End
      End
      Begin VB.Frame frmPropiedades_Armas 
         Caption         =   "Propiedades"
         Height          =   3855
         Left            =   -71280
         TabIndex        =   36
         Top             =   480
         Width           =   4695
         Begin VB.CommandButton cmdAplicar_Armas 
            Caption         =   "Aplicar"
            Height          =   360
            Left            =   120
            TabIndex        =   91
            ToolTipText     =   "Guarda los cambios realizados"
            Top             =   3360
            Width           =   2055
         End
         Begin VB.CommandButton cmdRestablecer_Armas 
            Caption         =   "Restablecer"
            Height          =   360
            Left            =   2280
            TabIndex        =   90
            Top             =   3360
            Width           =   2295
         End
         Begin VB.TextBox txtArma 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   72
            Top             =   1320
            Width           =   3495
         End
         Begin VB.TextBox txtArma 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   71
            Top             =   1680
            Width           =   3495
         End
         Begin VB.TextBox txtArma 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   70
            Top             =   2040
            Width           =   3495
         End
         Begin VB.TextBox txtArma 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   69
            Top             =   960
            Width           =   3495
         End
         Begin VB.TextBox txtArmaNombre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            TabIndex        =   68
            Top             =   570
            Width           =   3495
         End
         Begin VB.Label lblArmaNumeroResultado 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            Height          =   195
            Left            =   1080
            TabIndex        =   73
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblArmaNumero 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Numero:"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblArmaNombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblArmasNorte 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Norte"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   960
            Width           =   405
         End
         Begin VB.Label lblArmasEste 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Este"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   1320
            Width           =   315
         End
         Begin VB.Label lblArmasSur 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sur"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label lblArmasOeste 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Oeste"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   2040
            Width           =   435
         End
      End
      Begin VB.Frame frmPropiedades_Cascos 
         Caption         =   "Propiedades"
         Height          =   3855
         Left            =   -71280
         TabIndex        =   29
         Top             =   480
         Width           =   4695
         Begin VB.CommandButton cmdAplicar_Cascos 
            Caption         =   "Aplicar"
            Height          =   360
            Left            =   120
            TabIndex        =   87
            ToolTipText     =   "Guarda los cambios realizados"
            Top             =   3360
            Width           =   2055
         End
         Begin VB.CommandButton cmdRestablecer_Cascos 
            Caption         =   "Restablecer"
            Height          =   360
            Left            =   2280
            TabIndex        =   86
            Top             =   3360
            Width           =   2295
         End
         Begin VB.TextBox txtCascoNombre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            TabIndex        =   53
            Top             =   570
            Width           =   3495
         End
         Begin VB.TextBox txtCasco 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   52
            Top             =   960
            Width           =   3495
         End
         Begin VB.TextBox txtCasco 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   51
            Top             =   2040
            Width           =   3495
         End
         Begin VB.TextBox txtCasco 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   50
            Top             =   1680
            Width           =   3495
         End
         Begin VB.TextBox txtCasco 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   49
            Top             =   1320
            Width           =   3495
         End
         Begin VB.Label lblCascoNumeroResultado 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            Height          =   195
            Left            =   1080
            TabIndex        =   54
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblCascoNumero 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Numero:"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblCascoNombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblCascoNorte 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Norte"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   960
            Width           =   405
         End
         Begin VB.Label lblCascoEste 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Este"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   1320
            Width           =   315
         End
         Begin VB.Label lblCascoSur 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sur"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label lblCascoOeste 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Oeste"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   2040
            Width           =   435
         End
      End
      Begin VB.Frame frmPropiedades_Cabezas 
         Caption         =   "Propiedades"
         Height          =   3855
         Left            =   -71280
         TabIndex        =   22
         Top             =   480
         Width           =   4695
         Begin VB.CommandButton cmdAplicar_Cabezas 
            Caption         =   "Aplicar"
            Height          =   360
            Left            =   120
            TabIndex        =   85
            ToolTipText     =   "Guardar los cambios realizados"
            Top             =   3360
            Width           =   2055
         End
         Begin VB.CommandButton cmdRestablecer_Cabezas 
            Caption         =   "Restablecer"
            Height          =   360
            Left            =   2280
            TabIndex        =   84
            Top             =   3360
            Width           =   2295
         End
         Begin VB.TextBox txtCabeza 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   47
            Top             =   1320
            Width           =   3495
         End
         Begin VB.TextBox txtCabeza 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   46
            Top             =   1680
            Width           =   3495
         End
         Begin VB.TextBox txtCabeza 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   45
            Top             =   2040
            Width           =   3495
         End
         Begin VB.TextBox txtCabeza 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   44
            Top             =   960
            Width           =   3495
         End
         Begin VB.TextBox txtNombreCabeza 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            TabIndex        =   43
            Top             =   570
            Width           =   3495
         End
         Begin VB.Label lblCabezaNumeroResultado 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            Height          =   195
            Left            =   1080
            TabIndex        =   48
            Top             =   240
            Width           =   300
         End
         Begin VB.Label lblCabezaOeste 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Oeste"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   2040
            Width           =   435
         End
         Begin VB.Label lblCabezaSur 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sur"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label lblCabezaEste 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Este"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   1320
            Width           =   315
         End
         Begin VB.Label lblCabezaNorte 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Norte"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   405
         End
         Begin VB.Label lblCabezaNombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblCabezaNumero 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Numero:"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   360
         Index           =   0
         Left            =   6000
         TabIndex        =   15
         Top             =   4400
         Width           =   2415
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   360
         Index           =   0
         Left            =   3720
         TabIndex        =   14
         Top             =   4400
         Width           =   2175
      End
      Begin VB.CommandButton cmdNuevo_Cuerpos 
         Caption         =   "Nuevo"
         Height          =   360
         Left            =   120
         TabIndex        =   13
         Top             =   4400
         Width           =   1575
      End
      Begin VB.CommandButton cmdEliminar_Cuerpos 
         Caption         =   "Eliminar"
         Height          =   360
         Left            =   1920
         TabIndex        =   12
         Top             =   4400
         Width           =   1575
      End
      Begin VB.Frame frmPropiedades_Cuerpos 
         Caption         =   "Propiedades"
         Height          =   3855
         Left            =   3720
         TabIndex        =   2
         Top             =   480
         Width           =   4695
         Begin VB.TextBox txtCuerpo 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   108
            Top             =   2040
            Width           =   3495
         End
         Begin VB.TextBox txtCuerpo 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   107
            Top             =   1680
            Width           =   3495
         End
         Begin VB.TextBox txtCuerpo 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   106
            Top             =   1320
            Width           =   3495
         End
         Begin VB.TextBox txtCuerpo 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   105
            Top             =   960
            Width           =   3495
         End
         Begin EditorTDS.UpDownText txtCuerpoOffsetX 
            Height          =   315
            Left            =   1560
            TabIndex        =   103
            Top             =   2400
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            MaxValue        =   0
            MinValue        =   0
         End
         Begin EditorTDS.UpDownText txtCuerpoOffsetY 
            Height          =   310
            Left            =   1560
            TabIndex        =   102
            Top             =   2760
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            MaxValue        =   0
            MinValue        =   0
         End
         Begin VB.CommandButton cmdRestablecer_Cuerpos 
            Caption         =   "Restablecer"
            Height          =   360
            Left            =   2280
            TabIndex        =   83
            Top             =   3360
            Width           =   2295
         End
         Begin VB.CommandButton cmdAplicar_Cuerpos 
            Caption         =   "Aplicar"
            Height          =   360
            Left            =   120
            TabIndex        =   82
            ToolTipText     =   "Guarda los cambios realizados"
            Top             =   3360
            Width           =   2055
         End
         Begin VB.TextBox txtCuerpoNombre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            TabIndex        =   3
            Top             =   570
            Width           =   3495
         End
         Begin VB.Label lblNumeroCuerpoResultado 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            Height          =   195
            Left            =   1080
            TabIndex        =   17
            Top             =   240
            Width           =   300
         End
         Begin VB.Label lblNumeroCuerpo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Numero:"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblNombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   555
         End
         Begin VB.Label lblNorteCuerpo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Norte"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   960
            Width           =   405
         End
         Begin VB.Label lblEsteCuerpo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Este"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   1320
            Width           =   315
         End
         Begin VB.Label lblArriba 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sur"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label lblOesteCuerpo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Oeste"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   2040
            Width           =   435
         End
         Begin VB.Label lblOffsetCuerpoX 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Posición Cabeza X"
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   2500
            Width           =   1290
         End
         Begin VB.Label lblCuerpoOffsetY 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Posición Cabeza Y"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   2820
            Width           =   1290
         End
      End
      Begin EditorTDS.TreeConBuscador tree_Cuerpo 
         Height          =   3975
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   7011
      End
      Begin EditorTDS.TreeConBuscador tree_Cabezas 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   18
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   7011
      End
      Begin EditorTDS.TreeConBuscador tree_Cascos 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   19
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   7011
      End
      Begin EditorTDS.TreeConBuscador tree_Escudos 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   7011
      End
      Begin EditorTDS.TreeConBuscador tree_Armas 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   21
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   7011
      End
   End
End
Attribute VB_Name = "frmConfigurarPersonajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Cuando se inicia el formulario se guarda el estado actual del personaje (ver "Cancelar")
' y se establece como elementos seleccionados los atributos del charactual.

' Cuando se hace clic en una de las listas se establece como seleccionado el elemento al cual se le hizo
' clic y se refresca el personaje.
' Cada vez que se cambia alguno de las propiedades de un elemento, se actualiza el elemento seleccionado
' y se actualiza el personaje.
' Cuando se hace "Aplicar" se guardan las modificaciones en la estructura de elementos y se persiste
' Cuando se hace "Restablecer" se resetea el formulario al estado anterior re cargando los datos desde la Estructura
' del elemento correspondiente
' Si se pulsa "Aceptar" se mantienen al cerrar la ventana los cambios en el modo caminata.
' Si se pusa "Cancelar" el personaje vuelve al estado que tenia antes de iniciar el modo ventana.

'Textbox que esta actualemnte seleccionado
Private campoSeleccionado As TextBox

'Elementos que estan seleccionados
Private cuerpoSeleccionado As BodyData
Private cabezaSeleccionado As HeadData
Private cascoSeleccionado As HeadData
Private armaSeleccionado As WeaponAnimData
Private escudoSeleccionado As ShieldAnimData

'Estado del charindex antes de abrir la ventana
Private cuerpoBackupSeleccionado As BodyData
Private cabezaBackupSeleccionado As HeadData
Private cascoBackupSeleccionado As HeadData
Private armaBackupSeleccionado As WeaponAnimData
Private escudoBackupSeleccionado As ShieldAnimData

'Refresca el personaje con los elementos seleccionados
Private Sub actualizarPersonaje()
    CharList(UserCharIndex).body = cuerpoSeleccionado
    CharList(UserCharIndex).Head = cabezaSeleccionado
    CharList(UserCharIndex).casco = cascoSeleccionado
    CharList(UserCharIndex).arma = armaSeleccionado
    CharList(UserCharIndex).escudo = escudoSeleccionado
End Sub

'Guarda el estado del char del usuario
Private Sub guardarEstadoActualChar()

    cuerpoBackupSeleccionado = CharList(UserCharIndex).body
    cabezaBackupSeleccionado = CharList(UserCharIndex).Head
    cascoBackupSeleccionado = CharList(UserCharIndex).casco
    armaBackupSeleccionado = CharList(UserCharIndex).arma
    escudoBackupSeleccionado = CharList(UserCharIndex).escudo

End Sub

'Restablece el char actual con los elementos backapeados
Private Sub restablecerEstadoActualChar()

    CharList(UserCharIndex).body = cuerpoBackupSeleccionado
    CharList(UserCharIndex).Head = cabezaBackupSeleccionado
    CharList(UserCharIndex).casco = cascoBackupSeleccionado
    CharList(UserCharIndex).arma = armaBackupSeleccionado
    CharList(UserCharIndex).escudo = escudoBackupSeleccionado

End Sub

'Estabelce como elementos selecionados, los que ya tiene el personaje
Private Sub establecerSeleccionDefault()
    cuerpoSeleccionado = CharList(UserCharIndex).body
    cabezaSeleccionado = CharList(UserCharIndex).Head
    cascoSeleccionado = CharList(UserCharIndex).casco
    armaSeleccionado = CharList(UserCharIndex).arma
    escudoSeleccionado = CharList(UserCharIndex).escudo
End Sub

'Pone en el elemento seleccionado las propiedades que estan en el Editor
Private Sub actualizarCuerpoActual()
    Dim direccion As Byte
        
    cuerpoSeleccionado.nombre = Me.txtCuerpoNombre

    cuerpoSeleccionado.HeadOffset.X = CInt(Me.txtCuerpoOffsetX.value)
    cuerpoSeleccionado.HeadOffset.Y = CInt(Me.txtCuerpoOffsetY.value)
    
    For direccion = E_Heading.NORTH To E_Heading.WEST
        InitGrh cuerpoSeleccionado.Walk(direccion), CInt(val(Me.txtCuerpo(direccion).text))
    Next
    
    'Si hay modificaciones, activamos el boton para guardarlas
    Me.cmdAplicar_Cuerpos.Enabled = True
    Me.cmdRestablecer_Cuerpos.Enabled = True
End Sub

Private Sub actualizarCabezaActual()
    Dim direccion As Byte
    
    cabezaSeleccionado.nombre = Me.txtNombreCabeza
    
    For direccion = E_Heading.NORTH To E_Heading.WEST
        InitGrh cabezaSeleccionado.Head(direccion), CInt(val(Me.txtCabeza(direccion).text))
    Next
    
    'Si hay modificaciones, activamos el boton para guardarlas
    Me.cmdAplicar_Cabezas.Enabled = True
    Me.cmdRestablecer_Cabezas.Enabled = True
End Sub

Private Sub actualizarCascoActual()
    Dim direccion As Byte
    
    cascoSeleccionado.nombre = Me.txtCascoNombre
        
    For direccion = E_Heading.NORTH To E_Heading.WEST
        InitGrh cascoSeleccionado.Head(direccion), CInt(val(Me.txtCasco(direccion).text))
    Next
    
    'Si hay modificaciones, activamos el boton para guardarlas
    Me.cmdAplicar_Cascos.Enabled = True
    Me.cmdRestablecer_Cascos.Enabled = True
End Sub

Private Sub actualizarEscudoActual()
    Dim direccion As Byte
    
    escudoSeleccionado.nombre = Me.txtEscudoNombre
    
    For direccion = E_Heading.NORTH To E_Heading.WEST
        InitGrh escudoSeleccionado.ShieldWalk(direccion), CInt(val(Me.txtEscudo(direccion).text))
    Next
    
    'Si hay modificaciones, activamos el boton para guardarlas
    Me.cmdAplicar_Escudos.Enabled = True
    Me.cmdRestablecer_Escudos.Enabled = True
End Sub

Private Sub actualizarArmaActual()
    Dim direccion As Byte
    
    armaSeleccionado.nombre = Me.txtArmaNombre
        
    For direccion = E_Heading.NORTH To E_Heading.WEST
        InitGrh armaSeleccionado.WeaponWalk(direccion), CInt(val(Me.txtArma(direccion).text))
    Next
    
    'Si hay modificaciones, activamos el boton para guardarlas
    Me.cmdAplicar_Armas.Enabled = True
    Me.cmdRestablecer_Armas.Enabled = True
End Sub

Private Sub cmdActualizarNombres_Click()
    Dim elemento As Integer
    Dim i As Integer
    Dim nombre As String
    
    i = 1
    nombre = "Fantasma oscuro"
    For elemento = 501 To 501
            ' Ponemos el nombre
            HeadData(elemento).nombre = nombre & i
            
            GrhData(HeadData(elemento).Head(E_Heading.NORTH).GrhIndex).nombreGrafico = "Cabeza " & nombre & i & " Norte"
            GrhData(HeadData(elemento).Head(E_Heading.SOUTH).GrhIndex).nombreGrafico = "Cabeza " & nombre & i & " Sur"
            GrhData(HeadData(elemento).Head(E_Heading.WEST).GrhIndex).nombreGrafico = "Cabeza " & nombre & i & " Oeste"
            GrhData(HeadData(elemento).Head(E_Heading.EAST).GrhIndex).nombreGrafico = "Cabeza " & nombre & i & " Este"
            
            GrhData(HeadData(elemento).Head(E_Heading.NORTH).GrhIndex).esInsertableEnMapa = False
            GrhData(HeadData(elemento).Head(E_Heading.SOUTH).GrhIndex).esInsertableEnMapa = False
            GrhData(HeadData(elemento).Head(E_Heading.WEST).GrhIndex).esInsertableEnMapa = False
            GrhData(HeadData(elemento).Head(E_Heading.EAST).GrhIndex).esInsertableEnMapa = False
            
            
            Call Me_indexar_Graficos.actualizarEnIni(HeadData(elemento).Head(E_Heading.NORTH).GrhIndex)
            Call Me_indexar_Graficos.actualizarEnIni(HeadData(elemento).Head(E_Heading.SOUTH).GrhIndex)
            Call Me_indexar_Graficos.actualizarEnIni(HeadData(elemento).Head(E_Heading.WEST).GrhIndex)
            Call Me_indexar_Graficos.actualizarEnIni(HeadData(elemento).Head(E_Heading.EAST).GrhIndex)
             
            ' Actualizamos en disco
            Call Me_indexar_Cabezas.actualizarEnIni(elemento)
            i = i + 1
    Next
    
End Sub

' ****************************************************************************
'        METODOS PARA ACTUALIZAR EL ASPECTO CUANDO SE MODIFICA ALGUNA PROPIEDADES
Private Sub cmdAplicar_Cuerpos_Click()
    Dim numero As Integer
    Dim direccion As Byte
    
    numero = CInt(Me.lblNumeroCuerpoResultado)
    
    'Guardo en el Slot
    BodyData(numero).nombre = cuerpoSeleccionado.nombre
    BodyData(numero).HeadOffset = cuerpoSeleccionado.HeadOffset
    
    For direccion = E_Heading.NORTH To E_Heading.WEST
        InitGrh BodyData(numero).Walk(direccion), cuerpoSeleccionado.Walk(direccion).GrhIndex
    Next
    
    'Persisto el cambio
    Call Me_indexar_Cuerpos.actualizarEnIni(CLng(numero))
        
    Call ActualizarElementoEnlistaCuerpos(numero)
    
    Call ME_ControlCambios.SetHayCambiosSinActualiar("Cuerpos")
    
    Me.cmdAplicar_Cuerpos.Enabled = False
    Me.cmdRestablecer_Cuerpos.Enabled = False
End Sub
' ****************************************************************************
'        METODOS PARA ACTUALIZAR LA LISTA
Private Sub ActualizarElementoEnlistaCabezas(id As Integer)
    Dim direccion As Byte
    Dim numeroGrh As Integer
    
    'Actualizo el nombre
    Call Me.tree_Cabezas.cambiarNombre(CLng(id), id & " - " & HeadData(id).nombre)
        
    Call Me.tree_Cabezas.eliminarHijos(CLng(id))
        
    'Cargo las animaciones de cada perfil
    For direccion = E_Heading.NORTH To E_Heading.WEST
        numeroGrh = HeadData(id).Head(direccion).GrhIndex
        Call Me.tree_Cabezas.addString(CLng(numeroGrh), numeroGrh & " - " & GrhData(numeroGrh).nombreGrafico, CLng(id))
    Next
End Sub
Private Sub ActualizarElementoEnlistaCuerpos(id As Integer)
    Dim direccion As Byte
    Dim numeroGrh As Integer

    'Actualizo el nombre
    Call Me.tree_Cuerpo.cambiarNombre(CLng(id), id & " - " & BodyData(id).nombre)
    
    'Actualizo los hijos
    Call Me.tree_Cuerpo.eliminarHijos(CLng(id))
        
    'Cargo las animaciones de cada perfil
    For direccion = E_Heading.NORTH To E_Heading.WEST
        numeroGrh = BodyData(id).Walk(direccion).GrhIndex
        Call Me.tree_Cuerpo.addString(CLng(numeroGrh), numeroGrh & " - " & GrhData(numeroGrh).nombreGrafico, CLng(id))
    Next
End Sub
Private Sub ActualizarElementoEnlistaCascos(id As Integer)
    Dim direccion As Byte
    Dim numeroGrh As Integer

    'Actualizo el nombre
    Call Me.tree_Cascos.cambiarNombre(CLng(id), id & " - " & CascoAnimData(id).nombre)
    
    'Actualizo los hijos
    Call Me.tree_Cascos.eliminarHijos(CLng(id))
        
    'Cargo las animaciones de cada perfil
    For direccion = E_Heading.NORTH To E_Heading.WEST
        numeroGrh = CascoAnimData(id).Head(direccion).GrhIndex
        Call Me.tree_Cascos.addString(CLng(numeroGrh), numeroGrh & " - " & GrhData(numeroGrh).nombreGrafico, CLng(id))
    Next
End Sub
Private Sub ActualizarElementoEnlistaEscudos(id As Integer)
    Dim direccion As Byte
    Dim numeroGrh As Integer

    'Actualizo el nombre
    Call Me.tree_Escudos.cambiarNombre(CLng(id), id & " - " & ShieldAnimData(id).nombre)
    
    'Actualizo los hijos
    Call Me.tree_Escudos.eliminarHijos(CLng(id))
        
    'Cargo las animaciones de cada perfil
    For direccion = E_Heading.NORTH To E_Heading.WEST
        numeroGrh = ShieldAnimData(id).ShieldWalk(direccion).GrhIndex
        Call Me.tree_Escudos.addString(CLng(numeroGrh), numeroGrh & " - " & GrhData(numeroGrh).nombreGrafico, CLng(id))
    Next
End Sub

Private Sub ActualizarElementoEnlistaArmas(id As Integer)
    Dim direccion As Byte
    Dim numeroGrh As Integer

    'Actualizo el nombre
    Call Me.tree_Armas.cambiarNombre(CLng(id), id & " - " & WeaponAnimData(id).nombre)
    
    'Actualizo los hijos
    Call Me.tree_Armas.eliminarHijos(CLng(id))
        
    'Cargo las animaciones de cada perfil
    For direccion = E_Heading.NORTH To E_Heading.WEST
        numeroGrh = WeaponAnimData(id).WeaponWalk(direccion).GrhIndex
        Call Me.tree_Escudos.addString(CLng(numeroGrh), numeroGrh & " - " & GrhData(numeroGrh).nombreGrafico, CLng(id))
    Next
End Sub

Private Sub Cancelar()
    Call restablecerEstadoActualChar
    Unload Me
End Sub

Private Sub cmdAplicar_Armas_Click()
    Dim numero As Integer
    Dim direccion As Byte
    
    numero = CInt(Me.lblArmaNumeroResultado)
    
    'Guardo en el Slot
    WeaponAnimData(numero).nombre = armaSeleccionado.nombre
    
    For direccion = E_Heading.NORTH To E_Heading.WEST
        InitGrh WeaponAnimData(numero).WeaponWalk(direccion), armaSeleccionado.WeaponWalk(direccion).GrhIndex
    Next
    
    'Actualizo la lista
    Call ActualizarElementoEnlistaArmas(numero)
        
    'Persisto el cambio
    Call Me_indexar_Armas.actualizarEnIni(numero)
    
    Call ME_ControlCambios.SetHayCambiosSinActualiar("Armas")
    
    'Estado de los botones
    Me.cmdAplicar_Armas.Enabled = False
    Me.cmdRestablecer_Armas.Enabled = False
End Sub

Private Sub cmdEliminar_Armas_Click()
    Dim numeroCuerpo As Integer
    Dim confirma As VbMsgBoxResult
    Dim idElemento As Integer
    
    If Not Me.tree_Armas.obtenerValor() = "" Then
        
        idElemento = Me.tree_Armas.obtenerIDValor
        
        confirma = MsgBox("¿Está seguro de que desea eliminar el arma '" & Me.tree_Armas.obtenerValor & "'?", vbYesNo + vbExclamation, "Configurar Gráficos")
        
        If confirma = vbYes Then
            Call Me_indexar_Armas.eliminar(idElemento)
            
            'Lo borramos de la lista
            Call Me.tree_Armas.eliminarElemento(CLng(idElemento))
            
            Me.cmdEliminar_Armas.Enabled = False
        End If
    End If
End Sub

Private Sub cmdNuevo_Armas_Click()
    Dim nuevo As Integer
    Dim direccion As Integer
    Dim error As Boolean
    
    Me.cmdNuevo_Armas.Enabled = False
    
    'Obtengo el nuevo id
    nuevo = Me_indexar_Armas.nuevo
    
    If nuevo = -1 Then
        error = True
        MsgBox "No se ha podido obtener espacio para agregar una nueva arma. Por favor, intenta más tarde o contactate con un administrador del sistema.", vbExclamation
    End If
    
    If Not error Then
        'Lo agrego a la lista
        If Me.tree_Armas.seleccionarElemento(CLng(nuevo)) = False Then
            Call Me.tree_Armas.addString(CLng(nuevo), nuevo & " - ", 0)
            Call Me.tree_Armas.seleccionarElemento(CLng(nuevo))
        End If
    End If
    
    Me.cmdNuevo_Armas.Enabled = True
    'Cuando se haga clic en "Aplicar" se guarda
End Sub

Private Sub cmdAplicar_Cabezas_Click()
    Dim numero As Integer
    Dim direccion As Byte
    
    numero = CInt(Me.lblCabezaNumeroResultado)
    
    'Guardo en el Slot
    HeadData(numero).nombre = cabezaSeleccionado.nombre
    
    For direccion = E_Heading.NORTH To E_Heading.WEST
        InitGrh HeadData(numero).Head(direccion), cabezaSeleccionado.Head(direccion).GrhIndex
    Next
    
    'Actualizo la lista
    Call ActualizarElementoEnlistaCabezas(numero)
        
    'Persisto el cambio
    Call Me_indexar_Cabezas.actualizarEnIni(numero)
    
    Call ME_ControlCambios.SetHayCambiosSinActualiar("Cabezas")
    
    'Estado de los botones
    Me.cmdAplicar_Cabezas.Enabled = False
    Me.cmdRestablecer_Cabezas.Enabled = False
End Sub


Private Sub cmdEliminar_Cabezas_Click()
    Dim numeroCuerpo As Integer
    Dim confirma As VbMsgBoxResult
    Dim idElemento As Integer
    
    If Not Me.tree_Cabezas.obtenerIDValor > 0 Then
        
        idElemento = Me.tree_Cabezas.obtenerIDValor
        
        confirma = MsgBox("¿Está seguro de que desea eliminar la cabeza '" & Me.tree_Cabezas.obtenerValor & "'?", vbYesNo + vbExclamation)
        
        If confirma = vbYes Then
            Call Me_indexar_Cabezas.eliminar(idElemento)
            'Lo borramos de la lista
            Call Me.tree_Cabezas.eliminarElemento(CLng(idElemento))
            
            Me.cmdEliminar_Cabezas.Enabled = False
        End If
    End If
End Sub

Private Sub cmdNuevo_Cabezas_Click()
    Dim nuevo As Integer
    Dim error As Boolean
    
    Me.cmdNuevo_Cabezas.Enabled = False
    
    'Obtengo el nuevo id
    nuevo = Me_indexar_Cabezas.nuevo
    
    If nuevo = -1 Then
        error = True
        MsgBox "No se ha podido obtener espacio para agregar una nueva cabeza. Por favor, intenta más tarde o contactate con un administrador del sistema.", vbExclamation
    End If
    
    If Not error Then
        'Lo agrego a la lista
        If Me.tree_Cabezas.seleccionarElemento(CLng(nuevo)) = False Then
            Call Me.tree_Cabezas.addString(CLng(nuevo), nuevo & " - ", 0)
            Call Me.tree_Cabezas.seleccionarElemento(CLng(nuevo))
        End If
    End If
    
    Me.cmdNuevo_Cabezas.Enabled = True
    'Cuando se haga clic en "Aplicar" se guarda
End Sub

Private Sub cmdAplicar_cascos_Click()
    Dim numero As Integer
    Dim direccion As Byte
    
    numero = CInt(Me.lblCascoNumeroResultado)
    
    'Guardo en el Slot
    CascoAnimData(numero).nombre = cascoSeleccionado.nombre
    
    For direccion = E_Heading.NORTH To E_Heading.WEST
        InitGrh CascoAnimData(numero).Head(direccion), cascoSeleccionado.Head(direccion).GrhIndex
    Next
    
    'Actualizo la lista
    Call ActualizarElementoEnlistaCascos(numero)
    
    'Persisto el cambio
    Call Me_indexar_Cascos.actualizarEnIni(numero)
    
    Call ME_ControlCambios.SetHayCambiosSinActualiar("Cascos")
    
    'Estado de los botones
    Me.cmdAplicar_Cascos.Enabled = False
    Me.cmdRestablecer_Cascos.Enabled = False
End Sub

Private Sub cmdEliminar_Cascos_Click()
    Dim numeroCuerpo As Integer
    Dim confirma As VbMsgBoxResult
    Dim idElemento As Integer
    
    If Not Me.tree_Cascos.obtenerValor() = "" Then
        
        idElemento = Me.tree_Cascos.obtenerIDValor
        
        confirma = MsgBox("¿Está seguro de que desea eliminar el casco '" & Me.tree_Cascos.obtenerValor & "'?", vbYesNo + vbExclamation, "Configurar Gráficos")
        
        If confirma = vbYes Then
            Call Me_indexar_Cascos.eliminar(idElemento)
            'Lo borramos de la lista
            Call Me.tree_Cascos.eliminarElemento(CLng(idElemento))
        
            Me.cmdEliminar_Cascos.Enabled = False
        End If
    End If
End Sub

Private Sub cmdNuevo_Cascos_Click()
    Dim nuevo As Integer
    Dim error As Boolean
    
    Me.cmdNuevo_Cascos.Enabled = False
    'Obtengo el nuevo id
    nuevo = Me_indexar_Cascos.nuevo
    
    If nuevo = -1 Then
        error = True
        MsgBox "No se ha podido obtener espacio para agregar un nuevo casco. Por favor, intenta más tarde o contactate con un administrador del sistema.", vbExclamation
    End If
    
    If Not error Then
        'Lo agrego a la lista
        If Me.tree_Cascos.seleccionarElemento(CLng(nuevo)) = False Then
            Call Me.tree_Cascos.addString(CLng(nuevo), nuevo & " - ", 0)
            Call Me.tree_Cascos.seleccionarElemento(CLng(nuevo))
        End If
    End If
    
    Me.cmdNuevo_Cascos.Enabled = True
    'Cuando se haga clic en "Aplicar" se guarda
End Sub

Private Sub cmdEliminar_Cuerpos_Click()
Dim confirma As VbMsgBoxResult
Dim idElemento As Integer

If Not Me.tree_Cuerpo.obtenerValor() = "" Then
    
    idElemento = Me.tree_Cuerpo.obtenerIDValor
    
    confirma = MsgBox("¿Está seguro de que desea eliminar el cuerpo '" & Me.tree_Cuerpo.obtenerValor & "'?", vbYesNo + vbExclamation, "Configurar Gráficos")
    
    If confirma = vbYes Then
        Call Me_indexar_Cuerpos.eliminar(idElemento)
        'Lo borramos de la lista
        Call Me.tree_Cuerpo.eliminarElemento(CLng(idElemento))
        
        Me.cmdEliminar_Cuerpos.Enabled = False
    End If
End If
End Sub

Private Sub cmdAplicar_Escudos_Click()
    Dim numero As Integer
    Dim direccion As Byte
    
    numero = CInt(Me.lblEscudoNumeroResultado)
    
    'Guardo en el Slot
    ShieldAnimData(numero).nombre = escudoSeleccionado.nombre
    
    For direccion = E_Heading.NORTH To E_Heading.WEST
        InitGrh ShieldAnimData(numero).ShieldWalk(direccion), escudoSeleccionado.ShieldWalk(direccion).GrhIndex
    Next
    
    'Actualizo la lista
    Call ActualizarElementoEnlistaEscudos(numero)
        
    'Persisto el cambio
    Call Me_indexar_Escudos.actualizarEnIni(numero)
    
    Call ME_ControlCambios.SetHayCambiosSinActualiar("Escudos")
    
    'Estado de los botones
    Me.cmdAplicar_Escudos.Enabled = False
    Me.cmdRestablecer_Escudos.Enabled = False
End Sub

Private Sub cmdEliminar_Escudos_Click()
    Dim confirma As VbMsgBoxResult
    Dim idElemento As Integer
    
    If Not Me.tree_Escudos.obtenerValor() = "" Then
        
        idElemento = Me.tree_Escudos.obtenerIDValor
        
        confirma = MsgBox("¿Está seguro de que desea eliminar el escudo '" & Me.tree_Escudos.obtenerValor & "'?", vbYesNo + vbExclamation, "Configurar Gráficos")
        
        If confirma = vbYes Then
            Call Me_indexar_Escudos.eliminar(idElemento)
            'Lo borramos de la lista
            Call Me.tree_Escudos.eliminarElemento(CLng(idElemento))
            
            Me.cmdEliminar_Escudos.Enabled = False
        End If
    End If
End Sub

Private Sub cmdNuevo_Escudos_Click()
    Dim nuevo As Integer
    Dim error As Boolean
    
    Me.cmdNuevo_Escudos.Enabled = False
    'Obtengo el nuevo id
    nuevo = Me_indexar_Escudos.nuevo
    
    If nuevo = -1 Then
        error = True
        MsgBox "No se ha podido obtener espacio para agregar un nuevo escudo. Por favor, intenta más tarde o contactate con un administrador del sistema.", vbExclamation
    End If
    
    If Not error Then
        'Lo agrego a la lista
        If Me.tree_Escudos.seleccionarElemento(CLng(nuevo)) = False Then
            Call Me.tree_Escudos.addString(CLng(nuevo), nuevo & " - ", 0)
            Call Me.tree_Escudos.seleccionarElemento(CLng(nuevo))
        End If
    End If
    
    Me.cmdNuevo_Escudos.Enabled = True
    'Cuando se haga clic en "Aplicar" se guarda
End Sub

Private Sub cmdNuevo_Cuerpos_Click()
    Dim nuevo As Integer
    Dim error As Boolean
    
    Me.cmdNuevo_Cuerpos.Enabled = False

    'Obtengo el nuevo id
    nuevo = Me_indexar_Cuerpos.nuevo
    
    If nuevo = -1 Then
        error = True
        MsgBox "No se ha podido obtener espacio para agregar un nuevo cuerpo. Por favor, intenta más tarde o contactate con un administrador del sistema.", vbExclamation
    End If
    
    If Not error Then
        'Lo agrego a la lista
        If Me.tree_Cuerpo.seleccionarElemento(CLng(nuevo)) = False Then
            Call Me.tree_Cuerpo.addString(CLng(nuevo), nuevo & " - ", 0)
            Call Me.tree_Cuerpo.seleccionarElemento(CLng(nuevo))
        End If
    End If
    
    Me.cmdNuevo_Cuerpos.Enabled = True
    'Cuando se haga clic en "Aplicar" se guarda
End Sub


Private Sub cmdOptionManoHabil_Click(Index As Integer)
    CharList(UserCharIndex).invh = Me.cmdOptionManoHabil(1).value
End Sub





Private Sub Form_Load()

Dim i As Integer
Dim direccion As Integer
Dim numeroGrh As Integer


'Sino esta el modo caminata activado lo activamos
If ME_Render.WalkMode = False Then
    Call ME_Render.ToggleWalkMode
End If

'Sino pude activarlo, salgo
If ME_Render.WalkMode = False Then Unload Me: Exit Sub

Call guardarEstadoActualChar

Call establecerSeleccionDefault

'******************************************************************************
'*********************** CUERPOS **********************************************
'Cargo la lista de cuerpos
Call Me.tree_Cuerpo.addString(0, "0 - Nada", 0)
    
For i = 1 To UBound(BodyData)

    If Me_indexar_Cuerpos.existe(i) Then
        
        Call Me.tree_Cuerpo.addString(CLng(i), CStr(i & " - " & BodyData(i).nombre), 0)
            
        'Cargo las animaciones de cada perfil
        For direccion = E_Heading.NORTH To E_Heading.WEST
            numeroGrh = BodyData(i).Walk(direccion).GrhIndex
            Call Me.tree_Cuerpo.addString(CLng(numeroGrh), numeroGrh & " - " & GrhData(numeroGrh).nombreGrafico, CLng(i))
        Next
    End If
Next i

'Inicialmente el formulario arranca desactivado
Call modPosicionarFormulario.setEnabledHijos(False, Me.frmPropiedades_Cuerpos, Me)
Me.cmdEliminar_Cuerpos.Enabled = False

'******************************************************************************
'*********************** CABEZAS **********************************************
'Cargo la lista de Cabezas
Call Me.tree_Cabezas.addString(0, "0 - Nada", 0)
    
For i = 1 To UBound(HeadData)

    If Me_indexar_Cabezas.existe(i) Then
        Call Me.tree_Cabezas.addString(CLng(i), CStr(i & " - " & HeadData(i).nombre), 0)
        
            'Cargo las animaciones de cada perfil
            For direccion = E_Heading.NORTH To E_Heading.WEST
                numeroGrh = HeadData(i).Head(direccion).GrhIndex
                Call Me.tree_Cabezas.addString(CLng(numeroGrh), numeroGrh & " - " & GrhData(numeroGrh).nombreGrafico, CLng(i))
            Next
    End If
Next i

'Inicialmente el formulario arranca desactivado
Call modPosicionarFormulario.setEnabledHijos(False, Me.frmPropiedades_Cabezas, Me)
Me.cmdEliminar_Cabezas.Enabled = False

'******************************************************************************
'*********************** CASCOS **********************************************
'Cargo la lista de Cascos
Call Me.tree_Cascos.addString(0, "0 - Nada ", 0)
    
For i = 1 To UBound(CascoAnimData)

    If Me_indexar_Cascos.existe(i) Then
        Call Me.tree_Cascos.addString(CLng(i), CStr(i & " - " & CascoAnimData(i).nombre), 0)
        
        'Cargo las animaciones de cada perfil
            For direccion = E_Heading.NORTH To E_Heading.WEST
                numeroGrh = CascoAnimData(i).Head(direccion).GrhIndex
                Call Me.tree_Cascos.addString(CLng(numeroGrh), numeroGrh & " - " & GrhData(numeroGrh).nombreGrafico, CLng(i))
            Next
    End If
   
Next i

'Inicialmente el formulario arranca desactivado
Call modPosicionarFormulario.setEnabledHijos(False, Me.frmPropiedades_Cascos, Me)
Me.cmdEliminar_Cascos.Enabled = False

'******************************************************************************
'*********************** ESCUDOS **********************************************
'Cargo la lista de Escudos
Call Me.tree_Escudos.addString(CLng(0), CStr("0 - Nada"), 0)

For i = 1 To UBound(ShieldAnimData)
    'Cargo las animaciones de cada perfil
    If Me_indexar_Escudos.existe(i) Then
        
        Call Me.tree_Escudos.addString(CLng(i), CStr(i & " - " & ShieldAnimData(i).nombre), 0)
            
        For direccion = E_Heading.NORTH To E_Heading.WEST
            numeroGrh = ShieldAnimData(i).ShieldWalk(direccion).GrhIndex
            Call Me.tree_Escudos.addString(CLng(numeroGrh), numeroGrh & " - " & GrhData(numeroGrh).nombreGrafico, CLng(i))
        Next
        
    End If
Next i

'Inicialmente el formulario arranca desactivado
Call modPosicionarFormulario.setEnabledHijos(False, Me.frmPropiedades_Escudos, Me)
Me.cmdEliminar_Escudos.Enabled = False

'******************************************************************************
'*********************** ARMAS **********************************************
'Cargo la lista de Armas

For i = LBound(WeaponAnimData) To UBound(WeaponAnimData)

    If Me_indexar_Armas.existe(i) Then
        Call Me.tree_Armas.addString(CLng(i), CStr(i & " - " & WeaponAnimData(i).nombre), 0)
        
        'Cargo las animaciones de cada perfil
        For direccion = E_Heading.NORTH To E_Heading.WEST
            numeroGrh = WeaponAnimData(i).WeaponWalk(direccion).GrhIndex
            Call Me.tree_Armas.addString(CLng(numeroGrh), numeroGrh & " - " & GrhData(numeroGrh).nombreGrafico, CLng(i))
        Next
    End If
   
Next i

'Inicialmente el formulario arranca desactivado
Call modPosicionarFormulario.setEnabledHijos(False, Me.frmPropiedades_Armas, Me)
Me.cmdEliminar_Armas.Enabled = False

'******************************************************************************
'******************************************************************************

'Cargo los grh posibles que tiene el juego
For i = LBound(GrhData) To UBound(GrhData)
    If GrhData(i).NumFrames > 0 Then
        Call Me.lstlGrhDisponibles.addString(i, i & " - " & GrhData(i).nombreGrafico)
    End If
Next

Me.txtCuerpoOffsetY.MaxValue = 1000
Me.txtCuerpoOffsetY.MinValue = -1000

Me.txtCuerpoOffsetX.MaxValue = 1000
Me.txtCuerpoOffsetX.MinValue = -1000

Me.cmdOptionManoHabil(1).value = (CharList(UserCharIndex).invh)
Me.cmdOptionManoHabil(0).value = Not Me.cmdOptionManoHabil(1).value
End Sub

' ****************************************************************************
'        BOTONES RESTABLECER
Private Sub cmdRestablecer_Cuerpos_Click()
    If Not Me.tree_Cuerpo.obtenerIDValor = 0 Then Call cargarCuerposEnEditor(Me.tree_Cuerpo.obtenerIDValor)
End Sub
Private Sub cmdRestablecer_Armas_Click()
    If Not Me.tree_Armas.obtenerIDValor = 0 Then Call cargarArmasEnEditor(Me.tree_Armas.obtenerIDValor)
End Sub
Private Sub cmdRestablecer_Cabezas_Click()
    If Not Me.tree_Cabezas.obtenerIDValor = 0 Then Call cargarCabezasEnEditor(Me.tree_Cabezas.obtenerIDValor)
End Sub
Private Sub cmdRestablecer_Cascos_Click()
    If Not Me.tree_Cascos.obtenerIDValor = 0 Then Call cargarCascosEnEditor(Me.tree_Cascos.obtenerIDValor)
End Sub
Private Sub cmdRestablecer_Escudos_Click()
    If Not Me.tree_Escudos.obtenerIDValor = 0 Then Call cargarEscudoEnEditor(Me.tree_Escudos.obtenerIDValor)
End Sub

Private Sub lstlGrhDisponibles_Desactivado()
    Call OcultarLista
End Sub

' ****************************************************************************
'        CUANDO SE SELECCIONA UN ELEMENTO DE LA LISTA
Private Sub tree_Armas_Change(valor As String, id As Integer, esPadre As Boolean)
    If esPadre Or id = 0 Then cargarArmasEnEditor (id)
End Sub

Private Sub tree_Cabezas_Change(valor As String, id As Integer, esPadre As Boolean)
    If esPadre Then cargarCabezasEnEditor (id)
End Sub

Private Sub tree_Cascos_Change(valor As String, id As Integer, esPadre As Boolean)
    If esPadre Then cargarCascosEnEditor (id)
End Sub

Private Sub tree_Cuerpo_Change(valor As String, id As Integer, esPadre As Boolean)
    If esPadre Then cargarCuerposEnEditor (id)
End Sub

Private Sub tree_Escudos_Change(valor As String, id As Integer, esPadre As Boolean)
    If esPadre Then cargarEscudoEnEditor (id)
End Sub

' ****************************************************************************
'        METODOS QUE CARGAN UN ELEMENTO EN EL EDITOR
Private Sub cargarArmasEnEditor(ByVal id As Integer)
    Me.lblArmaNumeroResultado = id

    Me.txtArmaNombre = WeaponAnimData(id).nombre

    Me.txtArma(E_Heading.NORTH) = WeaponAnimData(id).WeaponWalk(E_Heading.NORTH).GrhIndex & " - " & GrhData(WeaponAnimData(id).WeaponWalk(E_Heading.NORTH).GrhIndex).nombreGrafico
    Me.txtArma(E_Heading.SOUTH) = WeaponAnimData(id).WeaponWalk(E_Heading.SOUTH).GrhIndex & " - " & GrhData(WeaponAnimData(id).WeaponWalk(E_Heading.SOUTH).GrhIndex).nombreGrafico
    Me.txtArma(E_Heading.EAST) = WeaponAnimData(id).WeaponWalk(E_Heading.EAST).GrhIndex & " - " & GrhData(WeaponAnimData(id).WeaponWalk(E_Heading.EAST).GrhIndex).nombreGrafico
    Me.txtArma(E_Heading.WEST) = WeaponAnimData(id).WeaponWalk(E_Heading.WEST).GrhIndex & " - " & GrhData(WeaponAnimData(id).WeaponWalk(E_Heading.WEST).GrhIndex).nombreGrafico
    
    Call actualizarArmaActual
    Call actualizarPersonaje
    
    'Activamos el formulario de propiedades
    Call modPosicionarFormulario.setEnabledHijos(True, Me.frmPropiedades_Armas, Me)
    Me.cmdAplicar_Armas.Enabled = False
    Me.cmdRestablecer_Armas.Enabled = False
    Me.cmdEliminar_Armas.Enabled = True
        
End Sub
Private Sub cargarCabezasEnEditor(ByVal id As Integer)
    Me.lblCabezaNumeroResultado = id

    Me.txtNombreCabeza = HeadData(id).nombre
    
    Me.txtCabeza(E_Heading.NORTH) = HeadData(id).Head(E_Heading.NORTH).GrhIndex & " - " & GrhData(HeadData(id).Head(E_Heading.NORTH).GrhIndex).nombreGrafico
    Me.txtCabeza(E_Heading.SOUTH) = HeadData(id).Head(E_Heading.SOUTH).GrhIndex & " - " & GrhData(HeadData(id).Head(E_Heading.SOUTH).GrhIndex).nombreGrafico
    Me.txtCabeza(E_Heading.EAST) = HeadData(id).Head(E_Heading.EAST).GrhIndex & " - " & GrhData(HeadData(id).Head(E_Heading.EAST).GrhIndex).nombreGrafico
    Me.txtCabeza(E_Heading.WEST) = HeadData(id).Head(E_Heading.WEST).GrhIndex & " - " & GrhData(HeadData(id).Head(E_Heading.WEST).GrhIndex).nombreGrafico
    
    Call actualizarCabezaActual
    Call actualizarPersonaje
    
    'Activamos el formulario de propiedades
    Call modPosicionarFormulario.setEnabledHijos(True, Me.frmPropiedades_Cabezas, Me)
    Me.cmdAplicar_Cabezas.Enabled = False
    Me.cmdRestablecer_Cabezas.Enabled = False
    Me.cmdEliminar_Cabezas.Enabled = True
End Sub
Private Sub cargarCascosEnEditor(ByVal id As Integer)
    Me.lblCascoNumeroResultado = id

    Me.txtCascoNombre = CascoAnimData(id).nombre

    Me.txtCasco(E_Heading.NORTH) = CascoAnimData(id).Head(E_Heading.NORTH).GrhIndex & " - " & GrhData(CascoAnimData(id).Head(E_Heading.NORTH).GrhIndex).nombreGrafico
    Me.txtCasco(E_Heading.SOUTH) = CascoAnimData(id).Head(E_Heading.SOUTH).GrhIndex & " - " & GrhData(CascoAnimData(id).Head(E_Heading.SOUTH).GrhIndex).nombreGrafico
    Me.txtCasco(E_Heading.EAST) = CascoAnimData(id).Head(E_Heading.EAST).GrhIndex & " - " & GrhData(CascoAnimData(id).Head(E_Heading.EAST).GrhIndex).nombreGrafico
    Me.txtCasco(E_Heading.WEST) = CascoAnimData(id).Head(E_Heading.WEST).GrhIndex & " - " & GrhData(CascoAnimData(id).Head(E_Heading.WEST).GrhIndex).nombreGrafico
           
    Call actualizarCascoActual
    Call actualizarPersonaje
    
    'Activamos el formulario de propiedades
    Call modPosicionarFormulario.setEnabledHijos(True, Me.frmPropiedades_Cascos, Me)
    Me.cmdAplicar_Cascos.Enabled = False
    Me.cmdRestablecer_Cascos.Enabled = False
    Me.cmdEliminar_Cascos.Enabled = True
End Sub
Private Sub cargarCuerposEnEditor(ByVal id As Integer)
    'Actualizo la info en el editor
    Me.lblNumeroCuerpoResultado = id

    Me.txtCuerpoNombre = BodyData(id).nombre

    Me.txtCuerpoOffsetX.value = BodyData(id).HeadOffset.X
    Me.txtCuerpoOffsetY.value = BodyData(id).HeadOffset.Y
    
    Me.txtCuerpo(E_Heading.NORTH) = BodyData(id).Walk(E_Heading.NORTH).GrhIndex & " - " & GrhData(BodyData(id).Walk(E_Heading.NORTH).GrhIndex).nombreGrafico
    Me.txtCuerpo(E_Heading.SOUTH) = BodyData(id).Walk(E_Heading.SOUTH).GrhIndex & " - " & GrhData(BodyData(id).Walk(E_Heading.SOUTH).GrhIndex).nombreGrafico
    Me.txtCuerpo(E_Heading.EAST) = BodyData(id).Walk(E_Heading.EAST).GrhIndex & " - " & GrhData(BodyData(id).Walk(E_Heading.EAST).GrhIndex).nombreGrafico
    Me.txtCuerpo(E_Heading.WEST) = BodyData(id).Walk(E_Heading.WEST).GrhIndex & " - " & GrhData(BodyData(id).Walk(E_Heading.WEST).GrhIndex).nombreGrafico
    
    'Actualizo la vista
    Call actualizarCuerpoActual
    Call actualizarPersonaje
    
    'Activamos el formulario de propiedades
    Call modPosicionarFormulario.setEnabledHijos(True, Me.frmPropiedades_Cuerpos, Me)
    Me.cmdAplicar_Cuerpos.Enabled = False
    Me.cmdRestablecer_Cuerpos.Enabled = False
    Me.cmdEliminar_Cuerpos.Enabled = True
End Sub
Private Sub cargarEscudoEnEditor(ByVal id As Integer)
    Me.lblEscudoNumeroResultado = id

    Me.txtEscudoNombre = ShieldAnimData(id).nombre

    Me.txtEscudo(E_Heading.NORTH) = ShieldAnimData(id).ShieldWalk(E_Heading.NORTH).GrhIndex & " - " & GrhData(ShieldAnimData(id).ShieldWalk(E_Heading.NORTH).GrhIndex).nombreGrafico
    Me.txtEscudo(E_Heading.SOUTH) = ShieldAnimData(id).ShieldWalk(E_Heading.SOUTH).GrhIndex & " - " & GrhData(ShieldAnimData(id).ShieldWalk(E_Heading.SOUTH).GrhIndex).nombreGrafico
    Me.txtEscudo(E_Heading.EAST) = ShieldAnimData(id).ShieldWalk(E_Heading.EAST).GrhIndex & " - " & GrhData(ShieldAnimData(id).ShieldWalk(E_Heading.EAST).GrhIndex).nombreGrafico
    Me.txtEscudo(E_Heading.WEST) = ShieldAnimData(id).ShieldWalk(E_Heading.WEST).GrhIndex & " - " & GrhData(ShieldAnimData(id).ShieldWalk(E_Heading.WEST).GrhIndex).nombreGrafico
    
    Call actualizarEscudoActual
    Call actualizarPersonaje
    
    'Activamos el formulario de propiedades
    Call modPosicionarFormulario.setEnabledHijos(True, Me.frmPropiedades_Escudos, Me)
    Me.cmdAplicar_Escudos.Enabled = False
    Me.cmdRestablecer_Escudos.Enabled = False
    Me.cmdEliminar_Escudos.Enabled = True
End Sub
' ****************************************************************************
'        METODOS GLOBALES FORMULARIO
Private Sub Form_Unload(Cancel As Integer)
    If Cancel = 1 Then Call Cancelar
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Call Cancelar
End Sub

' ****************************************************************************
'        METODOS PARA ACTUALIZAR EL ASPECTO CUANDO SE MODIFICA ALGUNA PROPIEDADES
Private Sub txtArma_Change(Index As Integer)
    Call actualizarArmaActual
    Call actualizarPersonaje
End Sub

Private Sub txtArmaNombre_Change()
    Call actualizarArmaActual
End Sub

Private Sub txtCabeza_Change(Index As Integer)
    Call actualizarCabezaActual
    Call actualizarPersonaje
End Sub

Private Sub txtCascoNombre_Change()
    Call actualizarCascoActual
End Sub

Private Sub txtCuerpoNombre_Change()
    Call actualizarCuerpoActual
End Sub

Private Sub txtCuerpoOffsetX_Change(valor As Double)
    Call actualizarCuerpoActual
    Call actualizarPersonaje
End Sub

Private Sub txtCuerpoOffsetY_Change(valor As Double)
    Call actualizarCuerpoActual
    Call actualizarPersonaje
End Sub

Private Sub txtEscudo_Change(Index As Integer)
    Call actualizarEscudoActual
    Call actualizarPersonaje
End Sub

Private Sub txtCuerpo_Change(Index As Integer)
    Call actualizarCuerpoActual
    Call actualizarPersonaje
End Sub

Private Sub txtCasco_Change(Index As Integer)
    Call actualizarCascoActual
    Call actualizarPersonaje
End Sub

' ****************************************************************************
'                   METODOS PARA LA LISTA ESPECIAL

' Cuando se cambia el valor en la lista se actualiza en el textbox
Private Sub lstlGrhDisponibles_Change(valor As String, id As Integer)
     campoSeleccionado.text = valor
End Sub



' Metodo que activa la lista desplegable en elt extbox seleccionado. Si hay otro seleccionado, lo deslecciona
Private Sub MostrarEn(text As TextBox)

    If Not campoSeleccionado Is Nothing Then
        campoSeleccionado.visible = True
    End If
    
    Set campoSeleccionado = text
    
    Me.lstlGrhDisponibles.top = text.Container.top + text.top + text.Container.Container.top
    Me.lstlGrhDisponibles.left = text.Container.left + text.left + text.Container.Container.left
    Me.lstlGrhDisponibles.visible = True
    campoSeleccionado.visible = False
    
    'Establece el foco en la lista
    Me.lstlGrhDisponibles.SetFocus
End Sub

Private Sub OcultarLista()
    If Not campoSeleccionado Is Nothing Then
        Me.lstlGrhDisponibles.visible = False
        campoSeleccionado.visible = True
        
        Set campoSeleccionado = Nothing
    End If
End Sub
Private Sub lstlGrhDisponibles_DblClic()
    Call OcultarLista
End Sub
Private Sub lstlGrhDisponibles_LostFocus()
   Call OcultarLista
End Sub
Private Sub txtArma_GotFocus(Index As Integer)
    Call MostrarEn(Me.txtArma(Index))
End Sub
Private Sub txtCabeza_GotFocus(Index As Integer)
    Call MostrarEn(Me.txtCabeza(Index))
End Sub

Private Sub txtCasco_GotFocus(Index As Integer)
    Call MostrarEn(Me.txtCasco(Index))
End Sub

Private Sub txtCuerpo_GotFocus(Index As Integer)
    Call MostrarEn(Me.txtCuerpo(Index))
End Sub

Private Sub txtEscudo_GotFocus(Index As Integer)
    Call MostrarEn(Me.txtEscudo(Index))
End Sub

Private Sub txtEscudoNombre_Change()
    Call actualizarEscudoActual
End Sub

Private Sub txtNombreCabeza_Change()
    Call actualizarCabezaActual
End Sub
