VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfigurarGraficos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Graficos"
   ClientHeight    =   7830
   ClientLeft      =   8745
   ClientTop       =   4230
   ClientWidth     =   9465
   Icon            =   "frmIndexacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   522
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   631
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmPanelSecreto 
      Caption         =   "Panel secreto"
      Height          =   3735
      Left            =   3000
      TabIndex        =   37
      Top             =   2160
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CommandButton cmdActualizarOffSet 
         Caption         =   "Actualizar OffSet"
         Height          =   360
         Left            =   600
         TabIndex        =   66
         Top             =   3240
         Width           =   3495
      End
      Begin VB.CommandButton cmdGenerarDelete 
         Caption         =   "Generar Delete"
         Height          =   360
         Left            =   600
         TabIndex        =   47
         Top             =   2760
         Width           =   3495
      End
      Begin VB.CommandButton cmdGraficosSinUitlizar 
         Caption         =   "Graficos sin utilizar"
         Height          =   360
         Left            =   600
         TabIndex        =   45
         Top             =   2280
         Width           =   3495
      End
      Begin VB.CommandButton cmdHechizos 
         Caption         =   "Hechizos"
         Height          =   360
         Left            =   600
         TabIndex        =   44
         Top             =   1800
         Width           =   3495
      End
      Begin VB.CommandButton cmdAplicarNombresAHijos 
         Caption         =   "Extender nombres"
         Height          =   360
         Left            =   600
         TabIndex        =   43
         Top             =   1320
         Width           =   3495
      End
      Begin VB.CommandButton cmdBorrarMultiple 
         Caption         =   "Borrar"
         Height          =   360
         Left            =   600
         TabIndex        =   42
         Top             =   840
         Width           =   3495
      End
      Begin EditorTDS.UpDownText nroHasta 
         Height          =   375
         Left            =   3000
         TabIndex        =   41
         Top             =   360
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
         maxvalue        =   0
         minvalue        =   0
         enabled         =   -1  'True
      End
      Begin EditorTDS.UpDownText nroInicio 
         Height          =   375
         Left            =   1200
         TabIndex        =   38
         Top             =   360
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         maxvalue        =   0
         minvalue        =   0
         enabled         =   -1  'True
      End
      Begin VB.Label lblBorrarMultipleHasta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         Height          =   195
         Left            =   2280
         TabIndex        =   40
         Top             =   480
         Width           =   465
      End
      Begin VB.Label lblBorrarMultipleDesde 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
         Height          =   195
         Left            =   600
         TabIndex        =   39
         Top             =   480
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdResetGUUID 
      Caption         =   "R"
      Height          =   240
      Left            =   9000
      TabIndex        =   74
      Top             =   675
      Width           =   330
   End
   Begin VB.CommandButton cmdImportarConfiguracion 
      Height          =   330
      Left            =   9000
      Picture         =   "frmIndexacion.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "Importar Configuración de Gráficos"
      Top             =   15
      Width           =   375
   End
   Begin MSComDlg.CommonDialog oFile 
      Left            =   2640
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmAjusteGrafico 
      Caption         =   "¿Cómo se ajusta el gráfico a la grilla?"
      Height          =   930
      Left            =   3720
      TabIndex        =   63
      Top             =   2280
      Width           =   5655
      Begin VB.OptionButton optCentrado 
         Appearance      =   0  'Flat
         Caption         =   "Centrar en un Tile"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   76
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optCentrado 
         Appearance      =   0  'Flat
         Caption         =   "Ajustar a la Grilla"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   75
         Top             =   240
         Width           =   1575
      End
      Begin EditorTDS.UpDownText sliderOffsetY 
         Height          =   315
         Left            =   2880
         TabIndex        =   70
         Top             =   540
         Width           =   855
         _extentx        =   1508
         _extenty        =   556
         maxvalue        =   0
         minvalue        =   0
         enabled         =   0   'False
      End
      Begin EditorTDS.UpDownText sliderOffsetX 
         Height          =   315
         Left            =   1680
         TabIndex        =   68
         Top             =   540
         Width           =   855
         _extentx        =   1508
         _extenty        =   556
         maxvalue        =   0
         minvalue        =   0
         enabled         =   0   'False
      End
      Begin VB.OptionButton optCentrado 
         Appearance      =   0  'Flat
         Caption         =   "Personalizado: X"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   67
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblOffsetY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         Height          =   195
         Left            =   2640
         TabIndex        =   69
         Top             =   600
         Width           =   150
      End
   End
   Begin VB.CommandButton cmdToogleCheckedList 
      Height          =   270
      Left            =   2925
      Picture         =   "frmIndexacion.frx":200C
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Activar o desactivar selección multiple"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmdCortar32H 
      Caption         =   "Cortar H"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   2520
      TabIndex        =   55
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCortar32V 
      Caption         =   "Cortar V"
      Height          =   360
      Left            =   120
      TabIndex        =   54
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdDuplicar 
      Caption         =   "Duplicar"
      Height          =   360
      Left            =   1320
      TabIndex        =   48
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   3720
      TabIndex        =   36
      Top             =   7080
      Width           =   5655
   End
   Begin VB.CommandButton cmdEliminar_Graficos 
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   1920
      TabIndex        =   23
      Top             =   7080
      Width           =   1695
   End
   Begin EditorTDS.TreeConBuscador Listado 
      Height          =   6345
      Left            =   120
      TabIndex        =   22
      Top             =   75
      Width           =   3495
      _extentx        =   6165
      _extenty        =   11192
   End
   Begin VB.CheckBox chkGraficoInsertableMapa 
      Appearance      =   0  'Flat
      Caption         =   "Este gráfico se puede insertar en el mapa. (No es un movimiento de criatura, cuerpo, arma, hechizo o un item)"
      ForeColor       =   &H80000008&
      Height          =   350
      Left            =   3720
      TabIndex        =   15
      Top             =   960
      Width           =   5655
   End
   Begin VB.TextBox txtNombreGrafico 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      MaxLength       =   50
      TabIndex        =   9
      Top             =   360
      Width           =   4815
   End
   Begin VB.OptionButton optTipoIndex 
      Appearance      =   0  'Flat
      Caption         =   "Simple  ó"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   4
      Top             =   3330
      Width           =   940
   End
   Begin VB.OptionButton optTipoIndex 
      Appearance      =   0  'Flat
      Caption         =   "Animacion"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   5
      Top             =   3330
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevo_Graficos 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton Reestablecer 
      Caption         =   "Reestablecer"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   6660
      Width           =   2775
   End
   Begin VB.CommandButton Visualizar 
      Caption         =   "Aplicar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   6660
      Width           =   2775
   End
   Begin VB.TextBox Index 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Nos indica el numero de GrhIndex"
      Top             =   360
      Width           =   735
   End
   Begin VB.Frame frmTipoAnimacion 
      Height          =   3255
      Left            =   3720
      TabIndex        =   6
      Top             =   3360
      Width           =   5655
      Begin VB.Frame frmIndexSimple 
         BorderStyle     =   0  'None
         Height          =   3045
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   5415
         Begin EditorTDS.TextConListaConBuscador FileName 
            Height          =   280
            Left            =   1800
            TabIndex        =   60
            Top             =   120
            Width           =   3615
            _extentx        =   6376
            _extenty        =   503
            cantidadlineasamostrar=   0
         End
         Begin EditorTDS.TextConListaConBuscador txtEfectoPisada 
            Height          =   285
            Left            =   3120
            TabIndex        =   58
            Top             =   720
            Width           =   2295
            _extentx        =   4048
            _extenty        =   503
            cantidadlineasamostrar=   0
         End
         Begin VB.TextBox txtComplementoNormal 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   35
            Text            =   "-"
            Top             =   2730
            Width           =   2655
         End
         Begin EditorTDS.UpDownText pixelHeight 
            Height          =   315
            Left            =   1800
            TabIndex        =   32
            Top             =   1560
            Width           =   1095
            _extentx        =   1931
            _extenty        =   556
            maxvalue        =   2048
            minvalue        =   0
            enabled         =   -1  'True
         End
         Begin EditorTDS.UpDownText pixelWidth 
            Height          =   315
            Left            =   1800
            TabIndex        =   31
            Top             =   1200
            Width           =   1095
            _extentx        =   1931
            _extenty        =   556
            maxvalue        =   2048
            minvalue        =   0
            enabled         =   -1  'True
         End
         Begin EditorTDS.UpDownText sy 
            Height          =   315
            Left            =   1800
            TabIndex        =   30
            Top             =   840
            Width           =   1095
            _extentx        =   1931
            _extenty        =   556
            maxvalue        =   2048
            minvalue        =   0
            enabled         =   -1  'True
         End
         Begin EditorTDS.UpDownText sx 
            Height          =   315
            Left            =   1800
            TabIndex        =   29
            Top             =   480
            Width           =   1095
            _extentx        =   1931
            _extenty        =   556
            maxvalue        =   2048
            minvalue        =   0
            enabled         =   -1  'True
         End
         Begin VB.TextBox txtColorAdd 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   27
            Text            =   "-"
            Top             =   2350
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   24
            Text            =   "-"
            Top             =   1980
            Width           =   2655
         End
         Begin VB.Label lblTieneSombra 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SI/NO tiene sombra"
            Height          =   195
            Left            =   3120
            TabIndex        =   71
            Top             =   1200
            Width           =   1410
         End
         Begin VB.Label lblComplementos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Complementos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4080
            TabIndex        =   59
            Top             =   2040
            Width           =   1230
         End
         Begin VB.Label lblEfectoPisada 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Efecto al pisar:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3120
            TabIndex        =   57
            Top             =   480
            Width           =   1305
         End
         Begin VB.Label lblComplementoNormales 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3. Normales:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   34
            ToolTipText     =   "Mapeo normal. Incidencia de la luz en el gráfico."
            Top             =   2805
            Width           =   1080
         End
         Begin VB.Label lblColorAdd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2. ColorAdd:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   26
            ToolTipText     =   "Cada pixel se suma a la capa de abajo"
            Top             =   2445
            Width           =   1065
         End
         Begin VB.Label lblComplementoBlendOne 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1. BlendOne:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   25
            ToolTipText     =   "Cada pixel multiplica al de la capa de abajo."
            Top             =   2040
            Width           =   1125
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   5400
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pixeles alto:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   0
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   21
            ToolTipText     =   "Alto del gráfico medido en pixels"
            Top             =   1600
            Width           =   1020
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pixeles de Ancho:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   0
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   20
            ToolTipText     =   "Ancho del gráfico medido en pixels"
            Top             =   1240
            Width           =   1470
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Posición en Y:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   0
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   19
            ToolTipText     =   "Pixel Y donde comienza el gráfico dentro de la imágen."
            Top             =   900
            Width           =   1140
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Posición en X:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   0
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   18
            ToolTipText     =   "Pixel X donde comienza el gráfico dentro de la imágen."
            Top             =   555
            Width           =   1140
         End
         Begin VB.Label lblNumeroArchivo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Imágen donde está:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   0
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   17
            ToolTipText     =   "Número de la imágen agregada previamente donde se encuentra el gráfico"
            Top             =   150
            Width           =   1695
         End
      End
      Begin VB.Frame frmIndexAnimacion 
         BorderStyle     =   0  'None
         Height          =   1900
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   4335
         Begin EditorTDS.UpDownText Speed 
            Height          =   315
            Left            =   2160
            TabIndex        =   33
            Top             =   0
            Width           =   1095
            _extentx        =   1931
            _extenty        =   556
            maxvalue        =   10000
            minvalue        =   0
            enabled         =   -1  'True
         End
         Begin VB.TextBox Frames 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   1485
            Left            =   2160
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            ToolTipText     =   "Escribir los números de gráficos separados por un ""Enter"""
            Top             =   360
            Width           =   1095
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   4320
            Y1              =   2100
            Y2              =   2100
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tiempo TOTAL que dura:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   0
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   14
            ToolTipText     =   $"frmIndexacion.frx":2516
            Top             =   0
            Width           =   2055
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Numeros de graficos que integran la animacion"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   585
            Left            =   0
            TabIndex        =   13
            Top             =   360
            Width           =   1920
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblMilisegundos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "milisegundos"
            Height          =   195
            Left            =   3350
            TabIndex        =   12
            Top             =   30
            Width           =   900
         End
      End
      Begin VB.Label lblAdvertencia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "El tiempo de duración total en aquellas animaciones correspondientes a personajes o criaturas no es tenido en cuenta."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   795
         Left            =   360
         TabIndex        =   61
         Top             =   2280
         Width           =   5085
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame frmCapasDondeInsertable 
      Caption         =   "¿Donde se puede insertar?"
      Height          =   960
      Left            =   3720
      TabIndex        =   49
      Top             =   1335
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CheckBox Capa 
         Appearance      =   0  'Flat
         Caption         =   "Capa 5 (adorno para pared)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   62
         Top             =   600
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox Capa 
         Appearance      =   0  'Flat
         Caption         =   "Capa 4 (techo) "
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   53
         Top             =   600
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox Capa 
         Appearance      =   0  'Flat
         Caption         =   "Capa 3 (elemento alto)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3720
         TabIndex        =   52
         Top             =   240
         Value           =   1  'Checked
         Width           =   1900
      End
      Begin VB.CheckBox Capa 
         Appearance      =   0  'Flat
         Caption         =   "Capa 2"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   51
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox Capa 
         Appearance      =   0  'Flat
         Caption         =   "Capa 1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
   End
   Begin VB.Label lblIdentificadorGrafico 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "----------------------------------------------------------------"
      Height          =   195
      Left            =   6000
      TabIndex        =   73
      Top             =   690
      Width           =   3360
   End
   Begin VB.Label lblIdentificador 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Identificador gráfico base:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3720
      TabIndex        =   72
      Top             =   690
      Width           =   2235
   End
   Begin VB.Label lblEncontrados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Encontrados: X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   64
      Top             =   6420
      Width           =   1320
   End
   Begin VB.Label lblSecreto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   7680
      TabIndex        =   46
      Top             =   120
      Width           =   525
   End
   Begin VB.Label lblCantidadGraficos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad Graficos:"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   7560
      Width           =   9255
   End
   Begin VB.Label lblNombreIndex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del gráfico"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblNumeroIndex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3720
      TabIndex        =   7
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "frmConfigurarGraficos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GRH_ANIMACION As Byte = 1
Private Const GRH_SIMPLE As Byte = 0

Private cambiosPendientes As Boolean

' El Grh que estoy editando
Private TmpGrhIndexarNum As Integer
Private huboCambios As Boolean ' Si hice algun cambio esta variable se activa

Private Enum eTipoCorte
    corteHorizontal
    corteVertical
End Enum

Private Type usoRecurso
    Tipo As Integer
    ids As New Collection
End Type

Private Enum eTipoCentrado
    ajustarAGrilla = 1
    centrarEnTile = 2
    personalizado = 0
End Enum
            
Public WithEvents vwEditorGraficos As vw_EditorGraficos
Attribute vwEditorGraficos.VB_VarHelpID = -1


Private Sub Capa_Click(Index As Integer)
    habilitarBotonesAplicarRestablecer
End Sub

Private Sub chkGraficoInsertableMapa_Click()

    If Me.chkGraficoInsertableMapa = vbChecked Then
        Me.frmCapasDondeInsertable.Visible = True
    Else
        Me.frmCapasDondeInsertable.Visible = False
    End If

    habilitarBotonesAplicarRestablecer
End Sub

Private Sub cmdAceptar_Click()
    
    Unload Me
    
    'Si hubo cambios, tengo que recargar la lista de graficos del frmmain. No en cada cambio comun.
    If huboCambios Then CargarListaGraficosComunes

End Sub

Private Sub eliminar(ByVal idgrh As Long)
    Dim borrarTodos As Boolean
    Dim confirma As VbMsgBoxResult
    Dim loopFrame As Integer
    
    If GrhData(idgrh).NumFrames > 1 Then
        confirma = MsgBox("El elemento seleccionado es una animación. ¿Desea también eliminar los elementos que componen la animación?. Si elige que NO los graficos que lo componen pasaran a ser gráficos simples.", vbExclamation + vbYesNo, "Configurar Gráficos")
    
        If confirma = vbYes Then
            borrarTodos = True
        Else
            borrarTodos = False
        End If
    End If
    
    'Desactivo al vista previa
    TmpGrhIndexarNum = 0
    
    'Lo quito de la lista
    Call Listado.eliminarElemento(idgrh)
               
    If borrarTodos Then
        For loopFrame = 1 To GrhData(idgrh).NumFrames
        
            Call Me_indexar_Graficos.eliminar(GrhData(idgrh).Frames(loopFrame))
            
            'Lo quito de la lista
            Call Listado.eliminarElemento(GrhData(idgrh).Frames(loopFrame))
            
        Next loopFrame
    ElseIf GrhData(idgrh).NumFrames > 1 Then
        For loopFrame = 1 To GrhData(idgrh).NumFrames
            'Lo agrego a la lista como
            Call Listado.addString(GrhData(idgrh).Frames(loopFrame), GrhData(idgrh).Frames(loopFrame) & " - " & GrhData(GrhData(idgrh).Frames(loopFrame)).nombreGrafico, 0)
            'Marco que no pertenece a una animación
            GrhData(GrhData(idgrh).Frames(loopFrame)).perteneceAunaAnimacion = False
            ' Actualizo la info adicional
            Call Me_indexar_Graficos.actualizarEnIni(GrhData(idgrh).Frames(loopFrame))
        Next loopFrame
    End If
    
    'Borro el GRH Principal
    Call Me_indexar_Graficos.eliminar(idgrh)
    
End Sub

Private Sub cmdActualizarOffSet_Click()

'    Dim loopGrh As Integer
'
'    For loopGrh = Me.nroInicio.value To Me.nroHasta.value
'
'        If GrhData(loopGrh).centrarEn32 Then
'
'            Call calcularCentradoEnGrilla(GrhData(loopGrh))
'
'            Call Me_indexar_Graficos.actualizarEnIni(loopGrh)
'
'        End If
'
'        DoEvents
'        Me.cmdAplicarNombresAHijos.caption = loopGrh & "/" & grhCount
'
'    Next

End Sub

Private Sub cmdAplicarNombresAHijos_Click()
Dim loopGrh As Long
Dim loopFrame As Long

For loopGrh = Me.nroInicio.value To Me.nroHasta.value

    If Me_indexar_Graficos.existe(loopGrh) Then
    
        If GrhData(loopGrh).NumFrames > 1 Then
        
            If Len(GrhData(loopGrh).nombreGrafico) > 0 Then
            
                'Recorremos cada frame
                For loopFrame = 1 To GrhData(loopGrh).NumFrames
                    ' Si ya no tiene el nombre..
                    If Not GrhData(GrhData(loopGrh).Frames(loopFrame)).nombreGrafico = "(A) " & GrhData(loopGrh).nombreGrafico & " " & loopFrame Then
                        'Marco que no pertenece a una animación
                        GrhData(GrhData(loopGrh).Frames(loopFrame)).nombreGrafico = "(A) " & GrhData(loopGrh).nombreGrafico & " " & loopFrame
                        ' Actualizo la info adicional
                        Call Me_indexar_Graficos.actualizarEnIni(GrhData(loopGrh).Frames(loopFrame))
                    End If
                Next loopFrame
            End If
        End If
    End If
    
    DoEvents
    
    Me.cmdAplicarNombresAHijos.caption = loopGrh & "/" & grhCount
Next

End Sub

Private Sub cmdBorrarMultiple_Click()
    Dim vbresult As VbMsgBoxResult
    Dim loopGrh As Integer
    
    vbresult = MsgBox("¿Seguro que queres borrar estos gráficos?", vbExclamation + vbYesNo)
    
    ' Si puso no, salimos
    If vbresult = vbNo Then Exit Sub
    
    Me.cmdBorrarMultiple.Enabled = False
        
    For loopGrh = Me.nroInicio.value To Me.nroHasta.value
        If Me_indexar_Graficos.existe(loopGrh) Then
            Call eliminar(loopGrh)
            
            Me.cmdBorrarMultiple.caption = loopGrh & "/" & Me.nroHasta.value
    
            DoEvents
        End If
    Next
    
    Me.cmdBorrarMultiple.Enabled = True
    
    Me.cmdBorrarMultiple.caption = "Borrar"
    
    MsgBox "Elementos eliminados", vbInformation, Me.caption
End Sub

Private Function consultarRemplazo(ByVal imagen As Integer, texto As String) As Integer

    Dim nombreRecurso As String
    Dim remplazoImagen As Integer

    nombreRecurso = pakGraficos.Cabezal_GetFilenameName(imagen)
    remplazoImagen = CInt(val(InputBox(Replace$(texto, "#{recurso}", imagen & " - " & nombreRecurso), "Remplazar Imágen")))

    consultarRemplazo = remplazoImagen
End Function
Private Function DuplicarSimple(ByVal idgrh As Integer, nombre As String, Optional ByVal remplazo As Integer = 0) As Integer
    Dim nuevoGRH As Integer
    Dim remplazoGrh As Integer
    
    nuevoGRH = Me_indexar_Graficos.nuevo ' Obtengo un nuevo id
        
    If Not nuevoGRH = -1 Then
        GrhData(nuevoGRH) = GrhData(idgrh) 'Copio
        GrhData(nuevoGRH).nombreGrafico = nombre
        GrhData(nuevoGRH).ID = "" ' El ID se setea a nulo ya que no proviene de un Sprite
        
        ' ¿Hay un remplazo por defecto?
        If GrhData(idgrh).filenum > 0 Then
            If remplazo > 0 Then
                remplazoGrh = remplazo
            Else
                remplazoGrh = consultarRemplazo(GrhData(idgrh).filenum, "¿Por qué número de recurso desea remplazar al recurso '#{recurso}'?. Cancelar para dejarlo tal cual.")
            End If
            
            '¿Seleccionó un remplazo?
            If remplazoGrh > 0 Then
                GrhData(nuevoGRH).filenum = remplazoGrh
                GrhData(nuevoGRH).Frames(1) = nuevoGRH
            End If
        End If
        
        DuplicarSimple = nuevoGRH
    Else
        DuplicarSimple = -1 'Error...
        Exit Function
    End If

End Function

Private Function duplicar(ByVal idgrh As Integer, nombre As String, Optional ByVal remplazoImagen As Integer = -1) As Integer
    Dim backupFrames() As Integer
    Dim loopFrame As Integer
    Dim remplazoGrh As Integer
    Dim numeroImagen As Integer
    
    If GrhData(idgrh).NumFrames > 1 Then
    
        
        'Es una animacion. Tengo que duplicar cada frame
        
        ' Antes de duplicar cada frame me voy a fijar si todos están relacionados
        ' a la misma imagen. Si esto es así, voy a preguntarle al usuario si
        ' desea remplazar en todos los frames la imagen por otra.
        ' Así evitamos que se haga una vez por cada frame en el duplicar Simple.
        If Not remplazoImagen = -1 Then
            numeroImagen = GrhData(GrhData(idgrh).Frames(1)).filenum
             
            For loopFrame = 2 To GrhData(idgrh).NumFrames
                If Not numeroImagen = GrhData(GrhData(idgrh).Frames(loopFrame)).filenum Then
                    Exit For
                    numeroImagen = -1
                End If
            Next loopFrame
            
            '¿Todos los componentes tienen la misma imagne?. Pregunto si quiere hacer un remplazo rápido
            If Not numeroImagen = -1 Then
                remplazoImagen = consultarRemplazo(numeroImagen, "Todos los componentes de esta animación hacen referencia a la imágen '#{recurso}'. Si queres remplazar está imagen por otra en todos los componentes, ingresa el número de la nueva imagen. Si no queres remplazar automaticamente todos los componentes, pulsa 'Cancelar'.")
            End If
        End If
        
        ' Copio los frames
        ReDim backupFrames(1 To GrhData(idgrh).NumFrames) As Integer
        
        For loopFrame = 1 To GrhData(idgrh).NumFrames
            backupFrames(loopFrame) = DuplicarSimple(GrhData(idgrh).Frames(loopFrame), "(A) " & nombre & " " & loopFrame, remplazoImagen)
            
            If backupFrames(loopFrame) = -1 Then
                duplicar = -1 'Error no se pudo obtener. TODO. Deberiamos liberar los que ya estaban tomados
                Exit Function
            End If
        Next loopFrame

        ' Duplico el original y le remplazo los frames
        remplazoGrh = DuplicarSimple(idgrh, nombre)
        
        If Not remplazoGrh = -1 Then '¿Pude obtener un indice?
        
            'Remplazo los frames y guardo
            For loopFrame = 1 To GrhData(remplazoGrh).NumFrames
                GrhData(remplazoGrh).Frames(loopFrame) = backupFrames(loopFrame)
                Call Me_indexar_Graficos.actualizarEnIni(backupFrames(loopFrame))
            Next
    
            Call Me_indexar_Graficos.actualizarEnIni(remplazoGrh)
        
            duplicar = remplazoGrh
        Else
            'TODO. Deberiamos liberar los que ya estaban tomados
            duplicar = -1
        End If
    Else
        'Es un grafico comun, tengo que duplicar solo este grh
        remplazoGrh = DuplicarSimple(idgrh, nombre, remplazoImagen)
        
        If Not remplazoGrh = -1 Then
            Call Me_indexar_Graficos.actualizarEnIni(remplazoGrh)
        End If
        
        duplicar = remplazoGrh
    End If
    
End Function

Private Sub cortar(ByVal idElemento As Integer, ByVal pixelsAncho As Integer, ByVal pixelsAlto As Integer, ByVal tipoCorte As eTipoCorte, ByVal crearPredefinido As Boolean, ByVal capaPredefinido As Byte, ByVal aplicarBloqueo As Boolean)

    Dim Graficos() As Integer
    Dim nombreGrafico  As String
    Dim Trigger As Long
    Dim loopGrafico As Integer
    Dim capaDefault As Byte
    
    nombreGrafico = GrhData(idElemento).nombreGrafico
    
    ' ************** Corto realmente ***************** '
    If cortar_(idElemento, pixelsAncho, pixelsAlto, Graficos) Then
    
        ' Actualizo los graficos de la lista y lo inicializo
        For loopGrafico = LBound(Graficos) To UBound(Graficos)
        
            If Listado.existe(Graficos(loopGrafico)) Then
                Call Listado.cambiarNombre(Graficos(loopGrafico), Graficos(loopGrafico) & " - " & GrhData(Graficos(loopGrafico)).nombreGrafico)
            Else
                Call Listado.addString(Graficos(loopGrafico), Graficos(loopGrafico) & " - " & GrhData(Graficos(loopGrafico)).nombreGrafico, 0)
            End If
            
            ' Inicializo
            Init_grh_tutv Graficos(loopGrafico)
            
        Next loopGrafico
        
        ' Actualizo la cantidad de graficos
        actualizarCantidadgraficos
        
         'Actualizo la lista
        CargarListaGraficosComunes
        
         ' Listo
        MsgBox "Se partió el grafico '" & nombreGrafico & "' en " & (UBound(Graficos) - LBound(Graficos) + 1) & " partes.", vbInformation, Me.caption
    End If
    
    
    '  ************ Generar auto Preset *********** '
    If crearPredefinido Then
        Trigger = 0
    
        capaDefault = capaPredefinido
    
        ' ¿Quiere aplicar bloqueo
        If aplicarBloqueo Then Trigger = eTriggers.TodosBordesBloqueados
        
        ' Generamos
        If generarPresetDesdeGraficos(tipoCorte, Graficos, capaDefault, nombreGrafico, Trigger) Then
            ' Recargamos la lista y avisamos
            Call ME_presets.cargarListaPresets
            Call MsgBox("Se creó con éxito el predefinido '" & nombreGrafico & "'.", vbInformation, Me.caption)
        Else
            ' Avisamos
            Call MsgBox("Se produjo un error al crear el predefinido. Por favor, intente más tarde o contacte a un Administrador del Sistema.", vbExclamation, Me.caption)
        End If
    
    End If
    
End Sub
Private Function cortar_(ByVal idElemento As Integer, ByVal pixelsAncho As Integer, ByVal pixelsAlto As Integer, Graficos() As Integer) As Boolean

    Dim cantidadTilesAncho As Byte
    Dim cantidadTilesAlto As Byte
    Dim cantidadGraficos As Byte
    Dim loopAlto As Byte
    Dim loopAncho As Byte
    Dim loopGrafico As Byte
    Dim error As Boolean
    Dim nombreGrafico As String
    
    Dim baseX As Integer
    Dim baseY As Integer
    
    '¿En cuantas partes horizontales y verticales va a ser cortado?
    cantidadTilesAncho = GrhData(idElemento).pixelWidth / pixelsAncho
    cantidadTilesAlto = GrhData(idElemento).pixelHeight / pixelsAlto

    cantidadGraficos = (cantidadTilesAncho * cantidadTilesAlto)
    
     ' Creo los graficos necesarios
    ReDim backupFrames(1 To cantidadGraficos) As Integer
    
    ' Obtengo los nuevos slots para los graficos (la cantidad de graficos menos uno que es el que ya tengo)
    error = False
    For loopGrafico = 2 To cantidadGraficos
         backupFrames(loopGrafico) = Me_indexar_Graficos.nuevo
        ' ¿Lo pude obtener?
        If backupFrames(loopGrafico) = -1 Then
            error = True
            Exit For
        End If
    Next
    
    If error Then
        'Si se produjo un error. Elimino los graficos que cree al pedo
        For loopGrafico = 2 To cantidadGraficos
           ' ¿Lo pude obtener?
            If backupFrames(loopGrafico) > 0 Then Me_indexar_Graficos.eliminar (backupFrames(loopGrafico))
        Next
        
        MsgBox "No se pudieron crear la cantidad de gráficos necesarios. Por favor intente más tarde o avise a un administrador.", vbExclamation, Me.caption
        cortar_ = False
        Exit Function
    End If
    
    ' El 1 es el original
    backupFrames(1) = idElemento
    
    ' Ya tengo todos los frames
    nombreGrafico = GrhData(idElemento).nombreGrafico
    
    ' Esto va a ser para todos igual, además de las otras propiedades que ya tiene el grafico
    GrhData(idElemento).pixelWidth = pixelsAncho
    GrhData(idElemento).pixelHeight = pixelsAlto

    'Call Me_indexar_Graficos.calcularPropiedadesVariables(GrhData(idElemento), offsetNeto)
    
    ' Seteo el resto de los graficos. Lo que cambia es el nombre y el offset en X
    baseX = GrhData(idElemento).sx
    baseY = GrhData(idElemento).sy
    
    
    loopGrafico = 1
    For loopAlto = 1 To cantidadTilesAlto
        For loopAncho = 1 To cantidadTilesAncho
        
            GrhData(backupFrames(loopGrafico)) = GrhData(idElemento) ' Copio
            
            With GrhData(backupFrames(loopGrafico))
                ' El nombre es una letra para identificar la fila y un número para a columna
                .nombreGrafico = nombreGrafico & " " & Chr$(64 + loopAlto) & " " & loopAncho
                
                ' El offset
                .sy = baseY + pixelsAlto * (loopAlto - 1)
                .sx = baseX + pixelsAncho * (loopAncho - 1)
            End With
            
            'Call Me_indexar_Graficos.calcularPropiedadesVariables(GrhData(backupFrames(loopGrafico),))
             
            loopGrafico = loopGrafico + 1
        Next loopAncho
    Next loopAlto
    
    ' La lista de ids de graficos que genere
    ReDim Graficos(1 To cantidadGraficos) As Integer
    
    ' Guardo los graficos
    For loopGrafico = 1 To cantidadGraficos
    
         ' Guardo los graficos
        Call Me_indexar_Graficos.actualizarEnIni(backupFrames(loopGrafico))
        
        ' Guardamos el ID en el vector de devolucion
        Graficos(loopGrafico) = backupFrames(loopGrafico)
    Next
    
    cortar_ = True

End Function

' FUNCION AUXILIAR QUE CALCULA POR UNICA VEZ EL OFFSET EN X E Y
Private Sub calcularCentradoEnGrilla(Grh As GrhData)
      
Dim offsetParaAjuste As Position

Call Me_indexar_Graficos.obtenerOffsetAjustadoTile(Grh, offsetParaAjuste.x, offsetParaAjuste.y)

Grh.offsetX = offsetParaAjuste.x
Grh.offsetY = offsetParaAjuste.y
   
End Sub

Private Sub mostrarPanelConfiguracionSombra()
    ' Cargamos la vWindows
   ' GUI_Load vwEditorGraficos
    
    ' Seteamos
   ' vwEditorGraficos.SetGrafico TmpGrhIndexarNum
End Sub

Private Sub cmdCortar32H_Click()
    Dim cantidadTilesAlto As Byte
    Dim nombreGrafico As String
    Dim idElemento As Integer
    Dim resultado As VbMsgBoxResult
    Dim loopGrafico As Integer
    
    Dim ids() As Long
    Dim error As Boolean
    
    Dim crearPredefinido As Boolean
    Dim aplicarBloqueo As Boolean
    Dim capaPredefinido As Byte
    
    ids = obtenerIDS()

    If Not arrayEstaIniciado(ids) Then
        Call MsgBox("Tenes que seleccionar uno o más elementos a duplicar", vbExclamation, Me.caption)
        Exit Sub
    End If
    
    ids = obtenerIDS
    
    ' Chequeamos
    For loopGrafico = 0 To UBound(ids)
    
        idElemento = CInt(ids(loopGrafico))
    
        ' ¿Es un grafico simple? Solo se pueden partir graficos simples
        If GrhData(idElemento).NumFrames > 1 Then
            MsgBox "No se pueden cortar animaciones.", vbExclamation, Me.caption
            error = True
            Exit For
        End If
    
        cantidadTilesAlto = CByte(HelperMath.redondearHaciaArriba(GrhData(idElemento).pixelHeight / 32))
        nombreGrafico = GrhData(idElemento).nombreGrafico
    
        ' ¿Tiene sentido cortar?
        If cantidadTilesAlto = 1 Then
            MsgBox "No se pueden cortar gráficos que ocupen un solo tile.", vbExclamation, Me.caption
            Exit Sub
        End If
        
        ' Confirmamos...
        resultado = MsgBox("¿Estás seguro que queres cortar de forma horizontal el gráfico '" & nombreGrafico & "'?. Se van a crear " & cantidadTilesAlto & " gráficos.", vbExclamation + vbYesNo)
        If resultado = vbNo Then
            error = True
            Exit For
        End If
        
    Next
    
    If error Then
        Call MsgBox("El corte fue cancelado, no se aplicaron los cambios.", vbExclamation, Me.caption)
        Exit Sub
    End If
            
    ' Creo predefinido?
    resultado = MsgBox("¿Querés generar un predefinido que contenga la unión de todos los gráficos generados por CADA corte?. ¡ATENCION! Asegurate de estar en un mapa en blanco. No hagas esto en un mapa del mundo porque se va a dañar una parte de el.", vbQuestion + vbYesNo)
      
    crearPredefinido = (resultado = vbYes)
    aplicarBloqueo = False
    capaPredefinido = 0
    
    If crearPredefinido Then
        
        ' ¿Aplico bloqueo?
        resultado = MsgBox("¿Queres aplicar un bloqueo parcial?. Solo se aplicará en los tiles donde se insertan los gráficos y no en toda el área visual que este ocupa.", vbQuestion + vbYesNo)
        aplicarBloqueo = (resultado = vbYes)
        
        capaPredefinido = CByte(val(InputBox("¿Cuál es la Capa predefinida en donde queres que se pongan estos gráficos al momento de crear el Preset?. Un número entre 1 y " & CANTIDAD_CAPAS & ".", Me.caption)))

        If capaPredefinido <= 0 And capaPredefinido > CANTIDAD_CAPAS Then
            Call MsgBox("El número de capa que pusiste es incorrecto. Se cancela la creación del preset.", vbExclamation)
            Exit Sub
        End If
        
    End If
    
    ' Cortamos!
    For loopGrafico = 0 To UBound(ids)
        idElemento = ids(loopGrafico)
        Call cortar(idElemento, GrhData(idElemento).pixelWidth, 32, corteHorizontal, crearPredefinido, capaPredefinido, aplicarBloqueo)
    Next

    ' Avisamos que termino
    Call MsgBox("Cortar terminado", vbInformation, Me.caption)
    
End Sub

Private Sub cmdCortar32V_Click()
    Dim cantidadTilesAncho As Byte
    Dim nombreGrafico As String
    Dim idElemento As Integer
    Dim resultado As VbMsgBoxResult
    Dim loopGrafico As Integer
    Dim ids() As Long
    Dim error As Boolean
    
    ids = obtenerIDS()

    If Not arrayEstaIniciado(ids) Then
        Call MsgBox("Tenes que seleccionar uno o más elementos a duplicar", vbExclamation, Me.caption)
        Exit Sub
    End If
    
    ids = obtenerIDS
    
    ' Chequeamos
    For loopGrafico = 0 To UBound(ids)
    
        idElemento = CInt(ids(loopGrafico))
    
        ' ¿Es un grafico simple? Solo se pueden partir graficos simples
        If GrhData(idElemento).NumFrames > 1 Then
            MsgBox "No se pueden cortar animaciones.", vbExclamation, Me.caption
            error = True
            Exit For
        End If
    
        cantidadTilesAncho = GrhData(idElemento).pixelWidth \ 32
        nombreGrafico = GrhData(idElemento).nombreGrafico
    
        ' ¿Tiene sentido cortar?
        If cantidadTilesAncho = 1 Then
            MsgBox "No se pueden cortar gráficos que ocupen un solo tile.", vbExclamation, Me.caption
            error = True
            Exit For
        End If
        
        ' Confirmamos...
        resultado = MsgBox("¿Estás seguro que queres cortar de forma vertical el gráfico '" & nombreGrafico & "'?. Se van a crear " & cantidadTilesAncho & " gráficos.", vbExclamation + vbYesNo)
        If resultado = vbNo Then
            error = True
            Exit For
        End If
        
    Next
    
    If error Then
        Call MsgBox("El corte fue cancelado, no se aplicaron los cambios.", vbExclamation, Me.caption)
        Exit Sub
    End If
    
    
    Dim crearPredefinido As Boolean
    Dim aplicarBloqueo As Boolean
    Dim capaPredefinido As Byte
        
    ' Creo predefinido?
    resultado = MsgBox("¿Querés generar un predefinido que contenga la unión de todos los gráficos generados por CADA corte?. ¡ATENCION! Asegurate de estar en un mapa en blanco. No hagas esto en un mapa del mundo porque se va a dañar una parte de el.", vbQuestion + vbYesNo)
      
    crearPredefinido = (resultado = vbYes)
    aplicarBloqueo = False
    capaPredefinido = 0
    
    If crearPredefinido Then
        
        ' ¿Aplico bloqueo?
        resultado = MsgBox("¿Queres aplicar un bloqueo parcial?. Solo se aplicará en los tiles donde se insertan los gráficos y no en toda el área visual que este ocupa.", vbQuestion + vbYesNo)
        aplicarBloqueo = (resultado = vbYes)
        
        capaPredefinido = CByte(val(InputBox("¿Cuál es la Capa predefinida en donde queres que se pongan estos gráficos al momento de crear el Preset?. Un número entre 1 y " & CANTIDAD_CAPAS & ".", Me.caption)))

        If capaPredefinido <= 0 And capaPredefinido > CANTIDAD_CAPAS Then
            Call MsgBox("El número de capa que pusiste es incorrecto. Se cancela la creación del preset.", vbExclamation)
            Exit Sub
        End If
        
    End If
    
    ' Cortamos!
    For loopGrafico = 0 To UBound(ids)
        idElemento = ids(loopGrafico)
        Call cortar(idElemento, 32, GrhData(idElemento).pixelHeight, corteVertical, crearPredefinido, capaPredefinido, aplicarBloqueo)
    Next

    ' Avisamos que termino
    Call MsgBox("Cortar terminado", vbInformation, Me.caption)
End Sub

Private Function generarPresetDesdeGraficos(union As eTipoCorte, Graficos() As Integer, Capa As Byte, nombre As String, Trigger As Long) As Boolean
    Dim area As tAreaSeleccionada
    Dim tipoElementos As Tools
    Dim tilesHorizontal As Byte
    Dim tilesVertical As Byte
    
    Dim loopX As Integer
    Dim loopY As Integer
    Dim loopGrafico As Integer
    
    
    ' Vamos a contar la cantidad de tiles que ocupa horizontal o verticalmente
    If union = corteVertical Then   'Si es corte vertical....
        tilesHorizontal = redondearHaciaArriba(GrhData(Graficos(1)).pixelWidth / 32) * (UBound(Graficos))
        tilesVertical = 1
    Else 'Si es corte horizontal...
        tilesHorizontal = 1
        tilesVertical = redondearHaciaArriba(GrhData(Graficos(1)).pixelHeight / 32) * (UBound(Graficos))
    End If
    
    ' Establecemos el area que vamos a utilizar
    area.abajo = SV_Constantes.Y_MAXIMO_USABLE
    area.arriba = SV_Constantes.Y_MAXIMO_USABLE - tilesVertical + 1
    area.derecha = SV_Constantes.X_MAXIMO_USABLE
    area.izquierda = SV_Constantes.X_MAXIMO_USABLE - tilesHorizontal + 1
    area.invertidoHorizontal = False
    area.invertidoHorizontal = False
    
    ' Los elementos del preset
    tipoElementos = tool_grh
    
    ' Quiere ponerle algún trigger?
    If Trigger > 0 Then tipoElementos = tipoElementos Or tool_triggers

    ' Borro la parte que voy a utilizar para asegurarme que no interfiera otra cosa.
    Call Me_Tools_Seleccion.eliminar(area, tipoElementos)
        
    ' Agregamos los gráficos
    loopY = area.arriba
    loopGrafico = 1
    
    Do While loopY <= SV_Constantes.Y_MAXIMO_USABLE
        
        loopX = area.izquierda

        Do While loopX <= SV_Constantes.X_MAXIMO_USABLE
    
            ' Grafico
            mapdata(loopX, loopY).Graphic(Capa).GrhIndex = Graficos(loopGrafico)
            InitGrh mapdata(loopX, loopY).Graphic(Capa), Graficos(loopGrafico)
            
            ' Trigger
            mapdata(loopX, loopY).Trigger = Trigger
            
            loopX = loopX + 1
            loopGrafico = loopGrafico + 1
        Loop
        
        loopY = loopY + 1
    Loop

    ' Creamos
    generarPresetDesdeGraficos = Me_Tools_Seleccion.crearPresetDesdeMapa(area, nombre, tipoElementos)
    
    ' Borramos nuevamente
    Call Me_Tools_Seleccion.eliminar(area, tipoElementos)
End Function

Private Function obtenerIDS() As Long()
    Dim ids()  As Long
    
    '¿Seleccion multiple o simple?
    If Me.Listado.checked And Me.Listado.getCantidadSeleccionada > 0 Then
        ids = Me.Listado.obtenerIDSeleccionados
    ElseIf Me.Listado.obtenerIDValor > 0 Then
        ReDim ids(0)
        ids(0) = Me.Listado.obtenerIDValor
    Else
        Exit Function
    End If
    
    
    obtenerIDS = ids
End Function

Private Sub cmdDuplicar_Click()

    Dim ids()  As Long
    Dim loopElemento As Long
    Dim numeroImagen As Integer
    Dim remplazoImagen As Integer
    Dim cantidadseleccionada As Integer
    Dim buscar As String
    Dim remplazar As String
    Dim Grh As Integer
    Dim nuevoElemento As Integer
    Dim mensaje As String
        
    Dim respuesta As VbMsgBoxResult
      
    ids = obtenerIDS()

    If Not arrayEstaIniciado(ids) Then
        Call MsgBox("Tenes que seleccionar uno o más elementos a duplicar", vbExclamation, Me.caption)
        Exit Sub
    End If
    
    cantidadseleccionada = UBound(ids) + 1
    
    ' Confirmo
    If cantidadseleccionada = 1 Then
        respuesta = MsgBox("¿Estás seguro que queres duplicar el elemento '" & GrhData(ids(0)).nombreGrafico & "'?.", vbInformation + vbYesNo, Me.caption)
    Else
        respuesta = MsgBox("¿Estás seguro que queres duplicar los " & cantidadseleccionada & " elementos seleccionados?.", vbInformation + vbYesNo, Me.caption)
    End If
    
    If respuesta = vbNo Then Exit Sub
    
    ' Evito accion repetitiva
    If cantidadseleccionada > 1 Then
        ' Recorro todos los elementos seleccionados
        ' Me fijo si todos tienen la misma imagen
        ' Si la tienen le pregunto por cual la quiere remplazar
        ' Guardo el numero
        numeroImagen = GrhData(ids(LBound(ids))).filenum
        
        For loopElemento = LBound(ids) + 1 To UBound(ids)
            If Not numeroImagen = GrhData(ids(loopElemento)).filenum Then
                numeroImagen = -1
                Exit For
            End If
        Next
        
        If Not numeroImagen = -1 Then
             remplazoImagen = consultarRemplazo(numeroImagen, "Todos los elementos seleccionados hacen referencia a la imágen '#{recurso}'. Si queres remplazar está imagen por otra en todos los elementos duplicados, ingresá el número de la nueva imagen. Si no queres remplazar automaticamente todos los componentes e indicar la imagen de cada uno, pulsa 'Cancelar'.")
        End If
    End If

    ' Nuevo nombre
    If cantidadseleccionada > 1 Then
        buscar = InputBox("Desea remplazar automaticamente una parte del texto actual por otra?. Si es así, escriba el texto que desea remplazar:")
        
        If Not buscar = "" Then
            remplazar = InputBox("Ingresá el texto que lo remplazará:")
        Else
            remplazar = InputBox("Ingresá el nombre que tendrán todos los elementos duplicados:")
        End If
    Else
        buscar = ""
        remplazar = InputBox("¿Qué nombre le queres poner al gráfico que se va a crear en base al gráfico " & GrhData(ids(0)).nombreGrafico & "?.", Me.caption)
    End If
        
    ' Confirmo nuevamente
    If cantidadseleccionada > 1 Then
        mensaje = "Se van a duplicar " & cantidadseleccionada & " elementos."
    Else
        mensaje = "Se va a duplicar '" & GrhData(ids(0)).nombreGrafico & "'."
    End If
    
    If buscar = "" Then
        mensaje = mensaje & " Se pondrá el nombre " & remplazar
        If cantidadseleccionada > 1 Then
            mensaje = mensaje & " a todos los elementos."
        Else
            mensaje = mensaje & " al elemento nuevo."
        End If
    Else
        mensaje = mensaje & " Cada duplicado tendrá el nombre del elemento original pero se le cambiará donde dice '" & buscar & "' por '" & remplazar & "'. " & vbNewLine & vbNewLine & "Por ejemplo, el duplicado de  '" & GrhData(ids(0)).nombreGrafico & "' se llamará '" & Replace$(GrhData(ids(0)).nombreGrafico, buscar, remplazar) & "'."
    End If
    
    respuesta = MsgBox(mensaje & vbNewLine & vbNewLine & "¿Estás seguro?", vbInformation + vbYesNo, "Duplicar")
    
    If respuesta = vbNo Then Exit Sub
    
    ' Duplico
    For loopElemento = LBound(ids) To UBound(ids)
        
        Grh = ids(loopElemento)
            
        ' Duplico cada elemento, el duplicar recibe la imagen por defecto
        If Len(buscar) = 0 Then
            nuevoElemento = duplicar(Grh, remplazar, remplazoImagen)
        Else
            nuevoElemento = duplicar(Grh, Replace$(GrhData(Grh).nombreGrafico, buscar, remplazar, vbTextCompare), remplazoImagen)
        End If
        
        ' ¿Fallo?
        If nuevoElemento = -1 Then Exit For
        
        ' Agregamos el elemento nuevo a la lista
        Call agregarGrhALista(nuevoElemento)
        
        ' Lo inicializamos
        Init_grh_tutv nuevoElemento
    Next loopElemento
                    
    ' Aviso
    If nuevoElemento = -1 Then
        If cantidadseleccionada = 1 Then
            MsgBox "El elemento no pudo ser duplicado. Intenta nuevamente o contactá con el Administrador del Sistema.", vbExclamation, Me.caption
        Else
            MsgBox "Uno o más elementos no pudieron ser duplicados. Intenta nuevamente o contactá con el Administrador del Sistema.", vbExclamation, Me.caption
        End If
        Exit Sub
    End If
    
    If cantidadseleccionada > 1 Then
        MsgBox "Todos los elementos fueron duplicados exitosamente.", vbInformation, Me.caption
    Else
        MsgBox "Se creó el grafico " & nuevoElemento & " en base al " & GrhData(ids(0)).nombreGrafico & ".", vbInformation, Me.caption
    End If
    
    ' Actualizamos la cantidad de graficos
    Call actualizarCantidadgraficos
        
    'Actualizo la lista de graficos insertables
    Call CargarListaGraficosComunes
    
    ' Seleccionamos el ultimo elemento
    If nuevoElemento <> -1 Then Call Me.Listado.seleccionarElemento(nuevoElemento)
    
End Sub

Private Sub cmdEliminar_Graficos_Click()
    Dim confirma As VbMsgBoxResult
    Dim idgrh As Long
    
    If Not Me.Listado.obtenerValor() = "" Then
    
        confirma = MsgBox("¿Está seguro de que desea eliminar el elemento gráfico '" & Me.Listado.obtenerValor & "'?", vbYesNo + vbExclamation, "Configurar Gráficos")
    
        If confirma = vbYes Then
            
            'Codigo de eliminación de Grh
            idgrh = Me.Listado.obtenerIDValor
            
            Call eliminar(idgrh)
            
            'Desactivo el boton
            Me.cmdEliminar_Graficos.Enabled = False

            Call ejecutarControlCambios
            
            actualizarCantidadgraficos
        End If
    End If
End Sub

Private Sub cmdGenerarDelete_Click()
    Dim strDelete As String
    Dim archivo As Integer
    Dim loopGrh As Integer
   
    strDelete = "UPDATE recursos SET ESTADO='CONFIRMADO' WHERE tipo='RECURSO_IMAGEN' AND ID IN("
    
    For loopGrh = Me.nroInicio.value To Me.nroHasta.value
    
        If Me_indexar_Graficos.existe(loopGrh) Then
            strDelete = strDelete & loopGrh & ","
            
            DoEvents
            Me.cmdAplicarNombresAHijos.caption = loopGrh & "/" & grhCount
        End If
        
    Next
    
    strDelete = mid$(strDelete, 1, Len(strDelete) - 1) & ")"
 
    Debug.Print strDelete
    archivo = FreeFile
    
    Open "C:\salida_graficos.txt" For Output As #archivo
        Print #archivo, strDelete
    Close #archivo

  

End Sub

Private Sub cmdGraficosSinUitlizar_Click()

    'Creamos un array donde vamos a marcar si esta usando o no
    Dim graficosEstado() As usoRecurso
    ReDim graficosEstado(1 To pakGraficos.getCantidadElementos) As usoRecurso
    
    Dim loopFrame As Integer
    Dim loopGrh As Integer
    
    For loopGrh = LBound(graficosEstado) To UBound(graficosEstado)
        graficosEstado(loopGrh).Tipo = 0
    Next
    
    ' ¿Es un grafico?
    For loopGrh = LBound(GrhData) To UBound(GrhData)
            If Me_indexar_Graficos.existe(loopGrh) Then
                If GrhData(loopGrh).NumFrames = 1 Then
                    If GrhData(loopGrh).filenum > 0 Then
                        graficosEstado(GrhData(loopGrh).filenum).Tipo = 1
                        graficosEstado(GrhData(loopGrh).filenum).ids.Add loopGrh
                    End If
                ElseIf GrhData(loopGrh).NumFrames > 1 Then
                    For loopFrame = 1 To GrhData(loopGrh).NumFrames
                        graficosEstado(GrhData(GrhData(loopGrh).Frames(loopFrame)).filenum).Tipo = 1
                        graficosEstado(GrhData(GrhData(loopGrh).Frames(loopFrame)).filenum).ids.Add GrhData(loopGrh).Frames(loopFrame)
                    Next
                End If
                
                
                
            End If
    Next
    
    ' ¿Es un piso?
    For loopGrh = LBound(Tilesets) To UBound(Tilesets)
        If Me_indexar_Pisos.existe(loopGrh) Then
           
           
           If Tilesets(loopGrh).filenum > 0 Then
                graficosEstado(Tilesets(loopGrh).filenum).Tipo = 2
                graficosEstado(Tilesets(loopGrh).filenum).ids.Add loopGrh
           End If

            If Tilesets(loopGrh).Olitas > 0 Then
                graficosEstado(Tilesets(loopGrh).Olitas).Tipo = 3
                graficosEstado(Tilesets(loopGrh).Olitas).ids.Add loopGrh
            End If

            If Tilesets(loopGrh).stage_count > 1 Then
                For loopFrame = 1 To UBound(Tilesets(loopGrh).stages)
                    graficosEstado(Tilesets(loopGrh).stages(loopFrame)).Tipo = 2
                    graficosEstado(Tilesets(loopGrh).stages(loopFrame)).ids.Add loopGrh
                Next
            End If
            
            
            
        End If
    Next
    
    ' Estadisticas
    Dim cantidad As Integer
    Dim cantidadNoUsados As Integer
    
    cantidad = 0
    cantidadNoUsados = 0

    For loopFrame = 1 To UBound(graficosEstado)
        If graficosEstado(loopFrame).Tipo = 0 Then
            cantidad = cantidad + 1
            
            If pakGraficos.Cabezal_GetFileSize(loopFrame) > 0 Then
                cantidadNoUsados = cantidadNoUsados + 1
            End If
        End If
        
        
    Next
    
    ' Histograma deutilización
    ReDim histograma(0 To UBound(graficosEstado) \ 1000) As Integer
    
    For loopFrame = 1 To UBound(graficosEstado)
        If graficosEstado(loopFrame).Tipo = 0 Then
                histograma(loopFrame \ 1000) = histograma(loopFrame \ 1000) + 1
        End If
    Next
    
    
    Dim his As String
    his = ""
    
    For loopFrame = 0 To UBound(histograma)
        his = his & loopFrame & ":" & histograma(loopFrame) & vbNewLine
    Next
    
    MsgBox his & "Cantidad slots sin usar " & cantidad & ". Usados :" & pakGraficos.getCantidadElementos - cantidad & ". Graficos huerfanos: " & cantidadNoUsados


    'Extraer los gráficos no utilizados
    Dim respuesta As VbMsgBoxResult
    Dim numero As Integer
     
    respuesta = MsgBox("¿Extraer los graficos que no se utilizan?", vbYesNo, Me.caption)
   
    If respuesta = vbYes Then
        numero = 1
        For loopFrame = 1 To UBound(graficosEstado)
            If graficosEstado(loopFrame).Tipo = 0 Then
                If pakGraficos.Cabezal_GetFileSize(loopFrame) > 0 Then
                    Call pakGraficos.Extraer(loopFrame, app.Path & "\OUTPUT\" & loopFrame & ".png")
                    Debug.Print numero
                    numero = numero + 1
                End If
            End If
        Next
    End If

    respuesta = MsgBox("¿Generar DELETES?", vbYesNo, Me.caption)
   
    Dim strDelete As String
   
    strDelete = "UPDATE recursos SET ESTADO='CONFIRMADO' WHERE tipo='RECURSO_IMAGEN' AND ID IN("
   
    If respuesta = vbYes Then
        For loopFrame = 1 To UBound(graficosEstado)
            If Not graficosEstado(loopFrame).Tipo = 0 Then
                strDelete = strDelete & loopFrame & ","
            End If
        Next
        
        strDelete = mid$(strDelete, 1, Len(strDelete) - 1) & ")"
    
        Dim archivo As Integer
        archivo = FreeFile
    
        Open "C:\salida" For Output As #archivo
            Print #archivo, strDelete
        Close #archivo
        
    End If
    
    
     respuesta = MsgBox("¿Reacomodar?", vbYesNo, Me.caption)
   
   Dim nuevoid As Long
   Dim elementoTocado As Variant
   Dim INFOHEADER As INFOHEADER
   Dim Data() As Byte
   Dim loopPiso As Integer
   
    If respuesta = vbYes Then
        strDelete = ""
        For loopFrame = 5000 To UBound(graficosEstado)
            '¿Se usa?
            If Not graficosEstado(loopFrame).Tipo = 0 Then
                
                nuevoid = CDM.cerebro.SolicitarRecurso("RECURSO_IMAGEN")
             
                strDelete = strDelete & loopFrame & "-> " & nuevoid & ". Efectados: " & Coleccion_Join(graficosEstado(loopFrame).ids) & vbNewLine
                
                ' se usa en un grafico
                If graficosEstado(loopFrame).Tipo = 1 Then
                    For Each elementoTocado In graficosEstado(loopFrame).ids
                        GrhData(elementoTocado).filenum = nuevoid
                        Call Me_indexar_Graficos.actualizarEnIni(elementoTocado)
                    Next
                ElseIf graficosEstado(loopFrame).Tipo = 2 Then
                    ' Se usa en un piso
                    For Each elementoTocado In graficosEstado(loopFrame).ids
                    
                        If Tilesets(elementoTocado).filenum = loopFrame Then
                             Tilesets(elementoTocado).filenum = nuevoid
                        End If
                        
                       For loopPiso = 1 To UBound(Tilesets(elementoTocado).stages)
                            If Tilesets(elementoTocado).stages(loopPiso) = loopFrame Then
                                Tilesets(elementoTocado).stages(loopPiso) = nuevoid
                            End If
                        Next
        
                    
                        Call Me_indexar_Pisos.actualizarEnIni(CInt(elementoTocado))
                    Next
                ElseIf graficosEstado(loopFrame).Tipo = 3 Then
                    ' Se usa en una olita
                     For Each elementoTocado In graficosEstado(loopFrame).ids
                         Tilesets(elementoTocado).Olitas = nuevoid
                         Call Me_indexar_Pisos.actualizarEnIni(CInt(elementoTocado))
                     Next
                        
                    
                End If
               
                
                ' Pongo el viejo en el slot nuevo
                
                'Obtenemos el cabezal
                Debug.Print loopFrame
                Call pakGraficos.IH_Get(CInt(loopFrame), INFOHEADER)
                
                ' Obtenemos los datos
                Dim nombre As String
                Dim nuevoNombre As String * 32
                Dim borrar(0) As Byte
                
                nombre = Xor_String(INFOHEADER.originalname, INFOHEADER.cript)
                nuevoNombre = CStr(nuevoid) & mid$(nombre, InStr(1, nombre, "."))
                ' Blanqueo
                INFOHEADER.originalname = Space$(32)
                ' Asigno
                LSet INFOHEADER.originalname = Xor_String(nuevoNombre, INFOHEADER.cript)
                           
                Call pakGraficos.LeerIH(Data, INFOHEADER)
                ' Parcheamos

                Call pakGraficos.ParchearByteArray(Data, nuevoid, INFOHEADER)
                
                INFOHEADER.originalname = Xor_String("--------------------------------", INFOHEADER.cript)
                Call pakGraficos.ParchearByteArray(borrar, loopFrame, INFOHEADER)
            End If
        Next
        
    
        archivo = FreeFile
    
        Open "C:\salida.txt" For Output As #archivo
            Print #archivo, strDelete
        Close #archivo
        
        MsgBox "Graficos reacomodados"
    End If


End Sub

Private Sub cmdHechizos_Click()
    Dim loopHechizo As Integer
    
    For loopHechizo = LBound(FxData) To UBound(FxData)
        
        If FxData(loopHechizo).Animacion > 0 Then
            If GrhData(FxData(loopHechizo).Animacion).nombreGrafico = "" Then
                GrhData(FxData(loopHechizo).Animacion).nombreGrafico = FxData(loopHechizo).nombre
                
                Me_indexar_Graficos.actualizarEnIni (FxData(loopHechizo).Animacion)
            End If
        End If
    Next
End Sub

Private Sub cmdImportarConfiguracion_Click()
    ' Mostramos el formulario
    Call frmImportarConfiguracionGraficos.Show(, Me)
End Sub

Private Sub cmdResetGUUID_Click()
    Dim resultado As VbMsgBoxResult
    
    resultado = MsgBox("¿Estás seguro que queres RESETEAR el identificador base de este gráfico?. Tenés que estar muy seguro de esto porque se va a perder la relación con la información del generador de imagenes.", vbYesNo + vbQuestion, Me.caption)
    
    If resultado = vbYes Then
        GrhData(TmpGrhIndexarNum).ID = ""
        Me_indexar_Graficos.actualizarEnIni (TmpGrhIndexarNum)
        
        Me.lblIdentificadorGrafico.caption = ""
    End If
End Sub

Private Sub cmdToogleCheckedList_Click()
    Me.Listado.checked = Not Me.Listado.checked
End Sub

Private Sub FileName_change(valor As String, ID As Integer)
    Call habilitarBotonesAplicarRestablecer
End Sub

Private Sub cargarListaGraficos()
Dim i As Integer
Dim loopFrame As Integer

Listado.vaciar

' Cargo los gráficos disponibles
For i = 1 To grhCount
    With GrhData(i)
        If .perteneceAunaAnimacion = False Then
            If .NumFrames > 1 Then
                Call Listado.addString(CLng(i), i & " - " & .nombreGrafico, 0)
                
                For loopFrame = 1 To .NumFrames
                  Call Listado.addString(.Frames(loopFrame), .Frames(loopFrame) & " - " & GrhData(.Frames(loopFrame)).nombreGrafico, CLng(i))
                Next
            Else
                If .filenum Then
                    Call Listado.addString(CLng(i), i & " - " & .nombreGrafico, 0)
                End If
            End If
        End If
    End With
Next i

End Sub
Private Sub Form_Load()

' Permisos: Esto es grave! Deberia denunciarlo
If Not cerebro.Usuario.tienePermisos("CONFIG.GRAFICOS", ePermisosCDM.lectura) Then End

Dim i As Integer

Dim elementos() As modEnumerandosDinamicos.eEnumerado

If ME_ControlCambios.hayCambiosSinActualizarDe("Graficos") Then
    cambiosPendientes = True
Else
    cambiosPendientes = False
End If
    
Set vwEditorGraficos = New vw_EditorGraficos
    
Call cargarListaGraficos

Me.lblEncontrados.caption = "Encontrados: " & Listado.obtenerCantidadVisible
   
' Imagenes disponibles
elementos = modEnumerandosDinamicos.obtenerEnumeradosDinamicos("IMAGENES")

For i = LBound(elementos) To UBound(elementos)
    Call Me.FileName.addString(elementos(i).valor, elementos(i).valor & " - " & elementos(i).nombre)
Next

Me.FileName.CantidadLineasAMostrar = 12

' Efectos de pisada al pisar este grafico
elementos = modEnumerandosDinamicos.obtenerEnumeradosDinamicos("EFECTOS_PISADAS")
        
Call Me.txtEfectoPisada.limpiarLista

For i = LBound(elementos) To UBound(elementos)
    Call Me.txtEfectoPisada.addString(elementos(i).valor, elementos(i).valor & " - " & elementos(i).nombre)
Next

Me.txtEfectoPisada.CantidadLineasAMostrar = 10
    
' Actualizo el label que muestra la cantidad de graficos existentes
actualizarCantidadgraficos

' Configuraciones de botones y textos
Me.txtEfectoPisada.CantidadLineasAMostrar = 5

Me.nroInicio.MinValue = 1
Me.nroInicio.MaxValue = grhCount

Me.nroHasta.MinValue = Me.nroInicio.MinValue + 1
Me.nroHasta.MaxValue = grhCount

Set vwEditorGraficos = New vw_EditorGraficos

Call vwEditorGraficos.SetGrafico(0)

GUI_Load vwEditorGraficos

' Arranca todo inahbilitado
Call modPosicionarFormulario.setEnabledHijos(False, Me.frmTipoAnimacion, Me)

End Sub

Private Sub agregarGrhALista(i As Integer)
    Dim loopFrame As Integer
    
    With GrhData(i)
        If .perteneceAunaAnimacion = False Then
            If .NumFrames > 1 Then
                Call Listado.addString(CLng(i), i & " - " & .nombreGrafico, 0)
                
                For loopFrame = 1 To .NumFrames
                  Call Listado.addString(.Frames(loopFrame), .Frames(loopFrame) & " - " & GrhData(.Frames(loopFrame)).nombreGrafico, CLng(i))
                Next
            Else
                If .filenum Then
                    Call Listado.addString(CLng(i), i & " - " & .nombreGrafico, 0)
                End If
            End If
        End If
    End With
        

End Sub
Private Sub Form_Unload(Cancel As Integer)
    TmpGrhIndexarNum = 0
    
    Call GUI_Quitar(vwEditorGraficos)
    
    Me.Hide
    Me.Refresh
    
    Call Listado.vaciar
End Sub

Private Sub Frames_Change()
    Call habilitarBotonesAplicarRestablecer
End Sub

Private Sub habilitarBotonesAplicarRestablecer()
    Visualizar.Enabled = True
    Reestablecer.Enabled = True
End Sub

Public Sub seActualizoGraficos()
    ' Recargaamos la lista
    Call cargarListaGraficos
End Sub

Private Sub lblSecreto_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim clave As String
    
    If Button = 2 Then
        clave = InputBox("")
        
        If clave = "panelsecreto27" Then
            Me.frmPanelSecreto.Visible = True
        Else
            Beep
        End If
    Else
        Me.frmPanelSecreto.Visible = False
    End If
    
End Sub

Private Sub Listado_Change(valor As String, ID As Integer, esPadre As Boolean)
    If ID > 0 Then
        TmpGrhIndexarNum = ID
               
        ' Cargamos
        reestablece
        
        ' Vista previa
        Call vwEditorGraficos.SetGrafico(ID)
        
        'Activo el eliminar
        Me.cmdEliminar_Graficos.Enabled = True
        
         ' Activamos
        Call modPosicionarFormulario.setEnabledHijos(True, Me.frmTipoAnimacion, Me)
    End If
    
    Me.lblEncontrados.caption = "Encontrados: " & Listado.obtenerCantidadVisible
End Sub

Private Sub actualizarCantidadgraficos()

    Dim loopGrh As Integer
    
    Dim cantidad As Integer
    Dim animaciones As Integer
    Dim simples As Integer
    Dim sinNombre As Integer
    Dim pertenecenAAnim As Integer
    Dim insertables As Integer
    
    cantidad = 0 'Cantidad total de graficos
    animaciones = 0 ' Cantidad de animaciones
    simples = 0 ' Cantidad de elementos simples
    sinNombre = 0 'Cantidad de elementos sin nombre
    pertenecenAAnim = 0 ' Cantidad que pertenece a una animacion
    insertables = 0 ' Cantidad de graficos que se pueden insertar
    
    For loopGrh = 1 To grhCount
        If Me_indexar_Graficos.existe(loopGrh) Then
            
            If GrhData(loopGrh).NumFrames > 1 Then
                'Es una animacion
                animaciones = animaciones + 1
            Else
                If GrhData(loopGrh).perteneceAunaAnimacion Then
                    pertenecenAAnim = pertenecenAAnim + 1
                Else
                    simples = simples + 1
                End If
            
                cantidad = cantidad + 1
            End If
            
            If Not GrhData(loopGrh).perteneceAunaAnimacion Then
                If Len(GrhData(loopGrh).nombreGrafico) = 0 Then
                    sinNombre = sinNombre + 1
                End If
                
                If GrhData(loopGrh).esInsertableEnMapa Then
                    insertables = insertables + 1
                End If
            End If
        End If
    Next
    
    
    Me.lblCantidadGraficos = "Graficos: " & cantidad & ". Simples: " & simples & ". Animaciones: " & animaciones & " (" & pertenecenAAnim & ") . Sin nombre:" & sinNombre & ". Insertables: " & insertables
End Sub
Private Sub cmdNuevo_Graficos_Click()
    Dim nuevo As Integer
    Dim error As Boolean
    
    error = False
    Me.cmdNuevo_Graficos.Enabled = False
    
    nuevo = Me_indexar_Graficos.nuevo

    If nuevo = -1 Then
        error = True
        MsgBox "No se ha podido obtener espacio para agregar un nuevo grafico. Por favor, intenta más tarde o contactate con un administrador del sistema.", vbExclamation
    End If
    
    If Not error Then
        Listado.addString CLng(nuevo), nuevo & "-" & " (LIBRE)", 0
        Listado.seleccionarElemento (CLng(nuevo))
    End If
    
    Me.cmdNuevo_Graficos.Enabled = True

    actualizarCantidadgraficos
End Sub

Private Sub optCentrado_Click(Index As Integer)

    Dim SliderOffsetActivado As Boolean
    
    Select Case Index
    
        Case eTipoCentrado.centrarEnTile
        
            Me.sliderOffsetX.value = 0
            Me.sliderOffsetY.value = 0
            
            SliderOffsetActivado = False
        
        Case eTipoCentrado.ajustarAGrilla
            
            ' Ex Centrar en 32. El offset se ajusta para que el gráfico quede centrado
            SliderOffsetActivado = False
            
        Case eTipoCentrado.personalizado
            
            ' Dejo que ponga el offset a gusto
            SliderOffsetActivado = True
    
    End Select
    
    Me.sliderOffsetX.Enabled = SliderOffsetActivado
    Me.sliderOffsetY.Enabled = SliderOffsetActivado
    
    Call habilitarBotonesAplicarRestablecer
End Sub

Private Sub optTipoIndex_Click(Index As Integer)
    If Index = 0 Then
        If optTipoIndex(Index).value = True Then
            Me.frmIndexSimple.Visible = True
            Me.frmIndexAnimacion.Visible = False
        End If
    ElseIf Index = 1 Then
        If optTipoIndex(Index).value = True Then
            Me.frmIndexSimple.Visible = False
            Me.frmIndexAnimacion.Visible = True
        End If
    End If
    
End Sub

Private Sub pixelHeight_Change(valor As Double)
    Call habilitarBotonesAplicarRestablecer
End Sub

Private Sub pixelWidth_Change(valor As Double)
    Call habilitarBotonesAplicarRestablecer
End Sub

Private Sub Reestablecer_Click()
reestablece
End Sub



Private Sub Speed_Change(valor As Double)
    Call habilitarBotonesAplicarRestablecer
End Sub

Private Sub sx_Change(valor As Double)
    Call habilitarBotonesAplicarRestablecer
End Sub

Private Sub sy_Change(valor As Double)
    Call habilitarBotonesAplicarRestablecer
End Sub

Sub Aplicar()
On Error GoTo errorH
    Dim n As Integer
    Dim T() As String   'TempString
    Dim nf As Integer   'NumFrames
    Dim tf() As Long    'TempFrames
    Dim resultado As VbMsgBoxResult
    
    Dim numeroGrhFrame As Integer
    Dim GRH_BACKUP As GrhData
    Dim loopFrame As Byte
    Dim loopCapa As Byte
    
    Dim Offset As Position
    
    GRH_BACKUP = GrhData(TmpGrhIndexarNum)
    
    If InStr(1, Me.txtNombreGrafico, Me_indexar_Graficos.SEPARADOR_PROPIEDADES) Then
        Call MsgBox("No es posible usar la letra '" & Me_indexar_Graficos.SEPARADOR_PROPIEDADES & "' en el nombre del archivo.", vbOKOnly + vbExclamation)
        Exit Sub
    End If
    
    With GrhData(TmpGrhIndexarNum)
    
        ' Es un grafico simple
        If Me.optTipoIndex(GRH_SIMPLE).value Then
        
            If GRH_BACKUP.NumFrames > 1 Then
                resultado = MsgBox("¿Estas seguro que queres cambiar el elemento de ser una Animación a un Gráfíco Simple?. Se perderá la información de los frames.", vbYesNo + vbExclamation, "Cuidado")
                If resultado = vbNo Then
                    Exit Sub
                End If
            End If
            
            .filenum = val(FileName.obtenerIDValor)
            
            'TO-DO ¿Estas posiciones estan dentro del tamaño?
            .sx = CInt(sx.value)
            .sy = CInt(sy.value)
            
            'TO-DO ¿Es un tamaño válido?
            .pixelWidth = CInt(pixelWidth.value)
            .pixelHeight = CInt(pixelHeight.value)
        
            ' Sonido
            .EfectoPisada = Me.txtEfectoPisada.obtenerIDValor
        
            If .pixelWidth = 0 Or .pixelHeight = 0 Then GoTo errorH
            
            ' Redimensiono los frames a 1
            ReDim .Frames(1 To 1)
            
            .NumFrames = 1
            .Frames(1) = TmpGrhIndexarNum
            .Speed = 0
        
        ' ¿Es una animacion?
        ElseIf Me.optTipoIndex(GRH_ANIMACION).value = True Then
        
            ' ¿Tiene más de un frame?
            If GRH_BACKUP.filenum > 0 Then
                resultado = MsgBox("¿Estas seguro que queres cambiar el elemento de ser un Gráfico simple a una Animación?. Se perderá la información que poseen las animaciones simples.", vbYesNo + vbExclamation, "Cuidado")
                If resultado = vbNo Then
                    Exit Sub
                End If
            End If
            
            'Valido GrhIndex correspondiente a frames
            T = Split(Frames.text, vbCr)
        
            For n = 0 To UBound(T)
                
                numeroGrhFrame = val(T(n))
                
                If numeroGrhFrame <= grhCount Then
                    If numeroGrhFrame > 0 Then
                    If GrhData(numeroGrhFrame).NumFrames > 0 Then
                        nf = nf + 1
                        ReDim Preserve tf(nf)
                        tf(nf) = numeroGrhFrame
                    Else
                        MsgBox "El elemento número " & numeroGrhFrame & " que se esta intentando utilizar como frame no existe.", vbExclamation
                        Exit Sub
                    End If
                    End If
                Else
                    MsgBox "El elemento número " & numeroGrhFrame & " que se esta intentando utilizar como frame no existe.", vbExclamation
                    Exit Sub
                End If
                
            Next n
            
            .Speed = CInt(Speed.value)
            
            If .Speed = 0 Then
                MsgBox "El tiempo que pasa entre cada gráfico es 0. Eso no puede ser. Pasaria infinitamente rápido que la persona no se podría dar cuenta.", vbExclamation
                Exit Sub
            End If
            
            'Actualizo
            .NumFrames = nf
            
            ReDim .Frames(1 To nf)
            
            For n = 1 To .NumFrames
                .Frames(n) = tf(n)
                Debug.Print tf(n)
             Next
                     
            
            .pixelWidth = GrhData(.Frames(1)).pixelWidth
            .pixelHeight = GrhData(.Frames(1)).pixelHeight
                   
            ' La pisada es la del gráfico 1, si fuese por cada gráfico seria un lio y rara vez existen animaciones
            ' en las que el personaje pueda pisar
            .EfectoPisada = GrhData(.Frames(1)).EfectoPisada
        End If

                        
        .nombreGrafico = Me.txtNombreGrafico
        .esInsertableEnMapa = (Me.chkGraficoInsertableMapa.value = 1)
        
        If .esInsertableEnMapa Then
            For loopCapa = 1 To CANTIDAD_CAPAS
                .Capa(loopCapa) = (Me.Capa(loopCapa).value = vbChecked)
            Next loopCapa
        End If
        
    End With

            
    If Me.optCentrado(eTipoCentrado.centrarEnTile).value = True Then
        Call Me_indexar_Graficos.obtenerOffsetNatural(GrhData(TmpGrhIndexarNum), Offset.x, Offset.y)
        GrhData(TmpGrhIndexarNum).offsetX = Offset.x
        GrhData(TmpGrhIndexarNum).offsetY = Offset.y
    ElseIf Me.optCentrado(eTipoCentrado.ajustarAGrilla).value = True Then
        Call Me_indexar_Graficos.obtenerOffsetAjustadoTile(GrhData(TmpGrhIndexarNum), Offset.x, Offset.y)
        GrhData(TmpGrhIndexarNum).offsetX = Offset.x
        GrhData(TmpGrhIndexarNum).offsetY = Offset.y
    Else
        Call Me_indexar_Graficos.establecerOffsetBruto(GrhData(TmpGrhIndexarNum), Me.sliderOffsetX.value, Me.sliderOffsetY.value)
    End If
                
    'Actualizo Graficos.ini
    Call Me_indexar_Graficos.actualizarEnIni(TmpGrhIndexarNum)
    huboCambios = True
    
    'Actualizo la lista
    Call Listado.cambiarNombre(TmpGrhIndexarNum, TmpGrhIndexarNum & " - " & Me.txtNombreGrafico)
        
    If GRH_BACKUP.NumFrames > 1 Then
        Call Listado.eliminarHijos(TmpGrhIndexarNum)
    End If
    
    With GrhData(TmpGrhIndexarNum)
    
        If .NumFrames > 1 Then
            For loopFrame = 1 To .NumFrames
                'Si antes no era parte de la animación...
                If Not existeEnArray(.Frames(loopFrame), GRH_BACKUP.Frames) Then
                    GrhData(.Frames(loopFrame)).perteneceAunaAnimacion = True
                    Call Me_indexar_Graficos.actualizarEnIni(.Frames(loopFrame))
                    Call Listado.eliminarElemento(.Frames(loopFrame))
                End If
                
                Call Listado.addString(.Frames(loopFrame), .Frames(loopFrame) & " - " & GrhData(.Frames(loopFrame)).nombreGrafico, TmpGrhIndexarNum)
            Next

        End If
        
        If GRH_BACKUP.NumFrames > 1 Then
            For loopFrame = 1 To UBound(GRH_BACKUP.Frames)
                'Hay un elemento que ya no pertenece a la animacióm
                If Not existeEnArray(GRH_BACKUP.Frames(loopFrame), .Frames) Then
                    GrhData(GRH_BACKUP.Frames(loopFrame)).perteneceAunaAnimacion = False
                    Call Me_indexar_Graficos.actualizarEnIni(GRH_BACKUP.Frames(loopFrame))
                    Call Listado.addString(GRH_BACKUP.Frames(loopFrame), GRH_BACKUP.Frames(loopFrame) & " - " & GrhData(GRH_BACKUP.Frames(loopFrame)).nombreGrafico, 0)
                End If
            Next
        End If
        
    End With

    ' Re inicializamos el grafico
    Init_grh_tutv TmpGrhIndexarNum
    
    Call ejecutarControlCambios
        
    'Actualizo la lista
    CargarListaGraficosComunes
    'Actualizo la informacion del formulario
    reestablece
    'Botonera
    Visualizar.Enabled = False
    Reestablecer.Enabled = False

    
    GRH_BACKUP = GrhData(TmpGrhIndexarNum)
    Exit Sub
errorH:
GrhData(TmpGrhIndexarNum) = GRH_BACKUP
MsgBox "Error al aplicar los cambios, comprobá los datos."
End Sub




Sub reestablece()
    Dim n As Integer
    Dim loopCapa As Byte
    Dim C1 As Integer
    Dim C2 As Integer
    Dim C3 As Integer
   

   
    ' Cargamos los datos en el textbox
    With GrhData(TmpGrhIndexarNum)
        sx.value = .sx
        sy.value = .sy
        Index.text = TmpGrhIndexarNum
        FileName.seleccionarID (.filenum)
                
        pixelWidth.value = .pixelWidth
        pixelHeight.value = .pixelHeight
        Speed.value = .Speed
        Frames.text = vbNullString
        Me.txtNombreGrafico = .nombreGrafico
        
        Me.lblIdentificadorGrafico.caption = .ID
        
        If .esInsertableEnMapa Then
            Me.chkGraficoInsertableMapa.value = 1
            
            For loopCapa = 1 To CANTIDAD_CAPAS
                Me.Capa(loopCapa).value = IIf(.Capa(loopCapa), vbChecked, vbUnchecked)
            Next
            
        Else
            Me.chkGraficoInsertableMapa.value = 0
            
            For loopCapa = 1 To CANTIDAD_CAPAS
                Me.Capa(loopCapa).value = vbChecked
            Next
        End If
        
        If .NumFrames > 1 Then
            For n = 1 To .NumFrames
                Frames.text = Frames.text & .Frames(n) & vbCrLf
            Next n
            
            optTipoIndex(1).Enabled = True
            optTipoIndex(0).Enabled = True
            optTipoIndex(1).value = True
            Call optTipoIndex_Click(1)
        Else
        
            C1 = 0
            C2 = 0
            C3 = 0
            
            ' Nos aseguramos la carga de la textura
            Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(GrhData(TmpGrhIndexarNum).filenum)
    
            If .filenum Then Obtener_Texturas_Complementarias CInt(.filenum), C1, C2, C3
        
            Text1.text = IIf(C1 = 0, "-", C1 & " - " & pakGraficos.Cabezal_GetFileNameSinComplementos(C1))
            txtColorAdd.text = IIf(C2 = 0, "-", C2 & " - " & pakGraficos.Cabezal_GetFileNameSinComplementos(C2))
            txtComplementoNormal.text = IIf(C3 = 0, "-", C3 & " - " & pakGraficos.Cabezal_GetFileNameSinComplementos(C3))
        
            optTipoIndex(1).Enabled = True
            optTipoIndex(0).Enabled = True
            optTipoIndex(0).value = True
            Call optTipoIndex_Click(0)
                        
        End If
        
        Call setOffsetEnFormulario(GrhData(TmpGrhIndexarNum))
        
        Me.sliderOffsetX.MinValue = -32
        Me.sliderOffsetY.MinValue = -32
        
        Me.sliderOffsetX.MaxValue = 32
        Me.sliderOffsetY.MaxValue = 32
        
        Call setSombra(GrhData(TmpGrhIndexarNum))
        
        ' Efecto de Sonido al Pisar este Gráfico
        Me.txtEfectoPisada.seleccionarID (.EfectoPisada)
    End With
    
        
    Visualizar.Enabled = False
    Reestablecer.Enabled = False
End Sub

Private Sub setSombra(Grh As GrhData)

    ' ¿Sombra?
    If Grh.SombrasSize > 0 Then
        Me.lblTieneSombra = "SI tiene sombra"
    Else
        Me.lblTieneSombra = "NO tiene sombra"
    End If
        
End Sub

Private Sub setOffsetEnFormulario(Grh As GrhData)

    Dim offsetNatural As Position
    Dim OffsetAjustado As Position
    Dim offsetNeto As Position
   
    ' Calculamos que tipo de centrado tiene
    Call Me_indexar_Graficos.obtenerOffsetNatural(Grh, offsetNatural.x, offsetNatural.y)
    Call Me_indexar_Graficos.obtenerOffsetAjustadoTile(Grh, OffsetAjustado.x, OffsetAjustado.y)
       
    If Grh.offsetX = OffsetAjustado.x And Grh.offsetY = OffsetAjustado.y Then
        
        Me.optCentrado(eTipoCentrado.ajustarAGrilla).value = True
            
        Me.sliderOffsetX.value = 0
        Me.sliderOffsetY.value = 0
            
    ElseIf Grh.offsetX = offsetNatural.x And Grh.offsetY = offsetNatural.y Then
        
        Me.optCentrado(eTipoCentrado.centrarEnTile).value = True
            
        Me.sliderOffsetX.value = 0
        Me.sliderOffsetY.value = 0
    
    Else
        
        Me.optCentrado(eTipoCentrado.personalizado).value = True
            
        Call Me_indexar_Graficos.calcularOffsetNeto(Grh, offsetNeto.x, offsetNeto.y)
            
        Me.sliderOffsetX.value = offsetNeto.x
        Me.sliderOffsetY.value = offsetNeto.y
    End If
End Sub

Private Sub txtEfectoPisada_change(valor As String, ID As Integer)
    Call habilitarBotonesAplicarRestablecer
End Sub

Private Sub txtNombreGrafico_Change()
    habilitarBotonesAplicarRestablecer
End Sub

Private Sub Visualizar_Click()
    Call Aplicar
    Call actualizarCantidadgraficos
End Sub

Private Sub ejecutarControlCambios()
    If cambiosPendientes = False Then
        cambiosPendientes = True
    End If
End Sub

Private Sub vwEditorGraficos_Aplicar()

    ' Cargamos en este formulario la modificacion al offset
    Call setOffsetEnFormulario(GrhData(TmpGrhIndexarNum))
    
    ' Sombra
    Call setSombra(GrhData(TmpGrhIndexarNum))
    
    ' Guardamos
    Call Aplicar
End Sub
