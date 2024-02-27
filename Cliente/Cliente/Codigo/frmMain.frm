VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Object = "{50CBA22D-9024-11D1-AD8F-8E94A5273767}#8.7#0"; "TranImg2.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000040&
   BorderStyle     =   0  'None
   Caption         =   "Juego Tierras del Sur"
   ClientHeight    =   11520
   ClientLeft      =   360
   ClientTop       =   270
   ClientWidth     =   19380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
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
   HelpContextID   =   -1
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0E42
   ScaleHeight     =   768
   ScaleMode       =   0  'User
   ScaleWidth      =   1292
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   6720
      Top             =   2520
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   5
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   10920
      Width           =   6615
   End
   Begin VB.PictureBox MinimapUser 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   17160
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   7920
      Width           =   45
   End
   Begin VB.Timer TimerReRenderInv 
      Interval        =   200
      Left            =   8520
      Top             =   2160
   End
   Begin DevPowerTransImg.TransImg itemimg 
      Height          =   495
      Left            =   8760
      TabIndex        =   6
      Top             =   6600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      AutoSize        =   0   'False
      MaskColor       =   0
      Transparent     =   -1  'True
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2760
      Left            =   10740
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2580
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2790
      Left            =   11280
      ScaleHeight     =   186
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   11
      Top             =   2520
      Width           =   3435
   End
   Begin TDS_1.Caption label15 
      Height          =   180
      Left            =   10980
      TabIndex        =   7
      Top             =   9630
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   318
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAligmend =   1
      CaptionShadowed =   0   'False
   End
   Begin VB.Timer tNoche 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1680
      Top             =   2520
   End
   Begin VB.Timer Pasarsegundo 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   2520
   End
   Begin VB.Timer SoundFX 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1200
      Top             =   2520
   End
   Begin VB.TextBox SendGMSTXT 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   11010
      Visible         =   0   'False
      Width           =   9495
   End
   Begin VB.TextBox SendRMSTXT 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   11010
      Visible         =   0   'False
      Width           =   9495
   End
   Begin TDS_1.Caption label17 
      Height          =   180
      Left            =   13185
      TabIndex        =   8
      Top             =   8565
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   318
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAligmend =   1
      CaptionShadowed =   0   'False
   End
   Begin TDS_1.Caption label13 
      Height          =   180
      Left            =   10980
      TabIndex        =   9
      Top             =   7590
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   318
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAligmend =   1
      CaptionShadowed =   0   'False
   End
   Begin TDS_1.Caption label14 
      Height          =   180
      Left            =   10980
      TabIndex        =   10
      Top             =   8565
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   318
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAligmend =   1
      CaptionShadowed =   0   'False
   End
   Begin TDS_1.Caption Hpshp 
      Height          =   180
      Left            =   10980
      TabIndex        =   13
      Top             =   9630
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   318
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionGradientStart=   255
      CaptionGradientEnds=   8421631
      CaptionGradientchangeColor=   -1  'True
      CaptionGradientTyp=   2
   End
   Begin TDS_1.Caption label16 
      Height          =   180
      Left            =   13200
      TabIndex        =   14
      Top             =   9360
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   318
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAligmend =   1
      CaptionShadowed =   0   'False
   End
   Begin TDS_1.Caption stashp 
      CausesValidation=   0   'False
      Height          =   180
      Left            =   10980
      TabIndex        =   19
      Top             =   7590
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   318
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionGradientStart=   32896
      CaptionGradientEnds=   12648447
      CaptionGradientchangeColor=   -1  'True
      CaptionGradientTyp=   2
   End
   Begin TDS_1.Caption ManShp 
      CausesValidation=   0   'False
      Height          =   180
      Left            =   10980
      TabIndex        =   20
      Top             =   8565
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   318
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionGradientStart=   12632064
      CaptionGradientEnds=   16777152
      CaptionGradientchangeColor=   -1  'True
      CaptionGradientTyp=   2
   End
   Begin TDS_1.Caption Aguasp 
      Height          =   180
      Left            =   13185
      TabIndex        =   21
      Top             =   8580
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   318
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionINColor  =   0
      CaptionOUTColor =   255
      CaptionGradientStart=   8388608
      CaptionGradientEnds=   16744576
      CaptionGradientchangeColor=   -1  'True
      CaptionGradientTyp=   2
   End
   Begin TDS_1.Caption comidasp 
      CausesValidation=   0   'False
      Height          =   180
      Left            =   13185
      TabIndex        =   22
      Top             =   9660
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   318
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionGradientStart=   32768
      CaptionGradientEnds=   8454016
      CaptionGradientchangeColor=   -1  'True
      CaptionGradientTyp=   2
   End
   Begin VB.PictureBox Renderer 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10080
      Left            =   12360
      ScaleHeight     =   672
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   672
      TabIndex        =   30
      Top             =   1200
      Width           =   10080
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   14160
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   5040
         Width           =   390
      End
      Begin VB.Timer pasarMinuto 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   2040
         Top             =   2160
      End
   End
   Begin TDS_1.Caption lGrabando 
      Height          =   270
      Left            =   11280
      TabIndex        =   29
      Top             =   6240
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionShadowed =   0   'False
   End
   Begin TDS_1.Caption NumExp 
      Height          =   180
      Left            =   10950
      TabIndex        =   37
      Top             =   1590
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   318
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAligmend =   1
      CaptionShadowed =   0   'False
   End
   Begin TDS_1.Caption expshp 
      CausesValidation=   0   'False
      Height          =   165
      Left            =   10950
      TabIndex        =   36
      Top             =   1620
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   291
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionGradientStart=   4194368
      CaptionGradientEnds=   12583104
      CaptionGradientchangeColor=   -1  'True
      CaptionGradientTyp=   1
   End
   Begin VB.Image imgBotonMapa 
      Height          =   255
      Left            =   11280
      Top             =   11040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblIndicadorEscritura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   40
      Top             =   11160
      Width           =   195
   End
   Begin VB.Image Minimapa 
      Height          =   1935
      Left            =   16560
      Top             =   7440
      Width           =   2175
   End
   Begin VB.Image Curar 
      Height          =   495
      Left            =   10440
      Top             =   9480
      Width           =   495
   End
   Begin VB.Image CalcularPing 
      Height          =   255
      Left            =   12240
      Top             =   180
      Width           =   750
   End
   Begin VB.Label PING 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   13110
      TabIndex        =   38
      Top             =   180
      Width           =   375
   End
   Begin VB.Image Descansar 
      Height          =   495
      Left            =   10470
      Top             =   7440
      Width           =   495
   End
   Begin VB.Image Meditar 
      Height          =   495
      Left            =   10470
      Top             =   8400
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   16
      Left            =   14520
      MousePointer    =   99  'Custom
      Top             =   10065
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   15
      Left            =   14520
      MousePointer    =   99  'Custom
      Top             =   9255
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   14
      Left            =   14520
      MousePointer    =   99  'Custom
      Top             =   8445
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   13
      Left            =   14520
      MousePointer    =   99  'Custom
      Top             =   6000
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   12
      Left            =   12240
      MousePointer    =   99  'Custom
      Top             =   10140
      Width           =   390
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   11
      Left            =   13140
      MousePointer    =   99  'Custom
      Top             =   10140
      Width           =   390
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   10
      Left            =   12690
      MousePointer    =   99  'Custom
      Top             =   10140
      Width           =   390
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   9
      Left            =   11820
      MousePointer    =   99  'Custom
      Top             =   10140
      Width           =   390
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   8
      Left            =   11400
      MousePointer    =   99  'Custom
      Top             =   10140
      Width           =   390
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   7
      Left            =   10980
      MousePointer    =   99  'Custom
      Top             =   10140
      Width           =   390
   End
   Begin VB.Label lblClan 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<NOMBRE CLAN>"
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   11280
      TabIndex        =   35
      Top             =   1320
      Width           =   1950
   End
   Begin VB.Label GldLbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "999.999.999"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12960
      TabIndex        =   18
      Top             =   7560
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   6
      Left            =   14520
      MousePointer    =   99  'Custom
      Top             =   7620
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   5
      Left            =   14520
      MousePointer    =   99  'Custom
      ToolTipText     =   "Manual"
      Top             =   5190
      Width           =   555
   End
   Begin VB.Label LvlLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "45"
      ForeColor       =   &H0000C0C0&
      Height          =   195
      Left            =   10560
      TabIndex        =   34
      Top             =   1590
      Width           =   330
   End
   Begin VB.Label lblUserName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pom"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   270
      Left            =   11280
      TabIndex        =   33
      Top             =   960
      Width           =   1905
   End
   Begin VB.Image imgConsola 
      Height          =   300
      Left            =   14520
      MousePointer    =   99  'Custom
      Top             =   10995
      Width           =   450
   End
   Begin VB.Label lblLink 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   600
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   11040
      Width           =   9135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   660
      Left            =   480
      TabIndex        =   28
      Top             =   0
      Width           =   9000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   13620
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label IconoDyd 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   13020
      TabIndex        =   27
      Top             =   5640
      Width           =   330
   End
   Begin VB.Label IconoSeg 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   11160
      TabIndex        =   26
      Top             =   5640
      Width           =   330
   End
   Begin VB.Label FPS 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   11160
      TabIndex        =   25
      Top             =   180
      Width           =   375
   End
   Begin VB.Label Minimizar 
      BackStyle       =   0  'Transparent
      Height          =   345
      Left            =   14280
      TabIndex        =   24
      Top             =   120
      Width           =   420
   End
   Begin VB.Label lblCerrar 
      BackStyle       =   0  'Transparent
      Height          =   345
      Left            =   14850
      TabIndex        =   23
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Act"
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   11040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   510
      Index           =   0
      Left            =   13395
      MousePointer    =   99  'Custom
      Top             =   3030
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   525
      Index           =   1
      Left            =   13395
      MousePointer    =   99  'Custom
      Top             =   3900
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image CmdLanzar 
      Height          =   600
      Left            =   10725
      MouseIcon       =   "frmMain.frx":2537A
      MousePointer    =   99  'Custom
      Top             =   5460
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Image cmdInfo 
      Height          =   600
      Left            =   12870
      MousePointer    =   99  'Custom
      Top             =   5460
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   1980
      Width           =   1800
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12240
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   1980
      Width           =   1800
   End
   Begin VB.Image InvEqu 
      Height          =   4200
      Left            =   10320
      Top             =   1980
      Width           =   3840
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   14520
      MousePointer    =   99  'Custom
      Top             =   4425
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   14520
      MousePointer    =   99  'Custom
      Top             =   3645
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   2
      Left            =   14520
      MousePointer    =   99  'Custom
      Top             =   6810
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   3
      Left            =   14520
      MousePointer    =   99  'Custom
      Top             =   2880
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   4
      Left            =   14520
      MousePointer    =   99  'Custom
      Top             =   2160
      Width           =   555
   End
   Begin VB.Label Coord2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label19"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   10560
      TabIndex        =   5
      Top             =   10920
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   9840
      TabIndex        =   3
      Top             =   11040
      Width           =   735
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa 0 [0,0]"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   10560
      TabIndex        =   0
      Top             =   11010
      Width           =   3690
   End
   Begin VB.Image Image3 
      Height          =   540
      Index           =   0
      Left            =   12600
      Top             =   7350
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''''''''' TIMERS MARCHE
'n_sion@hotmail.com

' Hize un par de cambios con los timers:
'Se agregaron 4 nunevos timers al cliente.
'SoundFX,  los efectos de sonidos que antes se controlaba
'y mandaban desde el server ahora los controla el cliente
'Piquete, el timer de piquete que antes estaba en el server
'ahora lo pase al cliente. Funciona de la misma manera
'pero cuando es hora de encarcelalos el cliente avisa al server y
'el server encarcela.
'Pasar segundo. Por ahora solo funciona para los retos, es un peso
'que tiene que procesar y mandar el server.
'IntervaloLaburar, para el nuevo macro para crear objetos de
'carpinteria y herreria. Tambien agrege un modulo con los intervalos.
'INTERVALOS PARA REVISAR o DESACTIVADOS.
'Trainnign macro, desactivado para tds.
'Macro, no tiene nada adentro dsp veo si lo saco.
'Trabajar, lo mismo que el macro.
'Los otros 4 intervalos se necesitan.
'TOTAL INTERVALOS 11

Option Explicit
Public actual As Long
Public MouseBoton As Long
Public MouseShift As Long

Public Cmsgautomatico As Boolean
Public Pmsgautomatico  As Boolean
Public rapidomsj As Boolean
Public cantidad As Integer
Public puedechupar As Boolean
Public FotoString As String
Dim endEvent As Long
'Dim PuedeMacrear As Boolean
Public rdbuffer As String
Private antx  As Integer
Private anty As Integer
Private proba As Integer

'Resoluciones
Public tamanioBarraVida As Integer
Public tamanioBarraMana As Integer
Public tamanioBarraEnergia As Integer
Public tamanioBarraSed As Integer
Public tamanioBarraHambre As Integer
Public cantidadColumnas As Integer
Public tamanioBarraExp As Integer

Public HechizoSeleccionado As Byte

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public diferenciaClickDobleClick  As Long
Public rompeIntervaloDobleClick As Long
Public diferenciaClickDobleClickNula As Long
Public uRechazadas As Long
Public tiempoClicHechiLanzar As Long
Public tiempoClicHechizoLanzar As Long


' Anticheat
Public cantidadClicHechiLanzarRapidos As Integer
Public cantidadClicHechiLanzarSuperRapido As Integer
Public cantidadClicHechizoLanzarRapidos As Integer
Public cantidadClicHechizoLanzarSuperRapido As Integer
Public cantidadClicksInnecesarios As Integer

Public hizoClicInnecesario As Boolean

Public UmbralClicHechiLanzarRapidos As Integer
Public UmbralClicHechiLanzarSuperRapido As Integer
Public UmbralClicHechizoLanzarRapidos As Integer
Public UmbralClicHechizoLanzarSuperRapido As Integer
Public UmbralAntiLanzar As Integer

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Sub ReleaseCapture Lib "user32" ()

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Public comienzoMinutoCheat As Long

Public anteriorIndexLista As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_TRANSPARENT = &H20&

Private InterfaceImagenFondo As Integer
Private InterfaceImagenHechizos As Integer
Private InterfaceImagenInventario As Integer

Private WithEvents ventanaMiniMapa As vwMinimap
Attribute ventanaMiniMapa.VB_VarHelpID = -1

Private ventanaMiniMapaMaximizada As Boolean

Private Sub SetInterface1280x720()


InterfaceImagenFondo = 113
InterfaceImagenHechizos = 126
InterfaceImagenInventario = 125

'EnergiaNumero
frmMain.label13.width = 122
frmMain.label13.Height = 12
frmMain.label13.top = 484
frmMain.label13.left = 795

'EnergiaBarra
frmMain.stashp.width = 122
frmMain.stashp.Height = 12
frmMain.stashp.top = 484
frmMain.stashp.left = 795
frmMain.tamanioBarraEnergia = 122

'Atajo Descansar
frmMain.Descansar.width = 33
frmMain.Descansar.Height = 33
frmMain.Descansar.top = 474
frmMain.Descansar.left = 760
frmMain.Descansar.MousePointer = 99
frmMain.Descansar.MouseIcon = Me.CmdLanzar.MouseIcon

'Vida
frmMain.label15.width = 122
frmMain.label15.Height = 12
frmMain.label15.top = 608
frmMain.label15.left = 795

'VidaBarra
frmMain.Hpshp.width = 122
frmMain.Hpshp.Height = 12
frmMain.Hpshp.top = 608
frmMain.Hpshp.left = 795
frmMain.tamanioBarraVida = 122

'Atajo Curar
frmMain.Curar.width = 33
frmMain.Curar.Height = 33
frmMain.Curar.top = 600
frmMain.Curar.left = 760
frmMain.Curar.MousePointer = 99
frmMain.Curar.MouseIcon = Me.CmdLanzar.MouseIcon

'Mana
frmMain.label14.width = 122
frmMain.label14.Height = 12
frmMain.label14.top = 544
frmMain.label14.left = 795

'ManaBarra
frmMain.ManShp.width = 122
frmMain.ManShp.Height = 12
frmMain.ManShp.top = 544
frmMain.ManShp.left = 795
frmMain.tamanioBarraMana = 122

'Atajo Meditar
frmMain.Meditar.width = 33
frmMain.Meditar.Height = 33
frmMain.Meditar.top = 536
frmMain.Meditar.left = 760
frmMain.Meditar.MousePointer = 99
frmMain.Meditar.MouseIcon = Me.CmdLanzar.MouseIcon

'Hambre
frmMain.label16.width = 86
frmMain.label16.Height = 12
frmMain.label16.top = 606
frmMain.label16.left = 993

'HambreBarra
frmMain.comidasp.width = 86
frmMain.comidasp.Height = 12
frmMain.comidasp.top = 606
frmMain.comidasp.left = 993
frmMain.tamanioBarraHambre = 86

'Sed
frmMain.label17.width = 86
frmMain.label17.Height = 12
frmMain.label17.top = 540
frmMain.label17.left = 993

'SedBarra
frmMain.Aguasp.width = 86
frmMain.Aguasp.Height = 12
frmMain.Aguasp.top = 540
frmMain.Aguasp.left = 993
frmMain.tamanioBarraSed = 86

'MiniMapa
frmMain.Minimapa.width = 154
frmMain.Minimapa.Height = 154
frmMain.Minimapa.top = 490
frmMain.Minimapa.left = 1118
frmMain.Minimapa.MousePointer = 99
frmMain.Minimapa.MouseIcon = Me.CmdLanzar.MouseIcon

'Oro
frmMain.GldLbl.width = 115
frmMain.GldLbl.Height = 18
frmMain.GldLbl.top = 480
frmMain.GldLbl.left = 976

'Inventario
frmMain.Label4.width = 155
frmMain.Label4.Height = 30
frmMain.Label4.top = 104
frmMain.Label4.left = 760

'Hechizos
frmMain.Label7.width = 163
frmMain.Label7.Height = 30
frmMain.Label7.top = 104
frmMain.Label7.left = 921

'Nivel
frmMain.LvlLbl.width = 22
frmMain.LvlLbl.Height = 13
frmMain.LvlLbl.top = 75
frmMain.LvlLbl.left = 776

'Nivel
frmMain.LvlLbl.width = 22
frmMain.LvlLbl.Height = 13
frmMain.LvlLbl.top = 75
frmMain.LvlLbl.left = 776

'TirarOro
frmMain.Image3(0).width = 25
frmMain.Image3(0).Height = 30
frmMain.Image3(0).top = 475
frmMain.Image3(0).left = 952
frmMain.Image3(0).MousePointer = 99
frmMain.Image3(0).MouseIcon = Me.CmdLanzar.MouseIcon

'NombreMapa
frmMain.Coord2.width = 150
frmMain.Coord2.Height = 12
frmMain.Coord2.top = 645
frmMain.Coord2.left = 1120

'Coordenadas
frmMain.Coord.width = 150
frmMain.Coord.Height = 12
frmMain.Coord.top = 665
frmMain.Coord.left = 1120

'Experiencia
frmMain.expshp.width = 273
frmMain.expshp.Height = 11
frmMain.expshp.top = 78
frmMain.expshp.left = 802
frmMain.tamanioBarraExp = 273

'Experiencia Numero
frmMain.NumExp.width = 273
frmMain.NumExp.Height = 12
frmMain.NumExp.top = 77
frmMain.NumExp.left = 802

'NombreClan
frmMain.lblClan.width = 130
frmMain.lblClan.Height = 13
frmMain.lblClan.top = 56
frmMain.lblClan.left = 855

'Nick
frmMain.lblUserName.width = 127
frmMain.lblUserName.Height = 18
frmMain.lblUserName.top = 32
frmMain.lblUserName.left = 857

'ChatClan
frmMain.Image1(4).width = 37
frmMain.Image1(4).Height = 33
frmMain.Image1(4).top = 144
frmMain.Image1(4).left = 1138
frmMain.Image1(4).ToolTipText = "Chat del Clan"
frmMain.Image1(4).MousePointer = 99
frmMain.Image1(4).MouseIcon = Me.CmdLanzar.MouseIcon

'Estadisticas
frmMain.Image1(1).width = 37
frmMain.Image1(1).Height = 33
frmMain.Image1(1).top = 198
frmMain.Image1(1).left = 1138
frmMain.Image1(1).ToolTipText = "Estadisticas"
frmMain.Image1(1).MousePointer = 99
frmMain.Image1(1).MouseIcon = Me.CmdLanzar.MouseIcon

'ChatParty
frmMain.Image1(3).width = 37
frmMain.Image1(3).Height = 33
frmMain.Image1(3).top = 144
frmMain.Image1(3).left = 1210
frmMain.Image1(3).ToolTipText = "Chat de Party"
frmMain.Image1(3).MousePointer = 99
frmMain.Image1(3).MouseIcon = Me.CmdLanzar.MouseIcon

'Opciones
frmMain.Image1(0).width = 37
frmMain.Image1(0).Height = 33
frmMain.Image1(0).top = 198
frmMain.Image1(0).left = 1210
frmMain.Image1(0).ToolTipText = "Opciones"
frmMain.Image1(0).MousePointer = 99
frmMain.Image1(0).MouseIcon = Me.CmdLanzar.MouseIcon

'Manual
frmMain.Image1(5).width = 37
frmMain.Image1(5).Height = 33
frmMain.Image1(5).top = 253
frmMain.Image1(5).left = 1139
frmMain.Image1(5).ToolTipText = "Manual"
frmMain.Image1(5).MousePointer = 99
frmMain.Image1(5).MouseIcon = Me.CmdLanzar.MouseIcon

'AdministrarClan
frmMain.Image1(2).width = 37
frmMain.Image1(2).Height = 33
frmMain.Image1(2).top = 308
frmMain.Image1(2).left = 1139
frmMain.Image1(2).ToolTipText = "Clan"
frmMain.Image1(2).MousePointer = 99
frmMain.Image1(2).MouseIcon = Me.CmdLanzar.MouseIcon

'AdministrarParty
frmMain.Image1(6).width = 37
frmMain.Image1(6).Height = 33
frmMain.Image1(6).top = 308
frmMain.Image1(6).left = 1210
frmMain.Image1(6).ToolTipText = "Party"
frmMain.Image1(6).MousePointer = 99
frmMain.Image1(6).MouseIcon = Me.CmdLanzar.MouseIcon

'FPS
frmMain.FPS.width = 25
frmMain.FPS.Height = 17
frmMain.FPS.top = 702
frmMain.FPS.left = 1080

'PING
frmMain.PING.width = 25
frmMain.PING.Height = 17
frmMain.PING.top = 702
frmMain.PING.left = 815

'CalcularPing
frmMain.CalcularPing.width = 50
frmMain.CalcularPing.Height = 17
frmMain.CalcularPing.top = 702
frmMain.CalcularPing.left = 755
frmMain.CalcularPing.MousePointer = 99
frmMain.CalcularPing.MouseIcon = Me.CmdLanzar.MouseIcon

'ConsolaHistorica
frmMain.imgConsola.width = 30
frmMain.imgConsola.Height = 20
frmMain.imgConsola.top = 155
frmMain.imgConsola.left = 8

'Inventario + Hechizos
frmMain.InvEqu.width = 328
frmMain.InvEqu.Height = 289
frmMain.InvEqu.top = 104
frmMain.InvEqu.left = 760

'SeguroDrop
frmMain.IconoDyd.width = 22
frmMain.IconoDyd.Height = 27
frmMain.IconoDyd.top = 340
frmMain.IconoDyd.left = 990

'SeguroPersonaje
frmMain.IconoSeg.width = 22
frmMain.IconoSeg.Height = 27
frmMain.IconoSeg.top = 340
frmMain.IconoSeg.left = 841

'Lanzar
frmMain.CmdLanzar.width = 108
frmMain.CmdLanzar.Height = 40
frmMain.CmdLanzar.top = 339
frmMain.CmdLanzar.left = 795

'Info
frmMain.cmdInfo.width = 65
frmMain.cmdInfo.Height = 40
frmMain.cmdInfo.top = 339
frmMain.cmdInfo.left = 985

'MoverHechizoAbajo
frmMain.cmdMoverHechi(1).width = 28
frmMain.cmdMoverHechi(1).Height = 35
frmMain.cmdMoverHechi(1).top = 250
frmMain.cmdMoverHechi(1).left = 1027

'MoverHechizoArriba
frmMain.cmdMoverHechi(0).width = 28
frmMain.cmdMoverHechi(0).Height = 35
frmMain.cmdMoverHechi(0).top = 195
frmMain.cmdMoverHechi(0).left = 1027

'Hechizos
frmMain.hlst.width = 221
frmMain.hlst.Height = 184
frmMain.hlst.top = 145
frmMain.hlst.left = 798

'Inventario
frmMain.picInv.width = 256
frmMain.picInv.Height = 160
frmMain.picInv.top = 159
frmMain.picInv.left = 793

Call Init_Inventario(256, 160, 4404, 8)

'Render
frmMain.Renderer.width = 672
frmMain.Renderer.Height = 672
frmMain.Renderer.top = 24
frmMain.Renderer.left = 50

'Minimizar
frmMain.Minimizar.width = 15
frmMain.Minimizar.Height = 15
frmMain.Minimizar.top = 0
frmMain.Minimizar.left = 1170

'Cerrar
frmMain.lblCerrar.width = 15
frmMain.lblCerrar.Height = 15
frmMain.lblCerrar.top = 0
frmMain.lblCerrar.left = 1205

'Indicador RMSG/GMSG
frmMain.Label10.width = 49
frmMain.Label10.Height = 17
frmMain.Label10.top = 675
frmMain.Label10.left = 52

'Estadisticas (BOTON +)
frmMain.Label1.width = 20
frmMain.Label1.Height = 19
frmMain.Label1.top = 56
frmMain.Label1.left = 1055

'Escribir GMSG
frmMain.SendGMSTXT.width = 650
frmMain.SendGMSTXT.Height = 20
frmMain.SendGMSTXT.top = 699
frmMain.SendGMSTXT.left = 47

'Escribir Normal
frmMain.SendTxt.width = 650
frmMain.SendTxt.Height = 20
frmMain.SendTxt.top = 702
frmMain.SendTxt.left = 47

frmMain.lblIndicadorEscritura.top = frmMain.SendTxt.top - 4
frmMain.lblIndicadorEscritura.left = frmMain.SendGMSTXT.left - 15

'Escribir RMSG
frmMain.SendRMSTXT.width = 650
frmMain.SendRMSTXT.Height = 20
frmMain.SendRMSTXT.top = 699
frmMain.SendRMSTXT.left = 47

'GRABANDO
frmMain.lGrabando.width = 137
frmMain.lGrabando.Height = 18
frmMain.lGrabando.top = 0
frmMain.lGrabando.left = 52

'Atajos GM
frmMain.Label11.width = 25
frmMain.Label11.Height = 17
frmMain.Label11.top = 703
frmMain.Label11.left = 20

'Enlaces
frmMain.lblLink.width = 417
frmMain.lblLink.Height = 13
frmMain.lblLink.top = 3
frmMain.lblLink.left = 745

'Barra de Arrastre
frmMain.Label3.width = 537
frmMain.Label3.Height = 20
frmMain.Label3.top = 0
frmMain.Label3.left = 104

'Facebook
frmMain.Image1(7).width = 26
frmMain.Image1(7).Height = 28
frmMain.Image1(7).top = 650
frmMain.Image1(7).left = 830
frmMain.Image1(7).MousePointer = 99
frmMain.Image1(7).MouseIcon = Me.CmdLanzar.MouseIcon

'Instagram
frmMain.Image1(8).width = 26
frmMain.Image1(8).Height = 28
frmMain.Image1(8).top = 650
frmMain.Image1(8).left = 862
frmMain.Image1(8).MousePointer = 99
frmMain.Image1(8).MouseIcon = Me.CmdLanzar.MouseIcon

'YouTube
frmMain.Image1(9).width = 26
frmMain.Image1(9).Height = 28
frmMain.Image1(9).top = 650
frmMain.Image1(9).left = 892
frmMain.Image1(9).MousePointer = 99
frmMain.Image1(9).MouseIcon = Me.CmdLanzar.MouseIcon

'Twitch
frmMain.Image1(10).width = 26
frmMain.Image1(10).Height = 28
frmMain.Image1(10).top = 650
frmMain.Image1(10).left = 954
frmMain.Image1(10).MousePointer = 99
frmMain.Image1(10).MouseIcon = Me.CmdLanzar.MouseIcon

'Discord
frmMain.Image1(11).width = 26
frmMain.Image1(11).Height = 28
frmMain.Image1(11).top = 650
frmMain.Image1(11).left = 986
frmMain.Image1(11).MousePointer = 99
frmMain.Image1(11).MouseIcon = Me.CmdLanzar.MouseIcon

'Wiki
frmMain.Image1(12).width = 26
frmMain.Image1(12).Height = 28
frmMain.Image1(12).top = 650
frmMain.Image1(12).left = 923
frmMain.Image1(12).MousePointer = 99
frmMain.Image1(12).MouseIcon = Me.CmdLanzar.MouseIcon

'Misiones
frmMain.Image1(13).width = 37
frmMain.Image1(13).Height = 33
frmMain.Image1(13).top = 252
frmMain.Image1(13).left = 1211
frmMain.Image1(13).ToolTipText = "Misiones"
frmMain.Image1(13).MousePointer = 99
frmMain.Image1(13).MouseIcon = Me.CmdLanzar.MouseIcon

'Mapa del Mundo
frmMain.Image1(14).width = 37
frmMain.Image1(14).Height = 33
frmMain.Image1(14).top = 364
frmMain.Image1(14).left = 1139
frmMain.Image1(14).ToolTipText = "Mapa"
frmMain.Image1(14).MousePointer = 99
frmMain.Image1(14).MouseIcon = Me.CmdLanzar.MouseIcon

'Eventos Activos
frmMain.Image1(15).width = 37
frmMain.Image1(15).Height = 33
frmMain.Image1(15).top = 364
frmMain.Image1(15).left = 1211
frmMain.Image1(15).ToolTipText = "Eventos"
frmMain.Image1(15).MousePointer = 99
frmMain.Image1(15).MouseIcon = Me.CmdLanzar.MouseIcon

'Boton de Panico
frmMain.Image1(16).width = 37
frmMain.Image1(16).Height = 33
frmMain.Image1(16).top = 422
frmMain.Image1(16).left = 1173
frmMain.Image1(16).ToolTipText = "Denuncia Rápida"
frmMain.Image1(16).MousePointer = 99
frmMain.Image1(16).MouseIcon = Me.CmdLanzar.MouseIcon

frmMain.Coord.Visible = True
frmMain.Coord2.Visible = True

End Sub

Private Sub setInterface1024x768() 'TO DO

InterfaceImagenFondo = 201
InterfaceImagenHechizos = 202
InterfaceImagenInventario = 203

frmMain.Minimapa.Visible = False

'EnergiaNumero
frmMain.label13.width = 80
frmMain.label13.Height = 12
frmMain.label13.top = 506
frmMain.label13.left = 732

'EnergiaBarra
frmMain.stashp.width = 80
frmMain.stashp.Height = 12
frmMain.stashp.top = 506
frmMain.stashp.left = 732
frmMain.tamanioBarraEnergia = 80

'Atajo Descansar
frmMain.Descansar.width = 33
frmMain.Descansar.Height = 33
frmMain.Descansar.top = 496
frmMain.Descansar.left = 698
frmMain.Descansar.MousePointer = 99
frmMain.Descansar.MouseIcon = Me.CmdLanzar.MouseIcon

'Vida
frmMain.label15.width = 80
frmMain.label15.Height = 12
frmMain.label15.top = 642
frmMain.label15.left = 732

'VidaBarra
frmMain.Hpshp.width = 80
frmMain.Hpshp.Height = 12
frmMain.Hpshp.top = 642
frmMain.Hpshp.left = 732
frmMain.tamanioBarraVida = 80

'Atajo Curar
frmMain.Curar.width = 33
frmMain.Curar.Height = 33
frmMain.Curar.top = 632
frmMain.Curar.left = 698
frmMain.Curar.MousePointer = 99
frmMain.Curar.MouseIcon = Me.CmdLanzar.MouseIcon

'Mana
frmMain.label14.width = 80
frmMain.label14.Height = 12
frmMain.label14.top = 571
frmMain.label14.left = 732

'ManaBarra
frmMain.ManShp.width = 80
frmMain.ManShp.Height = 12
frmMain.ManShp.top = 571
frmMain.ManShp.left = 732
frmMain.tamanioBarraMana = 80

'Atajo Meditar
frmMain.Meditar.width = 33
frmMain.Meditar.Height = 33
frmMain.Meditar.top = 560
frmMain.Meditar.left = 698
frmMain.Meditar.MousePointer = 99
frmMain.Meditar.MouseIcon = Me.CmdLanzar.MouseIcon

'Hambre
frmMain.label16.width = 61
frmMain.label16.Height = 12
frmMain.label16.top = 644
frmMain.label16.left = 879

'HambreBarra
frmMain.comidasp.width = 61
frmMain.comidasp.Height = 12
frmMain.comidasp.top = 644
frmMain.comidasp.left = 879
frmMain.tamanioBarraHambre = 61

'Sed
frmMain.label17.width = 61
frmMain.label17.Height = 12
frmMain.label17.top = 571
frmMain.label17.left = 879

'SedBarra
frmMain.Aguasp.width = 61
frmMain.Aguasp.Height = 12
frmMain.Aguasp.top = 572
frmMain.Aguasp.left = 879
frmMain.tamanioBarraSed = 61

'Oro
frmMain.GldLbl.width = 75
frmMain.GldLbl.Height = 17
frmMain.GldLbl.top = 504
frmMain.GldLbl.left = 864

'Inventario
frmMain.Label4.width = 120
frmMain.Label4.Height = 33
frmMain.Label4.top = 132
frmMain.Label4.left = 688

'Hechizos
frmMain.Label7.width = 120
frmMain.Label7.Height = 33
frmMain.Label7.top = 132
frmMain.Label7.left = 816

'Nivel
frmMain.LvlLbl.width = 22
frmMain.LvlLbl.Height = 13
frmMain.LvlLbl.top = 106
frmMain.LvlLbl.left = 704

'TirarOro
frmMain.Image3(0).width = 25
frmMain.Image3(0).Height = 36
frmMain.Image3(0).top = 490
frmMain.Image3(0).left = 840
frmMain.Image3(0).MousePointer = 99
frmMain.Image3(0).MouseIcon = Me.CmdLanzar.MouseIcon

'NombreMapa
frmMain.Coord2.width = 246
frmMain.Coord2.Height = 16
frmMain.Coord2.top = 734
frmMain.Coord2.left = 704

'Coordenadas
frmMain.Coord.width = 246
frmMain.Coord.Height = 16
frmMain.Coord.top = 734
frmMain.Coord.left = 704

'Experiencia
frmMain.expshp.width = 200
frmMain.expshp.Height = 12
frmMain.expshp.top = 108
frmMain.expshp.left = 730
frmMain.tamanioBarraExp = 200

'Experiencia numero
frmMain.NumExp.width = 200
frmMain.NumExp.Height = 12
frmMain.NumExp.top = 106
frmMain.NumExp.left = 730

'NombreClan
frmMain.lblClan.width = 130
frmMain.lblClan.Height = 13
frmMain.lblClan.top = 88
frmMain.lblClan.left = 752

'Nick
frmMain.lblUserName.width = 127
frmMain.lblUserName.Height = 18
frmMain.lblUserName.top = 64
frmMain.lblUserName.left = 752

'ChatClan
frmMain.Image1(4).width = 37
frmMain.Image1(4).Height = 33
frmMain.Image1(4).top = 144
frmMain.Image1(4).left = 968
frmMain.Image1(4).ToolTipText = "Chat del Clan"
frmMain.Image1(4).MousePointer = 99
frmMain.Image1(4).MouseIcon = Me.CmdLanzar.MouseIcon

'Estadisticas
frmMain.Image1(1).width = 37
frmMain.Image1(1).Height = 33
frmMain.Image1(1).top = 243
frmMain.Image1(1).left = 968
frmMain.Image1(1).ToolTipText = "Estadisticas"
frmMain.Image1(1).MousePointer = 99
frmMain.Image1(1).MouseIcon = Me.CmdLanzar.MouseIcon

'ChatParty
frmMain.Image1(3).width = 37
frmMain.Image1(3).Height = 33
frmMain.Image1(3).top = 192
frmMain.Image1(3).left = 968
frmMain.Image1(3).ToolTipText = "Chat de Party"
frmMain.Image1(3).MousePointer = 99
frmMain.Image1(3).MouseIcon = Me.CmdLanzar.MouseIcon

'Opciones
frmMain.Image1(0).width = 37
frmMain.Image1(0).Height = 33
frmMain.Image1(0).top = 295
frmMain.Image1(0).left = 968
frmMain.Image1(0).ToolTipText = "Opciones"
frmMain.Image1(0).MousePointer = 99
frmMain.Image1(0).MouseIcon = Me.CmdLanzar.MouseIcon

'Manual
frmMain.Image1(5).width = 37
frmMain.Image1(5).Height = 33
frmMain.Image1(5).top = 346
frmMain.Image1(5).left = 968
frmMain.Image1(5).ToolTipText = "Manual"
frmMain.Image1(5).MousePointer = 99
frmMain.Image1(5).MouseIcon = Me.CmdLanzar.MouseIcon

'AdministrarClan
frmMain.Image1(2).width = 37
frmMain.Image1(2).Height = 33
frmMain.Image1(2).top = 454
frmMain.Image1(2).left = 968
frmMain.Image1(2).ToolTipText = "Clan"
frmMain.Image1(2).MousePointer = 99
frmMain.Image1(2).MouseIcon = Me.CmdLanzar.MouseIcon

'AdministrarParty
frmMain.Image1(6).width = 37
frmMain.Image1(6).Height = 33
frmMain.Image1(6).top = 508
frmMain.Image1(6).left = 968
frmMain.Image1(6).ToolTipText = "Party"
frmMain.Image1(6).MousePointer = 99
frmMain.Image1(6).MouseIcon = Me.CmdLanzar.MouseIcon

'FPS
frmMain.FPS.width = 25
frmMain.FPS.Height = 17
frmMain.FPS.top = 12
frmMain.FPS.left = 744

'PING
frmMain.PING.width = 25
frmMain.PING.Height = 17
frmMain.PING.top = 12
frmMain.PING.left = 874

'CalcularPing
frmMain.CalcularPing.width = 50
frmMain.CalcularPing.Height = 17
frmMain.CalcularPing.top = 12
frmMain.CalcularPing.left = 816
frmMain.CalcularPing.MousePointer = 99
frmMain.CalcularPing.MouseIcon = Me.CmdLanzar.MouseIcon

'ConsolaHistorica
frmMain.imgConsola.width = 30
frmMain.imgConsola.Height = 20
frmMain.imgConsola.top = 733
frmMain.imgConsola.left = 625
frmMain.imgConsola.MousePointer = 99
frmMain.imgConsola.MouseIcon = Me.CmdLanzar.MouseIcon

'  Mapa
frmMain.imgBotonMapa.width = 30
frmMain.imgBotonMapa.Height = 20
frmMain.imgBotonMapa.top = 730
frmMain.imgBotonMapa.left = 973
frmMain.imgBotonMapa.MousePointer = 99
frmMain.imgBotonMapa.MouseIcon = Me.CmdLanzar.MouseIcon
frmMain.imgBotonMapa.Visible = True

'Inventario + Hechizos
frmMain.InvEqu.width = 256
frmMain.InvEqu.Height = 280
frmMain.InvEqu.top = 130
frmMain.InvEqu.left = 688

'SeguroDrop
frmMain.IconoDyd.width = 22
frmMain.IconoDyd.Height = 27
frmMain.IconoDyd.top = 376
frmMain.IconoDyd.left = 868

'SeguroPersonaje
frmMain.IconoSeg.width = 22
frmMain.IconoSeg.Height = 27
frmMain.IconoSeg.top = 376
frmMain.IconoSeg.left = 744

'Lanzar
frmMain.CmdLanzar.width = 108
frmMain.CmdLanzar.Height = 40
frmMain.CmdLanzar.top = 364
frmMain.CmdLanzar.left = 715

'Info
frmMain.cmdInfo.width = 65
frmMain.cmdInfo.Height = 40
frmMain.cmdInfo.top = 364
frmMain.cmdInfo.left = 858

'MoverHechizoAbajo
frmMain.cmdMoverHechi(1).width = 28
frmMain.cmdMoverHechi(1).Height = 35
frmMain.cmdMoverHechi(1).top = 260
frmMain.cmdMoverHechi(1).left = 893

'MoverHechizoArriba
frmMain.cmdMoverHechi(0).width = 28
frmMain.cmdMoverHechi(0).Height = 34
frmMain.cmdMoverHechi(0).top = 202
frmMain.cmdMoverHechi(0).left = 893

'Hechizos
frmMain.hlst.width = 172
frmMain.hlst.Height = 184
frmMain.hlst.top = 172
frmMain.hlst.left = 717

'Inventario
frmMain.picInv.width = 192
frmMain.picInv.Height = 192
frmMain.picInv.top = 170
frmMain.picInv.left = 718

Call Init_Inventario(192, 192, 22308, 6)

'Render
frmMain.Renderer.width = 672
frmMain.Renderer.Height = 672
frmMain.Renderer.top = 46
frmMain.Renderer.left = 4

'Minimizar
frmMain.Minimizar.width = 28
frmMain.Minimizar.Height = 23
frmMain.Minimizar.top = 8
frmMain.Minimizar.left = 952

'Cerrar
frmMain.lblCerrar.width = 28
frmMain.lblCerrar.Height = 23
frmMain.lblCerrar.top = 8
frmMain.lblCerrar.left = 990

'Indicador RMSG/GMSG
frmMain.Label10.width = 49
frmMain.Label10.Height = 17
frmMain.Label10.top = 736
frmMain.Label10.left = 656

'Estadisticas (BOTON +)
frmMain.Label1.width = 20
frmMain.Label1.Height = 15
frmMain.Label1.top = 88
frmMain.Label1.left = 908

'Escribir GMSG
frmMain.SendGMSTXT.width = 633
frmMain.SendGMSTXT.Height = 20
frmMain.SendGMSTXT.top = 734
frmMain.SendGMSTXT.left = 24

'Escribir Normal
frmMain.SendTxt.width = 600
frmMain.SendTxt.Height = 20
frmMain.SendTxt.top = 734
frmMain.SendTxt.left = 24

frmMain.lblIndicadorEscritura.top = frmMain.SendTxt.top - 5
frmMain.lblIndicadorEscritura.left = frmMain.SendGMSTXT.left - 15

'Escribir RMSG
frmMain.SendRMSTXT.width = 633
frmMain.SendRMSTXT.Height = 20
frmMain.SendRMSTXT.top = 734
frmMain.SendRMSTXT.left = 24

'GRABANDO
frmMain.lGrabando.width = 137
frmMain.lGrabando.Height = 18
frmMain.lGrabando.top = 416
frmMain.lGrabando.left = 752

'Atajos GM
frmMain.Label11.width = 25
frmMain.Label11.Height = 17
frmMain.Label11.top = 736
frmMain.Label11.left = 0

'Enlaces
frmMain.lblLink.width = 609
frmMain.lblLink.Height = 13
frmMain.lblLink.top = 736
frmMain.lblLink.left = 40

'Barra de Arrastre
frmMain.Label3.width = 600
frmMain.Label3.Height = 44
frmMain.Label3.top = 0
frmMain.Label3.left = 32

'Facebook
frmMain.Image1(7).width = 26
frmMain.Image1(7).Height = 26
frmMain.Image1(7).top = 676
frmMain.Image1(7).left = 732
frmMain.Image1(7).MousePointer = 99
frmMain.Image1(7).MouseIcon = Me.CmdLanzar.MouseIcon

'Instagram
frmMain.Image1(8).width = 26
frmMain.Image1(8).Height = 26
frmMain.Image1(8).top = 676
frmMain.Image1(8).left = 760
frmMain.Image1(8).MousePointer = 99
frmMain.Image1(8).MouseIcon = Me.CmdLanzar.MouseIcon

'YouTube
frmMain.Image1(9).width = 26
frmMain.Image1(9).Height = 26
frmMain.Image1(9).top = 676
frmMain.Image1(9).left = 788
frmMain.Image1(9).MousePointer = 99
frmMain.Image1(9).MouseIcon = Me.CmdLanzar.MouseIcon

'Twitch
frmMain.Image1(10).width = 26
frmMain.Image1(10).Height = 26
frmMain.Image1(10).top = 676
frmMain.Image1(10).left = 846
frmMain.Image1(10).MousePointer = 99
frmMain.Image1(10).MouseIcon = Me.CmdLanzar.MouseIcon

'Discord
frmMain.Image1(11).width = 26
frmMain.Image1(11).Height = 26
frmMain.Image1(11).top = 676
frmMain.Image1(11).left = 876
frmMain.Image1(11).MousePointer = 99
frmMain.Image1(11).MouseIcon = Me.CmdLanzar.MouseIcon

'Wiki
frmMain.Image1(12).width = 26
frmMain.Image1(12).Height = 26
frmMain.Image1(12).top = 676
frmMain.Image1(12).left = 816
frmMain.Image1(12).MousePointer = 99
frmMain.Image1(12).MouseIcon = Me.CmdLanzar.MouseIcon

'Misiones
frmMain.Image1(13).width = 37
frmMain.Image1(13).Height = 33
frmMain.Image1(13).top = 400
frmMain.Image1(13).left = 968
frmMain.Image1(13).ToolTipText = "Misiones"
frmMain.Image1(13).MousePointer = 99
frmMain.Image1(13).MouseIcon = Me.CmdLanzar.MouseIcon

'Mapa del Mundo
frmMain.Image1(14).width = 37
frmMain.Image1(14).Height = 33
frmMain.Image1(14).top = 563
frmMain.Image1(14).left = 968
frmMain.Image1(14).ToolTipText = "Mapa"
frmMain.Image1(14).MousePointer = 99
frmMain.Image1(14).MouseIcon = Me.CmdLanzar.MouseIcon

'Eventos Activos
frmMain.Image1(15).width = 37
frmMain.Image1(15).Height = 33
frmMain.Image1(15).top = 617
frmMain.Image1(15).left = 968
frmMain.Image1(15).ToolTipText = "Eventos"
frmMain.Image1(15).MousePointer = 99
frmMain.Image1(15).MouseIcon = Me.CmdLanzar.MouseIcon

'Boton de Panico
frmMain.Image1(16).width = 37
frmMain.Image1(16).Height = 33
frmMain.Image1(16).top = 671
frmMain.Image1(16).left = 968
frmMain.Image1(16).ToolTipText = "Denuncia Rápida"
frmMain.Image1(16).MousePointer = 99
frmMain.Image1(16).MouseIcon = Me.CmdLanzar.MouseIcon

End Sub


Private Sub CalcularPing_Click()
    frmMain.PING = "Cargando"
    Call sSendData(1, Complejo.PING)
    PingPerformanceTimer.Time
End Sub

Private Sub cmdInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call procesarTeclaPresionada(Button + 1000)
End Sub

Private Sub CmdLanzar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If frmMain.MousePointer = 2 Then Exit Sub

If hlst.ListIndex <> -1 Then
    If hlst.list(hlst.ListIndex) = "Remover paralisis" Then
        If antx = X And anty = Y Then
        proba = proba + 1
        End If
        antx = X
        anty = Y
    End If
End If


End Sub

Private Sub CmdLanzar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Hace clic en hechizos y en lanzar en menos de 300?
Debug.Print GetTickCount - tiempoClicHechiLanzar

Dim diferencia As Long

diferencia = GetTickCount - tiempoClicHechiLanzar

If diferencia < UmbralClicHechiLanzarSuperRapido Then
    cantidadClicHechiLanzarSuperRapido = cantidadClicHechiLanzarSuperRapido + 1
ElseIf diferencia < UmbralClicHechiLanzarRapidos Then
    cantidadClicHechiLanzarRapidos = cantidadClicHechiLanzarRapidos + 1
End If

diferencia = GetTickCount - tiempoClicHechizoLanzar

If diferencia < 50 Then
    cantidadClicHechizoLanzarSuperRapido = cantidadClicHechizoLanzarSuperRapido + 1
ElseIf diferencia < 200 Then
    cantidadClicHechizoLanzarRapidos = cantidadClicHechizoLanzarRapidos + 1
End If

If hizoClicInnecesario Then
    cantidadClicksInnecesarios = cantidadClicksInnecesarios + 1
    hizoClicInnecesario = False
End If
'Mando HechiLanzar < 300 | HechiLanzar < 100 | HechizoLanzar < 200 | HechizoLanzar < 50 | ClicksInnecesarios
End Sub

Private Sub coord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Engine_Resolution.resolucionActual = RESOLUCION_43 Then
        frmMain.Coord2.Visible = True
        frmMain.Coord.Visible = False
    End If
End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)
If hlst.ListIndex = -1 Then Exit Sub

Select Case Index
Case 0 'subir
    If hlst.ListIndex = 0 Then Exit Sub
Case 1 'bajar
    If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
End Select

EnviarPaquete Paquetes.Moverhechi, Chr(Index + 1) & Chr(hlst.ListIndex + 1)

Select Case Index
Case 0 'subir
    hlst.ListIndex = hlst.ListIndex - 1
Case 1 'bajar
    hlst.ListIndex = hlst.ListIndex + 1
End Select
End Sub

Private Sub Coord2_Click()
If Coord2.tag <> "1" Then
Coord2.tag = "1"
Else
Coord2.tag = "0"
End If
End Sub

Private Sub Curar_Click()
    ProcesarComando ("/CURAR")
End Sub

Private Sub Descansar_Click()
    ProcesarComando ("/DESCANSAR")
End Sub

Private Sub Form_Click()
    If frmMain.SendTxt.Visible Then
        frmMain.SendTxt.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Dim i As Byte

If GetAsyncKeyState(KeyCode) <> -32767 Then
    DesabilitarTecla(KeyCode) = True
End If

If Shift = 1 Then DeAmuchos = True

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
GUI_Keypress KeyAscii
End Sub

Private Sub Form_LostFocus()
    Debug.Print "foco perdido"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call procesarTeclaPresionada(Button + 1000)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseBoton = Button
MouseShift = Shift
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun = True Then
    '    prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub hlst_Click()
Call CrearAccion(Chr(sPaquetes.accion) & 3 & hlst.ListIndex)

If hlst.ListIndex = anteriorIndexLista Then
    hizoClicInnecesario = True
End If

anteriorIndexLista = hlst.ListIndex

If frmMain.SendTxt.Visible Then
    frmMain.SendTxt.SetFocus
End If

End Sub

Private Sub hlst_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tiempoClicHechizoLanzar = GetTickCount
    Call procesarTeclaPresionada(Button + 1000)
End Sub

Private Sub IconoDyd_Click()
    ProcesarComando ("/DRAG")
End Sub

Private Sub IconoSeg_Click()
    ProcesarComando ("/SEG")
End Sub

Private Sub imgBotonMapa_Click()
    Call Minimapa_Click
End Sub

Private Sub Meditar_Click()
    ProcesarComando ("/MEDITAR")
End Sub

Private Sub imgConsola_Click()

If frmConsola.Visible = False Then
    Call CLI_Consola.CargarConsola
    Load frmConsola
    frmConsola.Show
Else
    Unload frmConsola
End If

End Sub

Private Sub Label18_Click()
If frmMain.MousePointer = 2 Then
If UserStats(SlotStats).UsingSkill = Magia Then
    UserStats(SlotStats).UsingSkill = 0
    frmMain.MousePointer = 0
    AddtoRichTextBox frmConsola.ConsolaFlotante, "Estás muy lejos para lanzar este hechizo.", 255, 0, 0, True, False, False
ElseIf UserStats(SlotStats).UsingSkill = Proyectiles Then
    UserStats(SlotStats).UsingSkill = 0
    frmMain.MousePointer = 0
    AddtoRichTextBox frmConsola.ConsolaFlotante, "Estás muy lejos para disparar.", 255, 0, 0, True, False, False
ElseIf UserStats(SlotStats).UsingSkill = Robar Then
    UserStats(SlotStats).UsingSkill = 0
    frmMain.MousePointer = 0
    AddtoRichTextBox frmConsola.ConsolaFlotante, "No puedes robar a esta distancia.", 255, 0, 0, True, False, False
Else
    UserStats(SlotStats).UsingSkill = 0
    frmMain.MousePointer = 0
End If
End If
End Sub
Private Sub Coord_Click()
    AddtoRichTextBox frmConsola.ConsolaFlotante, "Estas coordenadas son tu ubicaciÃ³n en el mapa. Utiliza la letra L para corregirla si esta no se corresponde con la del servidor por efecto del Lag.", 255, 255, 255, False, False, False
End Sub

Private Sub InvEqu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call procesarTeclaPresionada(Button + 1000)
End Sub

Private Sub lblCerrar_Click()
    prgRun = False
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngReturnValue As Long
    
    If Button = 1 And Not forzarFullScreen Then
            Call ReleaseCapture
            lngReturnValue = SendMessage(frmMain.hWnd, WM_NCLBUTTONDOWN, _
                                         HTCAPTION, 0&)
        End If
         
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call procesarTeclaPresionada(Button + 1000)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call procesarTeclaPresionada(Button + 1000)
End Sub

Private Sub lblLink_Click()
    Call CLI_CurrentLInk.clickLink
End Sub

Private Sub LvlLbl_Click()
If UserPasarNivel > 0 Then
    Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Nivel: " & UserLvl & " Experiencia: " & FormatNumber(UserExp, 0, vbFalse, vbFalse, vbTrue) & "/" & FormatNumber(UserPasarNivel, 0, vbFalse, vbFalse, vbTrue) & " (" & FormatNumber((CDbl(UserExp) * 100 / (UserPasarNivel)), 0) & "%)", 0, 200, 200, False, False, False)
End If
End Sub


Private Sub Minimapa_Click()

If ventanaMiniMapa Is Nothing Then
    Set ventanaMiniMapa = New vwMinimap
    
    ventanaMiniMapa.vWindow_SetPos MainViewWidth, MainViewHeight

    ventanaMiniMapa.SetMapa UserMap
    ventanaMiniMapa.SetMaximizado ventanaMiniMapaMaximizada
    GUI_Load ventanaMiniMapa
    ventanaMiniMapa.vWindow_Show
Else
    GUI_Quitar ventanaMiniMapa
    Set ventanaMiniMapa = Nothing
End If

End Sub

Private Sub Minimizar_Click()
Me.WindowState = 1
End Sub

Private Sub pasarMinuto_Timer()
Dim minutoReal As Long

'Anti speed hack del cheat engine
minutoReal = timeGetTime - comienzoMinutoCheat

If minutoReal > 65000 Then '65000 es menos de 1.1 * 60000 milesimas
    EnviarPaquete Lachiteo, "2" & LongToString(minutoReal)
End If

comienzoMinutoCheat = timeGetTime

cantidadClicHechiLanzarRapidos = 0

Call CLI_CurrentLInk.pasarTiempo

End Sub

Private Sub Pasarsegundo_Timer()
'/**********************************************
'**************** CBAY 2.0 *********************
'************ ANTI ACELERADORES ****************
'***********************************************
Static vecees As Byte
Static TimerB As Single


If UserPasos > 6 Then
    MsgBox "Has sido expulsado del juego por uso de aceleradores. Tipo 76", vbCritical, "cBay 2.0"
    End
End If

UserPasos = 0


If timer - TimerB < 0.9 Then
    If vecees > 10 Then
    MsgBox "Has sido expulsado del juego por uso de aceleradores. PRo 1", vbCritical, "cBay 2.0"
    prgRun = False
    Else
    vecees = vecees + 1
    End If
End If
TimerB = timer



'****************************************************
  
'\**********************************************
'**************** CBAY 2.0 *********************
'***********************************************

If Hour(Time) = 0 And Second(Time) = 0 Then
Puedeatacar = timer
NoPuedeChuparYuSarClick = timer
NoPuedeChuparYuSarU = timer
    UserPuedeRefrescar = timer
End If

frmMain.FPS = Engine.FPS


        
If uRechazadas > 80 Or diferenciaClickDobleClickNula > 40 Or rompeIntervaloDobleClick > 50 Then
    EnviarPaquete LaChiteo2, uRechazadas & "-" & diferenciaClickDobleClickNula & "-" & rompeIntervaloDobleClick
    uRechazadas = 0
    diferenciaClickDobleClickNula = 0
    rompeIntervaloDobleClick = 0
End If

If cantidadClicHechiLanzarRapidos + cantidadClicHechiLanzarSuperRapido * 2 + cantidadClicHechizoLanzarRapidos + cantidadClicHechizoLanzarSuperRapido * 2 > 5 Then
    EnviarPaquete Paquetes.Lachiteo, "1" & ByteToString(cantidadClicHechiLanzarRapidos) & ByteToString(cantidadClicHechiLanzarSuperRapido) & ByteToString(cantidadClicHechizoLanzarRapidos) & ByteToString(cantidadClicHechizoLanzarSuperRapido) & ByteToString(cantidadClicksInnecesarios)

    cantidadClicHechiLanzarRapidos = 0
    cantidadClicHechiLanzarSuperRapido = 0
    cantidadClicHechizoLanzarRapidos = 0
    cantidadClicHechizoLanzarSuperRapido = 0
    cantidadClicksInnecesarios = 0
            End If
        
If Grabando Then
    If SegVideo = 59 Then
        SegVideo = 0
        MinutoVideo = MinutoVideo + 1
    Else
        SegVideo = SegVideo + 1
    End If
    lGrabando.Caption = "Grabando video (" & IIf(MinutoVideo < 10, 0 & MinutoVideo, MinutoVideo) & ":" & IIf(SegVideo < 10, 0 & SegVideo, SegVideo) & ")"
End If

Call Consola.PassTimer
Call Consola_Clan.PassTimer

End Sub


Private Sub picInv_Click()
    diferenciaClickDobleClick = GetTickCount
        
    If frmMain.SendTxt.Visible Then
        frmMain.SendTxt.SetFocus
    End If
End Sub

Private Sub picInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errr
    If GetAsyncKeyState(Button) <> -32767 Then
        DesabilitarTecla(Button) = DesabilitarTecla(Button) + 2
    End If
    
    If itemElegido > 0 And itemElegido <= MAX_INVENTORY_SLOTS Then
    If Button = 2 And UserStats(SlotStats).UserEstado = 0 And UserInventory(itemElegido).GrhIndex > 0 Then
    Call CambiarCursor(frmMain, 1)
    MousePress = 1: ItemDragued = itemElegido
    
    Set itemimg.Picture = clsEnpaquetado_LeerIPicture(pakGraficos, (GrhData(UserInventory(itemElegido).GrhIndex).filenum))
    
    End If
    End If
    
    Call procesarTeclaPresionada(Button + 1000)
errr:
End Sub

Private Sub Renderer_DblClick()
If GUI_MouseDown(MouseBoton, MouseShift, MouseX, MouseY) = False Then
    If GUI_MouseUp(MouseBoton, MouseShift, MouseX, MouseY) = False Then
        If tx > 0 And ty > 0 Then EnviarPaquete Paquetes.ClickAccion, Chr(tx) & Chr(ty)
    End If
End If
End Sub

Private Sub Renderer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If GUI_MouseDown(Button, Shift, X, Y) = False Then
    If GetAsyncKeyState(Button) <> -32767 Then
        DesabilitarTecla(Button) = DesabilitarTecla(Button) + 2
    End If
    MouseBoton = Button
    MouseShift = Shift
End If

If frmMain.SendTxt.Visible Then
    SendTxt.SetFocus
Else
    Call procesarTeclaPresionada(Button + 1000)
End If
End Sub

Private Sub Renderer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If GUI_MouseMove(Button, Shift, X, Y) = False Then
    MouseX = X
    MouseY = Y
   ' Call ConvertCPtoTP(x, y, MouseTileX, MouseTileY)
End If
End Sub

Private Sub Socket1_Accept(SocketId As Integer)
    Debug.Print "Socket accept"
End Sub

Private Sub Socket1_Blocking(status As Integer, Cancel As Integer)
    Debug.Print "Socket Blocking"
End Sub

Private Sub Socket1_Cancel(status As Integer, response As Integer)
    Debug.Print "Socket cancel"
End Sub

Private Sub Socket1_Timeout(status As Integer, response As Integer)
    Debug.Print "Socket time out"
End Sub

Private Sub Renderer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If GUI_MouseUp(Button, Shift, X, Y) = False Then
    MouseBoton = Button
    MouseShift = Shift
End If
End Sub

Private Sub Socket1_Timer()
    Debug.Print "Socket timer"
End Sub

Private Sub SoundFX_Timer()
'on error GoTo HayError
Dim n As Integer

If RandomNumber(1, 150) < 12 Then
    
Select Case Terreno
    
Case "BOSQUE"
    n = RandomNumber(1, 100)
            Select Case Zona
                Case "CAMPO", "CIUDAD"
                    If Not bRain And Not bSnow Then
                        If n < 30 And n >= 15 Then
                            Call TocarMusica(21)
                        ElseIf n < 30 And n < 15 Then
                            Call TocarMusica(22)
                        ElseIf n >= 30 And n <= 35 Then
                            Call TocarMusica(28)
                        ElseIf n >= 35 And n <= 40 Then
                            Call TocarMusica(29)
                        ElseIf n >= 40 And n <= 45 Then
                            Call TocarMusica(34)
                        End If
                    End If
            End Select
End Select

End If


Exit Sub
hayError:
End Sub

Public Sub tNoche_Timer()
'If NocheAlpha < 140 Then NocheAlpha = NocheAlpha + 1
   On Error GoTo tNoche_Timer_Error

If DecryptStr(IntervaloPegarB, 0) <> UserStats(SlotStats).IntervaloPegar Then
    End
ElseIf DecryptStr(IntervaloNoChupClickB, 0) <> UserStats(SlotStats).IntervaloNoChupClick Then
    End
ElseIf DecryptStr(intervaloNoChupUB, 0) <> UserStats(SlotStats).intervaloNoChupU Then
    End
ElseIf DecryptStr(IntervaloLanzarFlechasB, 0) <> UserStats(SlotStats).IntervaloLanzarFlechas Then
    End
ElseIf DecryptStr(IntervaloLanzarMagiasB, 0) <> UserStats(SlotStats).IntervaloLanzarMagias Then
    End
End If
 
If DecryptStr(IntervaloSolapaLanzarB, 0) <> UserStats(SlotStats).IntervaloSolapaLanzar Then
    End
ElseIf DecryptStr(IntervaloSolapaLanzarSuperB, 0) <> UserStats(SlotStats).IntervaloSolapaLanzarSuper Then
    End
ElseIf DecryptStr(IntervaloHechizoLanzarB, 0) <> UserStats(SlotStats).IntervaloHechizoLanzar Then
    End
ElseIf DecryptStr(IntervaloHechizoLanzarSuperB, 0) <> UserStats(SlotStats).IntervaloHechizoLanzarSuper Then
    End
ElseIf DecryptStr(UmbralAlertaB, 0) <> UserStats(SlotStats).UmbralAlerta Then
    End
End If

Exit Sub
tNoche_Timer_Error:
End Sub
''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
If UserMeditar Then Exit Sub

If (itemElegido > 0 And itemElegido < MAX_INVENTORY_SLOTS + 1) Then
    If UserInventory(itemElegido).Amount = 1 Then
       EnviarPaquete Tirar, Chr(itemElegido) & Chr(1)
    Else
       If UserInventory(itemElegido).Amount > 1 Then
        frmCantidad.Show , frmMain
       End If
    End If
End If
End Sub

Private Sub AgarrarItem()
    EnviarPaquete Paquetes.Agarrar
End Sub

Private Function ttos(i As Single) As String
    ttos = LongToString(Int(i)) & Chr$(val((i - Fix(i)) * 100))
End Function

Private Function obtenerItem() As Single
    Dim i As Byte
    
    For i = 1 To Int(Rnd * 4)
        obtenerItem = timer
    Next
    obtenerItem = timer
End Function

Private Sub UsarItem()


 If itemElegido = 0 Then Exit Sub
        If UserInventory(itemElegido).Name = "Martillo de Herrero" Then
        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Haz click sobre el yunque...", 100, 100, 120, 0, 0)
        UserStats(SlotStats).UsingSkill = 16
        frmMain.MousePointer = 2
        Else
            If (itemElegido > 0) And (itemElegido < MAX_INVENTORY_SLOTS + 1) Then
                If InStr(1, UserInventory(itemElegido).Name, "Arco") > 0 Then
                    If timer - Puedeatacar > UserStats(SlotStats).IntervaloLanzarFlechas And UserInventory(itemElegido).Equipped = 1 Then
                    UserStats(SlotStats).UsingSkill = Proyectiles
                    frmMain.MousePointer = 2
                    Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
                     ElseIf UserInventory(itemElegido).Equipped = 0 Then
                    Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Debes equipar el arco para poder usarlo.", 65, 190, 156, 0, 0)
                    End If
                Else
                     If timer - NoPuedeChuparYuSarClick > UserStats(SlotStats).IntervaloNoChupClick Then
                        NoPuedeChuparYuSarClick = timer
                        EnviarPaquete Paquetes.Usar, Chr$(itemElegido)
                     Else
                       ' ANTICHEATTT
                        rompeIntervaloDobleClick = rompeIntervaloDobleClick + 1
                     End If
                End If
            End If
        End If
End Sub

Private Sub EquiparItem()
    If (itemElegido > 0) And (itemElegido < MAX_INVENTORY_SLOTS + 1) Then _
        EnviarPaquete Paquetes.Equipar, Chr(itemElegido)
End Sub

Private Sub cmdLanzar_Click()

'If DesabilitarTecla(MouseBoton) Then
'DesabilitarTecla(MouseBoton) = False: Exit Sub 'Anticheat
'End If

If frmMain.SendTxt.Visible Then
    frmMain.SendTxt.SetFocus
End If

If Not IScombate Then AddtoRichTextBox frmConsola.ConsolaFlotante, "¡¡No puedes lanzar hechizos si no estas en modo combate!!", 65, 190, 156, 0, 0: Exit Sub

' Chequeamos por las dudas que editen
If hlst.Visible = False Or hlst.Enabled = False Then Exit Sub

If UserStats(SlotStats).UserEstado = 0 Then
  If hlst.ListIndex < 0 Then AddtoRichTextBox frmConsola.ConsolaFlotante, "¡Debes seleccionar un hechizo!!", 65, 190, 156, 0, 0: Exit Sub
  If hlst.list(hlst.ListIndex) <> " (Vacio)" And (timer - Puedeatacar > UserStats(SlotStats).IntervaloLanzarMagias) Then
  HechizoSeleccionado = hlst.ListIndex + 1
'  EnviarPaquete Paquetes.LanzarHechizo, Chr()
  UserStats(SlotStats).UsingSkill = Magia
  frmMain.MousePointer = 2
  AddtoRichTextBox frmConsola.ConsolaFlotante, "Haz click sobre el objetivo...", 100, 100, 120, 0, 0
  End If
Else
    AddtoRichTextBox frmConsola.ConsolaFlotante, "¡¡Estás muerto!!", 65, 190, 156, 0, 0
End If
        
End Sub
Private Sub CmdInfo_Click()
    If hlst.ListIndex >= 0 Then
    If hlst = " (Vacio)" Then Exit Sub
    EnviarPaquete Paquetes.InfoHechizo, Chr(hlst.ListIndex + 1)
    End If
End Sub

Public Sub SetMapa(numero As Integer, Nombre As String)
    frmMain.Coord.Caption = NombreMapa
    frmMain.Minimapa.Stretch = True
    Call DameImagen(frmMain.Minimapa, BASE_INTERFACE_MAPAS + UserMap)
    
    If Not ventanaMiniMapa Is Nothing Then
        ventanaMiniMapa.SetMapa (numero)
    End If
End Sub

Private Sub Renderer_Click()
 
If DesabilitarTecla(MouseBoton) > 0 Then

    DesabilitarTecla(MouseBoton) = DesabilitarTecla(MouseBoton) - 1

    If proba > 0 Then
        EnviarPaquete Lachiteo, "0"
        proba = 0
    End If
End If

If MouseBoton >= 2 Then Exit Sub

Dim tiempo As Single

tiempo = obtenerItem

With UserStats(SlotStats)
    If Cartel Then Cartel = False
    If Not Comerciando And Not Bovedeando Then
    

       Call ConvertCPtoTP(MouseX, MouseY, 32, 32, tx, ty)
       
      ' Debug.Print tx, ty
       
        If ty > Y_MAXIMO_VISIBLE Or tx > X_MAXIMO_VISIBLE Then Exit Sub
                
        'MARCE TODO CUAL ES LA CONSTANTE QUE TE DICE EL ANCHO Y ALTO DE LA PANTALLA?
        If tx > UserPos.X + BORDE_TILES_INUTILIZABLE Or tx < UserPos.X - BORDE_TILES_INUTILIZABLE Or ty > UserPos.Y + BORDE_TILES_INUTILIZABLE Or ty < UserPos.Y - BORDE_TILES_INUTILIZABLE Then Exit Sub
        
        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then

                If .UsingSkill = 0 Then
                    EnviarPaquete Paquetes.ClickIzquierdo, ITS(tx) & ITS(ty)
                Else
                    Call CambiarCursor(frmMain, 0)
                    If .UsingSkill = Proyectiles Then
                        If IScombate = False Then
                            AddtoRichTextBox frmConsola.ConsolaFlotante, "¡¡No puedes lanzar flechas si no estas en modo combate!!", 65, 190, 156, 0, 0: .UsingSkill = 0: Exit Sub
                        ElseIf UserMeditar = True Then
                            AddtoRichTextBox frmConsola.ConsolaFlotante, "¡¡No puedes lanzar flechas si estas meditando!!", 65, 190, 156, 0, 0: .UsingSkill = 0: Exit Sub
                        Else
                            If tiempo - Puedeatacar > .IntervaloLanzarFlechas Then
                                Puedeatacar = tiempo
                                EnviarPaquete Paquetes.ClickSkill, ITS(tx) & ITS(ty) & Chr(.UsingSkill) & ttos(tiempo)
                            End If
                        End If
                    ElseIf .UsingSkill = Magia Then
                        If (tiempo - Puedeatacar > .IntervaloLanzarMagias) Then
                            Puedeatacar = tiempo
                            EnviarPaquete Paquetes.ClickSkill, ITS(tx) & ITS(ty) & Chr(.UsingSkill) & Chr(HechizoSeleccionado) & ttos(tiempo)
                        End If
                   Else
                        Puedeatacar = tiempo
                        EnviarPaquete Paquetes.ClickSkill, ITS(tx) & ITS(ty) & Chr(.UsingSkill) & ttos(tiempo)
                   End If
                   ' If UsingSkill = Magia Or UsingSkill = proyectiles Then UserCanAttack = 0
                   ' Debug.Print "USE EL PUTO SKILL"
                    .UsingSkill = 0
                End If
            End If
        ElseIf (MouseShift And 1) = 1 Then
            If MouseShift = 1 And MouseBoton = 1 Then
                Call sSendData(Paquetes.ComandosConse, Conse2.TELEP, "YO " & UserMap & " " & Int(tx) & " " & Int(ty))
            End If
        End If
    End If
End With


End Sub

Private Sub procesarTeclaPresionada(KeyCode As Integer)

Dim tiempo As Single


If UserPrivilegios > 0 Then
    Select Case KeyCode
        Case vbKeyG:
            If frmMain.Label11.Visible = True Then
                frmMain.Label10.Caption = "GMSG"
                If SendTxt.Visible Then Exit Sub
                
                If Not frmCantidad.Visible Then
                    SendGMSTXT.Visible = True
                    SendGMSTXT.SetFocus
                End If
            End If
        Case vbKeyK:
            If frmMain.Label11.Visible = True Then
            frmMain.Label10.Caption = "RMSG"
            
            If SendTxt.Visible Then Exit Sub
            
            If Not frmCantidad.Visible Then
                SendRMSTXT.Visible = True
                SendRMSTXT.SetFocus
                End If
            End If
        Case vbKeyI:
            If frmMain.Label11.Visible = True Then
            Call sSendData(Paquetes.ComandosConse, Conse1.CINVISIBLE)
            End If
        Case vbKeyW: '/TRABAJANDO
            If frmMain.Label11.Visible = True Then
            Call sSendData(Paquetes.ComandosConse, Conse1.TRABAJANDO)
            End If
        Case vbKeyP:
            If frmMain.Label11.Visible = True Then
            If UserPrivilegios = 0 Then Exit Sub
            frmPanelGm.Show , frmMain
            End If
    End Select
End If

Select Case KeyCode

    Case vbKeyConsolaClanes:
    
        If Consola_Clan.Activo Then
            Consola_Clan.Activo = False
            Consola_Clan.RemoveDialogs
            AddtoRichTextBox frmConsola.ConsolaFlotante, "Consola flotante de clanes desactivada.", 255, 200, 200, False, False, False
        Else
            Consola_Clan.Activo = True
            AddtoRichTextBox frmConsola.ConsolaFlotante, "Consola flotante de clanes activada.", 255, 200, 200, False, False, False
        End If

    Case vbKeyMusica:
        Call CLI_Audio.toogleMusica
    Case vbKeyAgarrarItem:
        Call AgarrarItem
    Case vbKeyModoCombate:
            If Istrabajando Then
                AddtoRichTextBox frmConsola.ConsolaFlotante, "No puedes combatir y trabajar al mismo tiempo.", 65, 190, 156, False, False, False
                Exit Sub
            Else
                IScombate = Not IScombate
                EnviarPaquete Paquetes.MCombate
                If IScombate Then
                AddtoRichTextBox frmConsola.ConsolaFlotante, "Has pasado al modo combate.", 65, 190, 156, False, False, False
                Else
                AddtoRichTextBox frmConsola.ConsolaFlotante, "Has salido del modo combate.", 65, 190, 156, False, False, False
                End If
                If UserStats(SlotStats).UsingSkill = Magia Then UserStats(SlotStats).UsingSkill = 0: frmMain.MousePointer = 0
            End If
            '[Wizard]
    Case vbKeyEquiparItem:
        Call EquiparItem
    Case vbKeyMostrarNombre:
    Nombres = Not Nombres
Case vbKeyDomar
        If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡No puedes domar una criatura si Estás muerto!!", 65, 190, 156, 0): Exit Sub
        If (timer - Puedeatacar >= UserStats(SlotStats).IntervaloPegar) And _
       (Not UserDescansar) And _
       (Not UserMeditar) Then
        UserStats(SlotStats).UsingSkill = 18
        frmMain.MousePointer = 2
        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Haz click sobre la criatura...", 100, 100, 120, 0, 0)
        Puedeatacar = timer
        End If
Case vbKeyOcultar:
    If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡Estas muerto!", 100, 100, 120, 0, 0): Exit Sub
    
    EnviarPaquete Paquetes.SkillSetOcultar
Case vbKeyTirarItem:
    Call TirarItem
Case vbKeyUsar:
    
    If itemElegido = 0 Then Exit Sub
    If itemElegido = FLAGORO Or itemElegido = 254 Then Exit Sub
    If UserInventory(itemElegido).Name = "Martillo de Herrero" Then
        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Haz click sobre el yunque...", 100, 100, 120, 0, 0)
        UserStats(SlotStats).UsingSkill = 16
        frmMain.MousePointer = 2
    Else
        tiempo = obtenerItem()

        If (itemElegido > 0) And (itemElegido < MAX_INVENTORY_SLOTS + 1) Then
            If InStr(1, UserInventory(itemElegido).Name, "Arco") > 0 Then
                If tiempo - Puedeatacar > UserStats(SlotStats).IntervaloLanzarFlechas And UserInventory(itemElegido).Equipped = 1 Then
                UserStats(SlotStats).UsingSkill = Proyectiles
                frmMain.MousePointer = 2
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
                ElseIf UserInventory(itemElegido).Equipped = 0 Then
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Debes equipar el arco para poder usarlo.", 65, 190, 156, 0, 0)
                End If
            Else
                If tiempo - NoPuedeChuparYuSarU > UserStats(SlotStats).intervaloNoChupU Then
                    NoPuedeChuparYuSarU = tiempo
                    If profileClicks Then
                        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "2", 65, 190, 156): Exit Sub
                    End If
                    EnviarPaquete Paquetes.Usar, Chr(itemElegido) & 1 & ttos(tiempo)
                End If
            End If
        End If
    End If
    '///////////////////////////////
Case vbKeyLag:
        If timer - UserPuedeRefrescar > 1 Then
            sSendData Paquetes.UNLAG
            UserPuedeRefrescar = timer
            Beep
        End If
Case vbKeyQ:
        If frmMain.Label11.Visible = True Then
            Call sSendData(Paquetes.ComandosConse, Conse1.SHOW_SOS)
        Else
            VerMapa = False
        End If
Case vbKeyPegar:
    tiempo = obtenerItem
    If (tiempo - Puedeatacar >= UserStats(SlotStats).IntervaloPegar) And _
       (Not UserDescansar) And _
       (Not UserMeditar) Then
            EnviarPaquete Paquetes.Pegar, ttos(tiempo)
            Debug.Print tiempo
            Puedeatacar = tiempo
    End If
Case vbKeyMeditar:
    ProcesarComando (" MEDITAR")
    Exit Sub
End Select
End Sub

Private Sub mostrarTextConsola()
    Me.lblLink.Visible = False

    Me.SendTxt.Visible = True
    
    Me.lblIndicadorEscritura.Visible = True

    frmMain.SendTxt.SetFocus
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'on error Resume
If DesabilitarTecla(KeyCode) Then
    DesabilitarTecla(KeyCode) = False
    uRechazadas = uRechazadas + 1
End If

If (Not SendTxt.Visible) And (Not SendTxt.Visible) And (Not SendGMSTXT.Visible) And (Not SendRMSTXT.Visible) Then
    Call procesarTeclaPresionada(KeyCode)
End If

Select Case KeyCode
    Case vbKeyReturn:
        If SendTxt.Visible Then Exit Sub
        If SendGMSTXT.Visible Then Exit Sub
        If SendRMSTXT.Visible Then Exit Sub
        If Not frmCantidad.Visible Then
            Call mostrarTextConsola
        End If
    Case vbKeyDelete:
    If SendTxt.Visible Then Exit Sub
    If Cmsgautomatico = True Then
        Cmsgautomatico = False
        Pmsgautomatico = False
        'Image1(4).Picture = Nothing
    Else
        If (CharList(UserCharIndex).flags And ePersonajeFlags.tieneClan) = 0 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡No perteneces a ningÃºn clan!", 65, 190, 156, 0): Exit Sub
        rapidomsj = True
        Cmsgautomatico = True
        Pmsgautomatico = False
        'Image1(3).Picture = Nothing
        'DameImagen Image1(4), 122
        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Todo lo que digas sera escuchado por tu clan. ", 0, 200, 200, False, False, False)
    End If
    Exit Sub
'*****************************************************************************
    Case vbKeyF1:
        
        If UserPrivilegios = 0 Then Exit Sub
        Me.lblLink.Visible = False
        If frmMain.Label11.Visible = True Then
            frmMain.Label11.Visible = False 'show sos y Trabajando
        Else
            frmMain.Label11.Visible = True 'show sos y Trabajando
        End If
        Exit Sub
'*****************************************************************************
    Case vbKeyPageDown
        If Pmsgautomatico = True Then
        Cmsgautomatico = False
        Pmsgautomatico = False
        Image1(3).Picture = Nothing
        Else
            If Not gh And Not Liderparty Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "No integras ninguna party. ", 0, 200, 200, False, False, False): Exit Sub
            Cmsgautomatico = False
            Pmsgautomatico = True
            rapidomsj = True
            'Image1(4).Picture = Nothing
            'DameImagen Image1(3), 121
            Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Lo que digas sera escuchado por los integrantes de la party. ", 0, 200, 200, False, False, False)
        End If
        Exit Sub
'*****************************************************************************
    Case vbKeyF3:
        EnviarPaquete Paquetes.iParty, ""
        Exit Sub
'*****************************************************************************
    Case vbKeyF4:
        ProcesarComando ("/SALIR")
        Exit Sub
'*****************************************************************************
    Case vbKeyF5:
        Call Retos.Show(vbModeless, frmMain)
        Exit Sub
'*****************************************************************************

'*****************************************************************************
    Case vbKeyF7:
        EnviarPaquete Paquetes.ccParty
        Call Partym.Show(vbModeless, frmMain)
        Exit Sub
'*****************************************************************************
     Case vbKeyF8:
        If Not Istrabajando Then
            If IScombate = True Then
            AddtoRichTextBox frmConsola.ConsolaFlotante, "No puedes trabajar en modo combate.", 255, 0, 0, True, False, False
            Else
                If UserStats(SlotStats).UserEstado = 1 Then Exit Sub
                EnviarPaquete Paquetes.MTrabajar, Chr(itemElegido)
            End If
        Else
        EnviarPaquete Paquetes.DejadeLaburar, ""
        Istrabajando = False
        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Has terminado de trabajar.", 0, 200, 200, False, False, False)
        End If
        Exit Sub
'*****************************************************************************
    Case vbKeyF9:
        Call frmOpciones.Show(vbModeless, frmMain)
        Exit Sub
'*****************************************************************************
    Case vbKeyF10
     If UserLvl > 13 Then
        If FotoDenuncia = 1 Then
            If (timer - FotoDenunciasTiempo) < 60 Then
            Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Haz alcanzado el mÃ¡ximo de envio de 1 FotoDenuncia por minuto. EsperÃ¡ unos instantes y volve a intentar.", 0, 200, 200, False, False, False)
            Else
                Call sSendData(Paquetes.EFotoDenuncia, 0, FotoString)
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "FotoDenuncia enviada correctamente.", 0, 200, 200, False, False, False)
                FotoDenunciasTiempo = timer
                FotoDenuncia = 0
            End If
        Else
            If (timer - FotoDenunciasTiempo) < 60 Then
            Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Haz alcanzado el mÃ¡ximo de envio de 1 FotoDenuncia por minuto. EsperÃ¡ unos instantes y volve a intentar.", 0, 200, 200, False, False, False)
            Else
            FotoString = GenerarFotoDenuncia
            If Not FotoString = "" Then
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "La FotoDenuncia fue sacada correctamente.", 0, 200, 200, False, False, False)
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Vuelva a presionar F10 para enviar la foto, ESC para cancelar.", 0, 200, 200, False, False, False)
                FotoDenuncia = 1
            Else
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Nadie te esta insultando. Las FotoDenuncias solo sirven para denunciar agravios.", 0, 200, 200, False, False, False)
            End If
            End If
        End If
      Else
            Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Debes ser nivel 14 o superior para poder usar este comando.", 0, 200, 200, False, False, False)
      End If
        Exit Sub
'*****************************************************************************
    Case vbKeyF11

            'If Not Grabando Then
            '        Call IniciarGrabacion
            'Else
            '        Call FinalizarGrabacion
            'End If
'*****************************************************************************
    Case vbKeyF12:
        Call CapturarPantalla
        Exit Sub
'*****************************************************************************
     Case vbKeyEnd
        ProcesarComando (" MEDITAR")
        Exit Sub
'*****************************************************************************
    Case vbKeyShift
        DeAmuchos = False
        
    Case vbKeyNumlock
        If Shift = 1 Then Exit Sub
        
        If MovimientoDefault = E_Heading.None Then
            If GetTimer - LastKeyPressTime < 200 Then
                MovimientoDefault = LastKeyPress
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Te mantienes caminando. Pulsa BLOQNUM para dejar de caminar.", 0, 200, 200, False, False, False)
            End If
        Else
            MovimientoDefault = E_Heading.None
        End If
'*****************************************************************************
    Case vbKeyEscape
        If FotoDenuncia = 1 Then
            FotoDenuncia = 0
            FotoString = ""
            Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "La FotoDenuncia fue cancelada.", 0, 200, 200, False, False, False)
            Exit Sub
        End If
'*****************************************************************************
End Select
End Sub

Private Sub Form_Load()
    itemimg.Transparent = True
        
    Me.width = Engine_Resolution.pixelesAncho * Screen.TwipsPerPixelX
    Me.Height = Engine_Resolution.pixelesAlto * Screen.TwipsPerPixelY
    
    Select Case Engine_Resolution.resolucionActual
        Case RESOLUCION_43
            Call setInterface1024x768
        Case RESOLUCION_169
            Call SetInterface1280x720
    End Select
    
    ' Mouse Pointer
    Me.lblLink.MousePointer = 99
    Me.lblLink.MouseIcon = Me.CmdLanzar.MouseIcon
    
    Me.Label7.MousePointer = 99
    Me.Label7.MouseIcon = Me.CmdLanzar.MouseIcon
    
    Me.Label4.MousePointer = 99
    Me.Label4.MouseIcon = Me.CmdLanzar.MouseIcon
    
    Me.imgConsola.MousePointer = 99
    Me.imgConsola.MouseIcon = Me.CmdLanzar.MouseIcon
    
    Me.cmdInfo.MousePointer = 99
    Me.cmdInfo.MouseIcon = Me.CmdLanzar.MouseIcon
    
    Me.cmdMoverHechi(0).MousePointer = 99
    Me.cmdMoverHechi(0).MouseIcon = Me.CmdLanzar.MouseIcon
    Me.cmdMoverHechi(1).MousePointer = 99
    Me.cmdMoverHechi(1).MouseIcon = Me.CmdLanzar.MouseIcon
    
    Me.Image1(5).MousePointer = 99
    Me.Image1(5).MouseIcon = Me.CmdLanzar.MouseIcon
    Me.Image1(5).ToolTipText = "Manual"
    
    Me.LvlLbl.MousePointer = 99
    Me.LvlLbl.MouseIcon = Me.CmdLanzar.MouseIcon
    
    Me.Label1.MousePointer = 99
    Me.Label1.MouseIcon = Me.CmdLanzar.MouseIcon
    
    Me.SendTxt.BackColor = RGB(48, 25, 17)
    
    Me.PING = "-"
    
    ' Imagenes de la interface
    DameImagenForm Me, InterfaceImagenFondo
    DameImagen InvEqu, InterfaceImagenInventario

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
    
    If Engine_Resolution.resolucionActual = RESOLUCION_43 Then
        If Coord2.tag <> "1" Then
            frmMain.Coord2.Visible = False
            frmMain.Coord.Visible = True
        Else
            frmMain.Coord2.Visible = True
            frmMain.Coord.Visible = False
        End If
        If LvlLbl.tag = "1" Then
            LvlLbl.tag = "0"
            LvlLbl.Caption = UserLvl
        End If
    Else
        frmMain.Coord.Visible = True
        frmMain.Coord2.Visible = True
    End If
    
    dibujar_tooltip_inv = 0
    inv_tooltip_counter = 0
                
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

Private Sub Image1_Click(Index As Integer)
    Sonido_Play (SND_CLICK)
    Select Case Index
        Case 0
            Call frmOpciones.Show(vbModeless, frmMain)
            frmOpciones.top = Me.top + (215 * Screen.TwipsPerPixelX)
            frmOpciones.left = Me.left + (251 * Screen.TwipsPerPixelY)
        Case 1
            If meves Then Exit Sub
            
            meves = True
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            EnviarPaquete Paquetes.CallForAtributos
            EnviarPaquete Paquetes.CallForSkill
            EnviarPaquete Paquetes.MEsteM
            EnviarPaquete Paquetes.FEST
            EnviarPaquete Paquetes.CallForFama
            
            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
        Case 2
                EnviarPaquete Paquetes.GuildInfo
        Case 3
            If Pmsgautomatico = True Then
            Cmsgautomatico = False
            Pmsgautomatico = False
            Image1(3).Picture = Nothing
            Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Dejas de ser escuchado por tu party.", 0, 200, 200, False, False, False)
            Else
            
            If Not gh And Not Liderparty Then
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "No integras ninguna party. ", 0, 200, 200, False, False, False)
                Exit Sub
            End If
            
            Cmsgautomatico = False
            Pmsgautomatico = True
            Image1(4).Picture = Nothing
            Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Todo lo que digas sera escuchado por tu party.", 0, 200, 200, False, False, False)
        End If
        
        Case 4
        
            If Cmsgautomatico = True Then
                Cmsgautomatico = False
                Pmsgautomatico = False
                Image1(4).Picture = Nothing
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Dejas de ser escuchado por tu clan. ", 0, 200, 200, False, False, False)
            Else
                If (CharList(UserCharIndex).flags And ePersonajeFlags.tieneClan) = 0 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡No perteneces a ningun clan!!", 65, 190, 156, 0): Exit Sub
                
                Cmsgautomatico = True
                Pmsgautomatico = False
                Image1(3).Picture = Nothing
                    
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Todo lo que digas sera escuchado por tu clan. ", 0, 200, 200, False, False, False)
            End If
        Case 5
            Call openUrl("https://wiki.tierrasdelsur.cc")
        Case 6
            EnviarPaquete Paquetes.ccParty
            Call Partym.Show(vbModeless, frmMain)
        Case 7
            Call openUrl("https://www.facebook.com/argentum.tds")
        Case 8
            Call openUrl("https://www.instagram.com/argentum.tds/")
        Case 9
            Call openUrl("https://www.youtube.com/channel/UCrvzB1ynJcxWcHKbeaoujKg")
        Case 10
            Call openUrl("https://www.twitch.tv/argentumtds")
        Case 11
            Call openUrl("https://discord.tierrasdelsur.cc/")
        Case 12
            Call openUrl("https://wiki.tierrasdelsur.cc/")
        Case 13
            Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "No tienes misiones asignadas.", 0, 200, 200, False, False, False)
        Case 14
            If frmMapa.Visible = False Then
                frmMapa.Visible = True
            Else
                frmMapa.Visible = False
            End If
        Case 15
            ProcesarComando ("/EVENTO")
        Case 16
            ProcesarComando ("/Denunciar Necesito Ayuda")
        
    End Select
End Sub

Private Sub Image3_Click(Index As Integer)
    Select Case Index
        Case 0
            itemElegido = FLAGORO
            If UserGLD > 0 Then
                frmCantidad.Show , frmMain
            End If
            'GUI_Load New Ventana_Tirar_Oro
    End Select
End Sub

Private Sub Label1_Click()

    If Not frmEstadisticas.Visible Then
    Call Image1_Click(1)
    End If
End Sub

Private Sub Label4_Click()
    
    Call Sonido_Play(SND_CLICK)
        
    If CmdLanzar.Visible Then
        DameImagen InvEqu, InterfaceImagenInventario
        
        ' Muestro el inventario
        picInv.Visible = True
        picInv.Enabled = True
      
        ' Iconos
        IconoSeg.Visible = True
        IconoSeg.Enabled = True
        IconoDyd.Visible = True
        IconoDyd.Enabled = True
        
        ' Oculto hechizos
        ' - Lista
        hlst.Visible = False
        hlst.Enabled = False
        
        cmdInfo.Visible = False
        cmdInfo.Enabled = False
        
        ' - Botones
        CmdLanzar.Visible = False
        CmdLanzar.Enabled = False
        
        ' - Mover hechizos
        cmdMoverHechi(0).Enabled = False
        cmdMoverHechi(1).Enabled = False
        cmdMoverHechi(0).Visible = False
        cmdMoverHechi(1).Visible = False
               
        Call CrearAccion(Chr(sPaquetes.accion) & 1)
        
        hizoClicInnecesario = False
    Else
        hizoClicInnecesario = True
    End If
    
End Sub

Private Sub Label7_Click()
    Call Sonido_Play(SND_CLICK)
    
    If picInv.Visible Then
        DameImagen InvEqu, InterfaceImagenHechizos

        ' Ocultamos el inventario
        picInv.Visible = False
        picInv.Enabled = False
        
        IconoSeg.Visible = False
        IconoSeg.Enabled = False
        IconoDyd.Visible = False
        IconoDyd.Enabled = False
        
        ' Mostramos hechizos
        ' - Lista
        hlst.Visible = True
        hlst.Enabled = True
        
        ' - Botones
        cmdInfo.Visible = True
        cmdInfo.Enabled = True
        CmdLanzar.Visible = True
        CmdLanzar.Enabled = True
        
        ' - Mover hechizos
        cmdMoverHechi(0).Enabled = True
        cmdMoverHechi(1).Enabled = True
        cmdMoverHechi(0).Visible = True
        cmdMoverHechi(1).Visible = True
            
        Call CrearAccion(Chr(sPaquetes.accion) & 2)
            hizoClicInnecesario = False
        Else
            hizoClicInnecesario = True
        End If
        
        tiempoClicHechiLanzar = GetTickCount()
End Sub
 
Private Sub picInv_DblClick()
    
Dim T As Single

T = obtenerItem

' Selecciono algo?
If itemElegido = 0 Then Exit Sub

' Revisamos por las dudas
If picInv.Enabled = False Or picInv.Visible = False Then Exit Sub
    
    If GetTickCount - diferenciaClickDobleClick < 16 Then
        diferenciaClickDobleClickNula = diferenciaClickDobleClickNula + 1
    End If
    
    If DesabilitarTecla(1) > 0 Then
        DesabilitarTecla(1) = DesabilitarTecla(1) - 2
        Exit Sub
    End If
    
If Not puedechupar Then Exit Sub

If UserInventory(itemElegido).Name = "Martillo de Herrero" Then

    Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Haz click sobre el yunque...", 100, 100, 120, 0, 0)
    UserStats(SlotStats).UsingSkill = 16
    frmMain.MousePointer = 2
    
Else
    If (itemElegido > 0) And (itemElegido < MAX_INVENTORY_SLOTS + 1) Then
    
        If InStr(1, UserInventory(itemElegido).Name, "Arco") > 0 Then
            
            If T - Puedeatacar > UserStats(SlotStats).IntervaloLanzarFlechas And UserInventory(itemElegido).Equipped = 1 Then
                UserStats(SlotStats).UsingSkill = Proyectiles
                frmMain.MousePointer = 2
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
            ElseIf UserInventory(itemElegido).Equipped = 0 Then
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Debes equipar el arco para poder usarlo.", 65, 190, 156, 0, 0)
            Else
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Indefinido", 65, 190, 156, 0, 0)
            End If
        
        Else
        
             If T - NoPuedeChuparYuSarClick > UserStats(SlotStats).IntervaloNoChupClick Then
                NoPuedeChuparYuSarClick = T
                If profileClicks Then
                    Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "3", 65, 190, 156): Exit Sub
                End If
                EnviarPaquete Paquetes.Usar, Chr$(itemElegido) & 0 & ttos(T)
             Else
                rompeIntervaloDobleClick = rompeIntervaloDobleClick + 1
             End If

        End If
    End If
End If
        
End Sub

Private Sub picInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If MousePress = 1 Then
        MousePressX = picInv.left + X - Renderer.left
        MousePressY = picInv.top + Y - Renderer.top
        
        Call ConvertPixelDragAndDrop(X, Y, MousePressPosX, MousePressPosY)
        If (Not MousePressPosX = -1 And Not MousePressPosY = -1) Then
            'Esta dentro de surface??
            itemimg.Visible = False
        Else
            If Not itemimg.Visible Then
                Set itemimg.Picture = clsEnpaquetado_LeerIPicture(pakGraficos, GrhData(UserInventory(itemElegido).GrhIndex).filenum)
                itemimg.Visible = True
            End If
            
            Call itemimg.Move(MousePressX + Renderer.left - 16, MousePressY + Renderer.top - 16)
        End If
    End If
    
    Call CLI_DibujarInventario.MouseMove(X, Y)
End Sub


Private Sub ConvertPixelDragAndDrop(pixelX As Single, pixelY As Single, ByRef tx As Integer, ByRef ty As Integer)
    Dim tempX As Integer
    Dim tempY As Integer
        
    tempX = picInv.left + pixelX
    tempY = picInv.top + pixelY
    
    If tempX >= Me.Renderer.left And tempX <= Me.Renderer.left + Me.Renderer.width And tempY >= Me.Renderer.top And tempY <= Me.Renderer.top + Me.Renderer.Height Then
        Call ConvertCPtoTP(tempX - Me.Renderer.left, tempY - Me.Renderer.top, 32, 32, tx, ty)
    Else
        tx = -1
        ty = -1
    End If
End Sub
Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button <> 2 Then
        puedechupar = True
        Sonido_Play SND_CLICK
        
        If ItemDragued > 0 Then
            ItemDragued = 0
            MousePress = 0
            itemimg.Visible = False
            If CursorPer = 1 Then
                frmMain.MouseIcon = LoadResPicture(101, vbResCursor)
                Exit Sub
            End If
        End If
        
        Call ItemClick(CInt(X), CInt(Y), picInv)
    'anti boton derecho para chupar..
    Else
        puedechupar = False
    End If

    If MousePress = 1 And Button = 2 And UserStats(SlotStats).UserEstado = 0 Then
    
        Call ConvertPixelDragAndDrop(X, Y, tx, ty)
        
        If Not tx = -1 And Not ty = -1 Then
                
           If CursorPer = 1 Then frmMain.MouseIcon = LoadResPicture(101, vbResCursor)
            If DeAmuchos Then
                frmCantidad.Show , frmMain
            Else
                EnviarPaquete Paquetes.DIClick, Chr$(tx) & Chr$(ty) & Chr$(ItemDragued) & ITS(1)
                ItemDragued = 0
            End If
            
            itemimg.Visible = False
            MousePress = 0
            
        Else 'Arroja algo sobre el inventario
            If CursorPer = 1 Then frmMain.MouseIcon = LoadResPicture(101, vbResCursor)
                MousePress = 0
                itemimg.Visible = False
                If ItemDragued > 0 Then 'Tiene seleccionado un item y lo quiere tirar sobre el inv?
                Call ItemClick(CInt(X), CInt(Y), picInv)
                    If itemElegido = ItemDragued Then 'Si es el mismo slot es al pedo!
                        ItemDragued = 0
                    Else
                        EnviarPaquete ChangeItemsSlot, Chr$(itemElegido) & Chr$(ItemDragued)
                        ItemDragued = 0
                    End If
                End If
        End If
    End If

End Sub

'Private Sub frmConsola.RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyEscape Then
'        If frmConsola.RecTxt.Visible Then
'            frmConsola.RecTxt.Visible = False
'        End If
'    End If
'End Sub

Private Sub SendTxt_Change()
    If Len(SendTxt.text) > 160 Then
        stxtbuffer = " "
    Else
        stxtbuffer = SendTxt.text
    End If
 End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    Dim key As Integer
    
    key = Asc(UCase(Chr$(KeyAscii)))
        
    If IScombate Then
        If (key = vbKeyNorte Or key = vbKeySur Or key = vbKeyEste Or key = vbKeyOeste Or key = vbKeyUsar) And Trim$(SendTxt.text) = "" Then
            KeyAscii = 0
        End If
    End If
    
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If left$(stxtbuffer, 1) = "/" Then
            ProcesarComando stxtbuffer
        'Shout
        ElseIf left$(stxtbuffer, 1) = "-" Then
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡Estas muerto!", 100, 100, 120, 0, 0): Exit Sub
            EnviarPaquete Paquetes.Gritar, right$(stxtbuffer, Len(stxtbuffer) - 1)

        'Whisper
        ElseIf left$(stxtbuffer, 1) = "\" Then
            stxtbuffer = mid$(stxtbuffer, 2)
            If InStr(1, stxtbuffer, " ") > 0 Then
             If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡Estas muerto!", 100, 100, 120, 0, 0): Exit Sub
            EnviarPaquete Paquetes.Susurrar, ReadField(1, stxtbuffer, Asc(" ")) & "," & right$(stxtbuffer, Len(stxtbuffer) - (Len(ReadField(1, stxtbuffer, Asc(" "))) + 1))
            'If EstaPcAreaName(ReadField(1, stxtbuffer, Asc(" "))) Then
             '   Dialogos.CrearDialogo right$(stxtbuffer, Len(stxtbuffer) - (Len(ReadField(1, stxtbuffer, Asc(" "))) + 1)), UserCharIndex, RGB(0, 0, 255)
            'End If
            End If
        ElseIf left$(stxtbuffer, 1) = "+" Then
            EnviarPaquete Paquetes.FaccionMsg, stxtbuffer
        ElseIf Cmsgautomatico = True And LTrim(stxtbuffer) <> "" And stxtbuffer <> "-" Then
            ProcesarComando (" CMSG " & stxtbuffer)
            If rapidomsj = True Then Call DesactivarCMSG
        ElseIf Pmsgautomatico = True And LTrim(stxtbuffer) <> "" And stxtbuffer <> "-" Then
            ProcesarComando "/PMSG " & stxtbuffer
            If rapidomsj = True Then Call DesactivarPMSG
        ElseIf stxtbuffer = "DEBUGCLIC" Then
            profileClicks = True
        ElseIf stxtbuffer <> "" Then
            EnviarPaquete Paquetes.Hablar, stxtbuffer
        End If

        stxtbuffer = vbNullString
        SendTxt.text = vbNullString
        KeyCode = 0
        SendTxt.Visible = False
        Me.lblIndicadorEscritura.Visible = False
    End If
End Sub
'/GMSG
Private Sub SendGMSTXT_Change()
    stxtbuffergmsg = SendGMSTXT.text
End Sub

Private Sub SendGMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbuffergmsg <> "" Then
            Call sSendData(Paquetes.ComandosConse, SemiDios2.GMSG, stxtbuffergmsg)
        End If
        frmMain.Label10 = ""
        stxtbuffergmsg = ""
        SendGMSTXT.text = ""
        KeyCode = 0
        Me.SendGMSTXT.Visible = False
    End If
End Sub

Private Sub SendGMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

'/RMSG
Private Sub SendRMSTXT_Change()
    stxtbufferrmsg = SendRMSTXT.text
End Sub

Private Sub SendRMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbufferrmsg <> "" Then
            Call sSendData(Paquetes.ComandosConse, Conse2.RMSG, stxtbufferrmsg)
        End If
        frmMain.Label10 = ""
        stxtbufferrmsg = ""
        SendRMSTXT.text = ""
        KeyCode = 0
        Me.SendRMSTXT.Visible = False
    End If
End Sub

Private Sub SendRMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub
Private Sub Socket1_Connect()
    Debug.Print "Conectamos socket normal"
    Call enviarHash
End Sub

Private Sub limpiarForm()

    ' Si estaba filmando.. lo guardo
    If Grabando Then FinalizarGrabacion
    
    ' Si estaba boveando o comerciando, oculto los formularios
    If Bovedeando = True Then
        Bovedeando = False
        Unload frmBancoObj
    End If
    
    If Comerciando = True Then
        Comerciando = False
        Unload frmComerciarUsu
    End If
    
    ' Desactivo los efectos y el antipiquete
    SoundFX.Enabled = False
    pasarMinuto.Enabled = False
    
    ' Dialogos en Pantalla
    Dialogos.RemoveAllDialogs
    rdbuffer = ""
    
    Call Consola.RemoveDialogs
    Call Consola_Clan.RemoveDialogs
    
    frmMain.Label11.Visible = False
    
    ' Formulario de Game Masters
    frmMain.Label11.Visible = False
    
    frmMain.SoundFX.Enabled = False
    
    If Not ventanaMiniMapa Is Nothing Then
        GUI_Quitar ventanaMiniMapa
        Set ventanaMiniMapa = Nothing
    End If
    
    ' Oculto el resto de los formularis
    Dim i As Integer
    On Local Error Resume Next
    For i = 0 To Forms.count - 1
        If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name Then
            Unload Forms(i)
        End If
    Next i
    On Local Error GoTo 0

    Me.PING = "-"
End Sub


Private Sub Socket1_Disconnect()
    Debug.Print "Desconectamos socket"
    
    EstadoConexion = E_Estado.Ninguno
    
    ' Si me estaba conectando...
    If EstadoLogin = E_MODO.PantallaCreacion Or EstadoLogin = E_MODO.IngresarPersonaje Then
        ' Me fallo esta IP. La marco como fallida.
        If Not TCP.recibiPaquete Then
            recibiPaquete = False
            Call modLogin.agregarIpNoAccesible(Socket1.HostAddress)
            Socket1.Cleanup
            'frmConnect.conectar
        End If
        Exit Sub
    End If
    
    Call limpiarForm
    
    ' Desconectado
    Connected = False

    ' Limpiamos socket
    Socket1.Cleanup
    
    frmConnect.MousePointer = 99

    If EstadoLogin = PantallaCreacion Or EstadoLogin = CreandoPersonaje Or EstadoLogin = CrearPersonajeSeteado Then
        Call modDibujarInterface.mostrarError(0, "No es posible conectarse con el servidor. Intente nuevamente.")
    Else
       frmConnect.Visible = True
       Call modDibujarInterface.Show
       frmMain.Visible = False
    End If
 
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    
    Debug.Print ("Error en socket " & ErrorString)
    
    EstadoConexion = E_Estado.Ninguno
    
    ' Si me estaba conectando...
    If EstadoLogin = E_MODO.PantallaCreacion Or EstadoLogin = E_MODO.IngresarPersonaje Then
        ' Me fallo esta IP. La marco como fallida.
        If Not TCP.recibiPaquete Then
            recibiPaquete = False
            Call modLogin.agregarIpNoAccesible(Socket1.HostAddress)
            Socket1.Cleanup
           ' frmConnect.conectar
        End If
        Exit Sub
    End If
    
    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    ElseIf ErrorCode = 24061 Then
        Call modDibujarInterface.mostrarError(0, "No se ha podido conectar con el servidor. Consulte el estado del mismo en www.tierrasdelsur.cc .")
        Exit Sub
    Else
        Call modDibujarInterface.mostrarError(0, ErrorString & ".")
    End If
   
    frmConnect.MousePointer = 1
    response = 0
    
    frmMain.Socket1.Disconnect
End Sub

Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
    Dim longitud As Integer
    Dim lastPos As Integer
    Dim bytesEnCola As Integer
    
    Dim RD As String
    Call Socket1.read(RD, DataLength)
    
    RD = DecryptStr(RD, CryptOffs)  'byGorlok
    
    RD = rdbuffer & RD
    bytesEnCola = Len(RD)
      
    lastPos = 1
    
    Do
    
        longitud = Asc(mid$(RD, lastPos, 1)) + 1
              
        If longitud = 256 Then 'La longitud esta partida en dos partes?. Si es 256 quiere decir que si.
            'Tengo la otra parte o todavia no llego?
            'Aparte de tener la otra parte voy a ver si tengo posibilidades de que este el paquete completo
            'Se que al menos voy a tener un paquete de 256 caracteres. Si hay menos, seguro que esta incompleto
            'Y no tiene sentido que obtenga la longitud real y fijarme si esta todo el paquete en la estructura
            'siguiente. Me ahorro dos ifs, una resta y una asignacion :p
            'Tengo 3 bytes de la longitud + 256 que al menos tienen que estar -1 porque lastpost empieza en 1
            If bytesEnCola - lastPos > 257 Then
                longitud = STI(RD, lastPos + 1) + longitud
                lastPos = lastPos + 2
            Else
                Exit Do
            End If
        End If
        
        'Tengo el paquete completo?
        If bytesEnCola - lastPos >= longitud Then
            Call ProcesarPaquete(mid$(RD, lastPos + 1, longitud))
            lastPos = lastPos + longitud + 1
            
        Else
            If longitud > 255 Then lastPos = lastPos - 2
            Exit Do
        End If
        
    Loop Until bytesEnCola < lastPos
    'Sale del Do cuando no tiene mas que procesar porque no tiene o porque falta
    'Me quedo con la cantidad que no procece
    rdbuffer = right$(RD, bytesEnCola - lastPos + 1)
End Sub
Sub TocarMusica(numero As Integer)
Call Sonido_Play(numero)
End Sub

Public Sub DesactivarCMSG()
    Cmsgautomatico = False
    Pmsgautomatico = False
    Image1(4).Picture = Nothing
    rapidomsj = False
End Sub
     
Public Sub DesactivarPMSG()
    Cmsgautomatico = False
    Pmsgautomatico = False
    Image1(3).Picture = Nothing
    rapidomsj = False
End Sub

Private Sub VentanaMinimapa_Cerrar(maximizado As Boolean)
    GUI_Quitar ventanaMiniMapa
    Set ventanaMiniMapa = Nothing
    ventanaMiniMapaMaximizada = maximizado
End Sub
