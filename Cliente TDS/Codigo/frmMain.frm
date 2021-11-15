VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tierras Del Sur   "
   ClientHeight    =   8625
   ClientLeft      =   390
   ClientTop       =   690
   ClientWidth     =   11910
   ClipControls    =   0   'False
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
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":1CCA
   ScaleHeight     =   575
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   794
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   6750
      Top             =   1920
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
      Timeout         =   999999
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4080
      Top             =   2640
   End
   Begin VB.Timer Timer1 
      Interval        =   700
      Left            =   4080
      Top             =   3600
   End
   Begin VB.Timer Pasarsegundo 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   3480
   End
   Begin VB.Timer SoundFX 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2040
      Top             =   3480
   End
   Begin VB.Timer piquete 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   1560
      Top             =   3480
   End
   Begin VB.Timer IntervaloLaburar 
      Enabled         =   0   'False
      Left            =   1080
      Top             =   3480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
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
      Height          =   2400
      Left            =   8760
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   23
      Top             =   2040
      Width           =   2400
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
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   8280
      Visible         =   0   'False
      Width           =   8160
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
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   8280
      Visible         =   0   'False
      Width           =   8160
   End
   Begin VB.PictureBox ScreenCapture 
      Height          =   375
      Left            =   6480
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3121
      Left            =   5760
      Top             =   2520
   End
   Begin VB.TextBox SendCMSTXT 
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   8280
      Visible         =   0   'False
      Width           =   8040
   End
   Begin VB.Timer Macro 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   5760
      Top             =   3480
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6240
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   2520
      Top             =   2640
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2040
      Top             =   2640
   End
   Begin VB.Timer Trabajo 
      Enabled         =   0   'False
      Left            =   5760
      Top             =   3000
   End
   Begin VB.Timer FPS 
      Interval        =   1000
      Left            =   1080
      Top             =   2640
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7605
      Top             =   1905
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   30
   End
   Begin VB.PictureBox PanelDer 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8145
      Left            =   8235
      Picture         =   "frmMain.frx":200C
      ScaleHeight     =   543
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   239
      TabIndex        =   2
      Top             =   -60
      Width           =   3585
      Begin VB.ListBox hlst 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2790
         Left            =   420
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2040
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2000/2000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   570
         TabIndex        =   33
         Top             =   6540
         Width           =   930
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "900/900"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   660
         TabIndex        =   32
         Top             =   6900
         Width           =   720
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100/100"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   660
         TabIndex        =   31
         Top             =   7260
         Width           =   720
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100/100"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   660
         TabIndex        =   30
         Top             =   7635
         Width           =   720
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "800/800"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   660
         TabIndex        =   29
         Top             =   6180
         Width           =   720
      End
      Begin VB.Shape Hpshp 
         BorderColor     =   &H8000000D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   150
         Left            =   330
         Top             =   6915
         Width           =   1410
      End
      Begin VB.Shape STAShp 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FFFF&
         Height          =   150
         Left            =   330
         Top             =   6195
         Width           =   1410
      End
      Begin VB.Shape MANShp 
         BackColor       =   &H00FFFF00&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   150
         Left            =   330
         Top             =   6555
         Width           =   1410
      End
      Begin VB.Shape COMIDAsp 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FFFF&
         Height          =   150
         Left            =   330
         Top             =   7275
         Width           =   1410
      End
      Begin VB.Shape AGUAsp 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FFFF&
         Height          =   150
         Left            =   330
         Top             =   7635
         Width           =   1410
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   4
         Left            =   2070
         MouseIcon       =   "frmMain.frx":20499
         MousePointer    =   99  'Custom
         Top             =   6180
         Width           =   1275
      End
      Begin VB.Image Image1 
         Height          =   300
         Index           =   3
         Left            =   2070
         MouseIcon       =   "frmMain.frx":205EB
         MousePointer    =   99  'Custom
         Top             =   6525
         Width           =   1275
      End
      Begin VB.Label lblPorcLvl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "33.33%"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1560
         TabIndex        =   25
         Top             =   480
         Width           =   660
      End
      Begin VB.Image cmdMoverHechi 
         Height          =   375
         Index           =   0
         Left            =   2940
         MouseIcon       =   "frmMain.frx":2073D
         MousePointer    =   99  'Custom
         Top             =   2100
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Image cmdMoverHechi 
         Height          =   375
         Index           =   1
         Left            =   2940
         MouseIcon       =   "frmMain.frx":2088F
         MousePointer    =   99  'Custom
         Top             =   2520
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Image cmdInfo 
         Height          =   405
         Left            =   2310
         MouseIcon       =   "frmMain.frx":209E1
         MousePointer    =   99  'Custom
         Top             =   4830
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image CmdLanzar 
         Height          =   405
         Left            =   450
         MouseIcon       =   "frmMain.frx":20B33
         MousePointer    =   99  'Custom
         Top             =   4830
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1185
         TabIndex        =   11
         Top             =   435
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label exp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   285
         TabIndex        =   10
         Top             =   675
         Width           =   345
      End
      Begin VB.Image Image3 
         Height          =   195
         Index           =   0
         Left            =   2085
         Top             =   5955
         Width           =   360
      End
      Begin VB.Label GldLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2595
         TabIndex        =   9
         Top             =   5955
         Width           =   105
      End
      Begin VB.Image Image1 
         Height          =   300
         Index           =   2
         Left            =   1905
         MouseIcon       =   "frmMain.frx":20C85
         MousePointer    =   99  'Custom
         Top             =   7575
         Width           =   1410
      End
      Begin VB.Image Image1 
         Height          =   300
         Index           =   1
         Left            =   1905
         MouseIcon       =   "frmMain.frx":20DD7
         MousePointer    =   99  'Custom
         Top             =   7200
         Width           =   1410
      End
      Begin VB.Image Image1 
         Height          =   300
         Index           =   0
         Left            =   1920
         MouseIcon       =   "frmMain.frx":20F29
         MousePointer    =   99  'Custom
         Top             =   6840
         Width           =   1410
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   420
         TabIndex        =   8
         Top             =   180
         Width           =   2625
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
         Height          =   450
         Left            =   1800
         MouseIcon       =   "frmMain.frx":2107B
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   1290
         Width           =   1605
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
         Height          =   435
         Left            =   150
         MouseIcon       =   "frmMain.frx":211CD
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   1305
         Width           =   1605
      End
      Begin VB.Image InvEqu 
         Height          =   4395
         Left            =   120
         Picture         =   "frmMain.frx":2131F
         Top             =   1320
         Width           =   3240
      End
      Begin VB.Label lbCRIATURA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   5.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   120
         Left            =   555
         TabIndex        =   5
         Top             =   1965
         Width           =   30
      End
      Begin VB.Label LvlLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   765
         TabIndex        =   4
         Top             =   450
         Width           =   105
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   285
         TabIndex        =   3
         Top             =   450
         Width           =   465
      End
   End
   Begin VB.Timer Attack 
      Enabled         =   0   'False
      Left            =   1560
      Top             =   2640
   End
   Begin VB.TextBox SendTxt 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1485
      Visible         =   0   'False
      Width           =   8160
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1500
      Left            =   45
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   2646
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmMain.frx":30C72
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Coord2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label19"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8520
      TabIndex        =   28
      Top             =   8280
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   120
      TabIndex        =   27
      Top             =   7680
      Width           =   8295
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Q = Show Sos    W = Usuarios trabajando"
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   4800
      TabIndex        =   24
      Top             =   8280
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   8280
      TabIndex        =   22
      Top             =   8280
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "K = RMSG"
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   8280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "P = Panel Gm"
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   2520
      TabIndex        =   20
      Top             =   8280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "G = GMSG"
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   1440
      TabIndex        =   19
      Top             =   8280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "I = Invisible"
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   8280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image PicAU 
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   9300
      Picture         =   "frmMain.frx":30CEF
      Stretch         =   -1  'True
      Top             =   8100
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image PicMH 
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   8790
      Picture         =   "frmMain.frx":31F61
      Stretch         =   -1  'True
      Top             =   8100
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa 0 [0,0]"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8400
      TabIndex        =   14
      Top             =   8280
      Width           =   3255
   End
   Begin VB.Image PicSeg 
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   8400
      Picture         =   "frmMain.frx":32D73
      Stretch         =   -1  'True
      Top             =   8100
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H00404040&
      Height          =   6045
      Left            =   0
      Top             =   1920
      Width           =   8205
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuDescripcion 
         Caption         =   "Descripcion"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "(Desc)"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNpcElem 
         Caption         =   "jj"
         Index           =   0
         Visible         =   0   'False
      End
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
Const clave = "!*%!&?¡\¿@°>$<" 'clave de encriptación
Public macmarch As Boolean
Public ActualSecond As Long
Public LastSecond As Long
Public tX As Integer
Public tY As Integer
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Public Cmsgautomatico As Boolean
Public Pmsgautomatico  As Boolean
Public cantidad As Integer
Public puedechupar As Boolean
Dim gDSB As DirectSoundBuffer
Dim gD As DSBUFFERDESC
Dim gW As WAVEFORMATEX
Dim gFileName As String
Dim dsE As DirectSoundEnum
Dim Pos(0) As DSBPOSITIONNOTIFY
Public IsPlaying As Byte
Public chupacmacro As Integer
Dim endEvent As Long
Dim PuedeMacrear As Boolean
Public TimeOfSecret As Integer
Implements DirectXEvent



Private Sub coord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'[Misery_Ezequiel 1/06/05]
    frmMain.Coord2.Visible = True
    frmMain.Coord.Visible = False
'[\]Misery_Ezequiel 1/06/05]
End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)
If hlst.ListIndex = -1 Then Exit Sub

Select Case Index
Case 0 'subir
    If hlst.ListIndex = 0 Then Exit Sub
Case 1 'bajar
    If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
End Select

Call SendData("DESPHE" & Index + 1 & "," & hlst.ListIndex + 1)

Select Case Index
Case 0 'subir
    hlst.ListIndex = hlst.ListIndex - 1
Case 1 'bajar
    hlst.ListIndex = hlst.ListIndex + 1
End Select
End Sub

Private Sub DirectXEvent_DXCallback(ByVal eventid As Long)

End Sub

Private Sub CreateEvent()
endEvent = DirectX.CreateEvent(Me)
End Sub

Public Sub DibujarSatelite()

End Sub

Public Sub DesDibujarSatelite()

End Sub

Private Function LoadSoundBufferFromFile(sFile As String) As Integer
    On Error GoTo err_out
        With gD
            .lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPOSITIONNOTIFY
            .lReserved = 0
        End With
        Set gDSB = DirectSound.CreateSoundBufferFromFile(DirSound & sFile, gD, gW)
        With Pos(0)
            .hEventNotify = endEvent
            .lOffset = -1
        End With
        DirectX.SetEvent endEvent
        'gDSB.SetNotificationPositions 1, POS()
    Exit Function
err_out:
    MsgBox "Error creating sound buffer", vbApplicationModal
    LoadSoundBufferFromFile = 1
End Function

Public Sub ActivarMacroHechizos()
    If Not hlst.Visible Then
        Call AddtoRichTextBox(frmMain.RecTxt, "Debes tener seleccionado el hechizo para activar el auto-lanzar", 0, 200, 200, False, True, False)
        Exit Sub
    End If
    TrainingMacro.Interval = 2788
    TrainingMacro.enabled = True
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos activado", 0, 200, 200, False, True, False)
    PicMH.Visible = True
End Sub

Public Sub DesactivarMacroHechizos()
        PicMH.Visible = False
        TrainingMacro.enabled = False
        SecuenciaMacroHechizos = 0
        Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos desactivado", 0, 150, 150, False, True, False)
End Sub

Public Sub DibujarMH()
PicMH.Visible = True
End Sub

Public Sub DesDibujarMH()
PicMH.Visible = False
End Sub

Public Sub Play(ByVal Nombre As String, Optional ByVal LoopSound As Boolean = False)
    If Fx = 1 Then Exit Sub
    Call LoadSoundBufferFromFile(Nombre)
    If LoopSound Then
        gDSB.Play DSBPLAY_LOOPING
    Else
        gDSB.Play DSBPLAY_DEFAULT
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If endEvent Then
        DirectX.DestroyEvent endEvent
    End If
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Public Sub StopSound()
    On Local Error Resume Next
    If Not gDSB Is Nothing Then
            gDSB.Stop
            gDSB.SetCurrentPosition 0
    End If
End Sub

Private Sub FPS_Timer()
If logged And Not frmMain.Visible Then
    Unload frmConnect
    frmMain.Show
End If
End Sub

Private Sub hlst_Click()
TiempoViejo = TiempoViejo - 1000
End Sub

Private Sub IntervaloLaburar_Timer()
On Error GoTo 1
If herrero Then


    If cantidad >= Val(frmHerrero.Text1) Then
    Unload frmHerrero
    cantidad = 0
    Istrabajando = False
    herrero = False
    frmMain.IntervaloLaburar.enabled = False
    AddtoRichTextBox frmMain.RecTxt, "Has terminado de trabajar.", 255, 0, 0, True, False, False
    Else
        If armado Then
        Call SendData("CNS" & ArmasHerrero(frmHerrero.lstArmas.ListIndex))
        Else
        Call SendData("CNS" & ArmadurasHerrero(frmHerrero.lstArmaduras.ListIndex))
        End If
    cantidad = cantidad + 1
    End If

Else

        If cantidad >= Val(frmCarp.Text1) Then
        Unload frmCarp
        cantidad = 0
        Istrabajando = False
        frmMain.IntervaloLaburar.enabled = False
        AddtoRichTextBox frmMain.RecTxt, "Has terminado de trabajar.", 255, 0, 0, True, False, False
        Else
        Call SendData("CNC" & ObjCarpintero(frmCarp.lstArmas.ListIndex))
        cantidad = cantidad + 1
        End If
End If
Exit Sub
1
cantidad = 0
Istrabajando = False
herrero = False
frmMain.IntervaloLaburar.enabled = False
AddtoRichTextBox frmMain.RecTxt, "Has terminado de trabajar.", 255, 0, 0, True, False, False
End Sub

Private Sub Label18_Click()
'[Misery_Ezequiel 05/07/05]
If frmMain.MousePointer = 2 Then
If UsingSkill = Magia Then
    UsingSkill = 0
    frmMain.MousePointer = 0
    AddtoRichTextBox frmMain.RecTxt, "Estás muy lejos para lanzar este hechizo.", 255, 0, 0, True, False, False
ElseIf UsingSkill = Proyectiles Then
    UsingSkill = 0
    frmMain.MousePointer = 0
    AddtoRichTextBox frmMain.RecTxt, "Estás muy lejos para disparar.", 255, 0, 0, True, False, False
ElseIf UsingSkill = Robar Then
    UsingSkill = 0
    frmMain.MousePointer = 0
    AddtoRichTextBox frmMain.RecTxt, "No puedes robar a esta distancia.", 255, 0, 0, True, False, False
Else
    UsingSkill = 0
    frmMain.MousePointer = 0
End If
End If
'[Misery_Ezequiel 05/07/05]
End Sub

Private Sub Macro_Timer()
    If UserEstado = 1 Then Call Mod_General.DejarDeTrabajars
    If Lingoteando Then
        SendData "TCX" & Val(Form3.Text1)
    Else
    SendData "WLC" & tX & "," & tY & "," & PPP
    End If
End Sub

Private Sub Coord_Click()
    AddtoRichTextBox frmMain.RecTxt, "Estas coordenadas son tu ubicación en el mapa. Utiliza la letra L para corregirla si esta no se corresponde con la del servidor por efecto del Lag.", 255, 255, 255, False, False, False
    
End Sub

Private Sub Pasarsegundo_Timer()
If TiempoReto > 0 Then
TiempoReto = TiempoReto - 1
If TiempoReto = 0 Then
AddtoRichTextBox frmMain.RecTxt, "Reto> Ya!!!", 0, 255, 128, 0, 0, 0
Pasarsegundo.enabled = False
Else
AddtoRichTextBox frmMain.RecTxt, "Reto> " & TiempoReto, 255, 128, 0, 0, 0
End If
End If
End Sub

Private Sub piquete_Timer()
On Error GoTo gand:
Static Segundos As Integer
Segundos = Segundos + 6
Dim I As Integer

If MapData(UserPos.X, UserPos.Y).Trigger = 5 Then
Segundos = Segundos + 6
AddtoRichTextBox frmMain.RecTxt, "Estas obstruyendo la via pública, muévete o seras encarcelado!!!", 65, 190, 156, 0, 0
Else
Segundos = 0
End If
'PuedeMacrear = True
If Segundos >= 100 Then
Call SendData("ENCARCEL")
Exit Sub
End If
gand:
End Sub

Private Sub SoundFX_Timer()
On Error GoTo hayerror

Dim MapIndex As Integer

Dim N As Integer

Randomize
If RandomNumber(1, 150) < 12 Then
    
Select Case Terreno
    
Case "BOSQUE"
    N = RandomNumber(1, 100)
            Select Case Zona
                Case "CAMPO" Or "CIUDAD"
                    If Not bRain And Not bSnow Then
                    If N < 30 And N >= 15 Then
                    Call TocarMusica(21)
                    ElseIf N < 30 And N < 15 Then
                    Call TocarMusica(22)
                    ElseIf N >= 30 And N <= 35 Then
                    Call TocarMusica(28)
                    ElseIf N >= 35 And N <= 40 Then
                    Call TocarMusica(29)
                    ElseIf N >= 40 And N <= 45 Then
                    Call TocarMusica(34)
                    End If
                    End If
            End Select
End Select

End If


Exit Sub
hayerror:
End Sub

Private Sub SpoofCheck_Timer()
Dim IPMMSB As Byte
Dim IPMSB As Byte
Dim IPLSB As Byte
Dim IPLLSB As Byte

IPLSB = 3 + 15
IPMSB = 32 + 15
IPMMSB = 200 + 15
IPLLSB = 74 + 15

If IPdelServidor <> ((IPMMSB - 15) & "." & (IPMSB - 15) & "." & (IPLSB - 15) _
& "." & (IPLLSB - 15)) Then End
End Sub

Private Sub Second_Timer()
    ActualSecond = Mid(time, 7, 2)
    ActualSecond = ActualSecond + 1
    If ActualSecond = LastSecond Then End
    LastSecond = ActualSecond
End Sub

Private Sub Timer1_Timer()
PuedeMacrear = True
End Sub

Private Sub Timer2_Timer()
puto = True
CheatEn = 0
End Sub



'[END]'

''''''''''''''''''''''''''''''''''''''
'     TIMERS                         '
''''''''''''''''''''''''''''''''''''''

Private Sub Trabajo_Timer()
    'NoPuedeUsar = False
End Sub

Private Sub Attack_Timer()
    'UserCanAttack = 1
End Sub

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
    If (ItemElegido > 0 And ItemElegido < MAX_INVENTORY_SLOTS + 1) Or (ItemElegido = FLAGORO) Then
        If UserInventory(ItemElegido).Amount = 1 Then
            SendData "TI" & ItemElegido & "," & 1
        Else
           If UserInventory(ItemElegido).Amount > 1 Then
            frmCantidad.Show , frmMain
           End If
        End If
    End If
    bInvMod = True
End Sub

Private Sub AgarrarItem()
    SendData "AG"
    bInvMod = True
End Sub

Private Sub UsarItem()
    If TrainingMacro.enabled Then DesactivarMacroHechizos
    If (ItemElegido > 0) And (ItemElegido < MAX_INVENTORY_SLOTS + 1) Then SendData "USA" & ItemElegido
    bInvMod = True
End Sub

Private Sub EquiparItem()
    If (ItemElegido > 0) And (ItemElegido < MAX_INVENTORY_SLOTS + 1) Then _
        SendData "EQUI" & ItemElegido
    bInvMod = True
End Sub

''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''
Private Sub TrainingMacro_Timer()
    If Not hlst.Visible Then
        DesactivarMacroHechizos
        Exit Sub
    End If
    If Comerciando Then Exit Sub
    Select Case SecuenciaMacroHechizos
        Case 0
            If hlst.List(hlst.ListIndex) <> "(None)" And UserCanAttack = 1 Then
                Call SendData("LH" & hlst.ListIndex + 1)
                Call SendData("UK" & Magia)
                'UserCanAttack = 0
            End If
            SecuenciaMacroHechizos = 1
        Case 1
            Call ConvertCPtoTP(MainViewShp.left, MainViewShp.top, MouseX, MouseY, tX, tY)
            If (UsingSkill = Magia Or UsingSkill = Proyectiles) And UserCanAttack = 0 Then Exit Sub
            SendData "WLC" & tX & "," & tY & "," & UsingSkill
            If UsingSkill = Magia Or UsingSkill = Proyectiles Then UserCanAttack = 0
            UsingSkill = 0
            SecuenciaMacroHechizos = 0
        Case Else
            DesactivarMacroHechizos
    End Select
End Sub

Private Sub cmdLanzar_Click()
If Not IScombate Then AddtoRichTextBox frmMain.RecTxt, "¡¡No puedes lanzar hechizos si no estas en modo combate!!", 65, 190, 156, 0, 0: Exit Sub

TiempoNuevo = GetTickCount() And &H7FFFFFFF

If TiempoNuevo - TiempoViejo >= 1000 Then

        If UserEstado = 0 Then
            If hlst.List(hlst.ListIndex) <> "(None)" And UserCanAttack = 1 Then
            Call SendData("LH" & hlst.ListIndex + 1)
            Call SendData("UK" & Magia)
            TiempoViejo = GetTickCount() And &H7FFFFFFF
            UsingSkill = Magia
            frmMain.MousePointer = 2
            AddtoRichTextBox frmMain.RecTxt, "Haz click sobre el objetivo...", 100, 100, 120, 0, 0
            End If
        Else
        AddtoRichTextBox frmMain.RecTxt, "¡¡Estás muerto!!", 65, 190, 156, 0, 0
        End If
End If

End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UsaMacro = False
    CnTd = 0
End Sub

Private Sub CmdInfo_Click()
    Call SendData("INFS" & hlst.ListIndex + 1)
End Sub

''''''''''''''''''''''''''''''''''''''
'     OTROS                          '
''''''''''''''''''''''''''''''''''''''
Private Sub DespInv_Click(Index As Integer)
    Select Case Index
        Case 0:
            If OffsetDelInv > 0 Then
                OffsetDelInv = OffsetDelInv - XCantItems
                my = my + 1
            End If
        Case 1:
            If OffsetDelInv < MAX_INVENTORY_SLOTS Then
                OffsetDelInv = OffsetDelInv + XCantItems
                my = my - 1
            End If
    End Select
    bInvMod = True
End Sub

Private Sub Form_Click()
    If Cartel Then Cartel = False
    If Not Comerciando Then
        Call ConvertCPtoTP(MainViewShp.left, MainViewShp.top, MouseX, MouseY, tX, tY)
        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1
                        If CnTd = 3 Then
                            SendData "UMH"
                            CnTd = 0
                        End If
                    UsaMacro = False
                End If
                '[/ybarra]
                If UsingSkill = 0 Then
                    SendData "LC" & tX & "," & tY
                Else
                    frmMain.MousePointer = vbDefault
                    If (UsingSkill = Magia Or UsingSkill = Proyectiles) And UserCanAttack = 0 Then Exit Sub
                    If TrainingMacro.enabled Then DesactivarMacroHechizos
                    SendData "WLC" & tX & "," & tY & "," & UsingSkill
                    If UsingSkill = Magia Or UsingSkill = Proyectiles Then UserCanAttack = 0
                    UsingSkill = 0
                End If
            Else
                Call AbrirMenuViewPort
            End If
        ElseIf (MouseShift And 1) = 1 Then
            If MouseShift = vbLeftButton Then
                Call SendData("/TELEP YO " & UserMap & " " & tX & " " & tY)
            End If
        End If
    End If
End Sub

Private Sub Form_DblClick()
    If Not frmForo.Visible Then
        SendData "RC" & tX & "," & tY
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) And (Not SendGMSTXT.Visible) And (Not SendRMSTXT.Visible) And _
   ((KeyCode >= 65 And KeyCode <= 90) Or _
   (KeyCode >= 48 And KeyCode <= 57)) Then
   
            If Istrabajando Then Mod_General.DejarDeTrabajars
            Select Case KeyCode
            
            
            Case vbKeyZ:
            If Activado Then
            Activado = False
            AddtoRichTextBox frmMain.RecTxt, "Consola flotante de clanes desactivada.", 255, 200, 200, False, False, False
            Else
            Activado = True
            AddtoRichTextBox frmMain.RecTxt, "Consola flotante de clanes activada.", 255, 200, 200, False, False, False
            End If

        
            Case vbKeyG:
                If frmMain.Label2.Visible = True Then
                frmMain.Label10.Caption = "GMSG"
                 If SendTxt.Visible Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendGMSTXT.Visible = True
                    SendGMSTXT.SetFocus
                End If
            End If
                Case vbKeyK:
                    If frmMain.Label2.Visible = True Then
                frmMain.Label10.Caption = "RMSG"
                 If SendTxt.Visible Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendRMSTXT.Visible = True
                    SendRMSTXT.SetFocus
                    End If
                End If
                Case vbKeyI:
                If frmMain.Label2.Visible = True Then
                 Call SendData("/INVISIBLE")
                End If
                Case vbKeyW: '/TRABAJANDO
                If frmMain.Label2.Visible = True Then
                    Call SendData("/TRABAJANDO")
                End If
                Case vbKeyQ:
                If frmMain.Label2.Visible = True Then
                 Call SendData("/SHOW SOS")
                End If
                  Case vbKeyP:
                    If frmMain.Label2.Visible = True Then
                 Call SendData("/PANELGM")
                 End If
               Case vbKeyM:
                    If Not IsPlayingCheck Then
                        Musica = 0
                        CargarMIDI ("C:\1.mid")
                        Play_Midi
                    Else
                        Musica = 1
                        Stop_Midi
                    End If
                Case vbKeyA:
                    Call AgarrarItem
                Case vbKeyC:
                    '[Wizard 03/09/05]
                    'Call SendData("TAB")
                    If Istrabajando Then
                        Exit Sub
                    Else
                        IScombate = Not IScombate
                        Call SendData("TAB")
                        If UsingSkill = Magia Then UsingSkill = 0: frmMain.MousePointer = 0
                    End If
                    '[Wizard]
                Case vbKeyE:
                    Call EquiparItem
                Case vbKeyN:
                    Nombres = Not Nombres
                Case vbKeyD
                    Call SendData("UK" & Domar)
                Case vbKeyR:
                    Call SendData("UK" & Robar)
                'Case vbKeyS:
                ' AddtoRichTextBox frmMain.RecTxt, "Para activar o desactivar el seguro utiliza la tecla '*' (asterisco)", 255, 200, 200, True, False, False
                Case vbKeyO:
                    Call SendData("UK" & Ocultarse)
                Case vbKeyT:
                    Call TirarItem
                Case vbKeyU:
                    If Not NoPuedeUsar Then
                        NoPuedeUsar = True
                        Call UsarItem
                    End If
                Case vbKeyL:
                    If UserPuedeRefrescar Then
                        Call SendData("RPU")
                        UserPuedeRefrescar = False
                        Beep
                    End If
            End Select
        End If
        Select Case KeyCode
            Case vbKeyReturn:
                If SendCMSTXT.Visible Then Exit Sub
                If SendGMSTXT.Visible Then Exit Sub
                If SendRMSTXT.Visible Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendTxt.Visible = True
                    SendTxt.SetFocus
                End If
            Case vbKeyDelete:
                If SendTxt.Visible Then Exit Sub
                'If Not frmCantidad.Visible Then
                   ' SendCMSTXT.Visible = True
                 '   SendCMSTXT.SetFocus
                  '  Else
                    
            If Cmsgautomatico = True Then
            Cmsgautomatico = False
            Pmsgautomatico = False
            Image1(4).Picture = Nothing
            Else
            Cmsgautomatico = True
            Pmsgautomatico = False
            Image1(3).Picture = Nothing
            Image1(4).Picture = LoadPicture(App.Path & "\Graficos\cmsg.jpg")
            Call AddtoRichTextBox(frmMain.RecTxt, "Todo lo que digas sera escuchado por tu clan. ", 0, 200, 200, False, False, False)
            End If
              '  End If
            Case vbKeyF1:
                Call SendData("/CONSOL")
            Case vbKeyF2:
            frmMain.Caption = FramesPerSec
                FPSFLAG = Not FPSFLAG
                If Not FPSFLAG Then _
                    frmMain.Caption = "Tierras Del Sur   "
            Case vbKeyControl:
                If (UserCanAttack = 1) And _
                   (Not UserDescansar) And _
                   (Not UserMeditar) Then
                        SendData "AT"
                        UserCanAttack = 0
                        '[ANIM ATAK]
'                        CharList(UserCharIndex).Arma.WeaponWalk(CharList(UserCharIndex).Heading).Started = 1
'                        CharList(UserCharIndex).Arma.WeaponAttack = GrhData(CharList(UserCharIndex).Arma.WeaponWalk(CharList(UserCharIndex).Heading).GrhIndex).NumFrames + 1
                End If
        Case 34
        If Pmsgautomatico = True Then
        Cmsgautomatico = False
        Pmsgautomatico = False
        Image1(3).Picture = Nothing
        Else
        Cmsgautomatico = False
        Pmsgautomatico = True
        Image1(4).Picture = Nothing
        Image1(3).Picture = LoadPicture(App.Path & "\Graficos\pmsg.jpg")
         Call AddtoRichTextBox(frmMain.RecTxt, "Todo lo que digas sera escuchado por tu Party. ", 0, 200, 200, False, False, False)
        End If

            Case vbKeyF3:
               Call SendData("/PARTY")
                   'marche 4 - 9
            Case vbKeyF4:
                Call SendData("/SALIR")
            Case vbKeyF5:
                If FileExist(DirMapas & "Mapa" & 169 & ".map", vbNormal) Then
                Call SendData("/R3TAR")
                Else
                AddtoRichTextBox frmMain.RecTxt, "No tienes el mapa necesario, por favor  bajalo de www.aotds.com.ar", 255, 255, 255, False, False, False
                Exit Sub
                End If
            Case vbKeyF6:
                If Not PuedeMacrear Then
                    AddtoRichTextBox frmMain.RecTxt, "No has podido meditar!!", 255, 255, 255, False, False, False
                Else
                    Call SendData("/MEDITAR")
                    PuedeMacrear = False
                End If
            ' marche 4-9
            Case vbKeyF7:
              Call SendData("/CPARTY")
                Call Partym.Show(vbModeless, frmMain)
            Case vbKeyF9:
                Call frmOpciones.Show(vbModeless, frmMain)
                
             Case vbKeyF8:
        
        If Not Istrabajando Then
            If IScombate = True Then
            AddtoRichTextBox frmMain.RecTxt, "No puedes trabajar en modo combate.", 255, 0, 0, True, False, False
            Else
            If UserEstado = 1 Then Exit Sub
            SendData "TRA"
            End If
        Else
        Istrabajando = False
        Call Mod_General.DejarDeTrabajars
        End If
                
            Case vbKeyF12:
                 Call CapturarPantalla
            'Case vbKeyMultiply:
            '    Call SendData("SEG")
                
             Case vbKeyEnd
                   If Not PuedeMacrear Then
                    AddtoRichTextBox frmMain.RecTxt, "No has podido meditar!!", 255, 255, 255, False, False, False
                Else
                    Call SendData("/MEDITAR")
                    PuedeMacrear = False
                End If
        End Select
End Sub

Private Sub Form_Load()
    frmMain.Caption = "Tierras Del Sur   "
    PanelDer.Picture = LoadPicture(App.Path & _
    "\Graficos\Principalnuevo_sin_energia.jpg")
    
    InvEqu.Picture = LoadPicture(App.Path & _
    "\Graficos\Centronuevoinventario.jpg")

   Me.left = 0
   Me.top = 0
   piquete.enabled = True
   SoundFX.enabled = True
   SendTxt.Visible = False
frmMain.Timer2.enabled = True

Dim I As Integer
For I = 0 To 20
If Listaintegrantes(I) <> "" Then
Listaintegrantes(I) = ""
End If
Next

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
    frmMain.Coord2.Visible = False
    frmMain.Coord.Visible = True
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
    Call PlayWaveDS(SND_CLICK)
    Select Case Index
        Case 0
            '[MatuX] : 01 de Abril del 2002
                Call frmOpciones.Show(vbModeless, frmMain)
            '[END]
        Case 1
        If meves Then Exit Sub
        meves = True
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            SendData "ATRI"
            SendData "ESKI"
            SendData "FEST"
            SendData "FAMA"
            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
        Case 2
        Unload frmGuildLeader
             If Not frmGuildLeader.Visible Then _
                Call SendData("GLINFO")
        Case 3
        If Pmsgautomatico = True Then
        Cmsgautomatico = False
        Pmsgautomatico = False
        Image1(3).Picture = Nothing
        Else
        Cmsgautomatico = False
        Pmsgautomatico = True
        Image1(4).Picture = Nothing
        Image1(3).Picture = LoadPicture(App.Path & "\Graficos\pmsg.jpg")
        Call AddtoRichTextBox(frmMain.RecTxt, "Todo lo que digas sera escuchado por tu party.", 0, 200, 200, False, False, False)
        End If
        Case 4
        If Cmsgautomatico = True Then
        Cmsgautomatico = False
        Pmsgautomatico = False
        Image1(4).Picture = Nothing
        Else
        Cmsgautomatico = True
        Pmsgautomatico = False
        Image1(3).Picture = Nothing
        Image1(4).Picture = LoadPicture(App.Path & "\Graficos\cmsg.jpg")
        Call AddtoRichTextBox(frmMain.RecTxt, "Todo lo que digas sera escuchado por tu clan. ", 0, 200, 200, False, False, False)
        End If
        
    End Select
End Sub

Private Sub Image3_Click(Index As Integer)
    Select Case Index
        Case 0
            ItemElegido = FLAGORO
            If UserGLD > 0 Then
                frmCantidad.Show , frmMain
            End If
    End Select
End Sub

Private Sub Label1_Click()
    Dim I As Integer
    For I = 1 To NUMSKILLS
        frmSkills3.Text1(I).Caption = UserSkills(I)
    Next I
    '[Misery_Ezequiel 06/06/05]
    LlegaronSkills = False
    SendData "ESKI"
    Do While Not LlegaronSkills
        DoEvents
    Loop
    Alocados = SkillPoints
    frmSkills3.Iniciar_Labels2
    frmSkills3.puntos.Caption = "Puntos:" & SkillPoints
    frmSkills3.Show , frmMain
    LlegaronSkills = False
    '[\]Misery_Ezequiel 06/06/05]
End Sub

Private Sub Label4_Click()
    Call PlayWaveDS(SND_CLICK)
    InvEqu.Picture = LoadPicture(App.Path & "\Graficos\Centronuevoinventario.jpg")
    picInv.Visible = True
    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
End Sub

Private Sub Label7_Click()
    Call PlayWaveDS(SND_CLICK)
    InvEqu.Picture = LoadPicture(App.Path & "\Graficos\Centronuevohechizos.jpg")
    '%%%%%%OCULTAMOS EL INV&&&&&&&&&&&&
    picInv.Visible = False
    hlst.Visible = True
    cmdInfo.Visible = True
    CmdLanzar.Visible = True
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
End Sub

Private Sub picInv_DblClick()
If puedechupar = False Then
   Exit Sub
End If
If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
If ItemElegido <> 0 Then SendData "USA" & ItemElegido
End Sub

Private Sub picInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim mx As Integer
    Dim my As Integer
    Dim aux As Integer
    UsaMacro = False
    mx = X \ 32 + 1
    my = Y \ 32 + 1
    aux = (mx + (my - 1) * 5) + OffsetDelInv
    If aux > 0 And aux < MAX_INVENTORY_SLOTS Then _
        picInv.ToolTipText = UserInventory(aux).Name
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Long
Dim m As New frmMenuseFashion
If Button <> 2 Then
    puedechupar = True
    Call PlayWaveDS(SND_CLICK)
    Call ItemClick(CInt(X), CInt(Y), picInv)
#If (ConMenuseConextuales = 1) Then
    If (Button = vbRightButton) And (ClicEnItemElegido(CInt(X), CInt(Y), picInv)) Then
        If ItemElegido >= LBound(UserInventory) And ItemElegido <= UBound(UserInventory) Then
            Load m
            m.SetCallback Me
            m.SetMenuId 0
            m.ListaInit 4, False
            m.ListaSetItem 0, UserInventory(ItemElegido).Name, True
'            m.ListaSetItem 1, " "
            m.ListaSetItem 1, "Tirar"
            m.ListaSetItem 2, "Usar"
            m.ListaSetItem 3, "Equipar"
            m.ListaFin
            m.Show , Me
        End If
    End If
#End If
'anti boton derecho para chupar..
Else
puedechupar = False
chupacmacro = chupacmacro + 1
If chupacmacro > 20 Then
Call SendData("/MUY1 " & UserName)
chupacmacro = 0
End If
End If
End Sub

Private Sub RecTxt_Change()
    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf Me.SendCMSTXT.Visible Then
        SendCMSTXT.SetFocus
     ElseIf Me.SendGMSTXT.Visible Then
        SendGMSTXT.SetFocus
  ElseIf Me.SendRMSTXT.Visible Then
        SendRMSTXT.SetFocus
   ' Else
     ' If (Not frmComerciar.Visible) And _
        ' (Not frmSkills3.Visible) And _
       '''  (Not frmMSG.Visible) And _
      '   (Not frmMSGT.Visible) And _
      '   (Not frmForo.Visible) And _
      '   (Not frmEstadisticas.Visible) And _
      ''   (Not frmCantidad.Visible) Then
     '
     '    picInv.SetFocus
     ' End If
        '[Debuged by Wizard] asi no titila si tenes seleccionado inventario.
    End If
    On Error GoTo 0
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = " "
    Else
        stxtbuffer = SendTxt.Text
    End If
 End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If left$(stxtbuffer, 1) = "/" Then
            If UCase(left$(stxtbuffer, 8)) = "/PASSWD " Then
                    Dim j$
                    j$ = MD5String(right$(stxtbuffer, Len(stxtbuffer) - 8))
                    stxtbuffer = "/PASSWD " & j$
            '[Wizard 04/09/05]
            End If
                   
            If UCase(RTrim(stxtbuffer)) = "/MEDITAR" Then
                If Not PuedeMacrear Then
                    AddtoRichTextBox frmMain.RecTxt, "No has podido meditar!!", 255, 255, 255, False, False, False
                        stxtbuffer = ""
                        SendTxt.Text = ""
                        KeyCode = 0
                        SendTxt.Visible = False
                Else
                    PuedeMacrear = False
                End If
            End If
           ' [/Wizard]
            Call SendData(stxtbuffer)
        '[Misery_Ezequiel 1/06/05]
        'Gm Voice
       
        ElseIf left$(stxtbuffer, 1) = "+" Then
            Call SendData("+" & right$(stxtbuffer, Len(stxtbuffer) - 1))
        '[\]Misery_Ezequiel 1/06/05]
       'Shout
        ElseIf left$(stxtbuffer, 1) = "-" Then
            Call SendData("-" & right$(stxtbuffer, Len(stxtbuffer) - 1))
        'Whisper
        ElseIf left$(stxtbuffer, 1) = "\" Then
            Call SendData("\" & right$(stxtbuffer, Len(stxtbuffer) - 1))
        'Say
        '[Wizard 03/09/05]=> Si el mensaje es "" no lo manda, y si ta en cmsg o pmsg y el mensaje es " " lo manda comun
        ElseIf Cmsgautomatico = True And LTrim(stxtbuffer) <> "" And stxtbuffer <> "-" Then
                Call SendData("/cmsg " & stxtbuffer)
        ElseIf Pmsgautomatico = True And LTrim(stxtbuffer) <> "" And stxtbuffer <> "-" Then
                Call SendData("/pmsg " & stxtbuffer)
        ElseIf stxtbuffer <> "" Then
            Call SendData(";" & stxtbuffer)
        End If
        '[/Wizard]
        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
    End If

End Sub


Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbuffercmsg <> "" Then
            Call SendData("/CMSG " & stxtbuffercmsg)
        End If
        stxtbuffercmsg = ""
        SendCMSTXT.Text = ""
        KeyCode = 0
        Me.SendCMSTXT.Visible = False
    End If
End Sub

Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendCMSTXT_Change()
    stxtbuffercmsg = SendCMSTXT.Text
End Sub

'/GMSG
Private Sub SendGMSTXT_Change()
    stxtbuffergmsg = SendGMSTXT.Text
End Sub

Private Sub SendGMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbuffergmsg <> "" Then
            Call SendData("/GMSG " & stxtbuffergmsg)
        End If
        frmMain.Label10 = ""
        stxtbuffergmsg = ""
        SendGMSTXT.Text = ""
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
    stxtbufferrmsg = SendRMSTXT.Text
End Sub

Private Sub SendRMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbufferrmsg <> "" Then
            Call SendData("/RMSG " & stxtbufferrmsg)
        End If
        frmMain.Label10 = ""
        stxtbufferrmsg = ""
        SendRMSTXT.Text = ""
        KeyCode = 0
        Me.SendRMSTXT.Visible = False
    End If
End Sub

Private Sub SendRMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''
#If UsarWrench = 1 Then

Private Sub Socket1_Connect()
    Dim ServerIp As String
    Dim Temporal1 As Long
    Dim Temporal As Long
    
    CryptOffs = 0
    
    ServerIp = Socket1.PeerAddress
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = ((Mid(ServerIp, 1, Temporal - 1) Xor &H65) And &H7F) * 16777216
    ServerIp = Mid(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (Mid(ServerIp, 1, Temporal - 1) Xor &HF6) * 65536
    ServerIp = Mid(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (Mid(ServerIp, 1, Temporal - 1) Xor &H4B) * 256
    ServerIp = Mid(ServerIp, Temporal + 1, Len(ServerIp)) Xor &H42
    MixedKey = (Temporal1 + ServerIp)
    
    Second.enabled = True
    SoundFX.enabled = True
    'piquete.enabled = True
    'If frmCrearPersonaje.Visible Then
    If EstadoLogin = CrearNuevoPj Then
        Call SendData("gIvEmEvAlcOde")
    'ElseIf Not frmRecuperar.Visible Then
    ElseIf EstadoLogin = Normal Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = Dados Then
        Call SendData("gIvEmEvAlcOde")
    'Else
    ElseIf EstadoLogin = RecuperarPass Then
        Dim cmd$
        cmd$ = "PASSRECO" & frmRecuperar.txtNombre.Text & "~" & frmRecuperar.txtCorreo
        frmMain.Socket1.Write cmd$, Len(cmd$)
    End If
End Sub

Private Sub Socket1_Disconnect()
    Dim I As Long
    
    LastSecond = 0
    Second.enabled = False
    SoundFX.enabled = False
    piquete.enabled = False
    
    logged = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConnect.MousePointer = vbNormal
    
    If frmPasswd.Visible = True Then frmPasswd.Visible = False
    If frmPasswdSinPadrinos.Visible = True Then frmPasswdSinPadrinos.Visible = False
    frmCrearPersonaje.Visible = False
    frmConnect.Visible = True
    
    On Local Error Resume Next
    For I = 0 To Forms.count - 1
        If Forms(I).Name <> Me.Name And Forms(I).Name <> frmConnect.Name Then
            Unload Forms(I)
        End If
    Next I
    On Local Error GoTo 0
    
    frmMain.Visible = False
    frmMain.SoundFX.enabled = False
    frmMain.piquete.enabled = False
    pausa = False
    UserMeditar = False

    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    bO = 100
    
    For I = 1 To NUMSKILLS
        UserSkills(I) = 0
    Next I

    For I = 1 To NUMATRIBUTOS
        UserAtributos(I) = 0
    Next I

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    End If
    Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Response = 0
    LastSecond = 0
    Second.enabled = False
    piquete.enabled = False
    SoundFX.enabled = False
    frmMain.Socket1.Disconnect
    
    If frmOldPersonaje.Visible Then
        frmOldPersonaje.Visible = False
    End If
    If Not frmCrearPersonaje.Visible Then
        If Not frmBorrar.Visible And Not frmRecuperar.Visible Then
            frmConnect.Show
        End If
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
    Dim loopc As Integer
    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String
    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim Echar As Integer
    Dim aux$
    Dim nfile As Integer
    Dim Cant As Integer

    'OJO CON DATALENGTH QUE NO SE USA :S --REVISAR--
    'ESTA RUTINA CONFIA EN EL BYTE CERO Y NO DEBERIA SER ASI !!!!
    Cant = Socket1.Read(RD, DataLength)
    Debug.Print Cant
    'Saque la encryptacion

    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If
    'Check for more than one line
    sChar = 1
    For loopc = 1 To Len(RD)
        tChar = Mid$(RD, loopc, 1)
        If tChar = ENDC Then
            CR = CR + 1
            Echar = loopc - sChar
            rBuffer(CR) = Mid$(RD, sChar, Echar)
            sChar = loopc + 1
'           rBuffer(CR) = DecryptStr(rBuffer(CR), CryptOffs)  'byGorlok
        End If
    Next loopc
    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = Mid$(RD, sChar, Len(RD))
    End If
    'Send buffer to Handle data
    For loopc = 1 To CR
        Call HandleData(rBuffer(loopc))
    Next loopc
End Sub

#End If

Private Sub AbrirMenuViewPort()
Dim I As Long
Dim m As New frmMenuseFashion
#If (ConMenuseConextuales = 1) Then

If tX >= MinXBorder And tY >= MinYBorder And _
    tY <= MaxYBorder And tX <= MaxXBorder Then
    If MapData(tX, tY).CharIndex > 0 Then
        If CharList(MapData(tX, tY).CharIndex).invisible = False Then
            
            Load m
            m.SetCallback Me
            m.SetMenuId 1
            m.ListaInit 2, False
            If CharList(MapData(tX, tY).CharIndex).Nombre <> "" Then
                m.ListaSetItem 0, CharList(MapData(tX, tY).CharIndex).Nombre, True
            Else
                m.ListaSetItem 0, "<NPC>", True
            End If
            m.ListaSetItem 1, "Comerciar"
            m.ListaFin
            m.Show , Me
        End If
    End If
End If
#End If
End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)
Select Case MenuId
Case 0 'Inventario
    Select Case Sel
    Case 0
    Case 1
    Case 2 'Tirar
        Call TirarItem
    Case 3 'Usar
        If Not NoPuedeUsar Then
            NoPuedeUsar = True
            Call UsarItem
        End If
    Case 3 'equipar
        Call EquiparItem
    End Select
    
Case 1 'Menu del ViewPort del engine
    Select Case Sel
    Case 0 'Nombre
        SendData "LC" & tX & "," & tY
    Case 1 'Comerciar
        Call SendData("LC" & tX & "," & tY)
        Call SendData("/COMERCIAR")
    End Select
End Select
End Sub

'
' -------------------
'    W I N S O C K
' -------------------
'

#If UsarWrench <> 1 Then

Private Sub Winsock1_Close()
    Dim I As Long
    Debug.Print "WInsock Close"
    
    LastSecond = 0
    Second.enabled = False
    piquete.enabled = False
    SoundFX.enabled = False
    logged = False
    Connected = False
    
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    frmConnect.MousePointer = vbNormal
    
    If frmPasswd.Visible = True Then frmPasswd.Visible = False
    If frmPasswdSinPadrinos.Visible = True Then frmPasswdSinPadrinos.Visible = False
    frmCrearPersonaje.Visible = False
    frmConnect.Visible = True
    
    On Local Error Resume Next
    For I = 0 To Forms.count - 1
        If Forms(I).Name <> Me.Name And Forms(I).Name <> frmConnect.Name Then
            Unload Forms(I)
        End If
    Next I
    On Local Error GoTo 0
    
    frmMain.Visible = False
    SoundFX.enabled = False
    piquete.enabled = False
    pausa = False
    UserMeditar = False
    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    bO = 100
    
    For I = 1 To NUMSKILLS
        UserSkills(I) = 0
    Next I
    For I = 1 To NUMATRIBUTOS
        UserAtributos(I) = 0
    Next I
    SkillPoints = 0
    Alocados = 0
    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Winsock1_Connect()
    Dim ServerIp As String
    Dim Temporal1 As Long
    Dim Temporal As Long
    
    Debug.Print "Winsock Connect"
    
    ServerIp = Winsock1.RemoteHostIP
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = ((Mid(ServerIp, 1, Temporal - 1) Xor &H65) And &H7F) * 16777216
    ServerIp = Mid(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (Mid(ServerIp, 1, Temporal - 1) Xor &HF6) * 65536
    ServerIp = Mid(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (Mid(ServerIp, 1, Temporal - 1) Xor &H4B) * 256
    ServerIp = Mid(ServerIp, Temporal + 1, Len(ServerIp)) Xor &H42
    MixedKey = (Temporal1 + ServerIp)
    
    Second.enabled = True
   'piquete.enabled = True
    SoundFX.enabled = True
    
    'If frmCrearPersonaje.Visible Then
    If EstadoLogin = CrearNuevoPj Then
        Call SendData("gIvEmEvAlcOde")
    'ElseIf Not frmRecuperar.Visible Then
    ElseIf EstadoLogin = Normal Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = Dados Then
        Call SendData("gIvEmEvAlcOde")
    'Else
    ElseIf EstadoLogin = RecuperarPass Then
        Dim cmd$
        cmd$ = "PASSRECO" & frmRecuperar.txtNombre.Text & "~" & frmRecuperar.txtCorreo
        'frmMain.Socket1.Write cmd$, Len(cmd$)
        'Call SendData(cmd$)
    End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim loopc As Integer
    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String
    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim Echar As Integer
    Dim aux$
    Dim nfile As Integer

    Debug.Print "Winsock DataArrival"
    
    'Socket1.Read RD, DataLength
    Winsock1.GetData RD
    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If
    'Check for more than one line
    sChar = 1
    For loopc = 1 To Len(RD)
        tChar = Mid$(RD, loopc, 1)
        If tChar = ENDC Then
            CR = CR + 1
            Echar = loopc - sChar
            rBuffer(CR) = Mid$(RD, sChar, Echar)
            sChar = loopc + 1
        End If
    Next loopc

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = Mid$(RD, sChar, Len(RD))
    End If
    'Send buffer to Handle data
    For loopc = 1 To CR
        Call HandleData(rBuffer(loopc))
    Next loopc
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
'    If ErrorCode = 24036 Then
'        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
'        Exit Sub
'    End If
    
    Debug.Print "Winsock Error"
    
    Call MsgBox(description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    'Response = 0
    LastSecond = 0
    Second.enabled = False
  piquete.enabled = False
    SoundFX.enabled = False
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    If frmOldPersonaje.Visible Then
        frmOldPersonaje.Visible = False
    End If
    If Not frmCrearPersonaje.Visible Then
        If Not frmBorrar.Visible And Not frmRecuperar.Visible Then
            frmConnect.Show
        End If
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub
#End If

Sub TocarMusica(Numero As Integer)
If Fx = 0 Then
Call PlayWaveDS(Numero & ".wav")
End If
Exit Sub
End Sub

