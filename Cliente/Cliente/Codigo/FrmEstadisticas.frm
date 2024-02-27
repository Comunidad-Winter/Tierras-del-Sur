VERSION 5.00
Begin VB.Form frmEstadisticas 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Estadisticas"
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6855
   Icon            =   "FrmEstadisticas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   6270
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image command1 
      Height          =   210
      Index           =   43
      Left            =   3240
      Top             =   5520
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   42
      Left            =   5880
      Top             =   5520
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sastreria"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   22
      Left            =   3600
      TabIndex        =   40
      Top             =   5520
      Width           =   645
   End
   Begin VB.Image Label7 
      Height          =   315
      Left            =   5905
      Top             =   5889
      Width           =   855
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   41
      Left            =   3240
      Top             =   5280
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   39
      Left            =   3240
      Top             =   5010
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   37
      Left            =   3240
      Top             =   4815
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   1
      Left            =   3240
      Top             =   585
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   35
      Left            =   3240
      Top             =   4560
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   33
      Left            =   3240
      Top             =   4350
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   31
      Left            =   3240
      Top             =   4095
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   29
      Left            =   3240
      Top             =   3870
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   27
      Left            =   3240
      Top             =   3630
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   25
      Left            =   3240
      Top             =   3405
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   23
      Left            =   3240
      Top             =   3150
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   21
      Left            =   3240
      Top             =   2925
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   19
      Left            =   3240
      Top             =   2700
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   17
      Left            =   3240
      Top             =   2445
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   15
      Left            =   3240
      Top             =   2220
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   13
      Left            =   3240
      Top             =   1980
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   11
      Left            =   3240
      Top             =   1740
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   9
      Left            =   3240
      Top             =   1515
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   7
      Left            =   3240
      Top             =   1275
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   5
      Left            =   3240
      Top             =   1050
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   3
      Left            =   3240
      Top             =   825
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   40
      Left            =   5880
      Top             =   5280
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   38
      Left            =   5880
      Top             =   5010
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   36
      Left            =   5880
      Top             =   4815
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   34
      Left            =   5880
      Top             =   4560
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   32
      Left            =   5880
      Top             =   4350
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   30
      Left            =   5880
      Top             =   4095
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   28
      Left            =   5880
      Top             =   3870
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   26
      Left            =   5880
      Top             =   3630
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   24
      Left            =   5880
      Top             =   3405
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   22
      Left            =   5880
      Top             =   3150
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   20
      Left            =   5880
      Top             =   2925
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   18
      Left            =   5880
      Top             =   2700
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   16
      Left            =   5880
      Top             =   2445
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   14
      Left            =   5880
      Top             =   2220
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   12
      Left            =   5880
      Top             =   1980
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   10
      Left            =   5880
      Top             =   1740
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   8
      Left            =   5880
      Top             =   1515
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   6
      Left            =   5880
      Top             =   1275
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   4
      Left            =   5880
      Top             =   1050
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   2
      Left            =   5880
      Top             =   825
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   0
      Left            =   5880
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   39
      Top             =   6000
      Width           =   5535
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   38
      Top             =   5580
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   37
      Top             =   5340
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   36
      Top             =   5100
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   35
      Top             =   4860
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   34
      Top             =   4620
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   33
      Top             =   4380
      Width           =   2475
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   21
      Left            =   3600
      TabIndex        =   32
      Top             =   5280
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   20
      Left            =   3600
      TabIndex        =   31
      Top             =   5040
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   19
      Left            =   3585
      TabIndex        =   30
      Top             =   4815
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   18
      Left            =   3585
      TabIndex        =   29
      Top             =   4575
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   17
      Left            =   3585
      TabIndex        =   28
      Top             =   4350
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   16
      Left            =   3585
      TabIndex        =   27
      Top             =   4110
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   15
      Left            =   3585
      TabIndex        =   26
      Top             =   3870
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   14
      Left            =   3585
      TabIndex        =   25
      Top             =   3645
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   13
      Left            =   3585
      TabIndex        =   24
      Top             =   3405
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   7
      Left            =   285
      TabIndex        =   23
      Top             =   3660
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   6
      Left            =   285
      TabIndex        =   22
      Top             =   3420
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   5
      Left            =   285
      TabIndex        =   21
      Top             =   3180
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   4
      Left            =   285
      TabIndex        =   20
      Top             =   2940
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   285
      TabIndex        =   19
      Top             =   2700
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   285
      TabIndex        =   18
      Top             =   2475
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   285
      TabIndex        =   17
      Top             =   2235
      Width           =   900
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   12
      Left            =   3585
      TabIndex        =   16
      Top             =   3165
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   11
      Left            =   3585
      TabIndex        =   15
      Top             =   2940
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   10
      Left            =   3585
      TabIndex        =   14
      Top             =   2700
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   9
      Left            =   3585
      TabIndex        =   13
      Top             =   2460
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   8
      Left            =   3585
      TabIndex        =   12
      Top             =   2235
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   7
      Left            =   3585
      TabIndex        =   11
      Top             =   1995
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   6
      Left            =   3585
      TabIndex        =   10
      Top             =   1755
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   5
      Left            =   3585
      TabIndex        =   9
      Top             =   1530
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   4
      Left            =   3585
      TabIndex        =   8
      Top             =   1290
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   3
      Left            =   3585
      TabIndex        =   7
      Top             =   1050
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   2
      Left            =   3585
      TabIndex        =   6
      Top             =   825
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   1
      Left            =   3585
      TabIndex        =   5
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   5
      Left            =   300
      TabIndex        =   4
      Top             =   1365
      Width           =   390
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   4
      Left            =   300
      TabIndex        =   3
      Top             =   1155
      Width           =   390
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   300
      TabIndex        =   2
      Top             =   945
      Width           =   390
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   300
      TabIndex        =   1
      Top             =   735
      Width           =   390
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   300
      TabIndex        =   0
      Top             =   510
      Width           =   390
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Selecionado As Integer
Private flags() As Integer



Public Sub Iniciar_Labels()
'Iniciamos los labels con los valores de los atributos y los skills
Dim i As Integer
For i = 1 To NUMATRIBUTOS
    Atri(i).Caption = AtributosNames(i) & ": " & UserAtributos(i)
Next
For i = 1 To NUMSKILLS
    Skills(i).Caption = SkillsNames(i) & ": " & UserSkills(i)
Next
Label4(1).Caption = "Asesino: " & UserReputacion.AsesinoRep
Label4(2).Caption = "Bandido: " & UserReputacion.BandidoRep
Label4(3).Caption = "Burgues: " & UserReputacion.BurguesRep
Label4(4).Caption = "Ladrón: " & UserReputacion.LadronesRep
Label4(5).Caption = "Noble: " & UserReputacion.NobleRep
Label4(6).Caption = "Plebe: " & UserReputacion.PlebeRep


If UserEstadisticas.faccion = eAlineaciones.caos Then
    Label4(7).ForeColor = &H8080F0
    Label4(7).Caption = "Ejército Escarlata"
ElseIf UserEstadisticas.faccion = eAlineaciones.Real Then
    Label4(7).ForeColor = &HC0C000
    Label4(7).Caption = "Ejército Índigo"
Else
    Label4(7).ForeColor = &HB0AAA3
    Label4(7).Caption = "Rebelde"
End If

With UserEstadisticas
    Label6(0).Caption = "Escarlatas matados: " & .criminalesMatados
    Label6(1).Caption = "Índigos matados: " & .ciudadanosMatados
    Label6(2).Caption = "Usuarios matados: " & .UsuariosMatados
    Label6(3).Caption = "Criaturas matadas: " & .NpcsMatados
    Label6(4).Caption = "Clase: " & .Clase
    Label6(5).Caption = "Tiempo restante en carcel: " & .PenaCarcel
End With
Label8.Caption = "Nivel: " & UserLvl & " Experiencia: " & FormatNumber(UserExp, 0, vbTrue, vbFalse, vbTrue) & "/" & FormatNumber(UserPasarNivel, 0, vbTrue, vbFalse, vbTrue) & " Skills Libres: " & SkillPoints

i = 0
'Flags para saber que skills se modificaron
ReDim flags(1 To NUMSKILLS)
'Cargamos el jpg correspondiente
Alocados = SkillPoints
If Alocados > 0 Then
    For i = 0 To NUMSKILLS * 2 - 1
        If i Mod 2 = 0 Then
             DameImagen Command1(i), 116
        Else
            DameImagen Command1(i), 115
        End If
        Command1(i).Visible = True
Next
End If
End Sub

Private Sub Command1_Click(Index As Integer)
Call Sonido_Play(SND_CLICK)
Dim indice
Dim CantidadSkill As Byte


If Index Mod 2 = 0 Then
    If Alocados > 0 Then
        indice = Index \ 2 + 1
        CantidadSkill = val(mid(Skills(indice).Caption, InStr(1, Skills(indice).Caption, ": ", vbBinaryCompare) + 2))
        If indice > NUMSKILLS Then indice = NUMSKILLS
        If CantidadSkill < MAXSKILLPOINTS Then
            CantidadSkill = CantidadSkill + 1
            flags(indice) = flags(indice) + 1
            Alocados = Alocados - 1
            Skills(indice) = SkillsNames(indice) & ": " & CantidadSkill
        End If
    End If
Else
    If Alocados < SkillPoints Then
        indice = Index \ 2 + 1
         CantidadSkill = val(mid(Skills(indice).Caption, InStr(1, Skills(indice).Caption, ": ", vbBinaryCompare) + 2))
        If CantidadSkill > 0 And flags(indice) > 0 Then
            CantidadSkill = CantidadSkill - 1
            Skills(indice) = SkillsNames(indice) & ": " & CantidadSkill
            flags(indice) = flags(indice) - 1
            Alocados = Alocados + 1
        End If
    End If
End If
Label8.Caption = "Nivel: " & UserLvl & " Experiencia: " & FormatNumber(UserExp, 0, vbTrue, vbFalse, vbFalse) & "/" & FormatNumber(UserPasarNivel, 0, vbTrue, vbFalse, vbTrue) & " Skills Libres: " & Alocados

'.Caption = "Puntos:" & Alocados
End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Index <> Selecionado Then
    If Index Mod 2 = 0 Then
    If Selecionado >= 0 Then DameImagen Command1(Selecionado), 116: Skills(Int(Selecionado / 2 + 1)).FontBold = False
    DameImagen Command1(Index), 118
    Else
    If Selecionado >= 0 Then DameImagen Command1(Selecionado), 115: Skills(Int(Selecionado / 2 + 1)).FontBold = False
    DameImagen Command1(Index), 117
    End If
    Selecionado = Index
    Skills(Int(Selecionado / 2 + 1)).FontBold = True
End If

End Sub

Private Sub Form_Load()
DameImagenForm Me, 102
Call CambiarCursor(frmEstadisticas)
'Alocados = SkillPoints
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Selecionado >= 0 Then
        If Selecionado Mod 2 = 0 Then
        Skills(Int(Selecionado / 2 + 1)).FontBold = False
        DameImagen Command1(Selecionado), 116
        'Command1(Selecionado).Picture = LoadPicture(App.Path & "\Graficos\+.jpg")
        Else
        Skills(Int(Selecionado / 2 + 1)).FontBold = False
        DameImagen Command1(Selecionado), 115
        End If
Selecionado = -1
End If

If Label7.tag = "1" Then
Label7.tag = 0
Label7.Picture = Nothing
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub Label7_Click()
Dim i As Integer
Dim cad As String
For i = 1 To NUMSKILLS
    cad = cad & ByteToString(flags(i))
Next

EnviarPaquete Paquetes.SkillMod, cad
If Alocados = 0 Then frmMain.Label1.Visible = False
SkillPoints = Alocados
Unload Me
meves = False
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label7.tag <> "1" Then
Label7.tag = 1
Call DameImagen(Me.Label7, 4)
End If
End Sub
