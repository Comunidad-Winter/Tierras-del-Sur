VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmInformeMapa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe del Mapa"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInformeMapa.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   9340
      _Version        =   393217
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmInformeMapa.frx":1CCA
   End
   Begin VB.Label lblInforme 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Informe de elementos en el mapa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3585
   End
End
Attribute VB_Name = "frmInformeMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.RichTextBox1.TextRTF = modInformeMapa.obtenerInforme()
End Sub
