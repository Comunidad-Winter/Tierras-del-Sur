VERSION 5.00
Begin VB.Form frmVertexShader 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cumbia shader ninja"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   4440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1200
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Compilar shaders"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   4560
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmVertexShader.frx":0000
      Left            =   240
      List            =   "frmVertexShader.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   8295
   End
   Begin VB.Label Label3 
      Caption         =   "Shader:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   "Pixel shader:"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Vertex shader:"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   4095
   End
End
Attribute VB_Name = "frmVertexShader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
    'Text1.text = PixelShaderCatalog(Combo1.listIndex).codigoVertexShader
End Sub

Private Sub Combo1_LostFocus()
If Combo1.listIndex >= 0 Then
Text1.text = PixelShaderCatalog(Combo1.listIndex).codigoVertexShader
Text2.text = PixelShaderCatalog(Combo1.listIndex).codigo
End If
End Sub

Private Sub Command1_Click()
    PixelShaderCatalog(Combo1.listIndex).codigoVertexShader = Text1.text
    PixelShaderCatalog(Combo1.listIndex).codigo = Text2.text
    Engine_PixelShaders.Engine_PixelShaders_EngineReiniciado
End Sub

