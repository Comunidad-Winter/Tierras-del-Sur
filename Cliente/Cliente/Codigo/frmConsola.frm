VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{50CBA22D-9024-11D1-AD8F-8E94A5273767}#8.7#0"; "TranImg2.ocx"
Begin VB.Form frmConsola 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   7245
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConsola.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   471
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   483
   ShowInTaskbar   =   0   'False
   Begin DevPowerTransImg.TransImg pictureCerrar 
      Height          =   240
      Left            =   4200
      TabIndex        =   2
      Top             =   3480
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      AutoSize        =   0   'False
      Transparent     =   -1  'True
   End
   Begin DevPowerTransImg.TransImg imgmover 
      Height          =   240
      Left            =   45
      TabIndex        =   1
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      AutoSize        =   0   'False
      MaskColor       =   0
      MousePointer    =   15
      Transparent     =   -1  'True
   End
   Begin RichTextLib.RichTextBox ConsolaFlotante 
      Height          =   6900
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   12171
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmConsola.frx":0E42
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
End
Attribute VB_Name = "frmConsola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Sub ConsolaFlotante_Change()
        ConsolaFlotante.Visible = True
        ConsolaFlotante.Refresh
End Sub

Private Sub Form_Load()
    Set Me.imgmover.Picture = LoadPicture(app.Path & "/Recursos/mover.jpg")
    Set Me.pictureCerrar.Picture = LoadPicture(app.Path & "/Recursos/cerrar.jpg")
    
    If ConsolaWidth > 0 And ConsolaHeight > 0 Then
        Dim top As Integer
        
        If ConsolaTop + ConsolaHeight > Screen.height Then
            Me.top = Screen.height - ConsolaHeight
        Else
            Me.top = ConsolaTop
        End If
        
        If ConsolaLeft + ConsolaWidth > Screen.width Then
            Me.left = Screen.width - ConsolaWidth
        Else
            Me.left = ConsolaLeft
        End If

        Me.width = ConsolaWidth
        Me.height = ConsolaHeight
    Else

   End If
End Sub

Private Sub Form_Resize()
    ConsolaFlotante.height = frmConsola.ScaleHeight
    ConsolaFlotante.width = frmConsola.ScaleWidth
    
    pictureCerrar.left = frmConsola.ScaleWidth - pictureCerrar.width - 25
    pictureCerrar.top = 8
    
    imgmover.left = frmConsola.ScaleWidth - imgmover.width - 50
    imgmover.top = 8
    ConsolaFlotante.SelStart = Len(ConsolaFlotante.text)
End Sub
Private Sub imgmover_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        Call ReleaseCapture
        Call SendMessage(frmConsola.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub

Private Sub pictureCerrar_Click()

    ConsolaTop = Me.top
    ConsolaLeft = Me.left
    ConsolaWidth = Me.width
    ConsolaHeight = Me.height
    
    Call Configuracion_Usuario.guardarConfiguracion
    Unload Me
End Sub

