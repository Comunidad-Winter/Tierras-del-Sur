VERSION 5.00
Object = "{50CBA22D-9024-11D1-AD8F-8E94A5273767}#8.7#0"; "TranImg2.ocx"
Begin VB.Form frmMapa 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "frmMapa"
   ClientHeight    =   9615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9165
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   641
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   611
   ShowInTaskbar   =   0   'False
   Begin DevPowerTransImg.TransImg pictureCerrar 
      Height          =   240
      Left            =   8775
      TabIndex        =   0
      Top             =   60
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      AutoSize        =   0   'False
      Transparent     =   -1  'True
   End
   Begin DevPowerTransImg.TransImg imgmover 
      Height          =   240
      Left            =   8400
      TabIndex        =   1
      Top             =   60
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      AutoSize        =   0   'False
      MaskColor       =   0
      MousePointer    =   15
      Transparent     =   -1  'True
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Sub Form_Load()
    DameImagenForm Me, 651
    Set Me.pictureCerrar.Picture = LoadPicture(app.Path & "/Recursos/cerrar.jpg")
    Set Me.imgmover.Picture = LoadPicture(app.Path & "/Recursos/mover.jpg")
End Sub

Private Sub imgmover_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        Call ReleaseCapture
        Call SendMessage(frmMapa.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub

Private Sub pictureCerrar_Click()
    Unload Me
End Sub
