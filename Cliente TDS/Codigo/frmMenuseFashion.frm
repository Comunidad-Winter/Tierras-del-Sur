VERSION 5.00
Begin VB.Form frmMenuseFashion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   1410
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   1410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Label2"
      ForeColor       =   &H80000014&
      Height          =   255
      Index           =   1
      Left            =   540
      TabIndex        =   2
      Top             =   900
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Label2"
      Height          =   195
      Index           =   0
      Left            =   540
      TabIndex        =   1
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1395
   End
End
Attribute VB_Name = "frmMenuseFashion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Super Menuse Fashion
'Creado por Alejandro "AlejoLp" Santos
'
'Util para evitar el bloqueo de la linea de ejecion
'normal de un programa, caso caracteristico al usar
'los menuses estandar de windows.
'********************Misery_Ezequiel 28/05/05********************'
Option Explicit

#If (ConMenuseConextuales = 1) Then

''** FUNCION CALLBACK **''

''Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)
''
''End Sub

''** CODIGO DE EJEMPLO DE USO **''

''Dim I As Long
''Dim M As New frmMenuseFashion
''
''Load M
''M.SetCallback Me
''M.SetMenuId 12
''M.ListaInit 3, False
''For I = 0 To 2
''    M.ListaSetItem I, "hgfsg " & I
''Next I
''M.ListaFin
''M.Show , Me

Private Type tMenuElemento
    Texto As String
    Bold As Boolean
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Dim DifX As Long, DifY As Long
Dim UltPos As Long
Dim Callback As Object
Dim MenuId As Long
Dim Elementos() As tMenuElemento
Dim mCantidad As Long
Dim MaxAncho As Long, MaxAlto As Long
Dim YaCargado As Long

Private Sub Form_Activate()
If Me.left + Me.Width > Screen.Width Then Me.left = Screen.Width - Me.Width
If Me.top + Me.Height > Screen.Height Then Me.top = Screen.Height - Me.Height
End Sub

Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_Load()
Dim p As POINTAPI
Dim I As Long

'Puajjjj :P Shhh... nadie me ve ^_^
If YaCargado <> 89345 Then
    YaCargado = 89345
    
    Call GetCursorPos(p)
    
    Me.left = p.X * Screen.TwipsPerPixelX
    Me.top = p.Y * Screen.TwipsPerPixelY
    
    DifX = Me.Width - Me.ScaleWidth
    DifY = Me.Height - Me.ScaleHeight
    
    MaxAncho = 0
    UltPos = -1
    ReDim Elementos(0 To 0)
    
    Label2(0).Font = Label1(0).Font
End If
End Sub

Private Sub Form_LostFocus()
Unload Me
End Sub

Private Sub Label1_Click(Index As Integer)
Me.Hide
On Local Error Resume Next
    Call Callback.CallbackMenuFashion(MenuId, Index)
On Local Error GoTo 0
Unload Me
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If UltPos <> Index Then
    If UltPos <> -1 Then
        Label1(UltPos).BackStyle = 0
        Label1(UltPos).ForeColor = Label2(0).ForeColor
    End If
    Label1(Index).BackStyle = 1
    Label1(Index).ForeColor = Label2(1).ForeColor
    UltPos = Index
End If
End Sub

Public Sub SetMenuId(ByVal MeID As Long)
MenuId = MeID
End Sub

Public Sub ListaInit(ByVal cantidad As Long, Optional ByVal Mantener As Boolean = False)
If cantidad >= 1 Then
    mCantidad = cantidad
    If Mantener = False Then
        ReDim Elementos(0 To cantidad - 1)
    Else 'true
        ReDim Preserve Elementos(0 To cantidad - 1)
    End If
End If
End Sub

Public Sub ListaSetItem(ByVal N As Long, ByVal Texto As String, Optional ByVal Bold As Boolean = False)
If N >= LBound(Elementos) And N <= UBound(Elementos) Then
    Elementos(N).Texto = Texto
    Elementos(N).Bold = Bold
    With Label1(0)
        .AutoSize = True
        .FontBold = Bold
        .Caption = Texto
        If .Width > MaxAncho Then MaxAncho = .Width
        If .Height > MaxAlto Then MaxAlto = .Height
        .AutoSize = False
    End With
End If
End Sub

Public Sub ListaFin()
Dim I As Long
MaxAncho = MaxAncho + 2 * Screen.TwipsPerPixelX
MaxAlto = MaxAlto + 2 * Screen.TwipsPerPixelX

For I = 0 To mCantidad - 1
    If I <> 0 Then
        Load Label1(I)
    End If
    
    Label1(I).Visible = True
    Label1(I).left = 0
    Label1(I).Width = MaxAncho
    Label1(I).Height = MaxAlto
    Label1(I).top = I * MaxAlto
    Label1(I).Caption = Elementos(I).Texto
    Label1(I).FontBold = Elementos(I).Bold
    Label1(I).BackStyle = 0
    Label1(I).BackColor = Label1(0).BackColor
Next I
Me.Height = (UBound(Elementos) - LBound(Elementos) + 1) * Label1(0).Height + DifY
Me.Width = MaxAncho + DifX
End Sub

Public Sub SetCallback(C As Object)
Set Callback = C
End Sub
#End If
'********************Misery_Ezequiel 28/05/05********************'
