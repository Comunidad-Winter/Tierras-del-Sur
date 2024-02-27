VERSION 5.00
Begin VB.UserControl UpDownText 
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   420
   ScaleWidth      =   915
   Begin VB.CommandButton cmd 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Index           =   0
      Left            =   610
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   210
      Width           =   300
   End
   Begin VB.CommandButton cmd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Index           =   2
      Left            =   610
      TabIndex        =   1
      Top             =   0
      Width           =   300
   End
   Begin VB.TextBox UpDownText 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Text            =   "0"
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "UpDownText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private valorMinimo As Double
Private valorMaximo As Double
Private valor As Double


Private anteriorValorScroll As Integer

Private Const BTN_MENOS As Byte = 0
Private Const BTN_MAS As Byte = 2

Private omitirChange As Boolean

Public Event change(valor As Double)
Private tagInfo As String


Private Function valorValido(valor As Double) As Boolean
    valorValido = (valor >= valorMinimo And valor <= valorMaximo)
End Function

Private Sub establecerValor(NuevoValor As Double)

    valor = NuevoValor
           
    cmd(BTN_MENOS).Enabled = True
    cmd(BTN_MAS).Enabled = True
    
    omitirChange = True
    UpDownText.text = valor
    omitirChange = False
    
    If valor <= valorMinimo Then
        cmd(BTN_MENOS).Enabled = False
    ElseIf valor >= valorMaximo Then
        cmd(BTN_MAS).Enabled = False
    End If

    RaiseEvent change(NuevoValor)
End Sub
Private Sub cmd_Click(Index As Integer)
    Dim NuevoValor As Double
    
    NuevoValor = valor + (Index - 1)
   
    Call establecerValor(NuevoValor)
End Sub

Private Sub actualizarcolor()
    If valorValido(val(UpDownText.text)) Then
        UpDownText.BackColor = vbWhite
    Else
        UpDownText.BackColor = vbYellow
    End If
End Sub

Private Sub UpDownText_Change()
    
    Call actualizarcolor
  
    
    If omitirChange Then Exit Sub
    Call establecerValor(val(UpDownText.text))
End Sub

Private Sub UpDownText_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
    Case vbKey0 To vbKey9
    Case vbKeyBack, vbKeyClear, vbKeyDelete
    Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
    Case vbKeySubtract, 45
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub UserControl_Resize()
    UpDownText.width = UserControl.width - cmd(BTN_MAS).width
    cmd(BTN_MAS).left = UserControl.width - cmd(BTN_MENOS).width + 1
    cmd(BTN_MENOS).left = UserControl.width - cmd(BTN_MENOS).width + 1
    
    UpDownText.height = UserControl.height - 1
    cmd(BTN_MAS).height = UpDownText.height / 2
    cmd(BTN_MENOS).height = UpDownText.height / 2
    
    cmd(BTN_MENOS).top = cmd(BTN_MAS).top + cmd(BTN_MAS).height + 1
    
End Sub

Property Get value() As Double
    value = valor
End Property

Property Get text() As String
    value = CStr(valor)
End Property

Property Let value(valor_ As Double)
     Call establecerValor(valor_)
End Property

Property Get MinValue() As Double
    MinValue = valorMinimo
    actualizarcolor
End Property

Property Let MinValue(valor As Double)
    valorMinimo = valor
    actualizarcolor
End Property

Property Get MaxValue() As Double
    MaxValue = valorMaximo
    actualizarcolor
End Property

Property Let MaxValue(valor As Double)
    valorMaximo = valor
    actualizarcolor
End Property

Public Property Get Enabled() As Boolean
   Enabled = UpDownText.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    UpDownText.Enabled = vNewValue
    cmd(BTN_MAS).Enabled = vNewValue
    cmd(BTN_MENOS).Enabled = vNewValue
End Property

Public Property Get tag() As String
   tag = tagInfo
End Property

Public Property Let tag(ByVal vNewValue As String)
    tagInfo = vNewValue
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "MaxValue", Me.MaxValue
    PropBag.WriteProperty "MinValue", Me.MinValue
    PropBag.WriteProperty "Enabled", Me.Enabled
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Me.MaxValue = CDbl(PropBag.ReadProperty("MaxValue", 0))
    Me.MinValue = CDbl(PropBag.ReadProperty("MinValue", 0))
    Me.Enabled = CBool(PropBag.ReadProperty("Enabled", True))
End Sub
