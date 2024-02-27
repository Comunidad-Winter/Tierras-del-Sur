VERSION 5.00
Begin VB.UserControl TextBoxConValidador 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   375
   ScaleWidth      =   1830
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "TextBoxConValidador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private valor As String
Private longitudMaxima As Integer
Private longitudMinima As Integer

Private tagInfo As String

Private Sub Text1_Change()
    Call establecerValor
End Sub

Private Function establecerValor()
    If Len(Text1.Text) >= longitudMinima And Len(Text1.Text) <= longitudMaxima Then
        Text1.BackColor = vbWhite
    Else
        Text1.BackColor = vbYellow
    End If
End Function
Private Sub UserControl_Resize()
    Text1.Width = UserControl.Width
    Text1.Height = UserControl.Height
End Sub

Property Get Text() As String
    Text = Text1.Text
End Property

Property Let Text(valor_ As String)
    Text1.Text = valor_
    Call establecerValor
End Property

Property Get MinLength() As Double
    MinLength = longitudMinima
End Property

Property Let MinLength(valor As Double)
    longitudMinima = valor
End Property

Property Get MaxLength() As Double
    MaxLength = longitudMaxima
End Property

Property Let MaxLength(valor As Double)
    longitudMaxima = valor
    Text1.MaxLength = longitudMaxima
End Property

Public Property Get Enabled() As Boolean
   Enabled = Text1.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    Text1.Enabled = vNewValue
End Property

Public Property Get tag() As String
   tag = tagInfo
End Property

Public Property Let tag(ByVal vNewValue As String)
    tagInfo = vNewValue
End Property
