VERSION 5.00
Begin VB.UserControl TextConListaConBuscador 
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1725
   ScaleWidth      =   2250
   Begin EditorTDS.ListaConBuscador ListaConBuscador 
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   3201
   End
   Begin VB.TextBox txtCampo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "TextConListaConBuscador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_cantidadLineas As Byte

Private tagInfo As String
Public Event Change(valor As String, id As Integer)

Public Property Let CantidadLineasAMostrar(ByVal cantidad As Byte)
    m_cantidadLineas = cantidad
End Property

Public Property Get CantidadLineasAMostrar() As Byte
    CantidadLineasAMostrar = m_cantidadLineas
End Property

Public Function seleccionarID(ByVal id As Integer) As Boolean
    If listaConBuscador.seleccionarID(id) Then
        txtCampo.Text = listaConBuscador.obtenerValor
        seleccionarID = True
    Else
        seleccionarID = False
    End If
End Function

Public Sub desseleccionar()
    listaConBuscador.deseleccionar
    txtCampo.Text = ""
End Sub

Public Sub limpiarLista()
    listaConBuscador.vaciar
End Sub

Private Sub ListaConBuscador_Change(valor As String, id As Integer)
    txtCampo.Text = valor
    
    RaiseEvent Change(valor, id)
End Sub

Public Function obtenerIDValor() As Integer
    obtenerIDValor = listaConBuscador.obtenerIDValor
End Function

Public Function obtenerValor() As String
     obtenerValor = listaConBuscador.obtenerValor
End Function

Private Sub ListaConBuscador_DblClic()
    Call OcultarLista
End Sub

Private Sub OcultarLista()
    listaConBuscador.visible = False
    redimensionar
End Sub

Private Sub ListaConBuscador_Click()

End Sub

Private Sub ListaConBuscador_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        OcultarLista
    End If
End Sub

Private Sub ListaConBuscador_LostFocus()
    OcultarLista
End Sub

Private Sub txtCampo_GotFocus()
    listaConBuscador.visible = True
    redimensionar
    'Establece el foco en la lista
    listaConBuscador.SetFocus
End Sub


Public Sub addString(ByVal id As Integer, ByVal contenido As String)
    Call listaConBuscador.addString(id, contenido)
End Sub

Private Sub redimensionar()
    Dim longitudLineas As Byte
    
    longitudLineas = IIf(m_cantidadLineas > listaConBuscador.getCantidadElementos, listaConBuscador.getCantidadElementos, m_cantidadLineas)
    
    If listaConBuscador.visible Then
        UserControl.Height = 210 * longitudLineas + txtCampo.Height
        listaConBuscador.Height = UserControl.Height
    Else
        UserControl.Height = txtCampo.Height
    End If
End Sub

Private Sub UserControl_Resize()
    listaConBuscador.Width = UserControl.Width
    txtCampo.Width = UserControl.Width
End Sub

Private Sub UserControl_Show()
    listaConBuscador.visible = False
    redimensionar
End Sub

Public Property Get Enabled() As Boolean
   Enabled = txtCampo.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    txtCampo.Enabled = vNewValue
End Property

Public Property Get tag() As String
    tag = tagInfo
End Property

Public Property Let tag(ByVal vNewValue As String)
    tagInfo = vNewValue
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "CantidadLineasAMostrar", Me.CantidadLineasAMostrar, 5
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_cantidadLineas = PropBag.ReadProperty("CantidadLineasAMostrar", 5)
End Sub
