VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Texto As clsGUIText
Private m_intOrden As Integer
Private m_StrData As String
Private m_bVisible As Boolean
Private m_bSeleccionado As Boolean


Public Property Get Orden() As Integer
    Orden = m_intOrden
End Property

Public Property Let Orden(ByVal intValue As Integer)
    m_intOrden = intValue
End Property

' Variable to hold 'Orden' property value
Public Property Get Data() As String
    Data = m_StrData
End Property

Public Property Let Data(ByVal StrValue As String)
    m_StrData = StrValue
End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_UserMemId = 0
    Visible = m_bVisible
End Property

Public Property Let Visible(ByVal bValue As Boolean)
    m_bVisible = bValue
End Property

Public Property Get seleccionado() As Boolean
    seleccionado = m_bSeleccionado
End Property

Public Property Let seleccionado(ByVal bValue As Boolean)
    m_bSeleccionado = bValue
End Property

Public Property Let texto(ByVal texto_ As String)
    m_Texto.text = texto_
End Property

Public Property Get texto() As String
    texto = m_Texto.text
End Property

Public Property Get clsGUIText() As clsGUIText
    Set clsGUIText = m_Texto
End Property

' Variable to hold 'seleccionado' property value
Private Sub Class_Initialize()
    m_bSeleccionado = False
    m_bVisible = True
    m_StrData = ""
    
    Set m_Texto = New clsGUIText
    m_Texto.text = ""
    m_intOrden = 0
End Sub
