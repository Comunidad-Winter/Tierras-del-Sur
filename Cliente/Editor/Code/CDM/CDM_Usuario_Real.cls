VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDM_Usuario_Real"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements CDM_Usuario

Private m_Nombre As String          ' Nombre del Usuario
Private m_ID As Long                ' Identificador unico numerico
Private m_Persona_Nombre As String  ' Nombre real de la persona
Private m_Correo As String          ' Correo electronico de la Persona

Private m_Permisos As Dictionary    ' Privilegios

Public Sub CDM_Usuario_crear(config As Dictionary)
    ' Datos Usuario
    m_Nombre = config.item("nombre")
    m_ID = CLng(val(config.item("id")))

    If VarType(config.item("permisos")) = 9 Then
        Set m_Permisos = config.item("permisos")
    End If
    
    
    ' Datos Persona
    m_Correo = config.item("correo")
    m_Persona_Nombre = config.item("persona_nombre")
End Sub

Private Sub reset()
    m_Nombre = ""
    m_Persona_Nombre = ""
    m_Correo = ""
    m_ID = 0
End Sub

Private Property Get CDM_Usuario_PersonaNombre() As String
    CDM_Usuario_PersonaNombre = m_Persona_Nombre
End Property

Public Function CDM_Usuario_tienePermisos(seccion As String, ByVal permisos As ePermisosCDM) As Boolean
    CDM_Usuario_tienePermisos = False
    
    If m_Permisos Is Nothing Then Exit Function
    
    If Not m_Permisos.Exists(seccion) Then Exit Function
    
    CDM_Usuario_tienePermisos = permisos And m_Permisos.item(seccion)
End Function
' Propiedades
Public Property Get CDM_Usuario_Correo() As String
    CDM_Usuario_Correo = m_Correo
End Property

Public Property Get CDM_Usuario_ID() As Long
    CDM_Usuario_ID = m_ID
End Property

Public Property Get CDM_Usuario_nombre() As String
    CDM_Usuario_nombre = m_Nombre
End Property
