VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDM_Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub crear(config As Dictionary)
End Sub

Public Function tienePermisos(seccion As String, ByVal permisos As ePermisosCDM) As Boolean
    If seccion = "EDITOR.CDM" Then
        tienePermisos = False
        Exit Function
    End If
   
   tienePermisos = True
End Function
' Propiedades
Public Property Get Correo() As String
    Correo = "test@test.com.ar"
End Property

Public Property Get PersonaNombre() As String
    PersonaNombre = ""
End Property

Public Property Get ID() As Long
    ID = 0
End Property

Public Property Get nombre() As String
    nombre = "Usuario Dummy"
End Property
