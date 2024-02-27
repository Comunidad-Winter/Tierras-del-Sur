VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CrearPersonajeDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Nombre As String
Public Genero As eGenero
Public Clase As eClass
Public Alineacion As Byte
Public Raza As eRaza
Public Contraseña As String

Public headId As Integer
Public bodyId As Integer
Public ropainteriorId As Integer
Public barbaId As Integer
Public peloId As Integer
