VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit


Private nombre As String
Private IDPJ As Long
Private fecha As Date
Private descripcion As String

Public Sub iniciar(NombreUser_ As String, IDUser_ As Long, descripcion_ As String, FechaSol_ As Date)
    nombre = NombreUser_
    IDPJ = IDUser_
    fecha = FechaSol_
    descripcion = descripcion_
End Sub
Public Function getNombre() As String
    getNombre = UCase(nombre)
End Function
Public Function getIDPJ() As Long
    getIDPJ = IDPJ
End Function

Public Function getFecha() As Date
    getFecha = fecha
End Function

Public Function getDescripcion() As String
    getDescripcion = descripcion
End Function


