VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iCondicionEvento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function puedeIngresarEquipo(integrantesIndex() As Integer) As Boolean
End Function

Public Sub asentarIngresoEquipo(integrantesIndex() As Integer)
End Sub
Public Sub removerEquipoInscripto(integrantes() As Long)
End Sub

Public Sub setMaximaMemoria(slots As Byte)
End Sub

Public Function obtenerDescripcion() As String
End Function
