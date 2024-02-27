VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iAccionEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/***********************************************************/
Public Function getDescripcion() As String
End Function

Public Function getNombre() As String
End Function

Public Function getTIPO() As Long
End Function
Public Function getNombreExtendido() As String
End Function
Public Function getID() As Integer
End Function
Public Sub setNombre(nombre As String)
End Sub
'**********************************************************/
'********** FUNCIONES DE PERSISTENCIA *********************/
Public Function persistir(archivoSalida As Integer)
End Function

Public Function cargar(archivoFuente As Integer)
End Function

Public Function generarAccionReal() As iAccion
End Function
