VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Inteligencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Sub determinarAtaque(NpcIndex As Integer, indexAtaque As Integer, ByRef tipoObjetivo As eTipoObjetivo, ByRef tipoAtaque As eTipoAtaque, ByRef subTipo As Byte)

End Sub

Public Sub determinarObjetivo(NpcIndex As Integer, ByRef Index As Integer, ByRef tipoObjetivo As eTipoObjetivo)

End Sub

Public Function determinarMovimiento(NpcIndex As Integer, tipoObjetivo As eTipoObjetivo, IndexObjetivo As Integer) As eHeading

End Function

