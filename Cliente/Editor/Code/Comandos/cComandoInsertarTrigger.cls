VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComandoInsertarTrigger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iComando

Private triggerAnterior As Long
Private TriggerAplicar As Long
Private posX As Byte
Private posY As Byte

Private Sub iComando_desHacer()

    MapData(posX, posY).Trigger = triggerAnterior
    
End Sub

Public Sub crear(x As Integer, y As Integer, Trigger As Long)
    posX = x
    posY = y
    TriggerAplicar = Trigger
End Sub
Private Sub iComando_hacer()

    triggerAnterior = MapData(posX, posY).Trigger
    MapData(posX, posY).Trigger = TriggerAplicar
    
End Sub

Private Function iComando_obtenerNombre() As String
    iComando_obtenerNombre = "Trigger " & ME_Tools_Triggers.obtenerDescripcionAbreviatura(TriggerAplicar) & _
                            " en (" & posX & "," & posY & ")"
End Function




