VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComandoInsertarAccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iComando

Private posX As Integer
Private posY As Integer
Private accion As iAccionEditor
Private accionAnterior As iAccionEditor

Public Sub crear(X As Integer, Y As Integer, accion_ As iAccionEditor)
    Set accion = accion_
    posX = X
    posY = Y
End Sub
Private Sub iComando_desHacer()
    Set MapData(posX, posY).accion = accionAnterior
End Sub
Private Sub iComando_hacer()
    
    Set accionAnterior = MapData(posX, posY).accion
    Set MapData(posX, posY).accion = accion

End Sub

Private Function iComando_obtenerNombre() As String
    
    If accion Is Nothing Then
        iComando_obtenerNombre = "Remover"
    Else
        iComando_obtenerNombre = "Insertar " & accion.GetNombre
    End If
    
    If posX > 0 And posY > 0 Then
        iComando_obtenerNombre = iComando_obtenerNombre & " en (" & posX & "," & posY & ")"
    Else
        iComando_obtenerNombre = iComando_obtenerNombre
    End If
End Function
