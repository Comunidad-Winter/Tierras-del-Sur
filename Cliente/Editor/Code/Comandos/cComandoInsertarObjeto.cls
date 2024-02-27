VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComandoInsertarObjeto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iComando

Private objetoAnterior As obj
Private objetoNuevo As obj
Private posx As Byte
Private posy As Byte


Public Sub crear(x As Integer, y As Integer, cantidad As Integer, indexObjeto As Integer)

    posx = x
    posy = y
    
    objetoNuevo.Amount = cantidad
    objetoNuevo.objIndex = indexObjeto

End Sub
Private Sub iComando_desHacer()
    Dim GrhIndex As Long
    
    MapData(posx, posy).OBJInfo = objetoAnterior
    
    If objetoAnterior.objIndex > 0 Then
        GrhIndex = ObjData(objetoAnterior.objIndex).GrhIndex
        InitGrh MapData(posx, posy).ObjGrh, GrhIndex
    Else
        MapData(posx, posy).ObjGrh.GrhIndex = 0
    End If
End Sub

Private Sub iComando_hacer()
    objetoAnterior = MapData(posx, posy).OBJInfo
    MapData(posx, posy).OBJInfo = objetoNuevo
    
    If objetoNuevo.objIndex > 0 Then
        InitGrh MapData(posx, posy).ObjGrh, ObjData(objetoNuevo.objIndex).GrhIndex
    Else
        MapData(posx, posy).ObjGrh.GrhIndex = 0
    End If
End Sub

Private Function iComando_obtenerNombre() As String
    Dim nombre As String
    
    If objetoNuevo.objIndex > 0 Then
        nombre = ObjData(objetoNuevo.objIndex).Name
        iComando_obtenerNombre = "Agregar " & nombre & "[" & objetoNuevo.Amount & "] en (" & posx & "," & posy & ")"
    Else
        If objetoAnterior.objIndex > 0 Then
            nombre = ObjData(objetoAnterior.objIndex).Name
            iComando_obtenerNombre = "Borrar " & nombre & "[" & objetoAnterior.Amount & "] en (" & posx & "," & posy & ")"
        End If
    End If
End Function
