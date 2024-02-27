VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComandoCompuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iComando

Private nombre As String
Private subComandos As Collection

Private Sub Class_Initialize()
    Set subComandos = New Collection
End Sub

Private Sub iComando_desHacer()
    Dim comando As iComando
    Dim nComando As Long
    
    'Se tiene que deshacer al revez de como se hace. Sino no se deshace correctamente en los casos
    'de que las acciones modifiquen el mismo tile en un mismo aspecto (por ejemplo un trigger).
    For nComando = subComandos.Count To 1 Step -1
        Set comando = subComandos.item(nComando)
        Call comando.desHacer
    Next
End Sub

Private Sub iComando_hacer()
    Dim comando As iComando
    
    For Each comando In subComandos
        Call comando.hacer
    Next
End Sub


Private Function iComando_obtenerNombre() As String
    iComando_obtenerNombre = nombre
End Function

Public Sub SetNombre(nombre_ As String)
    nombre = nombre_
End Sub
Public Sub agregarComando(comando As iComando)
    If Not comando Is Nothing Then
        Call subComandos.Add(comando)
    End If
End Sub

Public Sub crear(conjuntoDeComandos As Collection, nombreComando As String)
    nombre = nombreComando
    If Not conjuntoDeComandos Is Nothing Then
        Set subComandos = conjuntoDeComandos
    End If
End Sub

Private Sub iAccion_cargarParametros(parametros As Collection)
End Sub
