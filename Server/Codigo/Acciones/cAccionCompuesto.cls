VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAccionCompuesta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iAccion

Private Const vTRUE = 0
Private Const vFALSE = 1

Private subAcciones() As iAccion
Private ejecutarSiempre() As Byte


Private Sub Class_Initialize()
    ReDim ejecutarSiempre(1 To 1)
    ReDim subAcciones(1 To 1)
End Sub

Private Sub iAccion_cargar(archivoSalida As Integer)

End Sub

Private Sub iAccion_cargarParametros(parametros As Collection)

End Sub

Private Function iAccion_clonar() As iAccion
    Set iAccion_clonar = Nothing
End Function

Private Function iAccion_ejecutar(mapa As Integer, x_ As Byte, y_ As Byte) As Boolean

    Dim accion As Byte
    Dim resultado As Boolean
    
    For accion = 1 To UBound(subAcciones)
        
        resultado = subAcciones(accion).ejecutar(mapa, x_, y_)
        
        'Si la ejecución de la ultima accion no resulto verdadera
        'y la próxima acción necesitaba que resulte verdadera, entonces termino
        If Not resultado And ejecutarSiempre(accion) = vFALSE Then
            iAccion_ejecutar = False
            Exit For
        End If
    Next
    
    iAccion_ejecutar = True
End Function

Private Function iAccion_getID() As Long
    iAccion_getID = tipo_accion.COMPUESTA
End Function

Public Sub agregarHijo(accion As iAccion, ejecutarSiempre_ As Boolean)
    Dim cantidadActual As Byte
    
    cantidadActual = UBound(subAcciones())
    
    ReDim Preserve subAcciones(1 To cantidadActual + 1)
    ReDim Preserve ejecutarSiempre(1 To cantidadActual + 1)
    
    ejecutarSiempre(cantidadActual) = ejecutarSiempre_
    Set subAcciones(cantidadActual + 1) = accion
End Sub

Private Sub iAccion_persistir(archivoSalida As Integer)

End Sub
