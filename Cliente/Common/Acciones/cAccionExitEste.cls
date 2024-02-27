VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAccionExitEste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iAccion

Private MapaDestino As Integer 'Solo necesito sabaer al mapa al cual lo tengo que enviar

Public Sub crear(mapaDestino_ As Integer)
    MapaDestino = mapaDestino_
End Sub
Private Function iAccion_getID() As Long
    iAccion_getID = tipo_accion.EXIT_ESTE
End Function
Private Sub iAccion_persistir(archivoSalida As Integer)
    Put archivoSalida, , CInt(iAccion_getID)
    Put archivoSalida, , MapaDestino
End Sub
Private Sub iAccion_cargarParametros(parametros As Collection)
    Call crear(CInt(parametros.Item(1).getValor))
End Sub
Private Function iAccion_clonar() As iAccion
    Set iAccion_clonar = New cAccionExitNorte
End Function
