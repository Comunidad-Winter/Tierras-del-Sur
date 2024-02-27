VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAccionExitOeste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements iAccion

Private mapaDestino As Integer 'Solo necesito sabaer al mapa al cual lo tengo que enviar

Public Sub crear(mapaDestino_ As Integer)
    mapaDestino = mapaDestino_
End Sub

Private Sub iAccion_cargar(archivoSalida As Integer)
    Get archivoSalida, , mapaDestino
End Sub

Private Function iAccion_ejecutar(mapa As Integer, x_ As Byte, y_ As Byte) As Boolean
UserIndex = MapData(mapa, x_, y_).UserIndex

iAccion_ejecutar = SV_Desplazamientos.avanzarPersonajeOtroMapa(UserList(UserIndex), mapaDestino, X_MAXIMO_USABLE, CByte(UserList(UserIndex).pos.y))

End Function

Private Function iAccion_getID() As Long
    iAccion_getID = tipo_accion.EXIT_OESTE
End Function


