VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAccionExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements iAccion

Option Explicit

Private Mapa As Integer
Private X As Byte
Private Y As Byte
Private conEfecto As Byte
Private radio As Byte

Private Sub crear(MapaDestino As Integer, xDestino As Byte, YDestino As Byte, conEfecto_ As Byte, radio_ As Byte)
    Mapa = MapaDestino
    X = xDestino
    Y = YDestino
    radio = radio_
    conEfecto = conEfecto_
End Sub
Private Function iAccion_getID() As Long
    iAccion_getID = tipo_accion.EXIT_COMUN
End Function

Private Sub iAccion_persistir(archivoSalida As Integer)
    Put archivoSalida, , CInt(iAccion_getID)
    Put archivoSalida, , Mapa
    Put archivoSalida, , X
    Put archivoSalida, , Y
    Put archivoSalida, , conEfecto
    Put archivoSalida, , radio
End Sub
Private Sub iAccion_cargarParametros(parametros As Collection)
    Call crear(CInt(parametros.Item(1).getValor), CByte(parametros.Item(2).getValor), CByte(parametros.Item(3).getValor), CByte(parametros.Item(4).getValor), CByte(parametros.Item(5).getValor))
End Sub
Private Function iAccion_clonar() As iAccion
    Set iAccion_clonar = New cAccionExit
End Function
