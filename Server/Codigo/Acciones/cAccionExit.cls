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

Private mapa As Integer
Private x As Byte
Private y As Byte
Private radio As Byte
Private conEfecto As Byte

Public Sub crear(mapaDestino As Integer, xDestino As Byte, yDestino As Byte, radio_ As Byte, conEfecto_ As Byte)
        mapa = mapaDestino
        x = xDestino
        y = yDestino
        radio = radio_
        conEfecto = conEfecto_
End Sub

Private Function iAccion_ejecutar(mapa_ As Integer, x_ As Byte, y_ As Byte) As Boolean
Dim UserIndex As Integer
Dim xFinal As Byte
Dim yFinal As Byte

UserIndex = MapData(mapa_, x_, y_).UserIndex

xFinal = RandomNumberInt(x - radio, x + radio)
yFinal = RandomNumberInt(y - radio, y + radio)

iAccion_ejecutar = avanzarPersonajeOtroMapa(UserList(UserIndex), mapa, xFinal, yFinal)

If iAccion_ejecutar = True Then
    If conEfecto = 1 Then
        EnviarPaquete Paquetes.HechizoFX, ITS(UserList(UserIndex).Char.charIndex) & ByteToString(FXWARP) & ITS(0), UserIndex, ToPCArea, UserList(UserIndex).pos.map
    End If
End If
End Function

Private Sub iAccion_cargar(archivoSalida As Integer)
    Get archivoSalida, , mapa
    Get archivoSalida, , x
    Get archivoSalida, , y
    Get archivoSalida, , conEfecto
    Get archivoSalida, , radio
End Sub
