Attribute VB_Name = "Encriptacion"
Option Explicit
Private Enum SentidoRotacion
    ROTIzquierda = 0
    ROTDerecha = 1
End Enum

Public Function ProtoCrypt(ByVal s As String, ByVal KeY As Integer) As String
'esta funcion encripta un string
    ProtoCrypt = s


Exit Function
errorHandlerEncriptar:
'Call LogError("Error encriptando: " & s)
ProtoCrypt = s

End Function



