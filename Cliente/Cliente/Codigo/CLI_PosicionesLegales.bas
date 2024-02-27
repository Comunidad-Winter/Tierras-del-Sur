Attribute VB_Name = "CLI_PosicionesLegales"

Option Explicit

Public Function esPosicionJugable(ByVal x As Integer, ByVal y As Integer) As Boolean
    If x < X_MINIMO_JUGABLE Or x > X_MAXIMO_JUGABLE Or y < Y_MINIMO_JUGABLE Or y > Y_MAXIMO_JUGABLE Then
        esPosicionJugable = False
        Exit Function
    End If
    
    esPosicionJugable = True
End Function
