Attribute VB_Name = "HelperRandom"
Option Explicit

' Retorna un entero al azar entre el [Picho, Techo]. Ambos están incluidos.
Function RandomIntNumber(ByVal Piso As Integer, ByVal Techo As Integer) As Integer
    RandomIntNumber = Int((Techo - Piso + 1) * Rnd + Piso)
End Function

