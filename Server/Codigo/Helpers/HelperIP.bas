Attribute VB_Name = "HelperIP"
Option Explicit

' Transforma una direccion IP en formato entero a un string
Public Function longToIP(ByVal valor As Currency) As String

Dim x As Byte
Dim num As Long

For x = 1 To 4
    num = Int(valor / 256 ^ (4 - x))
    valor = valor - (num * 256 ^ (4 - x))

    If num > 255 Then
        longToIP = "0.0.0.0"
        Exit Function
    End If

    If x = 1 Then
        longToIP = num
    Else
        longToIP = longToIP & "." & num
    End If
Next
        
End Function

