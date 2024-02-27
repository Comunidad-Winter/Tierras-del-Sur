Attribute VB_Name = "HelperMatematicas"
Option Explicit

Private Const VB_MIN_BYTE As Byte = 0
Private Const VB_MIN_INT As Integer = -32768
Private Const VB_MIN_LONG As Long = -2147483648#

Private Const VB_MAX_BYTE As Byte = 255
Private Const VB_MAX_INT As Integer = 32767
Private Const VB_MAX_LONG As Long = 2147483647

Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180
Public Const RadianToDegree As Single = 57.2957795130823 '180 / Pi

Public Function mini(ByVal a As Integer, ByVal b As Integer) As Integer
    If a > b Then
        mini = b
    Else
        mini = a
    End If
End Function

Public Function maxi(ByVal a As Integer, ByVal b As Integer) As Integer
    If a > b Then
        maxi = a
    Else
        maxi = b
    End If
End Function


Public Function minl(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then
        minl = b
    Else
        minl = a
    End If
End Function

Public Function maxl(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then
        maxl = a
    Else
        maxl = b
    End If
End Function
Public Function mins(ByVal a As Single, ByVal b As Single) As Single
    If a > b Then
        mins = b
    Else
        mins = a
    End If
End Function

Public Function maxs(ByVal a As Single, ByVal b As Single) As Single
    If a > b Then
        maxs = a
    Else
        maxs = b
    End If
End Function

Public Function CByteSeguro(ByVal expresion As String) As Byte
    Dim temp As Single
    
    temp = val(expresion)
    
    '¿Esta dentro de los limites?
    If temp < VB_MIN_BYTE Or temp > VB_MAX_BYTE Then
        CByteSeguro = VB_MAX_BYTE
    Else
        CByteSeguro = CByte(temp)
    End If
End Function

Public Function Angulo(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer) As Single
    If x2 - x1 = 0 Then
        If y2 - y1 = 0 Then
            Angulo = 90
        Else
            Angulo = 270
        End If
    Else
        Angulo = Atn((y2 - y1) / (x2 - x1)) * RadianToDegree
        If (x2 - x1) < 0 Or (y2 - y1) < 0 Then Angulo = Angulo + 180
        If (x2 - x1) > 0 And (y2 - y1) < 0 Then Angulo = Angulo - 180
        If Angulo < 0 Then Angulo = Angulo + 360
    End If
End Function


Public Function CIntSeguro(ByVal expresion As String) As Integer
    Dim temp As Single
    
    temp = val(expresion)
    
    '¿Esta dentro de los limites?
    If temp < VB_MIN_INT Or temp > VB_MAX_INT Then
        CIntSeguro = VB_MAX_INT
    Else
        CIntSeguro = CInt(temp)
    End If
End Function

Sub AddtoVar(ByRef Var As Variant, ByVal Addon As Variant, ByVal max As Variant)
    'Le suma un valor a una variable respetando el maximo valor
    If Var >= max Then
        Var = max
    Else
        Var = Var + Addon
        If Var > max Then
            Var = max
        End If
    End If
End Sub

Public Sub RestToVar(ByRef Var As Integer, ByVal cantidad As Integer, ByVal min As Integer)
    'Le suma un valor a una variable respetando el maximo valor
    Var = Var - cantidad
    If Var < min Then
        Var = min
    End If
End Sub

Public Function Porcentaje(ByVal Total As Long, ByVal Porc As Long) As Long
    Porcentaje = (Total * Porc) / 100
End Function

Function distancia(wp1 As WorldPos, wp2 As WorldPos)
    'Encuentra la distancia entre dos WorldPos
    distancia = Abs(wp1.x - wp2.x) + Abs(wp1.y - wp2.y) + (Abs(wp1.map - wp2.map) * 100)
End Function

Function Distance(x1 As Variant, y1 As Variant, x2 As Variant, y2 As Variant) As Double
    'Encuentra la distancia entre dos puntos
    Distance = Sqr(((y1 - y2) ^ 2 + (x1 - x2) ^ 2))
End Function

Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

' Le sumo uno para que me devuelva un valor entre [MINIMO, MAXIMO.99999^]. Si me quedo con la pate entera, voy a tener la misma probabilidad para todos
Public Function RandomNumberInt(ByVal LowerBound As Integer, ByVal UpperBound As Integer) As Integer
    RandomNumberInt = Int(((UpperBound + 1) - LowerBound) * Rnd) + LowerBound
End Function

' Le sumo uno para que me devuelva un valor entre [MINIMO, MAXIMO.99999^]. Si me quedo con la pate entera, voy a tener la misma probabilidad para todos
Public Function RandomNumberByte(ByVal LowerBound As Byte, ByVal UpperBound As Byte) As Byte
    RandomNumberByte = Int(((UpperBound + 1) - LowerBound) * Rnd) + LowerBound
End Function

Public Function RandomNumberSingle(ByVal LowerBound As Single, ByVal UpperBound As Single) As Single
    'Generate random number
    RandomNumberSingle = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

