Attribute VB_Name = "HelperMath"
'ARCHIVO COMPARTIDO

Option Explicit

Public CDecimalSeparator As String * 1

Public Const pi             As Single = 3.14159265358979
Public Const Pi2            As Single = 6.28318530717959
Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180
Public Const RadianToDegree As Single = 57.2957795130823 '180 / Pi

Public Coseno(360)      As Single
Public Seno(360)        As Single
Public Alphas(255)      As Long

Public Sub Init_Math_Const()
Dim i As Integer
For i = 0 To 360
    Coseno(i) = Cos(i * DegreeToRadian)
    Seno(i) = Sin(i * DegreeToRadian)
Next i
For i = 0 To 255
    Alphas(i) = CLng("&H" & Hex$(i) & "000000")
Next i
End Sub

Public Function GetCurrencySymbol() As String
    GetCurrencySymbol = Replace(Replace(Replace(format(0, "Currency"), ".", ""), "0", ""), ",", "")
End Function

Public Function DecimalSeparator() As String
If CDecimalSeparator <> "," And CDecimalSeparator <> "." Then

    DecimalSeparator = mid$(1 / 2, 2, 1)
    
    If val("1" & DecimalSeparator & "9") <> 1.9 Then
        If val("1.9") = 1.9 Then
            DecimalSeparator = "."
        Else
            If val("1,9") = 1.9 Then
                DecimalSeparator = ","
            End If
        End If
    End If
have_err:
'Marce 'Marce 'Marce On error goto 0

    CDecimalSeparator = DecimalSeparator
Else
    DecimalSeparator = CDecimalSeparator
End If
    
End Function

Public Function CCVal(ByVal X As String) As Single
    On Error GoTo ja
    Dim ea As String
        ea = Replace(Replace(X, ".", ","), ",", DecimalSeparator)
        CCVal = CSng(Round(val(ea), 2))
    Exit Function
ja:
On Local Error GoTo ji
    CCVal = CSng(Round(val(ea), 2))
    Exit Function
ji:
    CCVal = 0
End Function

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
Public Function int32x32_int64(ByVal lLo As Long, ByVal lHi As Long) As Double
    Dim dLo As Double
    Dim dHi As Double
    
    If lLo < 0 Then
        dLo = (2 ^ 32) + lLo
    Else
        dLo = lLo
    End If
    If lHi < 0 Then
        dHi = (2 ^ 32) + lHi
    Else
        dHi = lHi
    End If
    
    int32x32_int64 = (dLo + (dHi * (2 ^ 32)))
End Function

Public Function CosInterp(ByVal y1 As Single, ByVal y2 As Single, ByVal mu As Single) As Single
'interpolación con coseno wachin
   Dim mu2 As Single
   'mu2 = (1 - Cos(mu * Pi)) / 2
   'CosInterp = y1 * (1 - mu2) + y2 * mu2
   mu2 = -(Cos(mu * pi) / 2) + 0.5
   CosInterp = y1 + mu2 * (y2 - y1)
End Function

Public Function Interp(ByVal y1 As Single, ByVal y2 As Single, ByVal mu As Single) As Single
'interpolación lineal
   Interp = y1 * (1 - mu) + y2 * mu
End Function

Public Function bounds(ByVal limite1 As Long, ByVal limite2 As Long, ByVal valor As Long) As Long
    Dim lmin As Long
    Dim lmax As Long
    
    lmin = minl(limite1, limite2)
    lmax = maxl(limite1, limite2)
    
    bounds = valor

    If bounds < lmin Then
        bounds = lmin
    Else
        If bounds > lmax Then
            bounds = lmax
        End If
    End If
End Function

Public Function boundsf(ByVal limite1 As Single, ByVal limite2 As Single, ByVal valor As Single) As Single
    Dim lmin As Single
    Dim lmax As Single
    
    lmin = minl(limite1, limite2)
    lmax = maxl(limite1, limite2)
    
    boundsf = valor

    If boundsf < lmin Then
        boundsf = lmin
    Else
        If boundsf > lmax Then
            boundsf = lmax
        End If
    End If
End Function

Public Function ASin(value As Double) As Double
    If Abs(value) <> 1 Then
        ASin = Atn(value / Sqr(1 - value * value))
    Else
        ASin = 1.5707963267949 * Sgn(value)
    End If
End Function

Public Function Atan2(ByVal X As Single, ByVal Y As Single) As Single
    If X Then
        Atan2 = Atn(Y / X) - (X > 0) * pi
    Else
        Atan2 = 1.5707963267949 + (Y > 0) * pi
    End If
End Function

Function ATAN_2(ByVal X As Double, ByVal Y As Double) As Single
Select Case Sgn(X)
    Case -1
        Select Case Sgn(Y)
            Case -1
                ATAN_2 = Atn(Y / X) - pi
            Case 0
                ATAN_2 = 0
            Case 1
                ATAN_2 = pi + Atn(Y / X)
        End Select
    Case 0
        Select Case Sgn(Y)
            Case -1
                ATAN_2 = -pi / 2
            Case 0
                ATAN_2 = 0
            Case 1
                ATAN_2 = pi / 2
        End Select
    Case 1
        ATAN_2 = Atn(Y / X)
End Select
End Function


Public Function angulo(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer) As Single
    If x2 - x1 = 0 Then
        If y2 - y1 = 0 Then
            angulo = 90
        Else
            angulo = 270
        End If
    Else
        angulo = Atn((y2 - y1) / (x2 - x1)) * RadianToDegree
        If (x2 - x1) < 0 Or (y2 - y1) < 0 Then angulo = angulo + 180
        If (x2 - x1) > 0 And (y2 - y1) < 0 Then angulo = angulo - 180
        If angulo < 0 Then angulo = angulo + 360
    End If
End Function

Public Sub Long2RGB(LongCol As Long, r As Single, g As Single, b As Single)
b = (LongCol And &HFF) / 255
g = ((LongCol And &HFF00) \ &H100) / 255
r = ((LongCol And &HFF0000) \ &H10000) / 255
End Sub

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Public Function interpolar(a!, b!, T!) As Single
    interpolar = a + T * (b - a)
End Function


Public Function redondearHaciaArriba(a As Single) As Long
    
    redondearHaciaArriba = Fix(a)
    
    If a > redondearHaciaArriba Then
        redondearHaciaArriba = redondearHaciaArriba + 1
    End If

End Function

Public Function AngleAndDistanceToCoordX(angle As Integer, distance As Integer) As Integer
    AngleAndDistanceToCoordX = (distance) * Math.Cos(angle * pi / 180)
End Function

Public Function AngleAndDistanceToCoordY(angle As Integer, distance As Integer) As Integer
    AngleAndDistanceToCoordY = (distance) * Math.Sin(angle * pi / 180)
End Function
