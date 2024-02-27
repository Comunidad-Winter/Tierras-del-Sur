Attribute VB_Name = "HelperDate"
Option Explicit

Public Function AbreviaturaANumero(abreviatura As String) As Byte

AbreviaturaANumero = 0

Select Case abreviatura
    Case "Dec"
        AbreviaturaANumero = 12
    Case "Nov"
        AbreviaturaANumero = 11
    Case "Oct"
        AbreviaturaANumero = 10
    Case "Sep"
        AbreviaturaANumero = 9
    Case "Aug"
        AbreviaturaANumero = 8
    Case "Jul"
        AbreviaturaANumero = 7
    Case "Jun"
        AbreviaturaANumero = 6
    Case "May"
        AbreviaturaANumero = 5
    Case "Apr"
        AbreviaturaANumero = 4
    Case "Mar"
        AbreviaturaANumero = 3
    Case "Feb"
        AbreviaturaANumero = 2
    Case "Jan"
        AbreviaturaANumero = 1
End Select

End Function
