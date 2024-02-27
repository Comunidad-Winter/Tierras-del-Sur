Attribute VB_Name = "HelperBuffer"
Option Explicit

'CSEH: Nada
Public Function DeCodify(ByVal Strin As String) As Long
If Len(Strin) > 4 Then GoTo errhandler

If Len(Strin) = 1 Then
    DeCodify = StringToByte(Strin, 1)
ElseIf Len(Strin) = 2 Then
    DeCodify = STI(Strin, 1)
Else
    DeCodify = StringToLong(Strin, 1)
End If
errhandler:
 End Function

Public Function eliminarTildesMayus(cadena As String) As String
    Dim modificada As String
    
    modificada = Replace$(cadena, "Á", "A")
    modificada = Replace$(modificada, "É", "E")
    modificada = Replace$(modificada, "Í", "I")
    modificada = Replace$(modificada, "Ó", "O")
    modificada = Replace$(modificada, "Ú", "U")
    
    eliminarTildesMayus = modificada
End Function

'CSEH: Nada
Public Function SingleTOString(i As Single) As String
    SingleTOString = LongToString(Int(i)) & Chr$(val((i - Fix(i)) * 100))
End Function

'CSEH: Nada
Public Function WriteString(valor As String) As String
    WriteString = Chr$(Len(valor)) & valor
End Function

'CSEH: Nada
Public Function ReadString(valor As String) As String
    ReadString = mid$(valor, 2, Asc(Left$(valor, 1)))
End Function

'---------------------------------------------------------------------------------------
' Procedure : Codify
' DateTime  : 18/02/2007 19:59
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
'CSEH: Nada
Public Function Codify(ByVal Strin As String, Optional ByVal Mode As Byte = 0) As String
'If val(Strin) > &HFFFFFFF Then GoTo errhandler
'Si Mode es 1, entonces el 0 es tomado como integer, si es 0 entonces es tomado como byte

If Mode = 0 Then
    If val(Strin) < 254 Then
        Codify = ByteToString(Strin)
    ElseIf val(Strin) < 16383 Then
        Codify = ITS(Strin)
    Else
        Codify = LongToString(Strin)
    End If
Else
    If val(Strin) < 255 And val(Strin) > 0 Then
        Codify = ByteToString(Strin)
    ElseIf val(Strin) < 16383 Then
        Codify = ITS(Strin)
    Else
        Codify = LongToString(Strin)
    End If
End If

End Function

'CSEH: Nada
Public Function StringToSingle(ByVal str As String, Start As Byte) As Single
   StringToSingle = StringToLong(str, Start) + (Asc(mid$(str, Start + 4, 1)) / 100)
End Function

'CSEH: Nada
Public Function LongToString(ByVal Var As Long) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    Dim temp As String
       
    'Convertimos a hexa
    temp = Hex$(Var)
    
    'Nos aseguramos tenga 8 Bytes de largo
    While Len(temp) < 8
        temp = "0" & temp
    Wend
    
    'Convertimos a string
    LongToString = Chr$(val("&H" & Left$(temp, 2))) & Chr$(val("&H" & mid$(temp, 3, 2))) & Chr$(val("&H" & mid$(temp, 5, 2))) & Chr$(val("&H" & mid$(temp, 7, 2)))
End Function

'CSEH: Nada
Public Function StringToLong(ByVal str As String, ByVal Start As Byte) As Long
    If Len(str) < Start - 3 Then Exit Function
    
    Dim tempstr As String
    Dim TempStr2 As String
    Dim tempstr3 As String
    
    'Tomamos los últimos 3 Bytes y convertimos sus valroes ASCII a hexa
    tempstr = Hex$(Asc(mid$(str, Start + 1, 1)))
    TempStr2 = Hex$(Asc(mid$(str, Start + 2, 1)))
    tempstr3 = Hex$(Asc(mid$(str, Start + 3, 1)))
    
    'Nos aseguramos todos midan 2 Bytes (los ceros a la izquierda cuentan por ser Bytes 2, 3 y 4)
    While Len(tempstr) < 2
        tempstr = "0" & tempstr
    Wend
    
    While Len(TempStr2) < 2
        TempStr2 = "0" & TempStr2
    Wend
    
    While Len(tempstr3) < 2
        tempstr3 = "0" & tempstr3
    Wend
    
    'Convertimos a una única cadena hexa
    StringToLong = CLng("&H" & Hex$(Asc(mid$(str, Start, 1))) & tempstr & TempStr2 & tempstr3)
End Function
'CSEH: Nada
Public Function ByteToString(ByVal Var As Byte) As String
    ByteToString = Chr$(Var)
End Function
