Attribute VB_Name = "HelperTCP"
Option Explicit

Public Function LongToString(ByVal Var As Long) As String
    Dim Temp As String
      
    'Convertimos a hexa
    Temp = Hex$(Var)
    
    'Nos aseguramos tenga 8 Bytes de largo
    While Len(Temp) < 8
        Temp = "0" & Temp
    Wend
    
    'Convertimos a string
    LongToString = Chr$(val("&H" & left$(Temp, 2))) & Chr$(val("&H" & mid$(Temp, 3, 2))) & Chr$(val("&H" & mid$(Temp, 5, 2))) & Chr$(val("&H" & mid$(Temp, 7, 2)))
Exit Function
ErrHandler:
LogError "LongToString:" & Var
End Function
Public Function StringToLong(ByVal Str As String, ByVal Start As Byte) As Long
    If Len(Str) < Start - 3 Then Exit Function
    
    Dim TempStr As String
    Dim tempstr2 As String
    Dim tempstr3 As String
    'Tomamos los últimos 3 Bytes y convertimos sus valroes ASCII a hexa
    TempStr = Hex$(Asc(mid$(Str, Start + 1, 1)))
    tempstr2 = Hex$(Asc(mid$(Str, Start + 2, 1)))
    tempstr3 = Hex$(Asc(mid$(Str, Start + 3, 1)))
    
    'Nos aseguramos todos midan 2 Bytes (los ceros a la izquierda cuentan por ser Bytes 2, 3 y 4)
    While Len(TempStr) < 2
        TempStr = "0" & TempStr
    Wend
    
    While Len(tempstr2) < 2
        tempstr2 = "0" & tempstr2
    Wend
    
    While Len(tempstr3) < 2
        tempstr3 = "0" & tempstr3
    Wend
    
    'Convertimos a una única cadena hexa
    StringToLong = CLng("&H" & Hex$(Asc(mid$(Str, Start, 1))) & TempStr & tempstr2 & tempstr3)
End Function

Public Function ByteToString(ByVal Var As Byte) As String
    ByteToString = Chr$(Var)
Exit Function

ErrHandler:
End Function
Public Function StringToByte(ByVal Str As String, ByVal Start As Byte) As Byte
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    If Len(Str) < Start Then Exit Function
    
    StringToByte = Asc(mid$(Str, Start, 1))
End Function
Public Function ITS(ByVal Var As Integer) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    'No aceptamos valores que utilicen los últimos bits, pues los usamos como flag para evitar chr$(0)s
    Dim Temp As String
       
    'Convertimos a hexa
    Temp = Hex$(Var)
    
    'Nos aseguramos tenga 4 Bytes de largo
    While Len(Temp) < 4
        Temp = "0" & Temp
    Wend
    
    'Convertimos a string
    ITS = Chr$(val("&H" & left$(Temp, 2))) & Chr$(val("&H" & right$(Temp, 2)))
Exit Function

ErrHandler:

End Function
Public Function STI(ByVal Str As String, ByVal Start As Byte) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    Dim TempStr As String
    
    'Asergurarse sea válido
    If Len(Str) < Start - 1 Then Exit Function
    'Convertimos a hexa el valor ascii del segundo Byte
    TempStr = Hex$(Asc(mid$(Str, Start + 1, 1)))
    
    'Nos aseguramos tenga 2 Bytes (los ceros a la izquierda cuentan por ser el segundo Byte)
    While Len(TempStr) < 2
        TempStr = "0" & TempStr
    Wend
    
    'Convertimos a integer
    STI = val("&H" & Hex$(Asc(mid$(Str, Start, 1))) & TempStr)
End Function

