Attribute VB_Name = "HelperStrings"
Option Explicit
'TODO : Si bien nunca estuvo allí, el mapa es algo independiente o a lo sumo dependiente del engine, no va acá!!!

Public Function ReadField(ByVal Pos As Integer, ByRef text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(text, lastPos + 1, Len(text) - lastPos)
    Else
        ReadField = mid$(text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function

Public Function FieldCount(ByRef text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, text, delimiter)
        count = count + 1
    Loop While curPos <> 0
    
    FieldCount = count
End Function

' Menduz: Esto es mas facil q la funcion de mierda que habian hecho los de ao con muchos on error goto y esas boludeces
Public Function CheckMailString(ByVal sString As String) As Boolean
    If Len(sString) < 3 Then Exit Function
    
    If InStr(1, sString, "@", vbBinaryCompare) = 0 Then Exit Function
    
    CheckMailString = True
End Function

Function QuitarDobleEspacios(ByRef cad As String) As String
    Dim Strin As String
    
    Strin = Trim(cad)

    Do While InStr(Strin, "  ")
        Strin = Replace(Strin, "  ", " ")
    Loop

    QuitarDobleEspacios = Strin
End Function


    
Private Function CMSValidateChar_(ByRef iAsc As Integer) As Boolean
CMSValidateChar_ = IIf( _
                    (iAsc >= 48 And iAsc <= 57) Or _
                    (iAsc >= 65 And iAsc <= 90) Or _
                    (iAsc >= 97 And iAsc <= 122) Or _
                    (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46), True, False)
End Function


Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)
For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    If (car < 97 Or car > 122) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
Next i
AsciiValidos = True
End Function

Function Mysql_Escape(texto As String) As Boolean
Dim i As Byte
For i = 1 To Len(texto)
    If Asc(mid(texto, i, 1)) = 34 Or Asc(mid(texto, i, 1)) = 39 Then
    Mysql_Escape = True
    Exit Function
    End If
Next
Mysql_Escape = False

End Function
Function DobleEspacios(UserName As String) As Boolean
Dim Antes As Boolean
Dim i As Byte

For i = 1 To Len(UserName)
    If mid(UserName, i, 1) = " " Then
        If Antes = True Then
            DobleEspacios = True
            Exit Function
        Else
        Antes = True
        End If
    Else
    Antes = False
    End If
Next
DobleEspacios = False
End Function

