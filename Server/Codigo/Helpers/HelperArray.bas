Attribute VB_Name = "HelperArray"
Option Explicit

'CSEH: Nada
Public Function existeEnArray(elemento As Long, array_() As Long) As Boolean
    Dim loopFrame As Long
    On Error GoTo error:
    For loopFrame = LBound(array_) To UBound(array_)
        If array_(loopFrame) = elemento Then
            existeEnArray = True
            Exit Function
        End If
    Next loopFrame
    existeEnArray = False
    Exit Function
error:
    existeEnArray = False
End Function


'CSEH: Nada
Public Function existeEnArrayString(elemento As String, array_() As String) As Boolean
    Dim loopFrame As Long
    On Error GoTo error:
    For loopFrame = LBound(array_) To UBound(array_)
        If array_(loopFrame) = elemento Then
            existeEnArrayString = True
            Exit Function
        End If
    Next loopFrame
    existeEnArrayString = False
    Exit Function
error:
    existeEnArrayString = False
End Function

'http://stackoverflow.com/questions/183353/how-do-i-determine-if-an-array-is-initialized-in-vb6
'CSEH: Nada
Public Function arrayEstaIniciado(vector As Variant) As Boolean

    On Error GoTo ProcError
    Dim lTmp As Long

    arrayEstaIniciado = False

    lTmp = UBound(vector) ' Acá puede saltar el error

    arrayEstaIniciado = (lTmp > -1)
    
    Exit Function
ProcError:
    arrayEstaIniciado = False
    'El error sera  "Subscript 'out of range", caso contrario el error es por otra cosa (no es un array)
    'If Not Err.Number = 9 Then Err.Raise (Err.Number)

End Function
