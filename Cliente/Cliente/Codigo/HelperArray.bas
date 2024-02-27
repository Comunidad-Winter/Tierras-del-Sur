Attribute VB_Name = "HelperArray"
Option Explicit

'http://stackoverflow.com/questions/183353/how-do-i-determine-if-an-array-is-initialized-in-vb6
Public Function arrayEstaIniciado(vector As Variant) As Boolean

    On Error GoTo ProcError
    Dim lTmp As Long

    arrayEstaIniciado = False

    lTmp = UBound(vector) ' Acá puede saltar el error

    arrayEstaIniciado = (lTmp > -1)
    
    Exit Function
ProcError:
    'El error sera  "Subscript 'out of range", caso contrario el error es por otra cosa (no es un array)
    If Not Err.Number = 9 Then Err.Raise (Err.Number)

End Function

Public Function Join(vector() As String, caracter As String) As String

    Dim loopC As Integer
    Dim cantidad As Integer
    
    ' ¿Cuantos?
    cantidad = UBound(vector) - LBound(vector)
    If LBound(vector) = 0 Then cantidad = cantidad + 1
    
    ' ¿Devuelvo?
    If cantidad = 0 Then
        Join = ""
    ElseIf cantidad = 1 Then
        Join = vector(LBound(vector))
    Else
        For loopC = LBound(vector) To UBound(vector)
            If Not loopC = UBound(vector) Then
                Join = Join & vector(loopC) & caracter
            Else
                Join = Join & vector(loopC)
            End If
        Next loopC
    End If
End Function

