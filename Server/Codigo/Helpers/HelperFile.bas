Attribute VB_Name = "HelperFile"
Option Explicit

'Se fija si existe el archivo
Function FileExist(file As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
    If Dir(file, FileType) = "" Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function

Public Function ReadField(pos As Integer, Text As String, SepASCII As Integer) As String
    Dim delimiter As String
    delimiter = Chr(SepASCII)
    Dim i As Long
        Dim LastPos As Long
        Dim CurrentPos As Long
        
        For i = 1 To pos
            LastPos = CurrentPos
            CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
        Next i
        
        If CurrentPos = 0 Then
            ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
        Else
            ReadField = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)
        End If
End Function
