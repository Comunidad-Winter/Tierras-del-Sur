Attribute VB_Name = "DataReader"
Option Explicit

Private data As String
Private currentPos As Integer

Public Sub setData(newData As String)
    data = newData
    currentPos = 1
End Sub

Public Function getCurrentPos() As Integer
    getCurrentPos = currentPos
End Function

Public Function getInteger() As Integer
    getInteger = STI(data, currentPos)
    currentPos = currentPos + 2
End Function

Public Function getByte() As Byte
    getByte = StringToByte(data, currentPos)
    currentPos = currentPos + 1
End Function

Public Function getLong() As Long
    getLong = StringToLong(data, currentPos)
    currentPos = currentPos + 4
End Function

Public Function tieneDatos() As Boolean
    tieneDatos = currentPos <= Len(data)
End Function

