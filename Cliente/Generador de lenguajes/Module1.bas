Attribute VB_Name = "Module1"
Option Explicit

'@@@@@@@@@@@@@@@@@@@ Twister Library Declarations @@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Private Declare Sub SetKeySize Lib "Twister" (ByVal nSize As Long)
Private Declare Sub SetKeyValue Lib "Twister" (ByVal nKey As Long, ByVal nValue As Long)
Private Declare Function Twist Lib "Twister" (ByVal lpBuffer As String, ByVal nSize As Long, ByVal nOffs As Long) As Long
Private Declare Function DeTwist Lib "Twister" (ByVal lpBuffer As String, ByVal nSize As Long, ByVal nOffs As Long) As Long
'@@@@@@@@@@@@@@@@@@@ Twister Library Declarations @@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public mensaje(1 To 385) As String
Dim buf As String
Private Const KEY_SIZE = 21
Public CryptOffs As Integer 'Gorlok

Public Sub CryptoInit()

    SetKeySize (KEY_SIZE)
    SetKeyValue 2, 1
    SetKeyValue 3, 5
    SetKeyValue 0, 1
    SetKeyValue -1, 2
    SetKeyValue 1, 11
    SetKeyValue -1, 1
    SetKeyValue -2, 2
    SetKeyValue 1, 3
    SetKeyValue 3, 2
    SetKeyValue 4, 3
    SetKeyValue 0, 1
    SetKeyValue 1, 9
    SetKeyValue 2, 7
    SetKeyValue 4, 2
    SetKeyValue 2, -1
    SetKeyValue -1, 1
    SetKeyValue 1, 1
    SetKeyValue 1, 4
    SetKeyValue -1, 3
    SetKeyValue 3, 1
    SetKeyValue 2, -1
    
End Sub

Public Function CryptStr(ByVal s$, ByRef Offset As Integer) As String
#If ENCRIPTADO = 1 Then
    Dim nLen As Integer
    nLen = Len(s$)
    Call Twist(s$, nLen, Offset)
    Offset = (Offset + nLen) Mod KEY_SIZE
#End If
    CryptStr = s$
End Function

Public Function DecryptStr(ByVal s$, ByRef Offset As Integer) As String
#If ENCRIPTADO = 1 Then
    Dim nLen As Integer
    nLen = Len(s$)
    Call DeTwist(s$, nLen, Offset)
    Offset = (Offset + nLen) Mod KEY_SIZE
#End If
    DecryptStr = s$
End Function

