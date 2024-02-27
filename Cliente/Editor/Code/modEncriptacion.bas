Attribute VB_Name = "modEncriptacion"
Option Explicit

'@@@@@@@@@@@@@@@@@@@ Twister Library Declarations @@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Private Declare Sub SetKeySize Lib "Twister" (ByVal nSize As Long)
Private Declare Sub SetKeyValue Lib "Twister" (ByVal nKey As Long, ByVal nValue As Long)
Private Declare Function Twist Lib "Twister" (ByVal lpBuffer As String, ByVal nSize As Long, ByVal nOffs As Long) As Long
Private Declare Function DeTwist Lib "Twister" (ByVal lpBuffer As String, ByVal nSize As Long, ByVal nOffs As Long) As Long
'@@@@@@@@@@@@@@@@@@@ Twister Library Declarations @@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Private Const KEY_SIZE = 29

Public Sub CryptoInit()
   If MD5File(app.Path & "\Twister.dll") <> "87b46c40883ce07de65f0bf8898c9aef" Then
      Call MsgBox("Cliente corrupto. Bájelo de nuevo desde www.tierrasdelsur.cc.", vbApplicationModal + vbCritical + vbOKOnly, "Error al ejecutar")
       End
    End If
    SetKeySize (21)
    SetKeyValue 1, 10
    SetKeyValue 2, 51
    SetKeyValue 3, 15
    SetKeyValue 4, 27
    SetKeyValue 5, 11
    SetKeyValue 6, 12
    SetKeyValue 7, 120
    SetKeyValue 8, -31
    SetKeyValue 9, 22
    SetKeyValue 10, 34
    SetKeyValue 11, 200
    SetKeyValue 12, 94
    SetKeyValue 13, 78
    SetKeyValue 14, 34
    SetKeyValue 15, -1
    SetKeyValue 16, 4
    SetKeyValue 17, 3
    SetKeyValue 18, 25
    SetKeyValue 19, 12
    SetKeyValue 20, 23
    SetKeyValue 21, 12
End Sub
Public Function CryptStr(ByVal s$, ByRef Offset As Integer) As String
    Dim nLen As Integer
    nLen = Len(s$)
    Call Twist(s$, nLen, Offset)
    Offset = (Offset + nLen) Mod KEY_SIZE
    CryptStr = s$
End Function

Public Function DecryptStr(ByVal s$, ByRef Offset As Integer) As String
    Dim nLen As Integer
    nLen = Len(s$)
    Call DeTwist(s$, nLen, Offset)
    Offset = (Offset + nLen) Mod KEY_SIZE
    DecryptStr = s$
End Function


