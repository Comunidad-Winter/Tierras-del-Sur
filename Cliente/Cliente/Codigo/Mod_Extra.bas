Attribute VB_Name = "CLI_Crypt"
Option Explicit

'@@@@@@@@@@@@@@@@@@@ Twister Library Declarations @@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Private Declare Sub SetKeySize Lib "Twister" (ByVal nSize As Long)
Private Declare Sub SetKeyValue Lib "Twister" (ByVal nKey As Long, ByVal nValue As Long)
Private Declare Function Twist Lib "Twister" (ByVal lpBuffer As String, ByVal nSize As Long, ByVal nOffs As Long) As Long
Private Declare Function DeTwist Lib "Twister" (ByVal lpBuffer As String, ByVal nSize As Long, ByVal nOffs As Long) As Long
'@@@@@@@@@@@@@@@@@@@ Twister Library Declarations @@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Private Const KEY_SIZE = 29

Public CryptOffs As Integer 'Gorlok

Public Sub CryptoInit()
   If MD5File(app.Path & "\Twister.dll") <> "87b46c40883ce07de65f0bf8898c9aef" Then
      Call MsgBox("Cliente corrupto. Bájelo de nuevo desde www.tierrasdelsur.cc.", vbApplicationModal + vbCritical + vbOKOnly, "Error al ejecutar")
       End
    End If
    SetKeySize (KEY_SIZE)
    SetKeyValue 0, 1
    SetKeyValue 1, 20
    SetKeyValue 2, 12
    SetKeyValue 3, 8
    SetKeyValue 7, 9
    SetKeyValue 4, 7
    SetKeyValue 5, 19
    SetKeyValue 6, 2
    SetKeyValue 8, -18
    SetKeyValue 9, -91
    SetKeyValue 10, 13
    SetKeyValue 11, 11
    SetKeyValue 12, 6
    SetKeyValue 13, 8
    SetKeyValue 14, 100
    SetKeyValue 15, 29
    SetKeyValue 16, 14
    SetKeyValue 17, 1
    SetKeyValue 18, 0
    SetKeyValue 19, 93
    SetKeyValue 20, 14
    SetKeyValue 21, 3
    SetKeyValue 22, 9
    SetKeyValue 23, 7
    SetKeyValue 24, 5
    SetKeyValue 25, -25
    SetKeyValue 26, 29
    SetKeyValue 27, 28
    SetKeyValue 28, 27
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

Public Sub CryptoInitInicial()
   If MD5File(app.Path & "\Twister.dll") <> "87b46c40883ce07de65f0bf8898c9aef" Then
      Call MsgBox("Cliente corrupto. Bájelo de nuevo desde www.tierrasdelsur.cc.", vbApplicationModal + vbCritical + vbOKOnly, "Error al ejecutar")
       End
    End If
    SetKeySize (21)
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
