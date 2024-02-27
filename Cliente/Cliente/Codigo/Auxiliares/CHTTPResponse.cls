VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHTTPResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : cHTTPResponse
'    Project    : CerebroDeMono2
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Option Explicit

Private intCode As Integer
Private strContenido As String
Private strBody As String
Private strHeader As String

Public Sub crear(header As String, body As String)


    
    Dim Space1 As Integer
    Dim Space2 As Integer
    Dim URL As String

    strBody = body
    strHeader = header

    Space1 = InStr(strHeader, " ") + 1
    Space2 = InStr(Space1, strHeader, " ")
    URL = mid$(strHeader, Space1, Space2 - Space1)
    
    intCode = CInt(val(URL))


End Sub
Public Property Get hasBody() As Boolean
    hasBody = Not (strContenido = "")
End Property

Public Property Get body() As String

    body = strBody

End Property

Public Property Get bodyJSON() As Object
    Set bodyJSON = JSON.parse(strBody)
End Property

Public Property Get Code() As Integer

    Code = intCode

End Property

