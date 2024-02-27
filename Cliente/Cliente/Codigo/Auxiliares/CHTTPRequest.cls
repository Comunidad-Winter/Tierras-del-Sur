VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHTTPRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strHost As String
Private strProxy As String
Private strPath As String
Private strBody As String
Private strUserAgent As String
Private strMimeBoundary As String
Private dicFormData As Object

Private dicFileNames  As collection
Private dicFileContents As collection
Private headers As Dictionary

Private enumMethod As eHttpMethod


Public Sub AddFile(pathArchivo As String, Nombre As String)

    Dim archivo As Variant
    Dim loopCaracter As Byte
    
    If strMimeBoundary = vbNullString Then
        For loopCaracter = 1 To 14
            strMimeBoundary = strMimeBoundary & Chr(97 + Round(Rnd() * 25))
        Next
    End If
    
    Call dicFileNames.Add(Nombre)
    Call dicFileContents.Add(pathArchivo)
    
End Sub

Public Sub AddFormData(strKey As String, ByVal StrValue As String)
    Call dicFormData.Add(strKey, StrValue)
End Sub

Private Sub Class_Initialize()
    method = httpGET
    
    strMimeBoundary = vbNullString

    Set dicFormData = CreateObject("Scripting.Dictionary")
    Set headers = CreateObject("Scripting.Dictionary")
    
    Set dicFileNames = New collection
    Set dicFileContents = New collection
End Sub

Public Sub reset()

    strMimeBoundary = vbNullString
    
    dicFormData.RemoveAll
    Set dicFileNames = New collection
    Set dicFileContents = New collection
    
End Sub


Private Sub Class_Terminate()
    
    Set dicFormData = Nothing
    Set dicFileNames = Nothing
    Set dicFileContents = Nothing

End Sub

Public Function GetQuery() As String

  Dim strFormData As String
  Dim StrValue As String
  Dim arrKeys
  Dim i As Integer
  Dim intFreeFile As Integer
  Dim sRefData As String
  Const sDblQts As String = """"
  Dim strBuffer As String
  
    If Not body = "" Then
        GetQuery = body
    ElseIf strMimeBoundary = vbNullString Then
    
        arrKeys = dicFormData.keys()
    
        strFormData = ""
        
        For i = 0 To UBound(arrKeys)
            StrValue = Replace(dicFormData(arrKeys(i)), "&", "&amp;")
            
            If enumMethod = httpGET Then
                StrValue = Replace(dicFormData(arrKeys(i)), " ", "%20")
            End If
            
            strFormData = strFormData & arrKeys(i) & "=" & StrValue
    
            'Next value
            If i < UBound(arrKeys) Then strFormData = strFormData & "&"
        Next i
    
        GetQuery = strFormData
    Else
    
        arrKeys = dicFormData.keys()
    
        sRefData = ""
        
        ' Variables
        sRefData = vbNullString
        sRefData = sRefData & "Content-Type: multipart/form-data; boundary=---------------------------" & strMimeBoundary & vbCrLf
        
        For i = 0 To UBound(arrKeys)
            StrValue = Replace(dicFormData(arrKeys(i)), "&", "&amp;")

           If enumMethod = httpGET Then
                StrValue = Replace(dicFormData(arrKeys(i)), " ", "%20")
            End If
            
            sRefData = sRefData & "-----------------------------" & strMimeBoundary & vbCrLf
            sRefData = sRefData & "Content-Disposition: form-data; name=" & sDblQts & arrKeys(i) & sDblQts & vbCrLf & vbCrLf
            sRefData = sRefData & StrValue & vbCrLf
        Next i
    
        ' Archivos
        intFreeFile = FreeFile

        For i = 1 To dicFileContents.count
        
            Open dicFileContents.item(i) For Binary As intFreeFile
                strBuffer = Space(LOF(intFreeFile))
                Get #intFreeFile, 1, strBuffer
            Close intFreeFile
        
            sRefData = sRefData & "-----------------------------" & strMimeBoundary & vbCrLf
            sRefData = sRefData & "Content-Disposition: form-data; name=" & sDblQts & dicFileNames.item(i) & sDblQts & "; filename=" & sDblQts & "A" & sDblQts & vbCrLf
            sRefData = sRefData & "Content-Type: application/x-zip-compressed" & vbCrLf & vbCrLf
            sRefData = sRefData & strBuffer & vbCrLf
            sRefData = sRefData & "-----------------------------" & strMimeBoundary & "--"
        Next

        GetQuery = sRefData
    End If
End Function

Public Sub addHeader(clave As String, valor As String)

    Call headers.Add(clave, valor)

End Sub

Public Function getHeaders() As String
    Dim header As Variant
    
    If Not strUserAgent = "" Then
        getHeaders = getHeaders & "User-Agent: " & strUserAgent & vbCrLf
    End If
    
    getHeaders = getHeaders & "Accept: */*" & vbCrLf & vbCrLf
   ' getHeaders = getHeaders & "Accept - Encoding: gzip , deflate, sdch" & vbCrLf
   ' getHeaders = getHeaders & "Accept-Language: es-AR,es;q=0.8" & vbCrLf
    
    If Not body = "" Then
        getHeaders = "Content-type: application/json" & vbCrLf
    End If
    
    For Each header In headers.keys
        getHeaders = getHeaders & CStr(header) & ": " & headers.item(header) & vbCrLf
    Next
    
    
    'If dicFileNames.count > 0 Then
    '    getHeaders = getHeaders & "Content-Type: multipart/form-data; boundary=---------------------------" & strMimeBoundary
    'Else
    '    getHeaders = getHeaders & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    'End If
End Function

Public Function getURL() As String
    getURL = strHost & "/" & strPath
End Function

Public Function getMethod() As String
    
    Select Case enumMethod
       Case httpGET
            getMethod = "get"
       Case httppost
            getMethod = "post"
       Case httpPUT
            getMethod = "put"
       Case httpDELETE
            getMethod = "delete"
       Case httpPATCH
            getMethod = "patch"
        Case httpUNLOCK
            getMethod = "unlock"
    End Select

End Function

Public Property Let Host(strInput As String)
    strHost = strInput
End Property

Public Property Let MimeBoundary(strInput As String)
    strMimeBoundary = strInput
End Property

Public Property Let method(strInput As eHttpMethod)
    enumMethod = strInput
End Property

Public Property Get body() As String
    body = strBody
End Property

Public Property Let body(strInput As String)
    strBody = strInput
End Property

Public Property Let Path(strInput As String)
    strPath = strInput
End Property

Public Property Let Proxy(strInput As String)
    strProxy = strInput
End Property

Public Property Let UserAgent(strInput As String)
    strUserAgent = strInput
End Property

