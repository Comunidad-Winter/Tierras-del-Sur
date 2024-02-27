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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : CHTTPRequest
'    Project    : CerebroDeMono2
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Option Explicit


Private strHost As String
Private strProxy As String
Private strPath As String
Private strUserAgent As String
Private strMimeBoundary As String
Private dicFormData As Object

Private dicFileNames  As Collection
Private dicFileContents As Collection

Private enumMethod As eHttpMethod


Public Sub AddFile(pathArchivo As String, nombre As String)

    Dim Archivo As Variant
    Dim loopCaracter As Byte
    
    If strMimeBoundary = vbNullString Then
        For loopCaracter = 1 To 14
            strMimeBoundary = strMimeBoundary & Chr(97 + Round(Rnd() * 25))
        Next
    End If
    
    Call dicFileNames.Add(nombre)
    Call dicFileContents.Add(pathArchivo)
    
End Sub

Public Sub AddFormData(strKey As String, ByVal strValue As String)
    Call dicFormData.Add(strKey, strValue)
End Sub

Private Sub Class_Initialize()
    
    method = httpGET
    
    strMimeBoundary = vbNullString

    Set dicFormData = CreateObject("Scripting.Dictionary")
    Set dicFileNames = New Collection
    Set dicFileContents = New Collection

End Sub

Public Sub reset()

    strMimeBoundary = vbNullString
    
    dicFormData.RemoveAll
    Set dicFileNames = New Collection
    Set dicFileContents = New Collection
    
End Sub


Private Sub Class_Terminate()
    
    Set dicFormData = Nothing
    Set dicFileNames = Nothing
    Set dicFileContents = Nothing

End Sub

Public Function GetQuery() As String

  Dim strFormData As String
  Dim strValue As String
  Dim arrKeys
  Dim i As Integer
  Dim intFreeFile As Integer
  Dim sRefData As String
  Const sDblQts As String = """"
  Dim strBuffer As String
  
    If strMimeBoundary = vbNullString Then
    
        arrKeys = dicFormData.keys()
    
        For i = 0 To UBound(arrKeys)
            strValue = Replace(dicFormData(arrKeys(i)), "&", "&amp;")
            
            If enumMethod = httpGET Then
                strValue = Replace(dicFormData(arrKeys(i)), " ", "%20")
            End If
            
            strFormData = strFormData & arrKeys(i) & "=" & strValue
    
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
            strValue = Replace(dicFormData(arrKeys(i)), "&", "&amp;")

           If enumMethod = httpGET Then
                strValue = Replace(dicFormData(arrKeys(i)), " ", "%20")
            End If
            
            sRefData = sRefData & "-----------------------------" & strMimeBoundary & vbCrLf
            sRefData = sRefData & "Content-Disposition: form-data; name=" & sDblQts & arrKeys(i) & sDblQts & vbCrLf & vbCrLf
            sRefData = sRefData & strValue & vbCrLf
        Next i
    
        ' Archivos
        intFreeFile = FreeFile

        For i = 1 To dicFileContents.count
        
            Open dicFileContents.item(i) For Binary As intFreeFile
                strBuffer = Space(LOF(intFreeFile))
                Get #intFreeFile, 1, strBuffer
            Close intFreeFile
        
            sRefData = sRefData & "-----------------------------" & strMimeBoundary & vbCrLf
            sRefData = sRefData & "Content-Disposition: form-data; name=" & sDblQts & "filename" & sDblQts & "; filename=" & sDblQts & "A" & sDblQts & vbCrLf
            sRefData = sRefData & "Content-Type: application/x-zip-compressed" & vbCrLf & vbCrLf
            sRefData = sRefData & strBuffer & vbCrLf
            sRefData = sRefData & "-----------------------------" & strMimeBoundary & "--"
        Next
        
        GetQuery = sRefData
    End If
End Function

Public Function getHeaders() As String


    getHeaders = getHeaders & "User-Agent: " & strUserAgent & vbCrLf
    getHeaders = getHeaders & "Accept: */*" & vbCrLf & vbCrLf
    getHeaders = getHeaders & "Accept - Encoding: gzip , deflate, sdch" & vbCrLf
    getHeaders = getHeaders & "Accept-Language: es-AR,es;q=0.8" & vbCrLf

    If dicFileNames.count > 0 Then
        getHeaders = getHeaders & "Content-Type: multipart/form-data; boundary=---------------------------" & strMimeBoundary
    Else
        getHeaders = getHeaders & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    End If
  
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

Public Property Let Path(strInput As String)
    strPath = strInput
End Property

Public Property Let Proxy(strInput As String)
    strProxy = strInput
End Property

Public Property Let UserAgent(strInput As String)
    strUserAgent = strInput
End Property
