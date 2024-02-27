VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Downloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Esta clase sirve para descargar páginas y ficheros por http."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Protocolo As String, Servidor As String, Objeto As String, Servicio As Long
Private URLCorrecta As Boolean, TipoConexion As Long
Private hInternetSession As Long, hInternetConnect As Long, hHttpOpenRequest As Long

Public Enum jrDownTipoAccion
    jrDownSoloInformacion = 0
    jrDownDescargar = 1
End Enum

'constantes de error
Private Const ERROR_URL As Long = 1
Private Const ERROR_INTERNETOPEN = 2
Private Const ERROR_INTERNETCONNECT = 3
Private Const ERROR_INTERNETOPENREQUEST = 4
Private Const ERROR_INTERNETSENDREQUEST = 5
Private Const ERROR_INTERNETQUERYINFO = 6
Private Const ERROR_INTERNETREADFILE = 7
Private Const ERROR_FICHERO = 8
Private Const ERROR_DESCARGA = 999
Private Const ERROR_CANCELADO = 998


'declaraciones del API
Const scUserAgent = "jrDownload"
Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Const INTERNET_OPEN_TYPE_DIRECT = 1
Const INTERNET_OPEN_TYPE_PROXY = 3
Const INTERNET_FLAG_RELOAD = &H80000000
Const HTTP_QUERY_CONTENT_TYPE = 1
Const HTTP_QUERY_CONTENT_LENGTH = 5
Const HTTP_QUERY_EXPIRES = 10
Const HTTP_QUERY_LAST_MODIFIED = 11
Const HTTP_QUERY_PRAGMA = 17
Const HTTP_QUERY_VERSION = 18
Const HTTP_QUERY_STATUS_CODE = 19
Const HTTP_QUERY_STATUS_TEXT = 20
Const HTTP_QUERY_RAW_HEADERS = 21
Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
Const HTTP_QUERY_FORWARDED = 30
Const HTTP_QUERY_SERVER = 37
Const HTTP_QUERY_USER_AGENT = 39
Const HTTP_QUERY_SET_COOKIE = 43
Const HTTP_QUERY_REQUEST_METHOD = 45
Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000
'Puertos por defecto
Const INTERNET_DEFAULT_FTP_PORT = 21
Const INTERNET_DEFAULT_GOPHER_PORT = 70
Const INTERNET_DEFAULT_HTTP_PORT = 80
Const INTERNET_DEFAULT_HTTPS_PORT = 443
Const INTERNET_DEFAULT_SOCKS_PORT = 1080
' Tipos de servicios
Const INTERNET_SERVICE_FTP = 1
Const INTERNET_SERVICE_GOPHER = 2
Const INTERNET_SERVICE_HTTP = 3
'funciones del API para internet
Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetConnect Lib "wininet.dll" Alias _
        "InternetConnectA" (ByVal InternetSession As Long, _
        ByVal sServerName As String, ByVal nServerPort As Integer, _
        ByVal sUsername As String, ByVal sPassword As String, _
        ByVal lService As Long, ByVal lFlags As Long, _
        ByVal lContext As Long) As Long
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias _
        "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As _
        String, ByVal sObjectName As String, ByVal sVersion As String, _
        ByVal sReferer As String, ByVal something As Long, ByVal lFlags _
        As Long, ByVal lContext As Long) As Long
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias _
        "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders _
        As String, ByVal lHeadersLength As Long, sOptional As Any, _
        ByVal lOptionalLength As Long) As Integer
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" _
        (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, _
        ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer
'para crear un nombre de fichero temporal y único
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
'para obtener descripción de errores del sistema
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
    (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
    Arguments As Long) As Long
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
'para obtener el tiempo de descarga
Private Declare Function GetTickCount Lib "kernel32" () As Long
'variables locales para almacenar los valores de las propiedades
Private mvarURL As String 'copia local
Private mvarHuboError As Boolean 'copia local
Private mvarStatusCode As String 'copia local
Private mvarStatusText As String 'copia local
Private mvarUsarProxy As String 'copia local
Private mvarNoUsarProxy As String 'copia local
Private mvarUsuario As String 'copia local
Private mvarPassword As String 'copia local
Private mvarFichero As String 'copia local
Private mvarContenidoDescargado As String 'copia local
Private mvarPuerto As Long 'copia local
Private mvarQContentType As Boolean 'copia local
Private mvarQContentLength As Boolean 'copia local
Private mvarQLastModified As Boolean 'copia local
Private mvarQVersion As Boolean 'copia local
Private mvarQRawHeaders As Boolean 'copia local
Private mvarQRawHeadersCrLf As Boolean 'copia local
Private mvarQForwarded As Boolean 'copia local
Private mvarQServer As Boolean 'copia local
Private mvarQRequestMethod As Boolean 'copia local
Private mvarQPragma As Boolean 'copia local
Private mvarQContentLengthStr As String 'copia local
Private mvarQContentTypeStr As String 'copia local
Private mvarQForwardedStr As String 'copia local
Private mvarQLastModifiedStr As String 'copia local
Private mvarQPragmaStr As String 'copia local
Private mvarQRawHeadersStr As String 'copia local
Private mvarQRequestMethodStr As String 'copia local
Private mvarQServerStr As String 'copia local
Private mvarQVersionStr As String 'copia local
Private mvarQRawHeadersCrLfStr As String 'copia local
Private mvarQExpires As Boolean 'copia local
Private mvarQExpiresStr As String 'copia local
Private mvarBytesBloqueDescarga As Long 'copia local


Private bytesRecibidosTotales As Long
'Para activar este evento, use RaiseEvent con la siguiente sintaxis:
'RaiseEvent Progreso[(arg1, arg2, ... , argn)]
Public Event Progreso(ByVal BytesTotales As Long, ByVal BytesRecibidos As Long, ByVal porcentaje As Double, ByVal SegundosTranscurridos As Long, ByVal SegundosRestantes As Double, BytesPorSegundo As Long, Cancelar As Boolean)

Public Property Let BytesBloqueDescarga(ByVal vData As Long)
Attribute BytesBloqueDescarga.VB_Description = "Devuelve o establece el tamaño de cada bloque que se descarga de internet. No conviene que sea demasiado grande porque se tardará más en poder cancelar una descarga y en mostrar el porcentaje descargado. "
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.BytesBloqueDescarga = 5
    mvarBytesBloqueDescarga = vData
End Property


Public Property Get BytesRecibidos() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.BytesBloqueDescarga
    BytesRecibidos = bytesRecibidosTotales
End Property

Public Property Get BytesBloqueDescarga() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.BytesBloqueDescarga
    BytesBloqueDescarga = mvarBytesBloqueDescarga
End Property




Private Function LastSystemError() As String
Dim sError As String * 500, lErrNum As Long, lErrMsg As Long

lErrNum = GetLastError()
lErrMsg = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, _
  ByVal 0&, lErrNum, 0, sError, Len(sError), 0)
LastSystemError = left(sError, InStr(sError, Chr(0) - 1))
End Function

Private Sub CierraConexiones()
If hHttpOpenRequest <> 0 Then InternetCloseHandle hHttpOpenRequest
If hInternetConnect <> 0 Then InternetCloseHandle hInternetConnect
If hInternetSession <> 0 Then InternetCloseHandle hInternetSession
End Sub

Public Property Let QExpiresStr(ByVal vData As String)
Attribute QExpiresStr.VB_Description = "Devuelve el valor de la cabecera EXPIRES si el valor de QExpires es true."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.QExpiresStr = 5
    mvarQExpiresStr = vData
End Property


Public Property Get QExpiresStr() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QExpiresStr
    QExpiresStr = mvarQExpiresStr
End Property



Public Property Let QExpires(ByVal vData As Boolean)
Attribute QExpires.VB_Description = "Devuelve o establece si queremos obtener el valor de la cabeceraEXPIRES."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.QExpires = 5
    mvarQExpires = vData
End Property


Public Property Get QExpires() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QExpires
    QExpires = mvarQExpires
End Property




Public Property Let QRawHeadersCrLfStr(ByVal vData As String)
Attribute QRawHeadersCrLfStr.VB_Description = "Devuelve el valor de la cabecera RAW_HEADERS_CRLF si la propiedad QRawHeadersCrLf es true."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.QRawHeadersCrLfStr = 5
    mvarQRawHeadersCrLfStr = vData
End Property


Public Property Get QRawHeadersCrLfStr() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QRawHeadersCrLfStr
    QRawHeadersCrLfStr = mvarQRawHeadersCrLfStr
End Property




Private Sub InicializaCabecerasStr()
mvarQContentLengthStr = ""
mvarQContentTypeStr = ""
mvarQForwardedStr = ""
mvarQLastModifiedStr = ""
mvarQPragmaStr = ""
mvarQRawHeadersStr = ""
mvarQRequestMethodStr = ""
mvarQServerStr = ""
mvarQVersionStr = ""
mvarQRawHeadersCrLfStr = ""
End Sub


Public Property Get QVersionStr() As String
Attribute QVersionStr.VB_Description = "Devuelve el valor de la cabecera VERSION si la propiedad QVersion es true."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QVersionStr
    QVersionStr = mvarQVersionStr
End Property





Public Property Get QServerStr() As String
Attribute QServerStr.VB_Description = "Devuelve el valor de la cabecera SERVER si la propiedad QServer es true."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QServerStr
    QServerStr = mvarQServerStr
End Property





Public Property Get QRequestMethodStr() As String
Attribute QRequestMethodStr.VB_Description = "Devuelve el valor de la cabecera REQUEST_METHOD si la propiedad QRequestMethod es true."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QRequestMethodStr
    QRequestMethodStr = mvarQRequestMethodStr
End Property





Public Property Get QRawHeadersStr() As String
Attribute QRawHeadersStr.VB_Description = "Devuelve el valor de la cabecera RAW_HEADERS si la propiedad QRawHeaders es true."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QRawHeadersStr
    QRawHeadersStr = mvarQRawHeadersStr
End Property





Public Property Get QPragmaStr() As String
Attribute QPragmaStr.VB_Description = "Devuelve el valor de la cabecera PRAGMA si la propiedad QPragma es true."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QPragmaStr
    QPragmaStr = mvarQPragmaStr
End Property





Public Property Get QLastModifiedStr() As String
Attribute QLastModifiedStr.VB_Description = "Devuelve el valor de la cabecera LAST_MODIFIED si la propiedad QLastModified es true."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QLastModifiedStr
    QLastModifiedStr = mvarQLastModifiedStr
End Property





Public Property Get QForwardedStr() As String
Attribute QForwardedStr.VB_Description = "Devuelve el contenido de la cabecera FORWARDED si la propiedad QForwarded es true."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QForwardedStr
    QForwardedStr = mvarQForwardedStr
End Property





Public Property Get QContentTypeStr() As String
Attribute QContentTypeStr.VB_Description = "Devuelve la información de la cabecera CONTENT_TYPE si la propiedad QContentType es true."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QContentTypeStr
    QContentTypeStr = mvarQContentTypeStr
End Property





Public Property Get QContentLengthStr() As String
Attribute QContentLengthStr.VB_Description = "Devuelve la información de la cabecera CONTENT_LENGTH, si la propiedad QContentLength as true."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QContentLengthStr
    QContentLengthStr = mvarQContentLengthStr
End Property



Public Property Let QPragma(ByVal vData As Boolean)
Attribute QPragma.VB_Description = "Devuelve o establece si deseamos recuperar la cabecera PRAGMA."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.QPragma = 5
    mvarQPragma = vData
End Property


Public Property Get QPragma() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QPragma
    QPragma = mvarQPragma
End Property



Public Property Let QRequestMethod(ByVal vData As Boolean)
Attribute QRequestMethod.VB_Description = "Devuelve o establece si deseamos recuperar la cabecera REQUEST_METHOD."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.QRequestMethod = 5
    mvarQRequestMethod = vData
End Property


Public Property Get QRequestMethod() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QRequestMethod
    QRequestMethod = mvarQRequestMethod
End Property



Public Property Let QServer(ByVal vData As Boolean)
Attribute QServer.VB_Description = "Devuelve o establece si deseamos recuperar la cabecera SERVER."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.QServer = 5
    mvarQServer = vData
End Property


Public Property Get QServer() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QServer
    QServer = mvarQServer
End Property



Public Property Let QForwarded(ByVal vData As Boolean)
Attribute QForwarded.VB_Description = "Devuelve o establece si deseamos recuperar la cabecera FORWARDED."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.QForwarded = 5
    mvarQForwarded = vData
End Property


Public Property Get QForwarded() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QForwarded
    QForwarded = mvarQForwarded
End Property



Public Property Let QRawHeadersCrLf(ByVal vData As Boolean)
Attribute QRawHeadersCrLf.VB_Description = "Devuelve o establece si deseamos recuperar la cabecera RAW_HEADERS_CRLF."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.QRawHeadersCrLf = 5
    mvarQRawHeadersCrLf = vData
End Property


Public Property Get QRawHeadersCrLf() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QRawHeadersCrLf
    QRawHeadersCrLf = mvarQRawHeadersCrLf
End Property



Public Property Let QRawHeaders(ByVal vData As Boolean)
Attribute QRawHeaders.VB_Description = "Devuelve o establece si deseamos recuperar la cabecera RAW_HEADERS."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.QRawHeaders = 5
    mvarQRawHeaders = vData
End Property


Public Property Get QRawHeaders() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QRawHeaders
    QRawHeaders = mvarQRawHeaders
End Property



Public Property Let QVersion(ByVal vData As Boolean)
Attribute QVersion.VB_Description = "Devuelve o establece si deseamos recuperar la cabecera VERSION."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.QVersion = 5
    mvarQVersion = vData
End Property


Public Property Get QVersion() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QVersion
    QVersion = mvarQVersion
End Property



Public Property Let QLastModified(ByVal vData As Boolean)
Attribute QLastModified.VB_Description = "Devuelve o establece si deseamos recuperar la cabecera LAST_MODIFIED"
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.QLastModified = 5
    mvarQLastModified = vData
End Property


Public Property Get QLastModified() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QLastModified
    QLastModified = mvarQLastModified
End Property



Public Property Let QContentLength(ByVal vData As Boolean)
Attribute QContentLength.VB_Description = "Devuelve o establece si deseamos recuperar la cabecera CONTENT_LENGTH"
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.QContentLength = 5
    mvarQContentLength = vData
End Property


Public Property Get QContentLength() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QContentLength
    QContentLength = mvarQContentLength
End Property



Public Property Let QContentType(ByVal vData As Boolean)
Attribute QContentType.VB_Description = "Devuelve o establece si deseamos recuperar la cabecera CONTENT_TYPE"
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.QContentType = 5
    mvarQContentType = vData
End Property


Public Property Get QContentType() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.QContentType
    QContentType = mvarQContentType
End Property



Public Sub Descargar(Optional TipoAccion As jrDownTipoAccion = jrDownDescargar)
Attribute Descargar.VB_Description = "Descarga la información o el contenido solicitado. Para determinar qué información descargar revise las propiedades que comienzan con Q."
Dim BytesTotales As Long, BytesRecibidos As Long, porcentaje As Double
Dim sBuffer As String, Res As Integer, UsarPuerto As Long, NumBloques As Long
Dim Fich As Long, Contenido As String, aux As String
Dim Cancelar As Boolean
Dim TiempoTranscurrido As Long, TiempoRestante As Double, TiempoInicio As Long, BytesSegundo As Double

On Error Resume Next

Dim ChunkSize As Long

ChunkSize = mvarBytesBloqueDescarga

'inicializo propiedades
InicializaCabecerasStr
mvarHuboError = False
mvarStatusCode = ""
mvarStatusText = ""
mvarContenidoDescargado = ""
Cancelar = False

'compruebo que la URL esté introdocida y sea sintácticamente correcta
URLCorrecta = ProcesaURL()
If Not URLCorrecta Then
    mvarHuboError = True
    Exit Sub
End If
'Creo buffer para recibir el fichero
sBuffer = Space(ChunkSize)
'Creo una conexión a internet
If TipoConexion = INTERNET_OPEN_TYPE_PROXY Then
    hInternetSession = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PROXY, mvarUsarProxy, mvarNoUsarProxy, 0)
Else
    hInternetSession = InternetOpen(scUserAgent, TipoConexion, vbNullString, vbNullString, 0)
End If
If hInternetSession = 0 Then
    ProcesaError ERROR_INTERNETOPEN
    Exit Sub
End If
'me conecto con el servidor
UsarPuerto = Puerto
hInternetConnect = InternetConnect(hInternetSession, Servidor, UsarPuerto, mvarUsuario, mvarPassword, Servicio, 0, 0)
If hInternetConnect = 0 Then
    ProcesaError ERROR_INTERNETCONNECT
    Exit Sub
End If
'abro una petición para el fichero solicitado
hHttpOpenRequest = HttpOpenRequest(hInternetConnect, "GET", Objeto, "HTTP/1.0", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
If hHttpOpenRequest = 0 Then
    ProcesaError ERROR_INTERNETOPENREQUEST
    Exit Sub
End If
'envío la petición
Res = HttpSendRequest(hHttpOpenRequest, vbNullString, 0, 0, 0)
If Res = 0 Then
    ProcesaError ERROR_INTERNETSENDREQUEST
    Exit Sub
End If
'miro la cabecera para ver si el fichero existe
If GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_STATUS_TEXT, sBuffer) Then
    mvarStatusText = sBuffer
Else
    ProcesaError ERROR_INTERNETQUERYINFO
    Exit Sub
End If
If GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_STATUS_CODE, sBuffer) Then
    mvarStatusCode = sBuffer
    If left(sBuffer, 1) <> "2" Then
        ProcesaError CLng(mvarStatusCode)
        Exit Sub
    End If
Else
    ProcesaError ERROR_INTERNETQUERYINFO
    Exit Sub
End If
'miro la longitud del contenido a descargar
If GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_CONTENT_LENGTH, sBuffer) Then
    If IsNumeric(sBuffer) Then
        BytesTotales = CLng(sBuffer)
    Else
        BytesTotales = -1
    End If
    If mvarQContentLength Then mvarQContentLengthStr = sBuffer
Else
    BytesTotales = -1
End If
'busco el resto de valores de la cabecera, si lo ha solicitado el usuario
If mvarQContentType Then
    If GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_CONTENT_TYPE, sBuffer) Then mvarQContentTypeStr = sBuffer
End If
If mvarQExpires Then
    If GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_EXPIRES, sBuffer) Then mvarQExpiresStr = sBuffer
End If
If mvarQLastModified Then
    If GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_LAST_MODIFIED, sBuffer) Then mvarQLastModifiedStr = sBuffer
End If
If mvarQPragma Then
    If GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_PRAGMA + HTTP_QUERY_FLAG_REQUEST_HEADERS, sBuffer) Then mvarQPragmaStr = sBuffer
End If
If mvarQVersion Then
    If GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_VERSION, sBuffer) Then mvarQVersionStr = sBuffer
End If
If mvarQRawHeaders Then
    If GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_RAW_HEADERS, sBuffer) Then mvarQRawHeadersStr = sBuffer
End If
If mvarQRawHeadersCrLf Then
    If GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_RAW_HEADERS_CRLF, sBuffer) Then mvarQRawHeadersCrLfStr = sBuffer
End If
If mvarQForwarded Then
    If GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_FORWARDED, sBuffer) Then mvarQForwardedStr = sBuffer
End If
If mvarQServer Then
    If GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_SERVER, sBuffer) Then mvarQServerStr = sBuffer
End If
If mvarQRequestMethod Then
    If GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_FLAG_REQUEST_HEADERS + HTTP_QUERY_REQUEST_METHOD, sBuffer) Then mvarQRequestMethodStr = sBuffer
End If
'si sólo queríamos información del archivo, ya acabamos
If TipoAccion = jrDownSoloInformacion Then
    CierraConexiones
    Exit Sub
End If
'si tengo que grabar un fichero
If mvarFichero <> "" Then
    'abro el fichero
    Fich = FreeFile()
    Open mvarFichero For Binary As Fich
Else
    'si conozco la longitud del archivo dimensiono el string porque se gana mucho en velocidad
    If BytesTotales <> -1 Then mvarContenidoDescargado = Space(BytesTotales)
End If
'leo el archivo de internet
'inicio el contador de tiempo
TiempoInicio = GetTickCount()
'inicio los bytes recibidos y el espacio a leer de cada vez
bytesRecibidosTotales = 0
Res = 1: BytesRecibidos = ChunkSize
sBuffer = Space(ChunkSize): NumBloques = 0
While Res <> 0 And BytesRecibidos <> 0 And Not Cancelar
    Res = InternetReadFile(hHttpOpenRequest, sBuffer, ChunkSize, BytesRecibidos)
    If Res = 0 Then
        mvarStatusText = LastSystemError()
        mvarStatusCode = "999"
    Else
        If BytesRecibidos > 0 Then
            aux = left(sBuffer, BytesRecibidos)
            'si estoy grabando un fichero...
            If mvarFichero <> "" Then
                Put Fich, , aux
            Else
                'si conozco el tamaño del archivo
                If BytesTotales <> -1 Then
                    Mid(mvarContenidoDescargado, (NumBloques * ChunkSize) + 1, BytesRecibidos) = aux
                    NumBloques = NumBloques + 1
                Else
                    mvarContenidoDescargado = mvarContenidoDescargado + aux
                End If
            End If
            bytesRecibidosTotales = bytesRecibidosTotales + BytesRecibidos
            TiempoTranscurrido = GetTickCount() - TiempoInicio
            BytesSegundo = bytesRecibidosTotales / (TiempoTranscurrido / 1000)
 
            'calcular el porcentaje descargado y lanzar el evento progreso
            If BytesTotales <> -1 Then
                porcentaje = (bytesRecibidosTotales * 100) / BytesTotales
                TiempoRestante = (BytesTotales / BytesSegundo) - (TiempoTranscurrido / 1000)
            Else
                porcentaje = 0
                TiempoRestante = 0
            End If
            If TiempoRestante < 0 Then TiempoRestante = 0
            RaiseEvent Progreso(BytesTotales, bytesRecibidosTotales, porcentaje, CLng(TiempoTranscurrido / 1000), CLng(TiempoRestante), CLng(BytesSegundo), Cancelar)
            DoEvents
        End If
    End If
Wend
'cierro el fichero
If mvarFichero <> "" Then Close Fich
'si el usuario canceló borro el fichero
If Cancelar Then
    If mvarFichero <> "" Then
        If Dir(mvarFichero) <> "" Then Kill mvarFichero
    End If
    ProcesaError ERROR_CANCELADO
    Exit Sub
Else
    'si acabó por un error borro el fichero (dejo la variable por si el contenido sirviera para algo)
    If Res = 0 Then
        If mvarFichero <> "" Then
            If Dir(mvarFichero) <> "" Then Kill mvarFichero
        End If
        ProcesaError ERROR_DESCARGA
        Exit Sub
    End If
End If
CierraConexiones
End Sub

Private Function DameDirectorio(Archivo As String) As String
Dim i As Long

'busco la última barra
i = InStrRev(Archivo, "\")
If i = 0 Then
    DameDirectorio = CurDir()
Else
    DameDirectorio = left(Archivo, i - 1)
End If
End Function

Private Function GetQueryInfo(ByVal hHttpRequest As Long, ByVal iInfoLevel As Long, Valor As String) As Boolean
Dim sBuffer As String * 1024, lBufferLength As Long

lBufferLength = Len(sBuffer)
GetQueryInfo = CBool(HttpQueryInfo(hHttpRequest, iInfoLevel, ByVal sBuffer, lBufferLength, 0))
lBufferLength = InStr(sBuffer, Chr(0))
Valor = left(sBuffer, lBufferLength - 1)
End Function



Public Property Let Puerto(ByVal vData As Long)
Attribute Puerto.VB_Description = "Si no se rellena se utilizará el puerto por defecto según el protocolo."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Puerto = 5
    mvarPuerto = vData
End Property


Public Property Get Puerto() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Puerto
If mvarPuerto = 0 Then
    Select Case Protocolo
        Case "http": Puerto = INTERNET_DEFAULT_HTTP_PORT
        Case "https": Puerto = INTERNET_DEFAULT_HTTPS_PORT
    End Select
Else
    Puerto = mvarPuerto
End If
End Property






Public Property Get ContenidoDescargado() As String
Attribute ContenidoDescargado.VB_Description = "Devuelve el contenido descargado de internet si se dejó en blanco la propiedad Fichero."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ContenidoDescargado
    ContenidoDescargado = mvarContenidoDescargado
End Property



Public Property Let Fichero(ByVal vData As String)
Attribute Fichero.VB_Description = "Devuelve o establece el nombre del fichero donde se guardará el contenido descargado. Si se deja en blanco el contenido descargado se devolverá  en la propiedad ContenidoDescargado."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Fichero = 5
    mvarFichero = vData
End Property


Public Property Get Fichero() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Fichero
    Fichero = Trim(mvarFichero)
End Property



Public Property Let Password(ByVal vData As String)
Attribute Password.VB_Description = "Devuelve o establece la password con la que se realizará la conexión."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Password = 5
    mvarPassword = vData
End Property


Public Property Get Password() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Password
    Password = mvarPassword
End Property



Public Property Let Usuario(ByVal vData As String)
Attribute Usuario.VB_Description = "Devuelve o establece el usuario con el que se realizará la conexión."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Usuario = 5
    mvarUsuario = vData
End Property


Public Property Get Usuario() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Usuario
    Usuario = mvarUsuario
End Property




Public Property Let NoUsarProxy(ByVal vData As String)
Attribute NoUsarProxy.VB_Description = "Establece la lista dedirecciones para las que no se usará el  proxy a usar. Si se deja vacía se usará la configuración predeterminada del sistema."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.NoUsarProxy = 5
    mvarNoUsarProxy = vData
    If mvarUsarProxy = "" And mvarNoUsarProxy = "" Then
        TipoConexion = INTERNET_OPEN_TYPE_PRECONFIG
    Else
        TipoConexion = INTERNET_OPEN_TYPE_PROXY
    End If
End Property


Public Property Get NoUsarProxy() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.NoUsarProxy
    NoUsarProxy = mvarNoUsarProxy
End Property



Public Property Let UsarProxy(ByVal vData As String)
Attribute UsarProxy.VB_Description = "Establece la lista de proxys a usar. Si se deja vacía se usará la configuración predeterminada del sistema."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.UsarProxy = 5
    mvarUsarProxy = vData
    If mvarUsarProxy = "" And mvarNoUsarProxy = "" Then
        TipoConexion = INTERNET_OPEN_TYPE_PRECONFIG
    Else
        TipoConexion = INTERNET_OPEN_TYPE_PROXY
    End If

End Property


Public Property Get UsarProxy() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.UsarProxy
    UsarProxy = mvarUsarProxy
End Property




Private Sub ProcesaError(Numero As Long)

mvarStatusCode = format(Numero, "000")

Select Case Numero
    Case ERROR_URL: mvarStatusText = "URL incorrecta."
    Case ERROR_INTERNETOPEN: mvarStatusText = "Error en InternetOpen."
    Case ERROR_INTERNETCONNECT: mvarStatusText = "Error en InternetConnect."
    Case ERROR_INTERNETOPENREQUEST: mvarStatusText = "Error en InternetOpenRequest."
    Case ERROR_INTERNETSENDREQUEST: mvarStatusText = "Error en InternetSendRequest."
    Case ERROR_INTERNETQUERYINFO: mvarStatusText = "Error en InternetQueryInfo."
    Case ERROR_INTERNETREADFILE: mvarStatusText = "Error en InternetReadFile."
    Case ERROR_FICHERO: mvarStatusText = "No se ha podido crear el fichero de destino."
    Case ERROR_DESCARGA: mvarStatusText = "Ha ocurrido un error durante la descarga."
    Case ERROR_CANCELADO: mvarStatusText = "Descarga cancelada por el usuario."
End Select
CierraConexiones
mvarHuboError = True
End Sub

Public Property Get StatusText() As String
Attribute StatusText.VB_Description = "Devuelve la descripción del error."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.StatusText
    StatusText = mvarStatusText
End Property





Public Property Get StatusCode() As String
Attribute StatusCode.VB_Description = "Devuelve el código de error."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.StatusCode
    StatusCode = mvarStatusCode
End Property





Public Property Get HuboError() As Boolean
Attribute HuboError.VB_Description = "Devuelve si la operación causó un error. Consultar StatusCode y StatusText para más información."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.HuboError
    HuboError = mvarHuboError
End Property



Public Property Let URL(ByVal vData As String)
Attribute URL.VB_Description = "Devuelve y establece la URL que queremos descargar"
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.URL = 5
Dim i As Long, j As Long

mvarURL = vData
URLCorrecta = ProcesaURL()
        
End Property

Public Property Get URL() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.URL
    URL = mvarURL
End Property



Private Function ProcesaURL() As Boolean
Dim i As Long, j As Long

On Error GoTo ProcesaURL_Err

ProcesaURL = False
'descomponemos la url en protocolo, servidor y objeto
'busco el protocolo
i = InStr(mvarURL, "://")
If i = 0 Then
    'si no existe asumimos que es http
    Protocolo = "http"
    i = 1
Else
    Protocolo = LCase(mid(mvarURL, 1, i - 1))
    i = i + 3
End If
'sólo permitimos http
Select Case Protocolo
    Case "http":
    Case "https":
    Case Else: ProcesaError ERROR_URL
End Select
'busco el servidor
j = InStr(i, mvarURL, "/")
If j = 0 Then j = Len(mvarURL) + 1

Servidor = mid(mvarURL, i, j - i)

' ¿Tiene puerto?
If InStr(1, Servidor, ":", vbTextCompare) Then
    Puerto = CInt(val(mid$(Servidor, InStrRev(Servidor, ":") + 1)))
    Servidor = left$(Servidor, InStr(1, Servidor, ":") - 1)
    
End If

i = j + 1

'busco el objeto a descargar
If i > Len(mvarURL) Then
    Objeto = vbNullString
Else
    Objeto = "/" & mid(mvarURL, i)
End If
ProcesaURL = True

ProcesaURL_End:
    Exit Function

ProcesaURL_Err:
    ProcesaError ERROR_URL
    Resume ProcesaURL_End

End Function

Private Sub Class_Initialize()
TipoConexion = INTERNET_OPEN_TYPE_PRECONFIG
mvarQContentLength = True
Servicio = INTERNET_SERVICE_HTTP
mvarBytesBloqueDescarga = 512
End Sub


Private Sub Class_Terminate()
CierraConexiones
End Sub


