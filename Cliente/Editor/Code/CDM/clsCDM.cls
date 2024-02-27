VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Descargador
Private WithEvents downloader As downloader
Attribute downloader.VB_VarHelpID = -1

' Eventos
Public Event compartido(Version As Long)
Public Event actualizado(Version As Long)
Public Event progreso(actual As Single, Maximo As Single)

' Variables
Private CDM_VersionActual           As Integer
Private CDM_Repositorio_Nombre      As String
Private CDM_UserSession             As String
Private CDM_Carpeta                 As String

' Variables temproales para el update
Private update_ArchivosUpdate As Collection
Private update_CantidadArchivosUpdateDescargados As Integer
Private update_bytesDescargados As Single
Private update_bytesTotales As Single

' Estado
Public estado As eEstadoCDM
Private ultimoError_ As String

' Objetos
Private navegador As NavegadorWeb
Private request As CHTTPRequest

' Constantes
#If Produccion = 0 Then
    Private Const URL = "https://cdm-dev.tierrasdelsur.cc"
#ElseIf Produccion = 1 Then
    Private Const URL = "https://cdm-dev.tierrasdelsur.cc"
#ElseIf Produccion = 2 Then
    Private Const URL = "https://cdm-dev.tierrasdelsur.cc"
#End If

Private Const PATH_LOGIN = "ingresar"
Private Const PATH_REPORTE_BUG = "bugs"
Private Const PATH_RECURSO = "recursos"
Private Const PATH_REPO = "r/"
Private Const PATH_LOG = "log"
Private Const PATH_COMMIT_BEGIN = "commit/begin"
Private Const PATH_COMMIT_END = "commit/end"
Private Const PATH_COMMIT_UPLOAD = "commit/upload"
Private Const PATH_UPDATE = "update"

' Usuario logueado
Private pUsuario As CDM_Usuario

' Devolvemos el usuario logueado
Public Property Get Usuario() As CDM_Usuario
    If estado = conectado Then
        Set Usuario = pUsuario
    Else
        Set Usuario = Nothing
    End If
End Property

Public Property Get ultimoError() As String
    ultimoError = ultimoError_
End Property

Public Property Let ultimoError(ByVal error As String)
    ultimoError_ = error
    If Len(error) > 0 Then
        Call Log("Error: " & error)
    End If
End Property

Private Sub limpiarUpdate()
    Set update_ArchivosUpdate = Nothing
    Set downloader = Nothing
    update_CantidadArchivosUpdateDescargados = 0
    update_bytesDescargados = 0
    update_bytesTotales = 0
End Sub
Public Sub iniciar(ControlInet As Inet, ControlTimer As VB.timer, UserAgent As String, repoNombre As String, carpetaTrabajo As String)

    'Creamos el navegador
    Set navegador = New NavegadorWeb
    Call navegador.crear(ControlInet, ControlTimer)
    
    'Creamos la solicitud
    Set request = New CHTTPRequest
    
    request.Host = URL
    request.UserAgent = UserAgent
    
    ' Variables
    CDM_Carpeta = carpetaTrabajo
    CDM_Repositorio_Nombre = repoNombre
    
    CDM_VersionActual = cargarVersionLocal()
    
    Call Log("Cerebro de Mono Iniciado")
End Sub

Private Function cargarVersionLocal() As Long
    cargarVersionLocal = val(GetVar(CDM_Carpeta & "cdm.ini", "REPOSITORIOS", CDM_Repositorio_Nombre))
End Function

Private Sub establecerVersionLocal(Version As Long)
    CDM_VersionActual = Version
    ' Actualizo el .ini
    Call WriteVar(CDM_Carpeta & "cdm.ini", "REPOSITORIOS", CDM_Repositorio_Nombre, Version)
End Sub

Public Property Get Version() As Integer
    Version = CDM_VersionActual
End Property

'Metodo: funcion / procedimiento donde se produjo el error.
'Error: Error que se produjo (Numero, descripcion, linea de codigo).
'Explicacion: Explicacion del usuario sobre lo que hizo antes de que aparezca el bug.
Public Function RepotarBug(metodo As String, error As String, explicacion As String, infopc As String) As Boolean
    Dim resultado As CHTTPResponse
       
    'Preparamos el request
    request.method = httppost
    request.Path = PATH_REPORTE_BUG
   
    request.AddFormData "token", CDM_UserSession
    request.AddFormData "metodo", metodo
    request.AddFormData "error", error
    request.AddFormData "explicacion", explicacion
    request.AddFormData "datospc", infopc
    
    'Ejecutamos
    Set resultado = navegador.ejecutarConsulta(request)
    
    Call request.reset
    
    RepotarBug = False
    
    'Obtenemos la respuesta
    If Not resultado Is Nothing Then
        If resultado.Code = 200 Then
            ultimoError = ""
            RepotarBug = True
        Else
            ultimoError = obtenerResponseErrorJSON(resultado)
        End If
    Else
        ultimoError = obtenerErrorNavegador
    End If
End Function

' Solicita para un tipo de recurso @var tipo, un identificador de recurso disponible
' Retorna el identificador de recurso o -1 si se produjo algun problema.
Public Function SolicitarRecurso(Tipo As String) As Long
    Dim resultado As CHTTPResponse
       
    'Preparamos el request
    request.method = httppost
    request.Path = PATH_RECURSO
   
    request.AddFormData "token", CDM_UserSession
    request.AddFormData "tipo", Tipo
    
    'Ejecutamos
    Set resultado = navegador.ejecutarConsulta(request)
    
    Call request.reset
    
    SolicitarRecurso = -1
    
    'Obtenemos la respuesta
    If Not resultado Is Nothing Then
        If resultado.Code = 200 Then
            SolicitarRecurso = CLng(val(resultado.bodyJSON().item("id")))
            ultimoError = ""
        Else
            ultimoError = obtenerResponseErrorJSON(resultado)
        End If
    Else
        ultimoError = obtenerErrorNavegador
    End If
End Function

' Esto no lo compilo cuando lo publico por las dudas de que editen memoria
#If Testeo = 1 Then

Public Function LoginDummy() As Boolean
    
    estado = conectado
    
    Set pUsuario = New CDM_Usuario
End Function

#End If

'Inicia session en el CDM
Public Function Login(user As String, Password As String) As Boolean

    Dim resultado As CHTTPResponse
    
    estado = Conectando
    
    'Preparamos el request
    request.method = httppost
    request.Path = PATH_LOGIN
   
    request.AddFormData "usuario", user
    request.AddFormData "clave", Password
    
    'Ejecutamos
    Set resultado = navegador.ejecutarConsulta(request)
    
    Call request.reset
    
    'Obtenemos la respuesta
    If Not resultado Is Nothing Then
        If resultado.Code = 200 Then
        
            Call cargarDatosDeSession(resultado)
            
            estado = conectado
            ultimoError = ""
            Login = True
        Else
            ultimoError = obtenerResponseErrorJSON(resultado)
            estado = error
            Login = False
        End If
    Else
        estado = error
        ultimoError = obtenerErrorNavegador
        Login = False
    End If
    
End Function

Private Function Repositorio_CommitArchivo(archivo As CDM_Archivo) As Boolean

    Dim resultado As CHTTPResponse
       
    'Preparamos el
    request.method = httppost
    request.Path = PATH_REPO & CDM_Repositorio_Nombre & "/" & PATH_COMMIT_UPLOAD
   
    request.AddFormData "token", CDM_UserSession
    request.AddFormData "origen", archivo.destino
    request.AddFormData "cambios", JSON.toString(archivo.cambios)
    request.AddFile archivo.archivo, "archivo"
    
    'Ejecutamos
    Set resultado = navegador.ejecutarConsulta(request)
    
    Call request.reset
    
    Repositorio_CommitArchivo = False
    
    'Obtenemos la respuesta
    If Not resultado Is Nothing Then
        If resultado.Code = 200 Then
            Repositorio_CommitArchivo = True
            ultimoError = ""
        Else
            ultimoError = obtenerResponseErrorJSON(resultado)
        End If
    Else
        ultimoError = obtenerErrorNavegador
    End If

End Function
Private Function Repositorio_Begin(comentario As String) As Boolean
    Dim resultado As CHTTPResponse
       
    'Preparamos el
    request.method = httppost
    request.Path = PATH_REPO & CDM_Repositorio_Nombre & "/" & PATH_COMMIT_BEGIN
   
    request.AddFormData "token", CDM_UserSession
    request.AddFormData "comentario", comentario
    request.AddFormData "miversion", CDM_VersionActual
    
    'Ejecutamos
    Set resultado = navegador.ejecutarConsulta(request)
    
    Call request.reset
    
    Repositorio_Begin = False
    
    'Obtenemos la respuesta
    If Not resultado Is Nothing Then
        If resultado.Code = 200 Then
            Repositorio_Begin = True
            ultimoError = ""
        Else
            ultimoError = obtenerResponseErrorJSON(resultado)
        End If
    Else
        ultimoError = obtenerErrorNavegador
    End If
End Function

' Retorna un diccionario con: fecha, usuario, comentario
Public Function Repositorio_ObtenerNovedades() As Collection

    Dim resultado As CHTTPResponse
       
    'Preparamos el
    request.method = httpGET
    request.Path = PATH_REPO & CDM_Repositorio_Nombre & "/" & PATH_LOG
   
    request.AddFormData "token", CDM_UserSession
    request.AddFormData "miversion", CDM_VersionActual
        
    'Ejecutamos
    Set resultado = navegador.ejecutarConsulta(request)
    
    Call request.reset
    
    Set Repositorio_ObtenerNovedades = Nothing
    
    'Obtenemos la respuesta
    If Not resultado Is Nothing Then
        If resultado.Code = 200 Then
            Set Repositorio_ObtenerNovedades = resultado.bodyJSON
            ultimoError = ""
        Else
            ultimoError = obtenerResponseErrorJSON(resultado)
        End If
    Else
        ultimoError = obtenerErrorNavegador
    End If
    
End Function

' Devuelve una lista con todos los archivos que el usuario se tiene que descargar
' para actualizar el Editor
Private Function Repositorio_Actualizar_ObtenerInfo() As Collection
    
    Dim resultado As CHTTPResponse
       
    'Preparamos el
    request.method = httpGET
    request.Path = PATH_REPO & CDM_Repositorio_Nombre & "/" & PATH_UPDATE
   
    request.AddFormData "token", CDM_UserSession
    request.AddFormData "miversion", CDM_VersionActual
        
    'Ejecutamos
    Set resultado = navegador.ejecutarConsulta(request)
    
    Call request.reset
    
    Set Repositorio_Actualizar_ObtenerInfo = Nothing
    
    'Obtenemos la respuesta
    If Not resultado Is Nothing Then
        If resultado.Code = 200 Then
            Set Repositorio_Actualizar_ObtenerInfo = resultado.bodyJSON
            ultimoError = ""
        Else
            ultimoError = obtenerResponseErrorJSON(resultado)
        End If
    Else
        ultimoError = obtenerErrorNavegador
    End If
    
End Function

Private Sub Repositorio_Actualizar_DescargarArchivo()
    Dim ubicacion As String
        
    ' ¿Termino?
    If update_CantidadArchivosUpdateDescargados = update_ArchivosUpdate.count Then
    
        Call Log("Descargas de archivos finalizadas.")

        Call Repositorio_Actualizar_Parchear
        
    Else
        ' Siguiente archivo a descargar
        update_CantidadArchivosUpdateDescargados = update_CantidadArchivosUpdateDescargados + 1
        
        ' Nommbre y ruta del archivo temporal
        ubicacion = HelperFiles.generarRandomNameFile(CDM_Carpeta & "TEMP", 20, "temp")
        
        Call Log("Descargando " & update_ArchivosUpdate(update_CantidadArchivosUpdateDescargados).archivo & " -> " & ubicacion & ".")
        
        ' Establesco las opciones del objeto que va a descargar los archivos
        downloader.Puerto = 3000
        downloader.URL = update_ArchivosUpdate(update_CantidadArchivosUpdateDescargados).archivo
        downloader.Fichero = ubicacion
        
        ' Descargo
        Call downloader.Descargar

        If Not downloader.StatusCode = 200 Then
            Call descargador_DownloadFailed(downloader.StatusCode, downloader.StatusText)
            Exit Sub
        End If
        
        ' Actualizo los bytes descargados
        update_bytesDescargados = update_bytesDescargados + downloader.BytesRecibidos
        
        ' Actualizo la ubicacion del archivo
        update_ArchivosUpdate(update_CantidadArchivosUpdateDescargados).archivo = ubicacion

        ' Actualizo el avance
        RaiseEvent progreso(CLng(update_bytesDescargados \ update_bytesTotales) * 100&, 100)
        DoEvents
        
        ' Continuamos con el sigueinte archivo
        Call Repositorio_Actualizar_DescargarArchivo
        
    End If
        
End Sub

' Esta procedimiento se ejecuta cuando se descargaron todos los archivos

Private Sub Repositorio_Actualizar_Parchear()
    Dim nuevaVersion As Long
    
    Call Log("Aplicando Cambios")
    
    ' Los parcheo
    nuevaVersion = versionador.Parchear(update_ArchivosUpdate)
    
    Call Log("Eliminando archivos temporales")
    ' Elimino los archivos temporales que descargue
    Call versionador.elimiarTemporales(update_ArchivosUpdate)
    Call limpiarUpdate
    
    ' ¿ok?
    If nuevaVersion > 0 Then
        ' Actualizo la variable de la version
        Call establecerVersionLocal(nuevaVersion)
    End If
    
    ' Listo
    RaiseEvent actualizado(nuevaVersion)
End Sub
Public Sub Repositorio_Actualizar()
    Dim archivo As CDM_Archivo
    Dim totalBytesDescargar As Long
    Dim infoUpdate As Collection
    Dim loopArchivo As Integer
    Dim encontramosArchivos As Boolean
    ' Obtengo la lista de de archivos que comprenden a la version ( URL del archivo, DESTINO, Version, Tamano)
    Set infoUpdate = Repositorio_Actualizar_ObtenerInfo
    
    
    ' ¿Pude?
    encontramosArchivos = False
    If Not infoUpdate Is Nothing Then
        If infoUpdate.count > 0 Then encontramosArchivos = True
    End If
    
    If Not encontramosArchivos Then
        ultimoError = "No se encontraron archivos para descargar en la actualización."
        RaiseEvent actualizado(0)
        Exit Sub
    End If
    
    ' Hago el inventario para arrancar a descargar
    update_bytesDescargados = 0 ' Contador de bytes actualizados
    update_CantidadArchivosUpdateDescargados = 0
    totalBytesDescargar = 0
    
    Set update_ArchivosUpdate = New Collection
    Set downloader = New downloader
    
    For loopArchivo = 1 To infoUpdate.count
        ' Obtengo la cantidad de bytes que tengo que descargar para hacer el progreso
        totalBytesDescargar = totalBytesDescargar + infoUpdate(loopArchivo).item("tamano")
        
        ' Creo el archivo
        Set archivo = New CDM_Archivo
                
        Call archivo.iniciar(infoUpdate(loopArchivo).item("ubicacion"), infoUpdate(loopArchivo).item("destino"), Nothing, infoUpdate(loopArchivo).item("version"))
        
        ' Agregamos a la lista de archivos que nos vamos a descargar
        Call update_ArchivosUpdate.Add(archivo)
    Next
    
    update_bytesTotales = totalBytesDescargar
    
    RaiseEvent progreso(0, 100)
    
    ' Comenzamos a descargar
    Call Repositorio_Actualizar_DescargarArchivo

End Sub
Private Function Repositorio_CommitEnd() As Long
    Dim resultado As CHTTPResponse
       
    'Preparamos el
    request.method = httppost
    request.Path = PATH_REPO & CDM_Repositorio_Nombre & "/" & PATH_COMMIT_END
   
    request.AddFormData "token", CDM_UserSession
    
    'Ejecutamos
    Set resultado = navegador.ejecutarConsulta(request)
    
    Call request.reset
    
    Repositorio_CommitEnd = 0
    
    'Obtenemos la respuesta
    If Not resultado Is Nothing Then
        If resultado.Code = 200 Then
            Repositorio_CommitEnd = resultado.bodyJSON.item("version")
        Else
            ultimoError = obtenerResponseErrorJSON(resultado)
        End If
    Else
        ultimoError = obtenerErrorNavegador
    End If
End Function

' Recibo los tipos de archivo que deseo comitear
Public Sub Repositorio_Compartir(tipoArchivos() As String, comentario As String, carpetaRoot As String)
    Dim total As Single
    Dim parcial As Single
    Dim archivos As Collection
    Dim archivo As CDM_Archivo
    Dim nuevaVersion As Long
    Dim commit As CDM_Commit
    
    ' Creamos el commit
    Set commit = versionador.generarCommit(tipoArchivos, carpetaRoot)
    commit.comentario = comentario
    
    ' Obtenemos los archivos reales
    Set archivos = commit.obtenerArchivos
    
    total = archivos.count + 2  ' Begin y Commit
    parcial = 0

    ' Inicio el commit
    Call Log("BEGIN COMPARTIR: " & commit.comentario & ".")
    
    If Not Repositorio_Begin(commit.comentario) Or archivos.count = 0 Then
        RaiseEvent compartido(0)
        Exit Sub
    End If
    
    ' Subo los archivos
    For Each archivo In commit.obtenerArchivos
        ' Subo el archivo
        Call Log("--- UPDATE BEGIN: " & archivo.archivo & "  -> " & archivo.destino & ". Cambios: " & JSON.toString(archivo.cambios))
        
        If Not Repositorio_CommitArchivo(archivo) Then
            RaiseEvent compartido(0)
            Exit Sub
        End If
        
        Call Log("--- UPDATE END")
        parcial = parcial + 1
        ' Informo el progreso
        RaiseEvent progreso(parcial, total)
        DoEvents
    Next
       
    nuevaVersion = Repositorio_CommitEnd()
    
    ' Limpiamos
    Call versionador.limpiar(commit, (nuevaVersion > 0))
    
    If nuevaVersion > 0 Then
        ' Guardamos
        Call establecerVersionLocal(nuevaVersion)
        Call Log("END COMPARTIR: Versión generada número " & nuevaVersion)
    Else
        Call Log("END COMPARTIR: Abortada.")
    End If
    
    ' Aviso de la nueva version que creo
    RaiseEvent compartido(nuevaVersion)

End Sub

'Inicia session en el CDM
Public Function estoyActualizado() As Boolean
    Dim infoRepo As Dictionary
    
    Set infoRepo = Repositorio_Info()
    
    If infoRepo Is Nothing Then
        estoyActualizado = False
        Exit Function
    End If
    
    If infoRepo.item("version") > CDM_VersionActual Then
        estoyActualizado = False
    Else
        estoyActualizado = True
    End If
End Function

' Obtenemos información de la versión del repositorio
Public Function Repositorio_Info() As Dictionary

    Dim resultado As CHTTPResponse
    
    'Preparamos el request
    request.method = httpGET
    request.Path = PATH_REPO & CDM_Repositorio_Nombre & "/info"
   
    request.AddFormData "token", CDM_UserSession
    request.AddFormData "miversion", CDM_VersionActual
        
    'Ejecutamos
    Set resultado = navegador.ejecutarConsulta(request)
    
    Call request.reset
    
    'Obtenemos la respuesta
    If Not resultado Is Nothing Then
        If resultado.Code = 200 Then
        
            Set Repositorio_Info = resultado.bodyJSON
        
            ultimoError = ""
        Else
            Set Repositorio_Info = Nothing
            ultimoError = obtenerResponseErrorJSON(resultado)
            estado = error
        End If
    Else
        Set Repositorio_Info = Nothing
        estado = error
        ultimoError = obtenerErrorNavegador
    End If
    
End Function
'*****************************************************************************
' Funciones Auxiliares
Private Sub cargarDatosDeSession(response As CHTTPResponse)

    Dim JSON As Object
    Set JSON = response.bodyJSON
    
    If JSON Is Nothing Then
        Exit Sub
    End If
    
    '
    CDM_UserSession = JSON.item("token")
    
    ' Creamos el usuario
    Set pUsuario = New CDM_Usuario_Real
    Call pUsuario.crear(JSON)
End Sub
Private Function obtenerResponseErrorJSON(response As CHTTPResponse) As String

    Dim JSON As Object
    Set JSON = response.bodyJSON
    
    ' ¿Hay respuesta?
    If JSON Is Nothing Then
        If Len(response.body) > 0 Then
            obtenerResponseErrorJSON = "La respuesta del servidor no es válida." & response.body
        Else
            obtenerResponseErrorJSON = "No se obtuvo una respuesta por parte del servidor."
        End If
        Exit Function
    End If
    
    ' ¿Tenemos el campo y es un string ?
    If Not VarType(JSON.item("mensaje")) = vbString Then
        obtenerResponseErrorJSON = "Se recibio una respuesta, en un formato válido, pero el error no está especificado." & response.body
        Exit Function
    End If
    
    ' Parseamos
    obtenerResponseErrorJSON = JSON.item("mensaje")
    
End Function

Private Function obtenerErrorNavegador() As String
    obtenerErrorNavegador = "Error " & navegador.error & ": " & navegador.errorDescripcion
End Function

Private Sub descargador_DownloadFailed(numero As Long, desc As String)
    ' Fallo la descarga
    ' Guardo el error. Elimino los archivos temporales
    Call limpiarUpdate
    ultimoError = "Falla al descargar archivo con actualizaciones (Error  " & numero & " " & desc & ")"
    RaiseEvent actualizado(0)
End Sub


Private Sub downloader_Progreso(ByVal BytesTotales As Long, ByVal BytesRecibidos As Long, ByVal porcentaje As Double, ByVal SegundosTranscurridos As Long, ByVal SegundosRestantes As Double, BytesPorSegundo As Long, Cancelar As Boolean)
    RaiseEvent progreso(minl((update_bytesDescargados + BytesRecibidos) / update_bytesTotales * 100, 100), 100)
End Sub

Private Sub Log(desc As String)
    Dim nFile As Integer
    
    nFile = FreeFile ' obtenemos un canal

    Open app.Path & "\cerebro_de_mono.log" For Append As #nFile
        desc = "[" & Now & "] - " & CDM_Repositorio_Nombre & " - " & desc
        Print #nFile, desc
        Debug.Print desc
    Close #nFile
End Sub
