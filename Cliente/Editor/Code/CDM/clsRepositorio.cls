VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRepositorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Event actualizaNovedades(novdades As Collection, error As String)
Public Event actualizado(archivos As Collection, version As Long, error As String)
Public Event compartido(version As Long, error As String)

Private WithEvents consolaDeWindows As clsConsolaWindows
Attribute consolaDeWindows.VB_VarHelpID = -1
Private usuario As String
Private clave As String
Private carpeta As String

Private accionPendiente  As String

Private contenido As Collection
Private archivos As Collection 'Archivos que se van a commitear

Private archivosRecibidos As Collection ' Archivos que actualice
Private versionActualizada As Long ' Version a la cual se actualizo por ultima vez


'Info temporal por asincronicidad
Private ultimoAlias As String
Private ultimoComentario As String

Private Const COMANDO_SVN As String = "\CDM\SVN\svn.exe"

Public Function obtenerTipoContenido() As Collection
    Set obtenerTipoContenido = contenido
End Function

Public Sub agregarTipoContenido(nombre As String)
    Call contenido.Add(nombre)
End Sub

Public Sub agregarArchivo(nombre As String)
    Call archivos.Add(nombre)
End Sub

Public Sub crear(usuario_ As String, clave_ As String, carpeta_ As String)
    usuario = usuario_
    clave = clave_
    carpeta = carpeta_
    
    Set contenido = New Collection
    Set archivos = New Collection
End Sub


Public Sub getUltimasModificaciones()

    Dim comando As String
    ' Con "log" obtemos las ultimas modificaciones
    ' Con "info" obtengo la version actual
    comando = Chr(34) & app.Path & COMANDO_SVN & Chr(34) & " log " & Chr(34) & Replace$(carpeta, "\", "/") & Chr(34) & " -r BASE:HEAD --xml --no-auth-cache --username=" & usuario & " --password=" & clave
    
    Debug.Print comando
    accionPendiente = "UM"
    
    Set consolaDeWindows = New clsConsolaWindows
    Call consolaDeWindows.RunCommand(frmMain.hwnd, comando)
    
    Debug.Print comando
End Sub
Private Sub archivosAgregados_(respuesta As String, error As String)
    
    Dim infoCommit As New Dictionary
    Dim comando As String
    
    'Error: Lo ignoro.
    
    'Genero el comentario en JSON
    infoCommit.item("usuario") = ultimoAlias
    infoCommit.item("comentario") = ultimoComentario
    
    accionPendiente = "CN"
    
    comando = Chr(34) & app.Path & COMANDO_SVN & Chr(34) & " commit " & Chr(34) & Replace$(carpeta, "\", "/") & "/" & Chr(34) & " --no-auth-cache --username=" & usuario & " --password=" & clave & " --message=" & Chr$(34) & Replace$(JSON.toString(infoCommit), Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
        
    ' Ejecuto el cmando
    Set consolaDeWindows = New clsConsolaWindows
    Call consolaDeWindows.RunCommand(frmMain.hwnd, comando)

    ' Libero la ifno del dictionary
    infoCommit.RemoveAll
    
    Debug.Print comando
End Sub

Public Sub compartirNovedades(usuarioAlias As String, comentario As String)

    Dim comando As String
    Dim archivo As Variant
    
    ' Guardo estos datos ya que los necesito en el contexto de otra funcion
    ultimoAlias = usuarioAlias
    ultimoComentario = comentario
    
    ' Agregamos los archivos al SVN (los mandamos todos, los que ya estan agregados los va a ignorar)
    comando = Chr(34) & app.Path & COMANDO_SVN & Chr(34) & " add "
    
    For Each archivo In archivos
        comando = comando & Chr(34) & archivo & Chr(34) & " "
    Next
    
    accionPendiente = "ADD"
        
    ' Ejecutamos el comando
    Set consolaDeWindows = New clsConsolaWindows
    Call consolaDeWindows.RunCommand(frmMain.hwnd, comando)

    Debug.Print comando
End Sub


Public Sub Actualizar()
    Dim comando As String
    
    comando = Chr(34) & app.Path & COMANDO_SVN & Chr(34) & " update " & Chr(34) & Replace$(carpeta, "\", "/") & "/" & Chr(34) & " --no-auth-cache --username=" & usuario & " --password=" & clave
    
    accionPendiente = "UP"
    Debug.Print comando
    Set consolaDeWindows = New clsConsolaWindows
    Call consolaDeWindows.RunCommand(frmMain.hwnd, comando)
    
    Debug.Print comando
End Sub

Public Function getInfo() As Dictionary

    Set getInfo = CDM.cerebro.Repositorio_Info
    
End Function

Public Function hayActualizaciones() As Boolean

End Function

Private Sub compartirNovedades_(respuesta As String, error As String)
    Dim version As String
    Dim comienzoVersion As Integer
    Dim lineas() As String
    Dim loopLinea As Integer
    
    If Len(error) > 0 Then
        RaiseEvent compartido(0, error)
        Exit Sub
    End If
    
    ' Busco la informacion que me importa
    lineas = Split(respuesta, vbNewLine)
    
    For loopLinea = LBound(lineas) To UBound(lineas)
        comienzoVersion = InStr(1, lineas(loopLinea), "Committed revision ") + Len("Committed revision ")
        
        If comienzoVersion > Len("Committed revision ") Then
            version = mid$(lineas(loopLinea), comienzoVersion, InStr(comienzoVersion, lineas(loopLinea), ".") - comienzoVersion)
        End If
    Next
    
    ' Ejecuto el evento
    RaiseEvent compartido(CLng(val(version)), error)
End Sub

' Se llama a actualizar
' Se llama a actualizar_obtenerVersionDeArchivo
' Se llama obtenidaVersionArchivo_
'    Llama a actualizado(archivos, version, error)
Private Sub obtenidaVersionArchivo_(respuesta As String, error As String)

    Dim doc As New MSXML2.DOMDocument
    Dim node As MSXML2.IXMLDOMNode
    
    Dim versiones As Collection
    Dim InfoUsuario As Dictionary
    Set versiones = New Collection
    
    Dim archivo As String
    Dim version As Long
    Dim archivoRepo As clsRepositorioArchivo
    
    If Len(error) > 0 Then
        RaiseEvent actualizado(Nothing, 0, error)
        Exit Sub
    End If
       
    If doc.loadXML(respuesta) = False Then
        RaiseEvent actualizado(archivosRecibidos, versionActualizada, "Error al interpretar version de archivos recibidos")
    Else
        Dim nodeList As MSXML2.IXMLDOMNodeList

        Set nodeList = doc.selectNodes("/info/entry")

        If Not nodeList Is Nothing Then
            
            ' Recorro todos los nodos, habra un nodo por archivo
            For Each node In nodeList
                ' Obtengo el nombre del archivo
                archivo = node.Attributes.getNamedItem("path").Text
                ' Obtenemos la version del archivo
                version = CLng(val(node.selectSingleNode("commit").Attributes.getNamedItem("revision").Text))
                                
                'Busco en la lista de archivos el archivo
                For Each archivoRepo In archivosRecibidos
                    If archivoRepo.nombre = archivo Then
                        archivoRepo.version = version
                        Exit For
                    End If
                Next
            
            Next node
            
        End If
        
        RaiseEvent actualizado(archivosRecibidos, versionActualizada, "")
    End If


End Sub

Private Sub actualizar_obtenerVersionDeArchivos(archivos As Collection)
    Dim comando As String
    Dim archivosString As String
    Dim elemento As clsRepositorioArchivo

    archivosString = ""
    For Each elemento In archivos
        archivosString = archivosString & Chr$(34) & app.Path & "\" & CStr(elemento.nombre) & Chr$(34) & " "
    Next
   
    comando = Chr(34) & app.Path & COMANDO_SVN & Chr(34) & " info " & archivosString & " --xml"
    
    accionPendiente = "IN"
    
    Debug.Print comando
    
    Set consolaDeWindows = New clsConsolaWindows
    Call consolaDeWindows.RunCommand(frmMain.hwnd, comando)
        
End Sub
Private Sub actualizar_(respuesta As String, error As String)
    Dim lineas() As String
    Dim loopLinea As Integer
    Dim version As Long
    Dim archivo As clsRepositorioArchivo
    
    If Len(error) > 0 Then
        RaiseEvent actualizado(Nothing, 0, error)
        Exit Sub
    End If
    
    lineas = Split(respuesta, vbNewLine)
    
    Set archivosRecibidos = New Collection
    
    For loopLinea = LBound(lineas) To UBound(lineas)
    
        If InStr(1, lineas(loopLinea), "At revision", vbTextCompare) > 0 Then
            version = val(mid$(lineas(loopLinea), InStr(1, lineas(loopLinea), "At revision", vbTextCompare) + Len("At revision")))
        ElseIf InStr(1, lineas(loopLinea), "Updating '", vbTextCompare) > 0 Then
            ' No me interesa saber la carpeta que esta actualizando
        ElseIf InStr(1, lineas(loopLinea), "Updated to revision", vbTextCompare) > 0 Then
             version = val(mid$(lineas(loopLinea), InStr(1, lineas(loopLinea), "Updated to revision", vbTextCompare) + Len("Updated to revision")))
        ElseIf Len(Trim$(lineas(loopLinea))) > 0 Then
            ' Obtenemos el nombre del archivo actualizado
            Set archivo = New clsRepositorioArchivo
        
            archivo.nombre = Trim(mid$(lineas(loopLinea), 2))
            archivo.version = -1
            
            Call archivosRecibidos.Add(archivo)
        End If
        
    Next loopLinea
    
    ' Tengo que obtener la version de cada archivo
    versionActualizada = version
    
    Call actualizar_obtenerVersionDeArchivos(archivosRecibidos)
End Sub
Private Sub getUltimasModificaciones_(respuesta As String, error As String)
    Dim doc As New MSXML2.DOMDocument
    Dim success As Boolean
    Dim versiones As Collection
    Dim InfoUsuario As Dictionary
    
    Set versiones = New Collection
    
    If Len(error) > 0 Then
        RaiseEvent actualizaNovedades(Nothing, error)
        Exit Sub
    End If
    
    success = doc.loadXML(respuesta)
   
    If success = False Then
        RaiseEvent actualizaNovedades(Nothing, "Error al interpretar respuesta.")
    Else
        Dim nodeList As MSXML2.IXMLDOMNodeList

        Set nodeList = doc.selectNodes("/log/logentry")

        If Not nodeList Is Nothing Then
            Dim node As MSXML2.IXMLDOMNode
            Dim infoVersion As clsRepositorioVersion
            Dim fecha As String
         
         
            For Each node In nodeList
                ' Cargamos los datos
                Set infoVersion = New clsRepositorioVersion
                
                fecha = Replace(node.selectSingleNode("date").Text, "T", " ")
                fecha = Replace(fecha, "Z", " ")
                fecha = Replace(fecha, "-", "/")
                fecha = mid$(fecha, 1, InStrRev(fecha, ".") - 1)
                            
                infoVersion.numero = CLng(val(node.Attributes.getNamedItem("revision").Text))
                infoVersion.fecha = CDate(fecha)
                
                Set InfoUsuario = JSON.parse(node.selectSingleNode("msg").Text)
                
                If Not InfoUsuario Is Nothing Then
                    infoVersion.usuario = InfoUsuario.item("usuario")
                    infoVersion.comentario = InfoUsuario.item("comentario")
                    
                    Call InfoUsuario.RemoveAll
                    Set InfoUsuario = Nothing
                Else
                    infoVersion.usuario = "?"
                    infoVersion.comentario = node.selectSingleNode("msg").Text
                End If
                ' Los agregamos
                Call versiones.Add(infoVersion)
            Next node
        End If
        
        RaiseEvent actualizaNovedades(versiones, "")
    End If

    
End Sub


Private Sub consolaDeWindows_Complete(stdOut As String, stdErr As String)
    If accionPendiente = "UM" Then
        Call getUltimasModificaciones_(stdOut, stdErr)
    ElseIf accionPendiente = "UP" Then
        Call actualizar_(stdOut, stdErr)
    ElseIf accionPendiente = "CN" Then
        Call compartirNovedades_(stdOut, stdErr)
    ElseIf accionPendiente = "ADD" Then
        Call archivosAgregados_(stdOut, stdErr)
    ElseIf accionPendiente = "IN" Then
        Call obtenidaVersionArchivo_(stdOut, stdErr)
    End If
    
End Sub
