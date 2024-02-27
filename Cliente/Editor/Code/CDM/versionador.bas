Attribute VB_Name = "versionador"
Option Explicit

Private Type tElementoVersionado
    nombre As String
    archivo As String
    Tipo As eTipoElemento
    fuente As Byte
End Type

Private Enum eTipoElementoFuente
    editor = 1
    server = 2
    cliente = 3
End Enum

Private Enum eTipoElemento
    ini = 1
    pack = 2
End Enum

Public Type tArchivoAlterado
    Tipo As String
    creados As New Collection
    modificados As New Collection
    eliminados As New Collection
    info As Object
End Type

' Lista de archivos versionados
Private elementosVersionados(1 To 24) As tElementoVersionado

Private Const CARPETA_CDM As String = "\CDM\Editor"
Private Const CARPETA_CDM_INI As String = "\CDM\Editor\Ini"
Private Const CARPETA_CDM_PACK As String = "\CDM\Editor\Pack"
Private Const CARPETA_CDM_TEMP As String = "\CDM\Temp"


Public Sub iniciar_versionador()
    
    'Lista de archivos versionados
    
    With elementosVersionados(1)
        .nombre = "PREDEFINIDO"
        .archivo = DBPath & "presets.ini"
        .Tipo = eTipoElemento.ini
        .fuente = eTipoElementoFuente.editor
    End With
    
    With elementosVersionados(2)
        .nombre = "ARMA"
        .archivo = DBPath & "armas.ini"
        .Tipo = eTipoElemento.ini
        .fuente = eTipoElementoFuente.editor
    End With
    
    With elementosVersionados(3)
        .nombre = "CABEZA"
        .archivo = DBPath & "cabezas.ini"
        .Tipo = eTipoElemento.ini
        .fuente = eTipoElementoFuente.editor
    End With
    
    With elementosVersionados(4)
        .nombre = "CASCO"
        .archivo = DBPath & "cascos.ini"
        .Tipo = eTipoElemento.ini
        .fuente = eTipoElementoFuente.editor
    End With
    
        With elementosVersionados(5)
        .nombre = "CUERPO"
        .archivo = DBPath & "cuerpos.ini"
        .Tipo = eTipoElemento.ini
        .fuente = eTipoElementoFuente.editor
    End With
    
    With elementosVersionados(6)
        .nombre = "EFECTO"
        .archivo = DBPath & "efectos.ini"
        .Tipo = eTipoElemento.ini
        .fuente = eTipoElementoFuente.editor
    End With
    
    With elementosVersionados(7)
        .nombre = "SONIDO"
        .archivo = DBPath & "sonidos.ini"
        .Tipo = eTipoElemento.ini
        .fuente = eTipoElementoFuente.editor
    End With
    
    With elementosVersionados(8)
        .nombre = "ENTIDAD"
        .archivo = DBPath & "entidades.ini"
        .Tipo = eTipoElemento.ini
        .fuente = eTipoElementoFuente.editor
    End With
    
    With elementosVersionados(9)
        .nombre = "ESCUDO"
        .archivo = DBPath & "escudos.ini"
        .Tipo = eTipoElemento.ini
        .fuente = eTipoElementoFuente.editor
    End With
        
    With elementosVersionados(10)
        .nombre = "GRAFICO"
        .archivo = DBPath & "graficos.ini"
        .Tipo = eTipoElemento.ini
        .fuente = eTipoElementoFuente.editor
    End With
    
    With elementosVersionados(11)
        .nombre = "PISO"
        .archivo = DBPath & "pisos.ini"
        .Tipo = eTipoElemento.ini
        .fuente = eTipoElementoFuente.editor
    End With
    
    With elementosVersionados(12)
        .nombre = "HECHIZO"
        .archivo = DBPath & "hechizos.dat"
        .Tipo = eTipoElemento.ini
        .fuente = eTipoElementoFuente.editor + eTipoElementoFuente.server
    End With
    
    With elementosVersionados(13)
        .nombre = "PROPIEDAD_MAPA"
        .archivo = DBPath & "mapas.dat"
        .Tipo = eTipoElemento.ini
        .fuente = eTipoElementoFuente.editor + eTipoElementoFuente.server
    End With
    
    With elementosVersionados(14)
        .nombre = "CRIATURA"
        .archivo = DBPath & "npcs.dat"
        .Tipo = eTipoElemento.ini
        .fuente = eTipoElementoFuente.editor + eTipoElementoFuente.server
    End With
    
    With elementosVersionados(15)
        .nombre = "OBJETO"
        .archivo = DBPath & "objetos.dat"
        .Tipo = eTipoElemento.ini
        .fuente = eTipoElementoFuente.editor + eTipoElementoFuente.server
    End With
    
    With elementosVersionados(16)
        .nombre = "RECURSO_IMAGEN"
        .archivo = DBPath & "Graficos.TDS"
        .Tipo = eTipoElemento.pack
        .fuente = eTipoElementoFuente.editor + eTipoElementoFuente.server
    End With

    With elementosVersionados(17)
        .nombre = "RECURSO_INTERFACE"
        .archivo = DBPath & "Interface.TDS"
        .Tipo = eTipoElemento.pack
        .fuente = eTipoElementoFuente.editor + eTipoElementoFuente.server
    End With
    
    With elementosVersionados(18)
        .nombre = "RECURSO_MAPA"
        .archivo = DBPath & "MapasME.TDS"
        .Tipo = eTipoElemento.pack
        .fuente = eTipoElementoFuente.editor + eTipoElementoFuente.server
    End With
    
    With elementosVersionados(19)
        .nombre = "RECURSO_SONIDO"
        .archivo = DBPath & "Sonidos.TDS"
        .Tipo = eTipoElemento.pack
        .fuente = eTipoElementoFuente.editor + eTipoElementoFuente.server
    End With
    
    With elementosVersionados(20)
        .nombre = "NIVEL"
        .archivo = DBPath & "niveles.dat"
        .Tipo = eTipoElemento.ini
        .fuente = eTipoElementoFuente.editor + eTipoElementoFuente.server
    End With
    
    With elementosVersionados(21)
        .nombre = "ARMADA_RANGO"
        .archivo = DBPath & "armada.dat"
        .Tipo = eTipoElemento.ini
        .fuente = eTipoElementoFuente.editor + eTipoElementoFuente.server
    End With
    
    With elementosVersionados(22)
        .nombre = "LEGION_RANGO"
        .archivo = DBPath & "legion.dat"
        .Tipo = eTipoElemento.ini
        .fuente = eTipoElementoFuente.editor + eTipoElementoFuente.server
    End With
    
    With elementosVersionados(23)
        .nombre = "ASPECTO"
        .archivo = DBPath & "pixels.dat"
        .Tipo = eTipoElemento.ini
        .fuente = eTipoElementoFuente.editor + eTipoElementoFuente.cliente
    End With
    
    With elementosVersionados(24)
        .nombre = "PISADA"
        .archivo = DBPath & "pisadas.ini"
        .Tipo = eTipoElemento.ini
        .fuente = eTipoElementoFuente.editor + eTipoElementoFuente.cliente
    End With
End Sub

' A partir del nombre del elemento versionado, obtiene el Slot donde se encuentra
Public Function obtenerTipo(ByVal nombre As String) As Integer
    Dim loopElemento As Integer
    
    For loopElemento = LBound(elementosVersionados) To UBound(elementosVersionados)
        
        If elementosVersionados(loopElemento).nombre = nombre Then
            obtenerTipo = loopElemento
            Exit Function
        End If

    Next

    obtenerTipo = -1
End Function

' Devuelvo el JSON correspondiente que tiene la informacion de los elementos modificados para el nombre
' de tipo indicado
Public Function obtenerInfoTipo(ByVal nombreTipo As String) As Collection
    Dim tipoElemento As Integer
    Dim strData As String
    Dim informacion As Variant
    Dim handle As Integer
    Dim archivo As String
    
    tipoElemento = obtenerTipo(nombreTipo)

    If tipoElemento > -1 Then
        'Abrimos el archivo
        archivo = elementosVersionados(tipoElemento).archivo & ".cdm"
        If Len(Dir$(archivo, vbArchive)) > 0 Then
            handle = FreeFile
            Open archivo For Input As #handle
            strData = Input$(LOF(handle), handle)
            Close #handle
        Else
            strData = ""
        End If
        
        Set informacion = JSON.parse(strData)
        
        If TypeName(informacion) = "Collection" Then
            Set obtenerInfoTipo = informacion
        Else
            Set obtenerInfoTipo = New Collection
        End If
        
    End If
End Function

' Almacena para el nombre del tipo indicado todas las modificaciones que se le hicieron
' a ese tipo de elemento versionado
Public Function guardarInfoTipo(ByVal Tipo As String, info As Collection)
    Dim archivo As String
    Dim handle As String
    Dim elemento As Integer
    handle = FreeFile
    
    elemento = obtenerTipo(Tipo)
    
    archivo = elementosVersionados(elemento).archivo & ".cdm"
    
    'Guardamos
    Open archivo For Output As #handle
        Print #handle, JSON.toString(info)
    Close #handle
End Function

Private Function obtenerPrioridadAccion(accion As String) As Byte
    If accion = "CREADO" Then
        obtenerPrioridadAccion = 6
    ElseIf accion = "ELIMINADO" Then
        obtenerPrioridadAccion = 10
    ElseIf accion = "MODIFICADO" Then
        obtenerPrioridadAccion = 3
    End If
End Function

'******************************************************************************
' MARCAS

' Se marca que el elemento identificado a través de @identificador
' del nombre del tipo indicado en @tipo fue MODIFICADO
Public Sub modificado(ByVal Tipo As String, ByVal identificador As Long, Optional ByVal nombre As String = "")
    
    Dim lista As Collection
    Dim nuevo As New Dictionary
    Set lista = obtenerInfoTipo(Tipo)
    
    If deboAgregar("MODIFICADO", identificador, lista) Then
       Call agregar("MODIFICADO", identificador, Tipo, lista, nombre)
    ElseIf recienCreado(identificador, lista) Then
        Call agregar("CREADO", identificador, Tipo, lista, nombre)
    End If
End Sub

' Se marca que el elemento identificado a través de @identificador
' del nombre del tipo indicado en @tipo fue CREADO
Public Sub creado(ByVal Tipo As String, ByVal identificador As Long, Optional ByVal nombre As String = "")
    
    Dim lista As Collection
    
    Set lista = obtenerInfoTipo(Tipo)
    
    If obtenerDato(identificador, lista) Is Nothing Then
        Call agregar("CREADO", identificador, Tipo, lista, nombre)
    End If
    
End Sub

' Se marca que el elemento identificado a través de @identificador
' del nombre del tipo indicado en @tipo fue eliminado
Public Sub eliminado(ByVal Tipo As String, ByVal identificador As Long, Optional ByVal nombre As String = "")
    
    Dim lista As Collection
    
    Dim nuevo As New Dictionary
    Set lista = obtenerInfoTipo(Tipo)
    
    If deboAgregar("ELIMINADO", identificador, lista) Then
        Call agregar("ELIMINADO", identificador, Tipo, lista, nombre)
    End If
End Sub

'******************************************************************************
'******************************************************************************

' Agrega una accion @accion que se le aplica al elemento identificado con @identificador, para determinado tipo de archivo versionado
' @lista contiene todas las acciones para ese tipo hasta el momento
Private Sub agregar(ByVal accion As String, ByVal identificador As Long, ByVal Tipo As String, lista As Collection, Optional ByVal nombre As String = "")
    Dim elemento As Dictionary

    Set elemento = obtenerDato(identificador, lista)
    
    If elemento Is Nothing Then
        Set elemento = New Dictionary
        
        Call elemento.Add("id", identificador)
        Call elemento.Add("accion", accion)
        Call elemento.Add("nombre", nombre)
        
        Call lista.Add(elemento)
    Else
        elemento.item("accion") = accion
        elemento.item("nombre") = nombre
    End If
    
    Call guardarInfoTipo(Tipo, lista)
End Sub

Private Function deboAgregar(ByVal accion As String, ByVal identificador As String, lista As Collection) As Boolean
    Dim elemento As Dictionary
    Dim agregar As Boolean
    Set elemento = obtenerDato(identificador, lista)
    
    agregar = False
    
    If elemento Is Nothing Then
       agregar = True
    ElseIf obtenerPrioridadAccion(accion) > obtenerPrioridadAccion(elemento.item("accion")) Then
        agregar = True
    End If
    
    deboAgregar = agregar
End Function

Private Function recienCreado(ByVal identificador As String, lista As Collection) As Boolean
    Dim elemento As Dictionary
    Set elemento = obtenerDato(identificador, lista)
    
    recienCreado = False
    
    If elemento Is Nothing Then
       recienCreado = False
    ElseIf elemento.item("accion") = "CREADO" Then
        recienCreado = True
    End If

End Function

' Devuelve informacion de todos los elementos versionados que han sido alterados
' indicando la cantidad de elementos creados, modificados y eliminados
' En @cantidad se indica la cantidad de elementos modificados en total
Public Sub obtenerArchivosAlterados(cantidad As Integer, archivos() As tArchivoAlterado)
    Dim loopElemento As Integer
    Dim infoTipo As Collection
    Dim total As Integer
    
    total = 0
       
    ' Recorremos cada una de las cosas versionables buscando las modificaciones
    For loopElemento = LBound(elementosVersionados) To UBound(elementosVersionados)
    
        Set infoTipo = obtenerInfoTipo(elementosVersionados(loopElemento).nombre)
        
        If Not infoTipo Is Nothing Then
        
            If infoTipo.count > 0 Then
                total = total + 1
                
                ReDim Preserve archivos(1 To total)
                
                archivos(total).Tipo = elementosVersionados(loopElemento).nombre
                Set archivos(total).info = infoTipo
                ' Tengo la info duplicada...
                Call modColeccion.buscarDonde("accion", "CREADO", infoTipo, vbNullString, archivos(total).creados)
                Call modColeccion.buscarDonde("accion", "MODIFICADO", infoTipo, vbNullString, archivos(total).modificados)
                Call modColeccion.buscarDonde("accion", "ELIMINADO", infoTipo, vbNullString, archivos(total).eliminados)
            End If
        End If
        
    
    Next loopElemento
    
    cantidad = total
End Sub

' Elimina archivos temporales generados por el versionador previo al commit respectivo
Public Function limpiar(commit As CDM_Commit, archivosCDM As Boolean)
     Dim archivos As Collection
     Dim tipoContenido As Variant
     Dim tipoElemento As Integer
     
     ' Eliminamos el archivo .cdm y el archivo temporal
     Set archivos = commit.obtenerArchivos
     
     For Each tipoContenido In archivos
        'Obtenemos info del tipo de archivo
        If archivosCDM Then
            tipoElemento = obtenerTipo(tipoContenido.destino)
            If Not tipoElemento = -1 Then Kill elementosVersionados(tipoElemento).archivo & ".cdm"
        End If
        ' Elimino el archivo temporal creado
        If Not tipoElemento = -1 Then Kill tipoContenido.archivo
     Next
End Function

' Elimina archivos temporales generados por el versionador previo al commit respectivo
Public Function elimiarTemporales(archivos As Collection)
    Dim archivo As CDM_Archivo
     
     ' Elimino el archivo temporal creado
     For Each archivo In archivos
        Kill archivo.archivo
     Next
End Function

Private Function getEmpaquetado(nombre As String) As clsEnpaquetado
    If nombre = "RECURSO_IMAGEN" Then
        Set getEmpaquetado = pakGraficos
    ElseIf nombre = "RECURSO_SONIDO" Then
        Set getEmpaquetado = pakSonidos
    ElseIf nombre = "RECURSO_MAPA" Then
        Set getEmpaquetado = pakMapasME
    ElseIf nombre = "RECURSO_INTERFACE" Then
        Set getEmpaquetado = pakGUI
    End If
End Function

Private Function generarIniPack(infoArchivo As tArchivoAlterado) As String
    Dim archivoDestino As String
    Dim archivoIni As cIniManager
    Dim elementoVersionado As Integer
    Dim numero As Variant
    
    elementoVersionado = obtenerTipo(infoArchivo.Tipo)
    
    ' Obtengo un archivo que ya no exista
    archivoDestino = HelperFiles.generarRandomNameFile(app.Path & CARPETA_CDM_TEMP, 20, "pini")
    
    Set archivoIni = New cIniManager
    Call archivoIni.Initialize(elementosVersionados(elementoVersionado).archivo)
    
    ' Extraigo los archivos
    For Each numero In infoArchivo.creados
        Call archivoIni.seccionAArchivo(CLng(numero.item("id")), archivoDestino)
    Next
    
    For Each numero In infoArchivo.modificados
         Call archivoIni.seccionAArchivo(CLng(numero.item("id")), archivoDestino)
    Next
    
    For Each numero In infoArchivo.eliminados
         Call archivoIni.seccionAArchivo(CLng(numero.item("id")), archivoDestino)
    Next
    
    generarIniPack = archivoDestino
End Function
Private Function generarPack(infoArchivo As tArchivoAlterado) As String
    Dim pack As clsEnpaquetado ' Empaquetado donde voy a obtener los datos
    Dim nuevoPack As clsEnpaquetado 'Empaquetado nuevo
    Dim nombreNuevoEmpaquetado As String 'Archivo que representa al empaquetado
    
    Dim numero As Variant 'Auxiliar para recorrer en foreach
    
    Dim mayorNumero As Integer 'Mayor número de elemento
    Dim INFOHEADER As INFOHEADER 'Informacion de un elemento
    Dim Data() As Byte 'Data correspondiente a un elemento
    
    ' Obtengo el empaquetado relacionado
    Set pack = getEmpaquetado(infoArchivo.Tipo)
            
    mayorNumero = 0
    
    ' Extraigo los archivos
    For Each numero In infoArchivo.creados
        If numero.item("id") > mayorNumero Then mayorNumero = numero.item("id")
    Next
    
    For Each numero In infoArchivo.modificados
        If numero.item("id") > mayorNumero Then mayorNumero = numero.item("id")
    Next
    
    ' Generamos el nombre del pack. [a-z] * 25
    nombreNuevoEmpaquetado = HelperFiles.generarRandomNameFile(app.Path & CARPETA_CDM_TEMP, 20, "pack")
    
    ' Creamos un empaquetado vacio
    Set nuevoPack = New clsEnpaquetado
    
    Call nuevoPack.CrearVacio(nombreNuevoEmpaquetado, mayorNumero)
    
    ' Vamos a completar este empaquetado con cada archivo
    For Each numero In infoArchivo.creados
        'Obtenemos el cabezal
        Call pack.IH_Get(CInt(numero.item("id")), INFOHEADER)
        ' Obtenemos los datos
        Call pack.LeerIH(Data, INFOHEADER)
        ' Parcheamos
        Call nuevoPack.ParchearByteArray(Data, numero.item("id"), INFOHEADER)
    Next
    
    For Each numero In infoArchivo.modificados
        'Obtenemos el cabezal
        Call pack.IH_Get(CInt(numero.item("id")), INFOHEADER)
        ' Obtenemos los datos
        Call pack.LeerIH(Data, INFOHEADER)
        ' Parcheamos
        Call nuevoPack.ParchearByteArray(Data, numero.item("id"), INFOHEADER)
    Next
          
    ' Retornamos donde se encuentra el pack generado
    generarPack = nombreNuevoEmpaquetado
End Function

Private Function obtenerTipoPorNombreArchivo(nombreArchivo As String) As Integer
    Dim loopElemento As Integer

    For loopElemento = LBound(elementosVersionados) To UBound(elementosVersionados)
        
        If UCase$(getNameFileInPath(elementosVersionados(loopElemento).archivo)) = UCase$(nombreArchivo) Then
            obtenerTipoPorNombreArchivo = loopElemento
            Exit Function
        End If

    Next loopElemento
    
    obtenerTipoPorNombreArchivo = -1
End Function

Private Sub OrdenarArchivos(archivos As Collection)

    Dim vItm As CDM_Archivo
    Dim i As Long, j As Long
    Dim vTemp As Object

    'Two loops to bubble sort
   For i = 1 To archivos.count - 1
        For j = i + 1 To archivos.count
            If archivos(i).Version > archivos(j).Version Then
                'store the lesser item
               Set vTemp = archivos(j)
               ' remove the lesser item
               archivos.Remove j
               ' re-add the lesser item before the
              ' greater item
               archivos.Add vTemp, , i
            End If
        Next j
    Next i

'test it
   For Each vItm In archivos
        Debug.Print vItm.Version
    Next vItm

End Sub

Public Function Parchear(archivos As Collection) As Long

    Dim archivo As CDM_Archivo
    Dim tipoElemento As Integer
    Dim empaquetado As clsEnpaquetado
    Dim carpetaDestino As String
    
    Call OrdenarArchivos(archivos)
       
    ' El ultimo es el mas grande
    Parchear = archivos.item(archivos.count).Version
       
    For Each archivo In archivos

        tipoElemento = obtenerTipo(archivo.destino)
        
        If tipoElemento > 0 Then
                        
            If elementosVersionados(tipoElemento).Tipo = eTipoElemento.ini Then
                Dim iniOriginal As cIniManager
                Dim iniParche As cIniManager
                
                ' Inicio
                Set iniOriginal = New cIniManager
                Set iniParche = New cIniManager
                
                Call iniOriginal.Initialize(elementosVersionados(tipoElemento).archivo)
                Call iniParche.Initialize(archivo.archivo)
                
                ' Copio de parche al original
                Call iniOriginal.copiar(iniParche)
                
                ' Guardo
                Call iniOriginal.DumpFile(elementosVersionados(tipoElemento).archivo)
                
                Set iniOriginal = Nothing
                Set iniParche = Nothing
            ElseIf elementosVersionados(tipoElemento).Tipo = eTipoElemento.pack Then
                Set empaquetado = getEmpaquetado(archivo.destino)

                Call empaquetado.Parchear(0, CStr(archivo.archivo))
            End If
        Else
            carpetaDestino = ProccessPath(HelperFiles.getPathFileInPath(archivo.destino))
            Debug.Print archivo.archivo; " -> "; carpetaDestino & HelperFiles.getNameFileInPath(archivo.destino)
            Call FileCopy(archivo.archivo, ProccessPath(HelperFiles.getPathFileInPath(archivo.destino)) & HelperFiles.getNameFileInPath(archivo.destino))
        End If
        
    Next

End Function

Public Function generarCommit(tiposArchivos() As String, carpetaRoot As String) As CDM_Commit
    
    'En tipoArchivos() están los tipos de archivos que me interesa compartir
    'En archivos() están todos los archivos alterados
    'En obtenerArchivosAlterados la lista de archivos que se van a commitear
    
    Dim tipoElemento As Integer
    Dim archivos() As tArchivoAlterado
    Dim archivoACommitear As String
    
    Dim numeroArchivoModificado As Byte
    Dim total As Integer
    Dim loopElemento As Integer

    ' Obtenemos la info de los archivos modificados
    Call obtenerArchivosAlterados(total, archivos)
    
    ' Creamos el commit
    Dim commit As CDM_Commit
    Dim archivo As CDM_Archivo
    
    Set commit = New CDM_Commit
       
    ' Commiteo las cosas que quiere commitear
    For loopElemento = LBound(tiposArchivos) To UBound(tiposArchivos)
            
        'Obtenemos info del tipo de archivo
        tipoElemento = obtenerTipo(tiposArchivos(loopElemento))
        
        If Not tipoElemento = -1 Then
        
            'Busco la posicion de este elemento en la lista de archivos
            For numeroArchivoModificado = LBound(archivos) To UBound(archivos)
                If archivos(numeroArchivoModificado).Tipo = elementosVersionados(tipoElemento).nombre Then
                    Exit For
                End If
            Next
            
            If elementosVersionados(tipoElemento).Tipo = eTipoElemento.ini Then
                ' Generamos el .pini
                archivoACommitear = generarIniPack(archivos(numeroArchivoModificado))
            Else
                ' Genero un .pak especial con los nuevos archivos
                archivoACommitear = generarPack(archivos(numeroArchivoModificado))
            End If
        
            ' Creo la instancia del archivo a commitea
            Set archivo = New CDM_Archivo
            Call archivo.iniciar(archivoACommitear, tiposArchivos(loopElemento), archivos(numeroArchivoModificado).info)
            
        Else
        
            If tiposArchivos(loopElemento) = "" Or Not FileExist(tiposArchivos(loopElemento)) Then
                MsgBox "Se produjo un error. Por algún motivo queres commitear un archivo ('" & tiposArchivos(loopElemento) & "') en blanco o que no existe. Fijate bien que seleccionaste y preguntale al Administrador.", vbExclamation
                Set archivo = Nothing
            Else
                ' Creo la instancia del archivo a commitea
                Set archivo = New CDM_Archivo
                Dim Temp As New Dictionary
                Call Temp.Add("version", HelperFiles.GetFileVersion(tiposArchivos(loopElemento)))
                Call archivo.iniciar(tiposArchivos(loopElemento), HelperFiles.obtenerPathRelativo(carpetaRoot, tiposArchivos(loopElemento)), Temp)
                Set Temp = Nothing
            End If
            
        End If
        
        ' Lo agregamos al commit
        If Not archivo Is Nothing Then Call commit.agregarArchivo(archivo)
            
    Next loopElemento
        
    Set generarCommit = commit
End Function

