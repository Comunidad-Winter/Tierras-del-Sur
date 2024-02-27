Attribute VB_Name = "ME_Configuracion_Usuario"
Option Explicit

Public OPath As String 'Carpeta de salida de datos generada pro el editor de mapas
Public DatosPath As String 'carpeta donde el map editor toma datos del juego
Public ClientPath As String
Public IniPath As String

Public zonaDefault As String

Private archivoIni As String

Public Sub cargarConfiguracionUsuario(archivoConfiguracion As String)

    If FileExist(archivoConfiguracion) Then
    
        'Guardo el archivo que abri
        archivoIni = archivoConfiguracion
        
        'Paths
        ClientPath = ProccessPath(GetPathIni(archivoConfiguracion, "MAP_EDITOR", "Path"))
        DatosPath = ProccessPath(GetPathIni(archivoConfiguracion, "MAP_EDITOR", "DatosPath"))
        OPath = ProccessPath(GetPathIni(archivoConfiguracion, "MAP_EDITOR", "OutputPath"))
        IniPath = ClientPath & "Init"
    Else
        MsgBox "El archivo de configuracion '" & archivoConfiguracion & " no existe. No se puede iniciar la aplicación sin este archivo", vbCritical, "Tierras del Sur Editor"
        End
    End If
End Sub

Public Function obtenerPreferenciaWorkSpace(ByRef variable As String) As String
     obtenerPreferenciaWorkSpace = GetVar(archivoIni, "PREFERENCIAS", variable)
End Function


Public Sub actualizarPreferenciaWorkSpace(ByRef variable As String, NuevoValor As String)
    'Rutas
    Call WriteVar(archivoIni, "PREFERENCIAS", variable, NuevoValor)
End Sub

Public Sub actualizarVariableConfiguracion(ByRef variable As String, NuevoValor As String)
    Call WriteVar(archivoIni, "MAP_EDITOR", variable, NuevoValor)
End Sub


Private Sub guardarConfiguracion(archivoConfiguracionSalida As String)

    'Generales
    Call WriteVar(archivoConfiguracionSalida, "MAP_EDITOR", "Path", ClientPath)
    Call WriteVar(archivoConfiguracionSalida, "MAP_EDITOR", "DatosPath", DatosPath)
    Call WriteVar(archivoConfiguracionSalida, "MAP_EDITOR", "OutputPath", OPath)
                    
End Sub

Public Sub crearDirectorios()

    If Not FolderExist(OPath & "Mapas") Then
        MkDir OPath & "Mapas"
    End If
    
    If Not FolderExist(OPath & "Mapas\Servidor") Then
        MkDir OPath & "Mapas\Servidor"
    End If
    
    If Not FolderExist(OPath & "Mapas\Cliente") Then
        MkDir OPath & "Mapas\Cliente"
    End If
    
    If Not FolderExist(OPath & "Imagenes") Then
        MkDir OPath & "Imagenes"
    End If
    
End Sub

Public Function DirGraficos() As String
    DirGraficos = ClientPath & "Graficos\"
End Function

Public Function DirSound() As String
    DirSound = ClientPath & "WAV\"
End Function

Public Function DirMidi() As String
    DirMidi = ClientPath
End Function

Public Function DirMapas() As String
    DirMapas = ClientPath
End Function

