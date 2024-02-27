Attribute VB_Name = "Configuracion_Usuario"
Option Explicit

Private archivoIni As String

Public ClientPath As String         ' Carpeta donde se encuentra instalado el cliente
Public IniPath As String            ' Carpeta donde se encuentran los archivos de indexacion
Public RecursosPath As String       ' Carpeta donde se encuentarn los graficos
'Opciones
Public lenguaje As String          ' Lenguaje del juego
Public invertiR As Byte             ' ¿Invertir parlantes?
Public oJPG As Byte                 ' ¿Guardar los screems en jpg?
Public CursorPer As Byte            ' Cursores personalizados?
Public Recpassword As Byte          ' Recueprar Password?
Public volumenMusica As Single      ' Volumen de la musica
Public VolumenF As Single           ' Volumen de los efectos
Public LimitarFPS As Byte           ' Limitar FPS
Public Musica As Boolean            ' ¿Escuchar musica?
Public EfectosSonidoActivados As Boolean

Public SonidoFinalizacionDopa As Boolean        ' Se escucha un sonido cuando se le esta por terminar la dopa

' Consola
Public ConsolaTop As Integer
Public ConsolaLeft As Integer
Public ConsolaHeight As Integer
Public ConsolaWidth As Integer

' Teclas
Public vbKeyMusica As Integer
Public vbKeyAgarrarItem As Integer
Public vbKeyTirarItem As Integer
Public vbKeyModoCombate As Integer
Public vbKeyEquiparItem As Integer
Public vbKeyMostrarNombre As Integer
Public vbKeyDomar As Integer
Public vbKeyOcultar As Integer
Public vbKeyUsar As Integer
Public vbKeyLag As Integer
Public vbKeyConsolaClanes As Integer
Public vbKeyNorte As Integer
Public vbKeySur As Integer
Public vbKeyEste As Integer
Public vbKeyOeste As Integer
Public vbKeyPegar As Integer
Public vbKeyMeditar As Integer

Public forzarFullScreen As Boolean  ' Ajusta resolucion del monitor?
Public ResolucionJuego As Integer

Public versionActual As Integer

Public Sub iniciarPaths()
    
    IniPath = app.Path & "\Recursos\" 'TODO: Iniciar Configuracion
    ClientPath = app.Path & "\"
    RecursosPath = app.Path & "\" & "Recursos\"
End Sub
Public Sub cargarConfiguracionUsuario()
    Dim configFile As String
    Dim lector As clsIniReader
    
    Set lector = New clsIniReader
        
    configFile = getConfigFilePath()
    
    ' Iniciamos
    Call lector.Initialize(configFile)
    
    Call AplicarConfiguracion(lector)
    
    Set lector = Nothing
End Sub

Private Function getConfigFilePath() As String
    Dim folder As String
    
    getConfigFilePath = app.Path & "\Configs.ini"
    
    Exit Function
    
    
    folder = getProgramDataFolderTierrasdelSur()
    ' Tratamos de guardar el config en el ProgramData de Windows donde no deberiamos tener problemas de Auth
    If Not folder = vbNullString Then
        getConfigFilePath = folder & "\Configs.ini"
    Else
        getConfigFilePath = app.Path & "\Configs.ini"
    End If
    
    If Not FileExist(getConfigFilePath, vbArchive) Then
        Call setDefaultConfig
        Call guardarConfigIni(getConfigFilePath)
    End If
    
End Function

Private Function getProgramData() As String
    On Error GoTo hayerror:
    
    getProgramData = CreateObject("Shell.Application").NameSpace(&H23).Self.Path
    
    Exit Function
hayerror:
    getProgramData = vbNullString
End Function

Private Function getProgramDataFolderTierrasdelSur() As String

    Dim programData As String
    Dim directorioTDS As String
    
    programData = getProgramData
    
    If programData = vbNullString Then
        getProgramDataFolderTierrasdelSur = vbNullString
        Exit Function
    End If
    
    ' Existe la carpeta commandata?
    directorioTDS = programData & "/Tierras del Sur/"

    If FolderExist(directorioTDS) Then
        getProgramDataFolderTierrasdelSur = directorioTDS
        Exit Function
    End If

    If crearDirectorio(directorioTDS) Then
        getProgramDataFolderTierrasdelSur = directorioTDS
        Exit Function
    End If
        
    getProgramDataFolderTierrasdelSur = vbNullString
End Function

Private Function crearDirectorio(directorio As String) As Boolean

On Error GoTo hayerror:

    Call MkDir(directorio)
    crearDirectorio = True
    
Exit Function

hayerror:
crearDirectorio = False
End Function

Private Sub setDefaultConfig()
    lenguaje = "es"
End Sub

Public Sub AplicarConfiguracion(config As clsIniReader)
    
    'Audio
    EfectosSonidoActivados = config.GetValue("INIT", "EFECTOS") = "1"
    Musica = config.GetValue("INIT", "MUSICA") = "1"
    invertiR = config.GetValue("INIT", "INVERTIR") = "1"
    volumenMusica = CSng(config.GetValue("INIT", "VOLUMENSONIDO"))
    VolumenF = CSng(config.GetValue("INIT", "VOLUMENFX"))
    
    'Video
    LimitarFPS = config.GetValue("INIT", "LIMITAR") = "1"
    
    SombrasHQ = config.GetValue("INIT", "SOMBRAS") = "1"
    cfgSoportaPointSprites = config.GetValue("INIT", "SPRITES") = "1"
          
    CambiarResolucion = False
    NoUsarSombras = False
    NoUsarLuces = False
    NoUsarParticulas = False
    AnimarAguatierra = True
    Optimizar_Textos = True
    UsarVSync = config.GetValueOrDefault("INIT", "VSYNC", "1")
    usaBumpMapping = True

    'Varios
    lenguaje = config.GetValueOrDefault("INIT", "LEN", "es")
    CursorPer = config.GetValue("INIT", "CPER")
    Recpassword = config.GetValue("INIT", "RPASSWORD")
    oJPG = config.GetValue("INIT", "JPG") = "1"
    forzarFullScreen = config.GetValue("INIT", "FORZARFULLSCREEN") = "1"
    
    SonidoFinalizacionDopa = config.GetValueOrDefaultInt("INIT", "SonidoFinalizacionDopa", 1) = 1
    
    ' Consola
    ConsolaTop = config.GetValueOrDefaultInt("INIT", "ConsolaTop", 0)
    ConsolaLeft = config.GetValueOrDefaultInt("INIT", "ConsolaLeft", 0)
    ConsolaHeight = config.GetValueOrDefaultInt("INIT", "ConsolaHeight", 0)
    ConsolaWidth = config.GetValueOrDefaultInt("INIT", "ConsolaWidth", 0)
    
    ' Teclas
    vbKeyMusica = config.GetValueOrDefaultInt("INIT", "vbKeyMusica", vbKeyM)
    vbKeyAgarrarItem = config.GetValueOrDefaultInt("INIT", "vbKeyAgarrarItem", vbKeyA)
    vbKeyModoCombate = config.GetValueOrDefaultInt("INIT", "vbKeyModoCombate", vbKeyC)
    vbKeyEquiparItem = config.GetValueOrDefaultInt("INIT", "vbKeyEquiparItem", vbKeyE)
    vbKeyMostrarNombre = config.GetValueOrDefaultInt("INIT", "vbKeyMostrarNombre", vbKeyN)
    vbKeyDomar = config.GetValueOrDefaultInt("INIT", "vbKeyDomar", vbKeyD)
    vbKeyUsar = config.GetValueOrDefaultInt("INIT", "vbKeyUsar", vbKeyU)
    vbKeyLag = config.GetValueOrDefaultInt("INIT", "vbKeyLag", vbKeyL)
    vbKeyConsolaClanes = config.GetValueOrDefaultInt("INIT", "vbKeyConsolaClanes", vbKeyZ)
    vbKeyNorte = config.GetValueOrDefaultInt("INIT", "vbKeyNorte", vbKeyUp)
    vbKeySur = config.GetValueOrDefaultInt("INIT", "vbKeySur", vbKeyDown)
    vbKeyEste = config.GetValueOrDefaultInt("INIT", "vbKeyEste", vbKeyRight)
    vbKeyOeste = config.GetValueOrDefaultInt("INIT", "vbKeyOeste", vbKeyLeft)
    vbKeyTirarItem = config.GetValueOrDefaultInt("INIT", "vbKeyTirarItem", vbKeyT)
    vbKeyPegar = config.GetValueOrDefaultInt("INIT", "vbKeyPegar", vbKeyControl)
    vbKeyOcultar = config.GetValueOrDefaultInt("INIT", "vbKeyOcultar", vbKeyO)
    vbKeyMeditar = config.GetValueOrDefaultInt("INIT", "vbKeyMeditar", vbKeyF6)
    
    versionActual = config.GetValueOrDefaultInt("INIT", "version", 0)
    
    '
    Dim tempResolucion As Integer
    
    tempResolucion = config.GetValueOrDefaultInt("INIT", "ResolucionJuego", -1)
    
    ' Buscamos la mejor resolucion para el usuario
    If tempResolucion = -1 Then
        If (Screen.width / Screen.Height < 1.4) Then
           ResolucionJuego = RESOLUCION_43
        Else
           ResolucionJuego = RESOLUCION_169
        End If
    Else
        ResolucionJuego = tempResolucion
    End If
    
    
End Sub

Public Sub guardarConfiguracion()

    Dim archivo As String

    archivo = getConfigFilePath()
    
    guardarConfigIni (archivo)
End Sub

Private Sub guardarConfigIni(archivo As String)
    Call WriteVar(archivo, "INIT", "Musica", IIf(Musica, 1, 0))
    Call WriteVar(archivo, "INIT", "Efectos", IIf(EfectosSonidoActivados, 1, 0))
    Call WriteVar(archivo, "INIT", "Invertir", val(invertiR))
    Call WriteVar(archivo, "INIT", "Limitar", val(LimitarFPS))
    Call WriteVar(archivo, "INIT", "Len", lenguaje)
    Call WriteVar(archivo, "INIT", "Cper", val(CursorPer))
    Call WriteVar(archivo, "INIT", "VolumenSonido", volumenMusica)
    Call WriteVar(archivo, "INIT", "VolumenFx", VolumenF)
    Call WriteVar(archivo, "INIT", "Rpassword", val(Recpassword))
    Call WriteVar(archivo, "INIT", "JPG", val(oJPG))
    Call WriteVar(archivo, "INIT", "FORZARFULLSCREEN", IIf(forzarFullScreen, 1, 0))
    Call WriteVar(archivo, "INIT", "VSYNC", IIf(UsarVSync, 1, 0))
    Call WriteVar(archivo, "INIT", "LEN", lenguaje)
    
    Call WriteVar(archivo, "INIT", "SonidoFinalizacionDopa", IIf(SonidoFinalizacionDopa, 1, 0))
    
    ' Consola
    Call WriteVar(archivo, "INIT", "ConsolaTop", ConsolaTop)
    Call WriteVar(archivo, "INIT", "ConsolaLeft", ConsolaLeft)
    Call WriteVar(archivo, "INIT", "ConsolaHeight", ConsolaHeight)
    Call WriteVar(archivo, "INIT", "ConsolaWidth", ConsolaWidth)
    
    ' Teclas
    Call WriteVar(archivo, "INIT", "vbKeyMusica", vbKeyMusica)
    Call WriteVar(archivo, "INIT", "vbKeyAgarrarItem", vbKeyAgarrarItem)
    Call WriteVar(archivo, "INIT", "vbKeyModoCombate", vbKeyModoCombate)
    Call WriteVar(archivo, "INIT", "vbKeyEquiparItem", vbKeyEquiparItem)
    Call WriteVar(archivo, "INIT", "vbKeyMostrarNombre", vbKeyMostrarNombre)
    Call WriteVar(archivo, "INIT", "vbKeyDomar", vbKeyDomar)
    Call WriteVar(archivo, "INIT", "vbKeyUsar", vbKeyUsar)
    Call WriteVar(archivo, "INIT", "vbKeyLag", vbKeyLag)
    Call WriteVar(archivo, "INIT", "vbKeyConsolaClanes", vbKeyConsolaClanes)
    Call WriteVar(archivo, "INIT", "vbKeyNorte", vbKeyNorte)
    Call WriteVar(archivo, "INIT", "vbKeySur", vbKeySur)
    Call WriteVar(archivo, "INIT", "vbKeyEste", vbKeyEste)
    Call WriteVar(archivo, "INIT", "vbKeyOeste", vbKeyOeste)
    Call WriteVar(archivo, "INIT", "vbKeyTirarItem", vbKeyTirarItem)
    Call WriteVar(archivo, "INIT", "vbKeyPegar", vbKeyPegar)
    Call WriteVar(archivo, "INIT", "vbKeyOcultar", vbKeyOcultar)
    Call WriteVar(archivo, "INIT", "vbKeyMeditar", vbKeyMeditar)
    
    ' Pantallas
    Call WriteVar(archivo, "INIT", "ResolucionJuego", ResolucionJuego)
End Sub

