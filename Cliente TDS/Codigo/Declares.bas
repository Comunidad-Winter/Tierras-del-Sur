Attribute VB_Name = "Mod_Declaraciones"
'********************Misery_Ezequiel 28/05/05********************'
Option Explicit
'MENSAJES
'********eLwE 15/05/05********
Public Mensaje() As String
Public Const NUMMENSAJES = 400
Public Mapa() As String
Public Const NumMapas = 184
Public Const NumArm As Byte = 10
Public RangoArmada() As String
Public Const NumCaos As Byte = 10
Public RangoCaos() As String
Public Coord As String
'********eLwE 15/05/05********
Public VolumeN As Integer
Public RawServersList As String
Public Cantidadsound As Integer
Public puto As Boolean
Public Type tServerInfo
    Ip As String
    Puerto As Integer
    desc As String
    PassRecPort As Integer
End Type

Public ServersLst() As tServerInfo
Public ServersRecibidos As Boolean

Public CurServer As Integer

Public CreandoClan As Boolean
Public ClanName As String
Public Site As String
Public CAlineacion As Byte


Public UserCiego As Boolean
Public UserEstupido As Boolean

Public NoRes As Boolean 'no cambiar la resolucion

Public Enum BTarget
bCabeza = 1
bPiernaIzquierda = 2
bPiernaDerecha = 3
bBrazoDerecho = 4
bBrazoIzquierdo = 5
bTorso = 6
End Enum

'Timers de GetTickCount
Public Const tAt = 2000
Public Const tUs = 600

Public Const PrimerBodyBarco As Byte = 84
Public Const UltimoBodyBarco As Byte = 87

Public Dialogos As New cDialogos
Public NumEscudosAnims As Integer

Public ArmasHerrero(0 To 100) As Integer
Public ArmadurasHerrero(0 To 100) As Integer
Public ObjCarpintero(0 To 100) As Integer

Public Versiones(1 To 7) As Integer

Public UsaMacro As Boolean
Public CnTd As Byte
Public SecuenciaMacroHechizos As Byte

'[KEVIN]
Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory
'[/KEVIN]

Public Tips() As String * 255
Public Const LoopAdEternum = 999

'[Misery_Ezequiel 10/07/05]
Public Const NUMCIUDADES As Byte = 4
'[\]Misery_Ezequiel 10/07/05]

'Direcciones
Public Enum PCardinales '[Wizard] Enumera los puntos cardinales
NORTH = 1
EAST = 2
SOUTH = 3
WEST = 4
End Enum


'Objetos

Public Const MAX_INVENTORY_OBJS = 10000
Public Const MAX_INVENTORY_SLOTS As Byte = 20
Public Const MAX_NPC_INVENTORY_SLOTS As Byte = 50
Public Const MAXHECHI As Byte = 35

Public Const NUMSKILLS As Byte = 21
Public Const NUMATRIBUTOS As Byte = 5
Public Const NUMCLASES As Byte = 15
Public Const NUMRAZAS As Byte = 5

Public Const MAXSKILLPOINTS As Byte = 100

Public Const FLAGORO = 777

Public Const FOgata = 1521
Public Enum Skl '[Wizard Enumeramos los skills]
Suerte = 1
Magia = 2
Robar = 3
Tacticas = 4
Armas = 5
Meditar = 6
Apuñalar = 7
Ocultarse = 8
Supervivencia = 9
Talar = 10
Comerciar = 11
Defensa = 12
Pesca = 13
Mineria = 14
Carpinteria = 15
Herreria = 16
Liderazgo = 17
Domar = 18
Proyectiles = 19
Wresterling = 20
Navegacion = 21
End Enum

Public Const FundirMetal As Byte = 88

'Inventario
Type Inventory
    OBJIndex As Integer
    Name As String
    GrhIndex As Integer
    '[Alejo]: tipo de datos ahora es Long
    Amount As Long
    '[/Alejo]
    Equipped As Byte
    Valor As Long
    ObjType As Integer
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
End Type

Type NpCinV
    OBJIndex As Integer
    Name As String
    GrhIndex As Integer
    Amount As Integer
    Valor As Long
    ObjType As Integer
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String
End Type

Type tReputacion 'Fama del usuario
    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    Promedio As Long
End Type

Type tEstadisticasUsu
    CiudadanosMatados As Long
    CriminalesMatados As Long
    UsuariosMatados As Long
    NpcsMatados As Long
    Clase As String
    PenaCarcel As Long
End Type

Public ListaRazas() As String
Public ListaClases() As String

Public Nombres As Boolean

Public MixedKey As Long

'User status vars
Public UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory
Global OtroInventario(1 To MAX_INVENTORY_SLOTS) As Inventory

Public UserHechizos(1 To MAXHECHI) As Integer

Public NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpCinV
Public NPCInvDim As Integer
Public UserMeditar As Boolean
Public UserName As String
Public UserPassword As String
Public UserMaxHP As Integer
Public UserMinHP As Integer
Public UserMaxMAN As Integer
Public UserMinMAN As Integer
Public UserMaxSTA As Integer
Public UserMinSTA As Integer
Public UserGLD As Long
Public UserLvl As Integer
Public UserPort As Integer
Public UserServerIP As String
Public UserCanAttack As Integer
Public UserEstado As Byte '0 = Vivo & 1 = Muerto
Public meves As Boolean
Public Cantidadlingo As Integer
Public Lingoteando As Boolean
Public TiempoNuevo As Long  '(??)
Public TiempoViejo As Long
Public UserPasarNivel As Long
Public UserExp As Long
Public UserReputacion As tReputacion
Public UserEstadisticas As tEstadisticasUsu
Public UserDescansar As Boolean
Public tipf As String
Public PrimeraVez As Boolean
Public FPSFLAG As Boolean
Public pausa As Boolean
Public IScombate As Boolean
Public Istrabajando As Boolean
Public PPP As String
Public UserParalizado As Boolean
Public UserNavegando As Boolean
Public UserHogar As String

'Barrin 29/9/03
Public PadrinoName As String
Public PadrinoPassword As String
Public UsandoSistemaPadrinos As Byte
Public PuedeCrearPjs As Integer
'<-------------------------NUEVO-------------------------->
Public Comerciando As Boolean
'<-------------------------NUEVO-------------------------->
Public UserClase As String
Public UserSexo As String
Public UserRaza As String
Public UserEmail As String

Public UserSkills() As Integer
Public SkillsNames() As String

Public UserAtributos() As Integer
Public AtributosNames() As String

Public Ciudades() As String
Public CityDesc() As String

Public Musica As Byte
Public Fx As Byte

Public SkillPoints As Integer
Public Alocados As Integer
Public flags() As Integer
Public Oscuridad As Integer
Public logged As Boolean
Public NoPuedeUsar As Boolean

'Barrin 30/9/03
Public UserPuedeRefrescar As Boolean

Public UsingSkill As Integer

Public MD5HushYo As String * 16

Public Enum E_MODO
    Normal = 1
    BorrarPj = 2
    CrearNuevoPj = 3
    Dados = 4
    RecuperarPass = 5
End Enum
Public EstadoLogin As E_MODO

'Server stuff
Public RequestPosTimer As Integer 'Used in main loop
Public stxtbuffer As String 'Holds temp raw data from server
Public stxtbuffercmsg As String 'Holds temp raw data from server
Public stxtbuffergmsg As String 'Holds temp raw data from server
Public stxtbufferrmsg As String 'Holds temp raw data from server
Public SendNewChar As Boolean 'Used during login
Public Connected As Boolean 'True when connected to server
Public DownloadingMap As Boolean 'Currently downloading a map from server
Public UserMap As Integer
Public Terreno As String
Public Zona As String
Public TiempoReto As Integer
'String contants
Public ENDC As String 'Endline character for talking with server
Public ENDL As String 'Holds the Endline character for textboxes
Public herrero As Boolean
Public armado As Boolean
'Control
Public prgRun As Boolean 'When true the program ends
Public finpres As Boolean

Public IPdelServidor As String
Public PuertoDelServidor As String

'********** FUNCIONES API ***********
Public Declare Function GetTickCount Lib "kernel32" () As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

Public Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type
'********************Misery_Ezequiel 28/05/05********************'
Public Function General_Var_Get(ByVal File As String, ByVal Main As String, ByVal var As String) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Get a var to from a text file
'*****************************************************************
    Dim l As Long
    Dim Char As String
    Dim sSpaces As String 'Input that the program will retrieve
    Dim szReturn As String 'Default value if the string is not found
    
    szReturn = ""
    
    sSpaces = Space$(5000)
    
    getprivateprofilestring Main, var, szReturn, sSpaces, Len(sSpaces), File
    
    General_Var_Get = RTrim$(sSpaces)
    General_Var_Get = left$(General_Var_Get, Len(General_Var_Get) - 1)
End Function
