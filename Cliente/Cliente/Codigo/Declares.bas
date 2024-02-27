Attribute VB_Name = "CLI_Declaraciones"
Option Explicit

Public DesabilitarTecla(0 To 255) As Integer 'Anticheat

Type ObjetosComercioSeguro
    Nombre As String
    index As Integer
    cantidad As Long
End Type

Type RecursoConstruccion
    index As Integer
    grhIndex As Integer
    cantidad As Integer
End Type

Type ObjetoConstruible
    index As Integer
    grhIndex As Integer
    recursosNecesarios() As RecursoConstruccion
End Type

Public objeto() As String
Public mensaje() As String
Public NpcsMensajes() As String
Public MensajesCompuestos() As String
Public Msgboxes() As String
Public mapa() As String
Public RangoArmada() As String
Public RangoCaos() As String
Public ArmasHerrero(0 To 100) As Integer
Public ArmadurasHerrero(0 To 100) As Integer
Public ObjCarpintero(0 To 100) As ObjetoConstruible
Public tips(1 To 15) As String * 255
Public ListaRazas() As String
Public ListaClases() As String
Public ListaGeneros() As String
Public SkillsNames() As String
Public AtributosNames() As String
Public Ciudades() As String
Public CityDesc() As String

Public UserPuedeRefrescar As Long

Public CreandoClan As Boolean
Public ClanName As String
Public Site As String
Public CAlineacion As Byte

Public Enum BTarget
    bCabeza = 1
    bPiernaIzquierda = 2
    bPiernaDerecha = 3
    bBrazoDerecho = 4
    bBrazoIzquierdo = 5
    bTorso = 6
End Enum

Public Dialogos As New clsDialogs

'Objetos
Public Const MAX_NPC_INVENTORY_SLOTS As Byte = 50

Type NpCinV
    OBJIndex As Integer
    Name As String
    grhIndex As Integer
    Amount As Integer
    valor As Long
    MinDef As Byte
    MaxDef As Byte
    OBJType As Integer
    MaxHit As Integer
    MinHit As Integer
End Type

Public NoPuedeChuparYuSarClick As Single
Public NoPuedeChuparYuSarU As Single
Public TiempoReto As Single

Public Enum E_Estado
    Ninguno
    Conectando
    conectado
End Enum

Public Enum E_MODO
    Ninguno
    IngresarPersonaje
    PantallaCreacion
    CreandoPersonaje
    CrearPersonajeSeteado
    Jugando
End Enum

Public EstadoLogin As E_MODO
Public EstadoConexion As E_Estado


'Server stuff
'Public RequestPosTimer As Integer 'Used in main loop
Public stxtbuffer As String 'Holds temp raw data from server
Public stxtbuffercmsg As String 'Holds temp raw data from server
Public stxtbuffergmsg As String 'Holds temp raw data from server
Public stxtbufferrmsg As String 'Holds temp raw data from server
Public Connected As Boolean 'True when connected to server

Public DeAmuchos As Boolean
Public MovimientoDefault As E_Heading
Public LastKeyPress As E_Heading
Public LastKeyPressTime As Double
'********** FUNCIONES API ***********
Public Declare Function GetTickCount Lib "kernel32" () As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Listaintegrantes(0 To 20) As String
Public Listasolicitudes(0 To 20) As String

Public LlegaronSkills As Boolean
Public LlegaronAtrib As Boolean
Public LlegoFama As Boolean

Public PacketNumber As Long
' Sistema de FotoDenuncias
Public FotoDenuncia As Byte
Public FotoDenunciasTiempo As Single
Public ultimoDenunciar As String


Public MinPacketNumber As Byte

Public Const STAT_MAXELV = 50


Public IntervaloNoChupClickB As String
Public intervaloNoChupUB As String
Public IntervaloPegarB As String
Public IntervaloLanzarMagiasB As String
Public IntervaloLanzarFlechasB As String

Public IntervaloSolapaLanzarB As String
Public IntervaloSolapaLanzarSuperB As String
Public IntervaloHechizoLanzarB As String
Public IntervaloHechizoLanzarSuperB As String
Public UmbralAlertaB As String


Public meves As Boolean

Public UserPassword As String
Public UserSexo As String
Public UserClase As Byte
Public UserRaza As String
Public UserHogar As String
Public UserPin As String
Public UserEmail As String
Public UserMac As String


Public UserRazaDesc(1 To 5) As String
