Attribute VB_Name = "Declaraciones"
Option Explicit

Public TrashCollector As New Collection

Public Const MAXSPAWNATTEMPS = 60
Public Const MAXUSERMATADOS = 9000000
Public Const LoopAdEternum = 999

Public Const LimiteNewbie = 15

Public Type tCabecera 'Cabecera de los con
    desc As String * 255
    Crc As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

Public Const FXWARP = 1

Public Const FXCURAR = 2

Public Type tRango
    minimo As Integer
    maximo As Integer
End Type

' Tiempo Carcel
Public Const TIEMPO_CARCEL_PIQUETE = 10

' <<<<<< Targets >>>>>>
Public Const uUsuarios = 1
Public Const uNPC = 2
Public Const uUsuariosYnpc = 3
Public Const uTerreno = 4

' <<<<<< Acciona sobre >>>>>>
Public Const uPropiedades = 1
Public Const uEstado = 2
Public Const uInvocacion = 4

Public Const DRAGON = 6
Public Const MATADRAGONES = 1

Public Const MAXUSERHECHIZOS = 35

Public Const EsfuerzoPescarPescador = 1
Public Const EsfuerzoPescarGeneral = 3

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

Public Const Guardias = 6

Public Const MAXREP = 50000000
Public Const MAXORO = 2000000000
Public Const MAXEXP = 9999999999#

Public Const MAXATRIBUTOS = 40
Public Const MINATRIBUTOS = 6

Public Const MAXNPCS = 10000
Public Const MAXCHARS = 10000

Public Const FX_TELEPORT_INDEX = 1

Public Const MIN_APUÑALAR = 10

'********** CONSTANTANTES ***********
Public Const NUMSKILLS = 21
Public Const NUMATRIBUTOS = 5
Public Const NUMRAZAS = 5

Public Const MAXSKILLPOINTS = 100
Public Const FLAGORO = 777

Public Enum eHeading
    Ninguno = 0
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

Public Const MAXMASCOTAS = 3


'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const vlASALTO = 100
Public Const vlASESINO = 1000
Public Const vlCAZADOR = 5
Public Const vlNoble = 5
Public Const vlLadron = 25
Public Const vlProleta = 2

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const iCuerpoMuerto = 8
Public Const iCabezaMuerto = 500
Public Const iCuerpoMuertoCrimi = 145
Public Const iCabezaMuertoCrimi = 501

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
'Skills
Public Enum eSkills
    ResistenciaMagica = 1
    Magia = 2
    Robar = 3
    tacticas = 4
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
    proyectiles = 19
    Wresterling = 20
    Navegacion = 21
End Enum

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    constitucion = 5
End Enum

Public Const AdicionalHPGuerrero = 2 'HP adicionales cuando sube de nivel
Public Const AdicionalHPCazador = 1
Public Const AdicionalSTLadron = 3
Public Const AdicionalSTLeñador = 23
Public Const AdicionalSTPescador = 20
Public Const AdicionalSTMinero = 25

'Sonidos
Public Const SOUND_BUMP = 1
Public Const SOUND_SWING = 2
Public Const SOUND_TALAR = 13
Public Const SOUND_PESCAR = 14
Public Const SOUND_MINERO = 15
Public Const SND_WARP = 3
Public Const SND_PUERTA = 5
Public Const SOUND_NIVEL = 6
Public Const SOUND_COMIDA = 7
Public Const SND_USERMUERTE = 11
Public Const SND_IMPACTO = 10
Public Const SND_IMPACTO2 = 12
Public Const SND_LEÑADOR = 13
Public Const SND_FOGATA = 14
Public Const SND_AVE = 21
Public Const SND_AVE2 = 22
Public Const SND_AVE3 = 34
Public Const SND_GRILLO = 28
Public Const SND_GRILLO2 = 29
Public Const SOUND_SACARARMA = 25
Public Const SOUND_SACARESPADA = 114
Public Const SND_ESCUDO = 37
Public Const MARTILLOHERRERO = 41
Public Const LABUROCARPINTERO = 42
Public Const SND_CREACIONCLAN = 44
Public Const SND_ACEPTADOCLAN = 43
Public Const SND_DECLAREWAR = 45
Public Const SND_BEBER = 46

'Objetos
Public Const MAX_INVENTORY_OBJS = 10000
Public Const MAX_INVENTORY_SLOTS = 20 ' Limite de inventario para criaturas
Public Const MAX_DROP = 5




'**************************************************************
'**************************************************************
'************************ TIPOS *******************************
'**************************************************************
'**************************************************************
Type tHechizo
    ClaseProhibida(1 To NUMCLASES) As eClases 'Clases que no pueden utilizarlo
    NeedStaff As Integer
    StaffAffected As Boolean
    nombre As String
    desc As String
    PalabrasMagicas As String 'Mensaje que aparece por la pantalla del juego
    HechizeroMsg As String 'Mensjae que se le envia al usuario que tiro el hechizo a consola
    TargetMsg As String 'Mensaje que se le envia al que recibio el hechizo
    PropioMsg As String 'Cuando el hechizo se lo tira a asi mismo
    tipo As Byte
    WAV As Integer 'Sonido al lanzarlo
    FXgrh As Integer
    loops As Byte
    SubeHP As Byte 'SEVA
    minHP As Integer 'Unificado en MIN-MODIFICADOR
    MaxHP As Integer 'Unificado en MAX-MODIFICADOR
    SubeHam As Byte 'SEVA
    minham As Integer 'Unificado en MIN-MODIFICADOR
    MaxHam As Integer 'Unificado en MAX-MODIFICADOR
    SubeSed As Byte 'SEVA
    MinSed As Integer 'Unificado en MIN-MODIFICADOR
    MaxSed As Integer 'Unificado en MAX-MODIFICADOR
    SubeAgilidad As Byte 'SEVA
    MinAgilidad As Integer 'SEVA. Unificado en MIN-MODIFICADOR
    MaxAgilidad As Integer 'SEVA. Unificado en MAX-MODIFICADOR
    SubeFuerza As Byte 'SEVA
    MinFuerza As Integer 'SEVA. Unificado en MIN-MODIFICADOR
    MaxFuerza As Integer 'SEVA. Unicicado en MAX-MODIFICADOR
    Invisibilidad As Byte 'SEVA
    Paraliza As Byte 'SEVA
    Inmoviliza As Byte 'SEVA
    RemoverParalisis As Byte 'SEVA
    CuraVeneno As Byte 'SEVA
    Envenena As Byte 'SEVA
    Revivir As Byte 'SEVA
    Mimetiza As Byte 'SEVA
    AgiUpAndFuer As Byte 'SEVA
    MinAgiFuer As Byte 'SEVA. Unificado en MIN-MODIFICADOR
    MaxAgiFuer As Byte 'SEVA. Unificado en MAX-MODIFICADOR
    RemueveInvisibilidadParcial As Byte 'SEVA
    NumNpc As Integer 'Numero del npc que invoca
    cant As Integer 'Cantidad de npcs de "NumNpcs" que invoca
    MinSkill As Integer 'Cantidad de skills necesarios para lanzar el hechizo. TODO Integer?
    ManaRequerido As Integer 'Mana requerido para lanzar
    ManaRequeridoPaladin As Integer 'Mana requerido para lanzar
    ManaRequeridoAsesino As Integer 'Mana requerido para lanzar
    ManaRequeridoBardo As Integer 'Mana requerido para lanzar
    StaRequerido As Integer 'Energia requerida para lanzar
    Target As Byte 'A quien ataca este hechizo (Usuarios, Npcs, Ambos, Terreno)
    manaPenalidad As Integer ' Cantidad de mana que pierde si erra el hechizo
    id As Integer
End Type

Type UserOBJ
    ObjIndex As Integer
    Amount As Integer
    Equipped As Byte
End Type


Type NPCObjectDrop
    ObjIndex As Integer
    Amount As Integer
    Probability As Integer
End Type

Type inventario
    Object(1 To 30) As UserOBJ
    ObjectDrop(1 To 5) As NPCObjectDrop
    WeaponEqpObjIndex As Integer
    WeaponEqpSlot As Byte
    ArmourEqpObjIndex As Integer
    ArmourEqpSlot As Byte
    EscudoEqpObjIndex As Integer
    EscudoEqpSlot As Byte
    CascoEqpObjIndex As Integer
    CascoEqpSlot As Byte
    MunicionEqpObjIndex As Integer
    MunicionEqpSlot As Byte
    HerramientaEqpObjIndex As Integer
    HerramientaEqpSlot As Integer
    AnilloEqpObjIndex As Integer
    AnilloEqpSlot As Byte
    BarcoObjIndex As Integer
    BarcoSlot As Byte
    BarcoEqpSlot As Byte
    NroItems As Integer
    NroItemsDrop As Integer
    CollarObjIndex As Integer
    BrasaleteEqpObjIndex As Integer
End Type

Type Position
    x As Integer
    y As Integer
End Type

Type WorldPos
    map As Integer
    x As Integer
    y As Integer
End Type

'Datos de user o npc
Type Char
    charIndex As Integer        ' Identificador Unico del Char (tanto para Usuarios como para Criaturas)
    Head As Integer             ' Cabeza
    Body As Integer             ' Cuerpo
    WeaponAnim As Integer       ' Arma
    ShieldAnim As Integer       ' Escudo
    CascoAnim As Integer        ' Casco
    FX As Integer               ' Efecto
    loops As Integer            ' Loops del Efecto
    heading As Byte             ' Hacia donde está mirando
End Type

Public Type ObjectoNecesario
    ObjIndex As Integer
    cantidad As Integer
End Type

'Tipos de objetos
Public Type ObjData
    Name As String 'Nombre del obj
    ObjType As Integer 'Tipo enum que determina cuales son las caract del obj
    subTipo As Integer 'Tipo enum que determina cuales son las caract del obj
    GrhIndex As Integer ' Indice del grafico que representa el obj
    tier As Byte    ' Nivel de poder de este item.
    GrhSecundario As Integer
    'Solo contenedores
    Apuñala As Byte
    QuitaEnergia As Integer
    HechizoIndex As Integer
    SkillCombate As Integer
    MineralIndex As Integer
    proyectil As Integer
    Municion As Integer
    StaffPower As Integer
    StaffDamageBonus As Integer
    Crucial As Byte
    Newbie As Integer
    TipoPocion As Byte
    MaxModificador As Integer
    MinModificador As Integer
    DuracionEfecto As Long
    MinSkill As Integer
    LingoteIndex As Integer
    LeñaIndex As Integer
    MinHIT As Integer 'Minimo golpe
    MaxHIT As Integer 'Maximo golpe
    minham As Integer 'Cuanto subo de hambre cuando se lo come
    MinSed As Integer 'Cuando subo de sed cuando se lo toma
    MinDef As Integer ' Armaduras
    MaxDef As Integer ' Armaduras
    Ropaje As Integer 'Indice del grafico del ropaje
    WeaponAnim As Integer ' Apunta a una anim de armas
    ShieldAnim As Integer ' Apunta a una anim de escudo
    CascoAnim As Integer
    valor As Long     ' Precio
    Cerrada As Integer
    Llave As Byte
    clave As Long 'si clave=llave la puerta se abre o cierra
    IndexAbierta As Integer
    IndexCerrada As Integer
    IndexCerradaLlave As Integer
    RazaEnana As Byte
    Genero As Byte
    Envenena As Byte
    Agarrable As Byte
        
    recursosNecesarios() As ObjectoNecesario
    premioReciclaje() As ObjectoNecesario
    
    SkillTacticass As Integer
    SkillDefe As Integer
    SkHerreria As Integer
    SkCarpinteria As Integer
    texto As String
    SkillM As Byte
    SkillMin As Byte
    
    clasesPermitidas As Long            'Clases que no tienen permitido usar este obj
    razas As Long
    alineacion As Byte
    
    Snd1 As Integer
    SeCae As Integer
    DefensaMagicaMax As Integer
    DefensaMagicaMin As Integer
    Refuerzo As Byte
    Ubicable As Byte
End Type

Public Type obj
    ObjIndex As Integer
    Amount As Integer
End Type

Public Const MAX_BANCOINVENTORY_SLOTS = 40

Type BancoInventario
    Object(1 To MAX_BANCOINVENTORY_SLOTS) As UserOBJ
    NroItems As Integer
End Type

'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'**  T I P O S   D E    N P C S **************************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Public Enum eTerrenoNPC
    Agua = 1
    Tierra = 2
    AguayTierra = 3
End Enum

Public Type NPCStats
    MaxHP As Long 'Vida
    minHP As Long 'Vida actual
    MaxHIT As Integer 'Golpe Maximo
    MinHIT As Integer 'Golpe Minimo
    Def As Integer 'Defensa
End Type

Type NpcCounters
    Paralisis As Long 'Cuando tiempo le falta antes de ser removido
    TiempoExistencia As Long 'Tiempo que le queda antes de morir
    TiempoUltimoAtaque As Long 'Tiempo en el que realizo el ultimo ataque
End Type

Public Type NPCFlags
    AfectaParalisis As Byte 'Si le afecta la paralisis
    Domable As Integer 'Puntos necesarios para domar al animal
    Respawn As Byte 'Si luego de matador respawenea
    ExpCount As Long 'Experiencia que le queda por entregar
    OldMovement As Byte 'Vieja actitud que tenia
    Terreno As eTerrenoNPC 'Terreno dondepuede andar el NPC
    BackUp As Byte 'Para npcs unicos. Si se debe backupear
    Paralizado As Byte '1 si es paralizado
    Inmovilizado As Byte '1 si esta inmovilizado
    Snd1 As Integer 'Sonido cuando la criatura ataca
    Snd2 As Integer 'Sonido cuando la criatura es atacada
    Snd3 As Integer 'Sonido cuando la criatura muere
End Type

Type tCriaturasEntrenador
    npcIndex As Integer
    NpcName As String
End Type

'<--------- New type for holding the pathfinding info ------>
Type NpcPathFindingInfo
    Path() As tVertice      ' This array holds the path
    Target As Position      ' The location where the NPC has to go
    PathLenght As Integer   ' Number of steps *
    CurPos As Integer       ' Current location of the npc
    TargetUser As Integer   ' UserIndex chased
    NoPath As Boolean       ' If it is true there is no path to the target location
    '* By setting PathLenght to 0 we force the recalculation
    '  of the path, this is very useful. For example,
    '  if a NPC or a User moves over the npc's path, blocking
    '  its way, the function NpcLegalPos set PathLenght to 0
    '  forcing the seek of a new path.
End Type
'<--------- New type for holding the pathfinding info ------>
Public Type npc
    Name As String 'Nombre
    Char As Char 'Define como se vera
    desc As String 'Descripcion
    NPCtype As Byte 'Tipo
    numero As Integer 'Numero del NPC
    TargetUserID As Long 'ID del personaje al cual tiene que atacar
    TargetNPCID As Long 'ID del npc al cual tiene que atacar
    Veneno As Byte 'Si el NPC puede envenenar a un usuario
    pos As WorldPos 'Posicion
    Orig As WorldPos 'Posicion donde nace
    Movement As Byte  'Actitud
    Attackable As Byte 'Si se le puede atacar
    InmuneAHechizos As Byte 'Si es inmune a hechizos
    PoderAtaque As Long 'TODO revisar si esto es long
    PoderEvasion As Long 'TODO revisar si esto es un long
    GiveEXP As Long 'Experiencia que da
    GiveGLD As Long 'Oro que da
    Stats As NPCStats
    flags As NPCFlags
    Contadores As NpcCounters
    NroSpells As Byte 'Cantidad de magias que puede lanzar
    Spells() As Integer  ' Magias disponibles para lanzar
    MaestroUser As Integer 'Quien es el dueño
    Hostil As Boolean ' ¿la criatura ataca a los usuarios por si sola?
    '<<<<<<Faccion>>>>>>>>
    faccion As eAlineaciones 'Faccion del npc
    '<<<<Comerciantes>>>>>
    Comercia As Byte 'Si es un NPC comerciante
    TipoItems As Integer 'Tipo de items que puede comerciar
    InvReSpawn As Byte 'Si cuando se le acaba el inventario vuelve a aparecer
    Inflacion As Long 'Porcentaje de mas que tine de precio
    Invent As inventario
    '<<<<Entrenadores>>>>>
    NroCriaturas As Integer
    Criaturas() As tCriaturasEntrenador
    MaestroNpc As Integer
    Mascotas As Integer
    '<<<<Pathfindig>>>>>>>
    PFINFO As NpcPathFindingInfo
    Inteligencia As Inteligencia
    '<<<Anti robo de npcs>
    UserIndexLucha As Integer 'Quien fue el ultimo que le pego
    UltimoGolpe As Long 'En que momento le pego
    
    Nivel As Byte ' Nivel de la criatura
    
    npcIndex As Integer ' Identificador unico de la criatura en el juego
    
End Type

Public listaRazas() As String
Public SkillsNames() As String

Public RecordUsuarios As Long

'Directorios
Public IniPath As String
Public MapPath As String
Public DatPath As String

'Bordes del mapa
Public NumUsers As Integer 'Numero de usuarios actual
Public NumUsersPremium As Integer 'Numero de usuarios premium


Public LastChar As Integer
Public NumChars As Integer
Public NumNPCs As Integer
Public NumObjDatas As Integer
Public NumeroHechizos As Integer
Public MaxUsers As Integer
Public haciendoBK As Boolean



Public EnPausa As Boolean

Public ProfilePaquetes As Boolean   ' Se va a guardar la marca de tiempo de los paquetes que se reciban
'*****************ARRAYS PUBLICOS*************************

Public NpcList() As npc 'NPCS

Public hechizos() As tHechizo
Public CharList() As Integer
Public ObjData() As ObjData
Public SpawnList() As tCriaturasEntrenador
Public ArmasHerrero() As Integer
Public ArmadurasHerrero() As Integer
Public ObjCarpintero() As Integer
'*********************************************************

Public Const Max_Distance As Byte = 12

Public GmsGroup As EstructurasLib.ColaConBloques
Public TrabajadoresGroup As EstructurasLib.ColaConBloques
Public Ayuda As EstructurasLib.ColaConBloques



Public Declare Function GetTickCount Lib "kernel32" () As Long
'Public VecesQuePasoPorDo As Single
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'Lista de clanes
Public Enum eRazas
    indefinido = 0 'Lo que seria el valor nulo
    Humano = 1
    Enano = 2
    Elfo = 4
    ElfoOscuro = 8
    Gnomo = 16
End Enum

Public Enum eGeneros
    indefinido = 0 'Lo que seria el valor nulo
    Hombre = 1
    Mujer = 2
End Enum


Public Type razasConfig
   nombre As String
End Type

Public denunciarActivado As Boolean ' Desactiva el anunciar

Public razasConfig(1 To NUMRAZAS) As razasConfig

Public Sub inicializarRazas()
    razasConfig(1).nombre = "HUMANO"
    razasConfig(2).nombre = "ENANO"
    razasConfig(3).nombre = "ELFO"
    razasConfig(4).nombre = "ELFO OSCURO"
    razasConfig(5).nombre = "GNOMO"
End Sub

Public Function razasToString(razas As Long) As String

    Dim loopRaza As Byte

    For loopRaza = 1 To NUMRAZAS
        If ((2 ^ (loopRaza - 1)) And razas) Then
            razasToString = razasToString & " " & razasConfig(loopRaza).nombre
        End If
    Next
    
End Function

Public Function razaConfigToEnum(configId As Byte) As eClases
    razaConfigToEnum = 2 ^ (configId - 1)
End Function


Public Function razaToConfigID(Raza As eRazas) As Long

    Select Case Raza
        Case eRazas.Humano
            razaToConfigID = 1
        Case eRazas.Enano
            razaToConfigID = 2
        Case eRazas.Elfo
            razaToConfigID = 3
        Case eRazas.ElfoOscuro
            razaToConfigID = 4
        Case eRazas.Gnomo
            razaToConfigID = 5
    End Select
        
        
End Function

Public Function razaToByte(Raza As String) As Byte

    Select Case UCase$(Raza)
        Case "HUMANO"
            razaToByte = eRazas.Humano
        Case "ENANO"
            razaToByte = eRazas.Enano
        Case "ELFO"
            razaToByte = eRazas.Elfo
        Case "ELFO OSCURO"
            razaToByte = eRazas.ElfoOscuro
        Case "GNOMO"
            razaToByte = eRazas.Gnomo
    End Select
        
End Function

Public Function generoToByte(Genero As String) As Byte
     Select Case UCase$(Genero)
        Case "HOMBRE"
            generoToByte = eGeneros.Hombre
        Case "MUJER"
            generoToByte = eGeneros.Mujer
     End Select
End Function

Public Function byteToGenero(Genero As eGeneros) As String
     Select Case Genero
        Case eGeneros.Hombre
            byteToGenero = "HOMBRE"
        Case eGeneros.Mujer
            byteToGenero = "MUJER"
     End Select
End Function


Public Function byteToRaza(Raza As eRazas) As String

    Select Case Raza
        Case eRazas.Humano
            byteToRaza = "HUMANO"
        Case eRazas.Enano
            byteToRaza = "ENANO"
        Case eRazas.Elfo
            byteToRaza = "ELFO"
        Case eRazas.ElfoOscuro
            byteToRaza = "ELFO OSCURO"
        Case eRazas.Gnomo
            byteToRaza = "GNOMO"
    End Select
        
End Function
