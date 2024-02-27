Attribute VB_Name = "modMiPersonaje"
Option Explicit

' *********** Constantes
Public Const MAX_INVENTORY_OBJS As Integer = 10000  ' Maxima canitdad de objetos por slot
Public Const MAX_INVENTORY_SLOTS As Byte = 30       ' Cantidad de Slots en Inventario
Public Const MAXHECHI As Byte = 35                  ' Cantidad de Slots en hechizos
Public Const MAXSKILLPOINTS As Byte = 100           ' Cantidad de puntos maximos en los skills
Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40  ' Cantidad de Slots en el inventario
Public Const NUMSKILLS As Byte = 21                 ' Cantidad de tipos Skils
Public Const NUMATRIBUTOS As Byte = 5               ' Cantidad de Atributos

' ********** Estructuras
Public Type tReputacion                             ' Reputacion del personaje
    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    
    Promedio As Long
End Type

Public Type tEstadisticasUsu                        ' Estadisticas
    CiudadanosMatados As Long
    CriminalesMatados As Long
    UsuariosMatados As Long
    NpcsMatados As Long
    Clase As String
    PenaCarcel As Long
End Type

'Inventario
Public Type Inventory                               ' Informacion en cada Slot del inventario
    OBJIndex As Integer
    Name As String
    GrhIndex As Integer
    Amount As Long
    Equipped As Byte
    valor As Long
    OBJType As Integer
    MaxDef As Integer
    MinDef As Integer
    MaxHit As Integer
    MinHit As Integer
End Type

' Stats
Public UserName As String           ' Nombre
Public UserMaxHP As Integer         ' Puntos de vida del personaje
Public UserMaxMAN As Integer        ' Puntos de mano del personaje
Public UserMaxSTA As Integer        ' PUntos de energia del personaje
Public UserMaxAGU As Byte           ' Puntos de sed del personaje
Public UserMinAGU As Byte           ' Puntos de sed actual
Public UserMaxHAM As Byte           ' Puntos de hambre del personaje
Public UserMinHAM As Byte           ' Puntos de hambre actual
Public UserGLD As Long              ' Oro en la billetera
Public UserLvl As Integer           ' Nivel
Public UserPasarNivel As Long       ' Experiencia necesaria para pasar de nivel
Public UserExp As Long              ' Experiencia actual
Public UserBody As Integer          ' Cuerpo
Public UserHead As Integer          ' Cabeza
Public UserPrivilegios As Byte      ' Privilegios del Personaje (User Comun, GM, Dios)

' Flags
Public UserSeguro As Boolean        ' ¿El usuario tiene el seguro activado?
Public UserDescansar As Boolean     ' ¿Esta descansando?
Public Bovedeando As Boolean        ' ¿El personaje esta usando la bovda?
Public Comerciando As Boolean       ' ¿El personaje esta comerciando?
Public UserMeditar As Boolean       ' ¿Esta Meditando?
Public UserMap As Integer           ' Mapa actual donde esta el personaje
Public UserNavegando    As Boolean  ' ¿El usuario esta navegando?
Public Istrabajando As Boolean      ' ¿Esta trabajando?
Public Liderparty As Boolean        ' ¿Es el lider de una party?
Public gh As Boolean                ' ¿Esta participando en una party?
Public IScombate As Boolean         ' ¿Esta en Modo Combate?
Public IsEnvenenado As Boolean      ' ¿Esta Envenenado?

Public UserReputacion As tReputacion
Public UserEstadisticas As tEstadisticasUsu

Public UserPos As Position          ' Posicion actual del personaj

' Inventario
Public UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory

' Boveda
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory

' Hechizos
Public UserHechizos(1 To MAXHECHI) As Integer
    
' Skills
Public UserSkills() As Integer
Public SkillPoints As Integer       ' Skills Libres

' Atributos
Public UserAtributos() As Integer

' Intervalos
Public Puedeatacar As Single

' CharIndex Actual
Public UserCharIndex As Integer

Public Type tUserStats
    UserMinHP As Integer
    UserMinMAN As Integer
    UserMinSTA As Integer
    UserEstado As Byte
    UserParalizado As Boolean
    UserCentinela As Boolean
    UsingSkill As Byte
    IntervaloNoChupClick As Single
    intervaloNoChupU As Single
    IntervaloPegar As Single
    IntervaloLanzarMagias As Single
    IntervaloLanzarFlechas As Single
End Type

Public SlotStats As Byte
Public UserStats() As tUserStats


Public Sub iniciar()
    Dim i As Integer
    
    ' Evitamos que puedan leer una dirección fija de memoria facilmente
    SlotStats = Int(Rnd * 12)

    ReDim UserStats(SlotStats)
    
    With UserStats(SlotStats)
    
        .IntervaloLanzarFlechas = 0
        .IntervaloLanzarMagias = 0
        .IntervaloNoChupClick = 0
        .intervaloNoChupU = 0
        .IntervaloPegar = 0
        
        .UserCentinela = False
        .UserEstado = False
        .UserMinHP = 0
        .UserMinMAN = 0
        .UserParalizado = False
        .UsingSkill = False
        
    End With
    
    ' Stats
    UserName = ""
    UserMaxHP = 0
    UserMaxMAN = 0
    UserMaxSTA = 0
    UserMaxAGU = 0
    UserMinAGU = 0
    UserMaxHAM = 0
    UserMinHAM = 0
    UserGLD = 0
    UserLvl = 0
    UserPasarNivel = 0
    UserExp = 0
    UserBody = 0
    UserHead = 0
    UserPrivilegios = 0

    ' Flags
    UserSeguro = False
    UserDescansar = False
    Bovedeando = False
    Comerciando = False
    UserMeditar = False
    UserMap = 0
    UserNavegando = False
    Istrabajando = False
    Liderparty = False
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
    
    SkillPoints = 0
End Sub

