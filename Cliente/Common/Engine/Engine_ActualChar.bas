Attribute VB_Name = "Engine_ActualChar"
'       __________________
'      / __________  ____ \
'     / |_   _|    \/  __\ \
'    /    | | | |\  \ `--.  \
'   /     | | | | | |`--. \  \
'  /      | | | |/  /\__/ /   \
'  \      |_| |____/\____/    /
'   \________________________/
Option Explicit

Public Const FLAGORO As Integer = 777

'-------------- User stats -------------------------
Public Enum eClass
    Mage = 1    'Mago
    Cleric      'Clérigo
    Warrior     'Guerrero
    Assasin     'Asesino
    Recolector       'Ladrón
    Bard        'Bardo
    Druid       'Druida
    Paladin     'Paladín
    Hunter      'Cazador
    Fisher      'Pescador
    Blacksmith  'Herrero
    Lumberjack  'Leñador
    Miner       'Minero
    Carpenter   'Carpintero
    Pirat       'Pirata
    Sastre = 20
End Enum

Enum eRaza
    Humano = 1
    Elfo
    ElfoOscuro
    Gnomo
    Enano
End Enum

Public Enum eSkill
    Suerte = 1
    Magia = 2
    Robar = 3
    Tacticas = 4
    armas = 5
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

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5
End Enum

Enum eGenero
    Hombre = 1
    Mujer = 2
End Enum

Public Enum PlayerType
    user = &H1
    Consejero = &H2
    SemiDios = &H4
    Dios = &H8
    Admin = &H10
    RoleMaster = &H20
    ChaosCouncil = &H40
    RoyalCouncil = &H80
End Enum

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserMoving As Byte

Public AddtoUserPos As D3DVECTOR2 ', Position 'Si se mueve
Public AddtoUserPosO As D3DVECTOR2 ',Position 'Si se mueve

Global OtroInventario(1 To MAX_INVENTORY_SLOTS) As Inventory

Public tx As Integer
Public ty As Integer

