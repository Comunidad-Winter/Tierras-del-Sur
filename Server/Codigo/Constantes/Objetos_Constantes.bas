Attribute VB_Name = "Objetos_Constantes"
Option Explicit

Public Const OBJTYPE_USEONCE = 1
Public Const OBJTYPE_WEAPON = 2
Public Const OBJTYPE_ARMOUR = 3
Public Const OBJTYPE_ARBOLES = 4
Public Const OBJTYPE_GUITA = 5
Public Const OBJTYPE_PUERTAS = 6
Public Const OBJTYPE_COFRES = 7
Public Const OBJTYPE_CARTELES = 8
Public Const OBJTYPE_LLAVES = 9
Public Const OBJTYPE_POCIONES = 11
Public Const OBJTYPE_BEBIDA = 13
Public Const OBJTYPE_LEÑA = 14
Public Const OBJTYPE_FOGATA = 15
Public Const OBJTYPE_HERRAMIENTAS = 18
Public Const OBJTYPE_TELEPORT = 19
Public Const OBJTYPE_YACIMIENTO = 22
Public Const OBJTYPE_MINERALES = 23
Public Const OBJTYPE_PERGAMINOS = 24
Public Const OBJTYPE_YUNQUE = 27
Public Const OBJTYPE_FRAGUA = 28
Public Const OBJTYPE_CUALQUIERA = 1000
Public Const OBJTYPE_INSTRUMENTOS = 26
Public Const OBJTYPE_BARCOS = 31
Public Const OBJTYPE_FLECHAS = 32
Public Const OBJTYPE_BOTELLAVACIA = 33
Public Const OBJTYPE_BOTELLALLENA = 34
Public Const OBJTYPE_MANCHAS = 35
Public Const OBJTYPE_ANILLOS = 36
Public Const OBJTYPE_COLLAR = 37
Public Const OBJTYPE_BRASALETE = 38
Public Const OBJTYPE_TRANSLADO = 39
Public Const OBJTYPE_VIAJES = 40

' Sub Categorias
Public Const OBJTYPE_ARMADURA = 0
Public Const OBJTYPE_CASCO = 1
Public Const OBJTYPE_ESCUDO = 2
Public Const OBJTYPE_CAÑA = 138
' ***********************************************************************************
' Index especifico

Public Const Barcafantasmal As Integer = 314
Public Const Barcaazul As Integer = 307
Public Const Barcaroja As Integer = 84
Public Const Barcaneutra As Integer = 208
Public Const Galeonfanstamal As Integer = 312
Public Const Galeongris As Integer = 225
Public Const GaleonAzul As Integer = 305
Public Const Galeonrojo As Integer = 306
Public Const Galeraazul As Integer = 309
Public Const Galeraroja As Integer = 310
Public Const Galeragris As Integer = 311
Public Const Galerafanstamal As Integer = 312


Public Const NingunEscudo = 2
Public Const NingunCasco = 2
Public Const NingunArma = 2

Public Const ObjArboles = 4
Public Const iORO = 12
Public Const DAGA = 15
Public Const Leña = 58
Public Const FOGATA = 63
Public Const iFragataFantasmal = 87
Public Const HACHA_LEÑADOR As Integer = 127
Public Const FOGATA_APAG = 136
Public Const PIQUETE_MINERO = 187
Public Const SERRUCHO_CARPINTERO = 198

Public Const ARMADURA_DE_CAZADOR As Integer = 360
Public Const ARMADURA_DE_CAZADOR_G As Integer = 612
Public Const ARMADURA_DE_CAZADOR_2 As Integer = 671

Public Const EQUIPO_INVERNAL_HH As Integer = 665
Public Const EQUIPO_INVERNAL_HM As Integer = 666
Public Const EQUIPO_INVERNAL_EG As Integer = 667

Public Const LingoteHierro = 386
Public Const LingotePlata = 387
Public Const LingoteOro = 388
Public Const MARTILLO_HERRERO = 389
Public Const EspadaMataDragonesIndex = 402
Public Const RED_PESCA = 543

Public Const SOMBRERO_DE_APRENDIZ As Integer = 621
Public Const SOMBRERO_DE_MAGO As Integer = 622

Public Const HACHA_DORADA As Integer = 630
Public Const ARBOL_DE_TEJO As Integer = 634
Public Const Leña_tejo = 642

Public Const ANILLO_RESISTENCIA = 644
Public Const ANILLO_RESISTENCIA_M1 = 128

Public Const ANILLO_PROTECCION = 645

Public Const LAUDMAGICO = 147
Public Const LAUDMAGICO_M1 = 643
Public Const LAUDESPECIAL = 659

Public Const CRUZMADERA = 149
Public Const CRUZTEJO = 862

Public Const VARA_FRESNO = 625
Public Const BASTON_NUDOSO = 624
Public Const BACULO_ENGARZADO = 623

Public Const FLAUTA_MAGICA = 540
Public Const ANILLO_PLATA = 137
Public Const ANILLO_PLATA_M1 = 648
Public Const ANILLO_PLATA_M2 = 49

Public Const ANILLOMAGICODRUIDA = 648
Public Const COLLAR = 859
Public Const PIQUETE_DE_ORO As Integer = 685


Public Const POCION_ROJA_NEWBIE = 461
Public Const POCION_AZUL_NEWBIE = 462
Public Const POCION_VIOLETA = 682
Public Const POCION_AMARILLA_NEWBIE = 650
Public Const POCION_VERDE_NEWBIE = 651

' Las Armaduras del Dragon no se cae
Public Const ARMADURA_DRAGON_H As Integer = 481 ' Hombre Elfo/Humano/Elfo Oscuro
Public Const ARMADURA_DRAGON_M As Integer = 482 ' Mujer Elfa/Humana/Elfa Oscura
Public Const ARMADURA_DRAGON_E As Integer = 483 ' Enanos / Gnomos


' Peces que se pueden obtener a través de la Pesca
Public Enum PECES_POSIBLES
    PESCADO1 = 139
    PESCADO2 = 544
    PESCADO3 = 545
    PESCADO4 = 546
    PESCADO5 = 732
End Enum

' Relacionados a los Minerales
Public Enum iMinerales
    hierrocrudo = 192
    platacruda = 193
    orocrudo = 194
    LingoteDeHierro = 386
    LingoteDePlata = 387
    LingoteDeOro = 388
End Enum

Public Enum ePociones
    Agilidad = 1
    Fuerza = 2
    Roja = 3
    Azul = 4
    Violeta = 5
    Negra = 6
    Energia = 7
End Enum
