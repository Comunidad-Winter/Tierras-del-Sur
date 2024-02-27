Attribute VB_Name = "SV_Constantes"
Option Explicit

'Cantidad de tiles
Public Const ANCHO_MAPA As Long = 214
Public Const ALTO_MAPA As Long = 214

'Tamaño real del mapa
Public Const X_MINIMO_VISIBLE  As Long = 1
Public Const Y_MINIMO_VISIBLE As Long = 1

Public Const X_MAXIMO_VISIBLE As Long = X_MINIMO_VISIBLE + ANCHO_MAPA - 1
Public Const Y_MAXIMO_VISIBLE As Long = Y_MINIMO_VISIBLE + ALTO_MAPA - 1

'Tamaño del borde (Esto es la calidad de tiles que el usuario tiene por delante + 1 (donde van los traspasos)
' Actual = ((21 - 1) / 2) + 1. 11
Public Const BORDE_TILES_INUTILIZABLE As Long = 11

'Este rectangulo (o cuadrado) tiene todas las coordenadas del mapa que visualmente no van a ser vistas desde
'otro mapa al copiar los bordes.
Public Const X_MINIMO_NO_VISIBLE_OTRO_MAPA As Long = BORDE_TILES_INUTILIZABLE + BORDE_TILES_INUTILIZABLE + 1
Public Const Y_MINIMO_NO_VISIBLE_OTRO_MAPA As Long = BORDE_TILES_INUTILIZABLE + BORDE_TILES_INUTILIZABLE + 1
Public Const X_MAXIMO_NO_VISIBLE_OTRO_MAPA As Long = ANCHO_MAPA - BORDE_TILES_INUTILIZABLE - BORDE_TILES_INUTILIZABLE
Public Const Y_MAXIMO_NO_VISIBLE_OTRO_MAPA As Long = ALTO_MAPA - BORDE_TILES_INUTILIZABLE - BORDE_TILES_INUTILIZABLE

'Tamaño absoluto Usablae del mapa. Entre estas coordenadas se podrán poner portales, acciones,eTriggers, etc.
Public Const X_MINIMO_USABLE As Long = X_MINIMO_VISIBLE + BORDE_TILES_INUTILIZABLE
Public Const Y_MINIMO_USABLE As Long = Y_MINIMO_VISIBLE + BORDE_TILES_INUTILIZABLE

Public Const X_MAXIMO_USABLE As Long = X_MAXIMO_VISIBLE - BORDE_TILES_INUTILIZABLE
Public Const Y_MAXIMO_USABLE As Long = Y_MAXIMO_VISIBLE - BORDE_TILES_INUTILIZABLE

'Coordenadas donde el usuario se puede mover (aca se incluye la linea de portales)
Public Const X_MINIMO_JUGABLE As Long = X_MINIMO_VISIBLE + BORDE_TILES_INUTILIZABLE - 1
Public Const Y_MINIMO_JUGABLE As Long = Y_MINIMO_VISIBLE + BORDE_TILES_INUTILIZABLE - 1

Public Const X_MAXIMO_JUGABLE As Long = X_MAXIMO_VISIBLE - BORDE_TILES_INUTILIZABLE + 1
Public Const Y_MAXIMO_JUGABLE As Long = Y_MAXIMO_VISIBLE - BORDE_TILES_INUTILIZABLE + 1


'Tamaño relativo
Public Const X_MINIMO_USABLE_RELATIVO As Long = 1
Public Const X_MAXIMO_USABLE_RELATIVO As Long = X_MAXIMO_USABLE - X_MINIMO_USABLE + 1

Public Const Y_MINIMO_USABLE_RELATIVO As Long = 1
Public Const Y_MAXIMO_USABLE_RELATIVO As Long = Y_MAXIMO_USABLE - Y_MINIMO_USABLE + 1


'Cantidad de tiles que tiene un mapa
Public Const TILES_POR_MAPA As Long = ANCHO_MAPA * ALTO_MAPA
