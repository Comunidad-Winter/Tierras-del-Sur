VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMapCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'LUCES
    Private pIntensidad_Del_Terreno(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE) As Byte         'Guarda la intensidad de la luz de un vertice del mapa
    Private pOriginalMapColor(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)      As BGRACOLOR_DLL   'Colores precalculados en el mapeditor
    Private pOriginalMapColorSombra(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE) As Long         'OriginalMapColor * Sombra
    Private pOriginalColorArray(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)     As Long         'BACKUP DE ResultColorArray (OriginalMapColorSombra * AMBIENTE)
'/LUCES

'AGUA DIFUMADA
    Private pOpacidadesAgua(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)         As Byte
'/AGUA DIFUMADA

'AGUA DINÁMICA(?)
    Private pModSuperWater(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)       As Byte
    Private pModSuperWaterDD(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)       As AUDT
    Private pModSuperWaterMM(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)       As AUDT
    Private pTileNumberWater(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)      As Byte
    Private pAguaBoxes(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)        As Box_Vertex
'/AUGA

'MONTAÑAS
    Private phMapData(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)             As AUDT
    Private pAlturaPie(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)             As Integer
    Private pAlturas(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)            As Integer
    Private pNormalData(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)           As mzVECTOR
    Private pAlturaAgua                                         As Integer
    Private pTexturaAgua                                        As Integer
    Private pSombra_Montañas(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)     As Byte
    Private pMapBoxes(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)            As Box_Vertex
'/MONTAÑAS
    
    Private pMapData(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)              As MapBlock
    Private pMapInfo As mapinfo
    
    Private Cargado As Boolean
    
    Private pMapNum As Long

Private Declare Sub mcpy Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'Private Declare Sub mcpy Lib "MZEngine.dll" Alias "memcpy_fast" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Sub GuardarMontañas()
    Dim Tamaño As Long

    Tamaño = LenB(hMapData(1, 1)) * TILES_POR_MAPA
    mcpy phMapData(1, 1), hMapData(1, 1), Tamaño
    
    Tamaño = LenB(AlturaPie(1, 1)) * TILES_POR_MAPA
    mcpy pAlturaPie(1, 1), AlturaPie(1, 1), Tamaño
    mcpy pAlturas(1, 1), Alturas(1, 1), Tamaño
    
    pAlturaAgua = AlturaAgua
    pTexturaAgua = TexturaAgua
    
    Tamaño = LenB(NormalData(1, 1)) * TILES_POR_MAPA
    mcpy pNormalData(1, 1), NormalData(1, 1), Tamaño

    Tamaño = TILES_POR_MAPA
    mcpy pSombra_Montañas(1, 1), Sombra_Montañas(1, 1), Tamaño

    Tamaño = BV_size * TILES_POR_MAPA
    mcpy pMapBoxes(1, 1), MapBoxes(1, 1), Tamaño
End Sub

Private Sub CargarMontañas()
    Dim Tamaño As Long

    Tamaño = LenB(hMapData(1, 1)) * TILES_POR_MAPA
    mcpy hMapData(1, 1), phMapData(1, 1), Tamaño
    
    Tamaño = LenB(AlturaPie(1, 1)) * TILES_POR_MAPA
    mcpy AlturaPie(1, 1), pAlturaPie(1, 1), Tamaño
    mcpy Alturas(1, 1), pAlturas(1, 1), Tamaño
    
    AlturaAgua = pAlturaAgua
    TexturaAgua = pTexturaAgua
    
    Tamaño = LenB(NormalData(1, 1)) * TILES_POR_MAPA
    mcpy NormalData(1, 1), pNormalData(1, 1), Tamaño

    Tamaño = TILES_POR_MAPA
    mcpy Sombra_Montañas(1, 1), pSombra_Montañas(1, 1), Tamaño

    Tamaño = BV_size * TILES_POR_MAPA
    mcpy MapBoxes(1, 1), pMapBoxes(1, 1), Tamaño
End Sub

Private Sub GuardarLuces()

    mcpy pIntensidad_Del_Terreno(1, 1), Intensidad_Del_Terreno(1, 1), TILES_POR_MAPA
    mcpy pOriginalMapColor(1, 1), OriginalMapColor(1, 1), TILES_POR_MAPA * 4
    mcpy pOriginalMapColorSombra(1, 1), OriginalMapColorSombra(1, 1), TILES_POR_MAPA * 4
    mcpy pOriginalColorArray(1, 1), OriginalColorArray(1, 1), TILES_POR_MAPA * 4

End Sub

Private Sub CargarLuces()

    mcpy Intensidad_Del_Terreno(1, 1), pIntensidad_Del_Terreno(1, 1), TILES_POR_MAPA
    mcpy OriginalMapColor(1, 1), pOriginalMapColor(1, 1), TILES_POR_MAPA * 4
    mcpy OriginalMapColorSombra(1, 1), pOriginalMapColorSombra(1, 1), TILES_POR_MAPA * 4
    mcpy OriginalColorArray(1, 1), pOriginalColorArray(1, 1), TILES_POR_MAPA * 4
    
End Sub

Private Sub GuardarAgua()

Dim Tamaño As Long

    Tamaño = TILES_POR_MAPA
    mcpy pOpacidadesAgua(1, 1), OpacidadesAgua(1, 1), Tamaño
    mcpy pModSuperWater(1, 1), ModSuperWater(1, 1), Tamaño
    mcpy pTileNumberWater(1, 1), TileNumberWater(1, 1), Tamaño
    
    Tamaño = LenB(ModSuperWaterDD(1, 1)) * ALTO_MAPA
    mcpy pModSuperWaterDD(1, 1), ModSuperWaterDD(1, 1), Tamaño
    mcpy pModSuperWaterMM(1, 1), ModSuperWaterMM(1, 1), Tamaño
    

    Tamaño = BV_size * ALTO_MAPA
    mcpy pAguaBoxes(1, 1), AguaBoxes(1, 1), Tamaño



End Sub

Private Sub CargarAgua()

Dim Tamaño As Long

    Tamaño = ALTO_MAPA
    mcpy OpacidadesAgua(1, 1), pOpacidadesAgua(1, 1), Tamaño
    mcpy ModSuperWater(1, 1), pModSuperWater(1, 1), Tamaño
    mcpy TileNumberWater(1, 1), pTileNumberWater(1, 1), Tamaño
    
    Tamaño = LenB(ModSuperWaterDD(1, 1)) * ALTO_MAPA
    mcpy ModSuperWaterDD(1, 1), pModSuperWaterDD(1, 1), Tamaño
    mcpy ModSuperWaterMM(1, 1), pModSuperWaterMM(1, 1), Tamaño
    

    Tamaño = BV_size * ALTO_MAPA
    mcpy AguaBoxes(1, 1), pAguaBoxes(1, 1), Tamaño

Water_Actualizar_Vertices = True

End Sub


Public Sub Guardar()
    GuardarMontañas
    GuardarLuces
    GuardarAgua
    
    mcpy pMapData(1, 1), mapdata(1, 1), LenB(mapdata(1, 1)) * ALTO_MAPA
    
    
    pMapInfo = mapinfo
    
    pMapNum = CurMap
    
    Cargado = True
    
    Debug.Print "MAPA "; CurMap; " CACHEADO EN MEMORIA."

End Sub

Public Function Cargar() As Boolean
If Cargado Then
    Debug.Print "MAPA "; CurMap; " CARGADO DESDE MEMORIA!."
    Dim x As Long
    Dim y As Long


    mcpy mapdata(1, 1), pMapData(1, 1), LenB(mapdata(1, 1)) * ALTO_MAPA
    
    ForzarRecalculoLuces = True
    
    CargarMontañas
    CargarAgua
    CargarLuces

    mapinfo = pMapInfo
    Cargar = True
    
    CurMap = pMapNum
    
    Engine_Landscape.Light_Update_Map = True
    cron_tiempo

Else
Cargar = False
MsgBox "CATASTROFE!, no hay mapa cargado en el clsCache"
End If
End Function
