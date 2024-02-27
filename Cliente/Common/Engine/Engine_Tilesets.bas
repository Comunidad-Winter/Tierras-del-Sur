Attribute VB_Name = "Engine_Tilesets"
Option Explicit

Public Tileset_Grh_Array(255) As Box_Vertex

Public Tilesets() As TilesetStruct
Public Tilesets_count As Integer

#If esMe Then
    Public Type TileSetInfoTile
        textura As Integer
        numero As Integer
    End Type
    
    Public Enum eFormatoTileSet
        formato_viejo = 0
        textura_simple = 1
        camino_chico = 2
        camino_grande_parte1 = 3
        camino_grande_parte2 = 4
        textura_agua = 5
        costa_tipo_1_parte1 = 6
        costa_tipo_1_parte2 = 7
        rocas_acuaticas = 8
    End Enum
#End If

Public Type TilesetStruct

    filenum         As Long
    stage_actual    As Byte
    stage_count     As Byte
    stage_mu        As Long
    stages()        As Long
    anim            As Long 'Tiempo total animacion; 0=estatico
    Nombre          As String
    Olitas          As Integer
    
    #If esMe Then
        ' Nuevo formato de efectos de sonido al pisar
        EfectoPisada(0 To 255) As Integer
        ' Nuevo formato de piso
        formato As eFormatoTileSet
        ' Referencia a la textura (y donde comienza)
        ' o en caso de ser camino_grande_parte2, la referencia a la grilla 1
        referencia As TileSetInfoTile
        matriz_transformacion() As TileSetInfoTile
    #End If
    
End Type

Public TilesetVersion As Long

Public Sub Init_Tilesets()
    Dim x&, y&, loopC&
    'ReDim Tileset_Grh_Array(0 To 255)
    For y = 0 To 15
        For x = 0 To 15
            With Tileset_Grh_Array(loopC)
                .tu0 = (x * 32) / 512!
                .tv0 = ((y + 1) * 32!) / 512!
                
                .tu1 = .tu0
                .tv1 = (y * 32) / 512!
                
                .tu2 = ((x + 1) * 32!) / 512!
                .tv2 = .tv0
                
                .tu3 = .tu2
                .tv3 = .tv1
                .rhw0 = 1
                .rhw1 = 1
                .rhw2 = 1
                .rhw3 = 1
            End With
            loopC = loopC + 1
        Next x
    Next y
End Sub


Public Sub AnimarTilesets()
    Dim loopC As Integer
    Dim i As Integer
    Dim TmpTime As Long
    TmpTime = GetTimer

    For loopC = 1 To Tilesets_count
        With Tilesets(loopC)
            If .anim Then
                i = Fix((TmpTime Mod .anim) / .anim * .stage_count) + 1
                
                '                                  Tick Mod TiempoAnim
                ' NumeroGrafico = Parte entera(  _______________________  * CantidadGraficos) + 1
                '                        T            TiempoAnim
                
                .filenum = .stages(i)
            End If
        End With
    Next loopC
End Sub
