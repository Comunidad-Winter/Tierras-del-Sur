Attribute VB_Name = "Me_Guardar_Cliente"
Option Explicit

Function Guardar_Mapa_CLI(ByVal SaveAs As String) As Boolean


Dim freeFileMap As Integer
Dim ByFlags As Integer

Dim loopC As Long
Dim y As Long
Dim x As Long
Dim tit As Integer

Dim offsetHeader As Long
Dim ResizeBackBufferY As Integer
Dim ResizeBackBufferX As Integer

Dim tempLuz As tLuzPropiedades
Dim posLuzX As Byte
Dim posLuzY As Byte
                        
Dim checkSum As String * 10
                
' Si el archivo existe lo eliminamos
If FileExist(SaveAs, vbNormal) = True Then Kill SaveAs
        
Guardar_Mapa_CLI = False

' Manejador para este archivo
freeFileMap = FreeFile

Open SaveAs For Binary As freeFileMap
Seek freeFileMap, 1
    
    '###################################################################################
    '########### HEADER

    Put freeFileMap, , header_m_3
    Put freeFileMap, , THIS_MAPA.numero
    
    '###################################################################################
    '########### PROPIEDADES
    offsetHeader = seek(freeFileMap)
    
    Put freeFileMap, , ResizeBackBufferX
    Put freeFileMap, , ResizeBackBufferY

    '###################################################################################
    '########### COLORES Y LUCES!
    
    'TODO Precalcular luces estáticas.
    Put freeFileMap, , OriginalMapColor
    Put freeFileMap, , Intensidad_Del_Terreno
    
    With mapinfo
        Put freeFileMap, , .BaseColor
        Put freeFileMap, , .ColorPropio
        
        Put freeFileMap, , .agua_tileset
        Put freeFileMap, , .agua_rect
        Put freeFileMap, , .agua_profundidad
        
        Put freeFileMap, , .UsaAguatierra
    End With

    ' En el cliente se guardan
    ' 0..4) Capas
    ' 5) Trigger (incluye bloqueo)
    ' 6) Piso
    ' 9) Alturas
    ' 10) Luz
    ' 7, 14, 8) Particula
    '
    
    For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
        For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
            With mapdata(x, y)
            
                ByFlags = 0
                                
                '0 a 4) Capas
                If .Graphic(1).GrhIndex Then ByFlags = ByFlags Or bitwisetable(1)
                If .Graphic(2).GrhIndex Then ByFlags = ByFlags Or bitwisetable(2)
                If .Graphic(3).GrhIndex Then ByFlags = ByFlags Or bitwisetable(3)
                If .Graphic(4).GrhIndex Then ByFlags = ByFlags Or bitwisetable(4)
                If .Graphic(5).GrhIndex Then ByFlags = ByFlags Or bitwisetable(0)
                
                '5) Trigger
                If .Trigger Then _
                    ByFlags = ByFlags Or bitwisetable(5)
                
                '6) Piso
                If .tile_texture Then _
                    ByFlags = ByFlags Or bitwisetable(6)
                
                '9) Alturas
                If Alturas(x, y) <> 0 Or AlturaPie(x, y) <> 0 Or hMapData(x, y).hs(0) <> 0 Then _
                    ByFlags = ByFlags Or bitwisetable(9)
                    
                '13) Luz
                If .luz Then
                    If DLL_Luces.Check(.luz) Then
                    
                        ' Tomamos las Propiedades
                        DLL_Luces.Get_Light .luz, posLuzX, posLuzY, tempLuz.LuzColor.r, tempLuz.LuzColor.g, tempLuz.LuzColor.b, tempLuz.LuzRadio, tempLuz.LuzBrillo, tempLuz.LuzTipo, tempLuz.luzInicio, tempLuz.luzInicio
                        
                        ' Ultimo chequeo antes de guardar la luz
                        If tempLuz.LuzRadio > 0 Then ByFlags = ByFlags Or bitwisetable(13)
                    End If
                End If
                
                '7)  Particula 0
                If Not .Particles_groups(0) Is Nothing Then ByFlags = ByFlags Or bitwisetable(7)
                
                '14) Particula 1
                If Not .Particles_groups(1) Is Nothing Then ByFlags = ByFlags Or bitwisetable(14)
                
                '8) Particula 2
                If Not .Particles_groups(2) Is Nothing Then ByFlags = ByFlags Or bitwisetable(8)
                
                ' Guardamos el Flag
                Put freeFileMap, , ByFlags
                
                ' ------------------------------------------------------------ '
                '           Guardado de Datos                                   '
                
                ' 6) Piso
                If ByFlags And bitwisetable(6) Then
                    Put freeFileMap, , .tile_texture
                    Put freeFileMap, , .tile_number
                End If
                
                ' 0 .. 4) Capas
                For loopC = 1 To CANTIDAD_CAPAS
                    If .Graphic(loopC).GrhIndex Then Put freeFileMap, , .Graphic(loopC).GrhIndex
                Next loopC
                
                ' Analizo el tamaño para ver si debo agrandar el area que el Juego analiza
                If ByFlags And bitwisetable(3) Then
                    If ResizeBackBufferY < GrhData(.Graphic(3).GrhIndex).pixelHeight Then ResizeBackBufferY = GrhData(.Graphic(3).GrhIndex).pixelHeight
                    If ResizeBackBufferX < GrhData(.Graphic(3).GrhIndex).pixelWidth Then ResizeBackBufferX = GrhData(.Graphic(3).GrhIndex).pixelWidth
                End If
                
                '5) Trigger
                If ByFlags And bitwisetable(5) Then Put freeFileMap, , .Trigger
                    
                '9) Altura
                If ByFlags And bitwisetable(9) Then
                    Put freeFileMap, , NormalData(x, y)     'Normal del vértice
                    Put freeFileMap, , Alturas(x, y)        'Alturas
                    Put freeFileMap, , AlturaPie(x, y)      'Altura Pie
                    tit = Round(hMapData(x, y).hs(0))
                    Put freeFileMap, , tit                  'Altura del vértice
                    Put freeFileMap, , .tile_orientation    'Orientación de la tile
                End If
                
                ' 13) Luz
                If ByFlags And bitwisetable(13) Then
                    Put freeFileMap, , tempLuz.LuzColor.b
                    Put freeFileMap, , tempLuz.LuzColor.g
                    Put freeFileMap, , tempLuz.LuzColor.r
                    Put freeFileMap, , tempLuz.LuzRadio
                    Put freeFileMap, , tempLuz.LuzBrillo
                    Put freeFileMap, , tempLuz.LuzTipo
                    Put freeFileMap, , tempLuz.luzInicio
                    Put freeFileMap, , tempLuz.luzFin
                End If
                
                '7)  Particula 0
                If ByFlags And bitwisetable(7) Then .Particles_groups(0).persistir freeFileMap
                '14) Particula 1
                If ByFlags And bitwisetable(14) Then .Particles_groups(1).persistir freeFileMap
                '8) Particula 2
                If ByFlags And bitwisetable(8) Then .Particles_groups(2).persistir freeFileMap
                
                ' CheckSum
                checkSum = "hola"
                Put freeFileMap, , checkSum
                 
            End With
        Next x
    Next y
    
    ' Actualizamos el Offset
    Seek freeFileMap, offsetHeader
    
    Put freeFileMap, , ResizeBackBufferX
    Put freeFileMap, , ResizeBackBufferY
    
    Close freeFileMap
    
    Guardar_Mapa_CLI = True
End Function

