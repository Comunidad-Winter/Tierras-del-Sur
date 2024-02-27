Attribute VB_Name = "Engine_Landscape"
' ESTE ARCHIVO ESTA COMPARTIDO POR TODOS LOS PROGRAMAS.


''
' @require Engine.bas
' @require Engine_Landscape_Water.bas


Option Explicit

Public Type ARGB_COLOR
    a As Byte
    r As Byte
    g As Byte
    b As Byte
End Type


Public Enum TipoLuces
    Luz_Normal = 1
    Luz_Animada = 2
    Luz_Incandecente = 4
    Luz_Fuego = 8
    Luz_Cuadrada = 16
End Enum

Public Type fRGBA_COLOR
    b As Single
    g As Single
    r As Single
    a As Single
End Type

Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal bytes As Long)
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef dest As Any, ByVal numbytes As Long)
Public Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Public Const Perspectiva As Single = 0.65

'LUCES
    Public Intensidad_Del_Terreno(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)      As Byte         'Guarda la intensidad de la luz de un vertice del mapa
    Public OriginalMapColor(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)           As BGRACOLOR_DLL 'Colores precalculados en el mapeditor
    Public OriginalMapColorSombra(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)     As Long         'OriginalMapColor * Sombra
    Public OriginalColorArray(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)          As Long         'BACKUP DE ResultColorArray (OriginalMapColorSombra * AMBIENTE)
    Public ResultColorArray(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)           As Long         'OriginalColorArray * LUCES DINÁMICAS * SOMBRAS
    'Public ResultColorArraySinSombra(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)    As Long        'OriginalColorArray * LUCES DINÁMICAS
'/LUCES

'MONTAÑAS
    'Altura de cada vertice del mapa
    Public hMapData(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)     As AUDT
    ' Altura de donde pisa el pj, o donde flotan las cosas, o donde el árbol vuela. Sirve para ahcer escaleras.
    Public AlturaPie(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)  As Integer
    'Es >0 si en la tile hay una altura distnta a cero.
    Public Alturas(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)    As Integer
    ' Almacena el vector normalizado de los triángulos del mapa, para calcular la sombra Intensidad_sombra = DOT(NORMALIZED•SOL_POS)
    Public NormalData(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)   As D3DVECTOR
    ' Altura del agua, no hay vuelta que darle.
    Public AlturaAgua                               As Integer
    Public TexturaAgua                              As Integer
    
    Public Sombra_Montañas(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE) As Byte
    
    Private PosicionSol                             As D3DVECTOR
    
    Public MapBoxes(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)    As Box_Vertex
'/MONTAÑAS

'FLAGS PARA RECALCULAR
    ' Setear TRUE si se quiere recalcular el (color del mapa * ambiente) * luces
    Public Light_Update_Map     As Boolean
    ' Setear TRUE si se quiere recalcular Color_y_ambiente_cache * luces
    Public Light_Update_Sombras As Boolean
    ' Saber si en el cliente recalcula todo o lo lee cacheado del archivo del mapa
    Public SombrasCacheadas     As Boolean
'/FLAGS PARA RECALCULAR



'COLOR FINAL = LERP[BRILLO_TERRENO,COLOR_PRESET.rgb,(SOMBRA * AMBIENTE.rgb)] * LUCES

Public ForzarRecalculoLuces As Boolean


Public Const LUZ_TIPO_FUEGO As Integer = 999

#If USAR_ENGINE_COM = 1 Then
    Public DLL_Terreno As Terreno
    Public DLL_Luces As LucesManager
#Else
    Public DLL_Luces As Engine_Luces
#End If


Public Sub CalcularNormales()
    Dim x%, y%
    
    Dim TmpVec As mzVECTOR
    
    DLL_Terreno.CalcularNormales VarPtr(NormalData(1, 1)), VarPtr(hMapData(1, 1))
    
    Light_Update_Sombras = True
    
    ColoresAgua_Redraw
    
    NormalesMapaNececitanActualizar = True
    
    'If Engine_General.SombrasHQ Then
        DLL_Terreno.SuavizarNormales VarPtr(NormalData(1, 1))
    'End If
End Sub


Public Sub MoverSol(ByVal x!, ByVal y!, ByVal Altura!)
    PosicionSol.x = x
    PosicionSol.y = y
    PosicionSol.z = Altura
    
    If Engine_General.NoUsarSombras = False Then
        Light_Update_Map = True
        Light_Update_Sombras = True
    End If
End Sub


Public Sub Init_Lights()
    #If USAR_ENGINE_COM = 1 Then
        Set DLL_Luces = New LucesManager
        DLL_Luces.iniciar 500, ResultColorArray(1, 1), ANCHO_MAPA
    #Else
        Set DLL_Luces = New Engine_Luces
    #End If
        Set DLL_Terreno = New Terreno
    
    DLL_Terreno.IniciarColores Intensidad_Del_Terreno(1, 1), VarPtr(OriginalMapColor(1, 1)), OriginalColorArray(1, 1), ResultColorArray(1, 1), OriginalMapColorSombra(1, 1), Sombra_Montañas(1, 1), Alturas(1, 1), ANCHO_MAPA
    
    
    Dim x As Long
    Dim y As Long
    
    For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
        For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
            With MapBoxes(x, y)
                .x0 = x * 32
                .x1 = .x0
                .y1 = y * 32
                .y0 = .y1 + 32
                .x2 = .x0 + 32
                .y2 = .y0
                .y3 = .y1
                .y2 = .y0
                .x3 = .x2
                .rhw0 = 1
                .rhw1 = 1
                .rhw2 = 1
                .rhw3 = 1
                
            End With
        Next
    Next

    color_mod_day_16.r = 128
    color_mod_day_16.g = 128
    color_mod_day_16.b = 128
End Sub

Public Sub Pre_Render_Lights()
    'Funcion exclusiva del editor de mapas
    
    'Recorre la lista de luces y las que no tengan ID las borra y cachea en el color del mapa
    
    Dim i       As Integer
    
    DLL_Luces.Iterador_Iniciar
    
    i = DLL_Luces.Iterar
    Do While i
        Pre_Render_Light i
        i = DLL_Luces.Iterar
    Loop

End Sub

Public Sub Pre_Render_Light(ByVal luz As Integer)

    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    Dim Color As RGBCOLOR
    Dim range As Byte
    Dim brillo As Byte
    Dim tipo As Integer
    Dim map_x As Byte
    Dim map_y As Byte
    
    If DLL_Luces.Get_Light(luz, map_x, map_y, r, g, b, range, brillo, tipo, 1, 1) Then
        If tipo = 0 Or True Then
            Color.r = r
            Color.g = g
            Color.b = b
            Pincelear_Mapa map_x, map_y, Color, range, brillo, 0, 0
            DLL_Luces.Quitar luz
        End If
    End If
    
End Sub

Public Sub Pincelear_Mapa(ByVal MapX As Integer, ByVal MapY As Integer, Color As RGBCOLOR, ByVal radio As Integer, Optional ByVal brillo As Byte = 255, Optional ByVal PixelPosX As Integer = 0, Optional ByVal PixelPosY As Integer = 0, Optional ByVal Intensidad As Byte = 255)

Dim x           As Integer
Dim y           As Integer
Dim mu          As Single
Dim tIntensidad As Integer

Dim rRGB        As RGBCOLOR
Dim RadioPlus   As Integer

Dim Cateto1!, Cateto2!, Hipotenusa!
Dim max_x%, max_y%, min_x%, min_y%

    RadioPlus = radio * 32
    
    min_y = MapY - radio - 1
    min_x = MapX - radio - 1
    max_y = MapY + radio + 1
    max_x = MapX + radio + 1
    
    'MARCE CHEQUEAR IFS
    If max_x > X_MAXIMO_VISIBLE Then
        max_x = X_MAXIMO_VISIBLE
    Else
        If (min_x < X_MINIMO_VISIBLE) Then min_x = X_MINIMO_VISIBLE
    End If
    
    If (min_y < Y_MINIMO_VISIBLE) Then
        min_y = Y_MINIMO_VISIBLE
    Else
        If (max_y > Y_MAXIMO_VISIBLE) Then max_y = Y_MAXIMO_VISIBLE
    End If
    
    If PixelPosX = 0 Then
        PixelPosX = MapX * 32 + 16
        PixelPosY = MapY * 32 + 16
    End If
    For y = min_y To max_y
        For x = min_x To max_x
        
            Cateto1 = PixelPosX - (x * 32)
            Cateto2 = PixelPosY - (y * 32)
            Hipotenusa = Sqr(Cateto1 * Cateto1 + Cateto2 * Cateto2) 'Obtengo la hipotenusa
            
            If (Hipotenusa <= RadioPlus) Then
                'COLOR
                
                mu = (Hipotenusa / RadioPlus) * (brillo / 255)
                
                rRGB.r = Interp(Color.r, OriginalMapColor(x, y).r, mu)
                rRGB.g = Interp(Color.g, OriginalMapColor(x, y).g, mu)
                rRGB.b = Interp(Color.b, OriginalMapColor(x, y).b, mu)
                
                If OriginalMapColor(x, y).r < rRGB.r Then _
                    OriginalMapColor(x, y).r = rRGB.r
                    
                If OriginalMapColor(x, y).g < rRGB.g Then _
                    OriginalMapColor(x, y).g = rRGB.g
                    
                If OriginalMapColor(x, y).b < rRGB.b Then _
                    OriginalMapColor(x, y).b = rRGB.b
                    
                mu = Hipotenusa / RadioPlus
                'BRILLO
                
                tIntensidad = Abs(Interp(Intensidad, Intensidad_Del_Terreno(x, y), mu))
                If tIntensidad > 255 Then tIntensidad = 255
                If tIntensidad > Intensidad_Del_Terreno(x, y) Then _
                    Intensidad_Del_Terreno(x, y) = tIntensidad
                    
                    'CopyMemory ResultColorArray(X, Y), OriginalMapColor(X, Y), 4
                    'ResultColorArray(X, Y) = ResultColorArray(X, Y) Or &HFF000000
            End If
        Next x
    Next y
    Light_Update_Map = True
Call CopyMemory(OriginalMapColorSombra(1, 1), OriginalMapColor(1, 1), TILES_POR_MAPA * 4)

End Sub

Public Sub map_render_light()
    Static LastLightUpdate As Long
    Dim Actualizar_Luces As Boolean
    
    Static AcumuladorTiempo As Single
    

    If ForzarRecalculoLuces Then
        LastLightUpdate = 0
        AcumuladorTiempo = 0
        Light_Update_Map = True
        Light_Update_Sombras = True
        SombrasCacheadas = False
        ForzarRecalculoLuces = False
    End If
    
    
    If LastLightUpdate + 32 < GetTimer Then
        LastLightUpdate = GetTimer
        
        'Marce On local error resume next
        DLL_Luces.PreProcesar AcumuladorTiempo
        'Marce 'Marce On local error goto 0
        AcumuladorTiempo = 0
        
        Actualizar_Luces = DLL_Luces.NeedUpdateLights Or Light_Update_Map
        Light_Update_Map = Light_Update_Map 'Or DLL_Luces.NeedUpdateMap
        
        If Light_Update_Map Then
           'Debug.Print "UPDATIE EL MAPA DE LUCES!"
            ' Multiplico la luz del día por la luz predefinida del terreno.
            DLL_Terreno.Refresh color_mod_day_16.r, color_mod_day_16.g, color_mod_day_16.b
            
            ' Está habilitada la configuración con sombras??
            If Engine_General.NoUsarSombras = False Then
                NormalesMapaNececitanActualizar = True
                Light_Update_Sombras = False
            End If
            
            'recalcular_colores_agua  'Copio los colores de las luces y les agrego el alpha apra el agua.
            
            ' HAgo un backup de resultcolorarray(base+ambiente+sombras) en este mopmento para usarla con las luces
            If (NoUsarLuces = False) Then Call CopyMemory(OriginalColorArray(1, 1), ResultColorArray(1, 1), TILES_POR_MAPA * 4)
            
            
            
            'recalcular_opacidades_agua
            Cachear_Tiles = True
            
        End If
        
        If Actualizar_Luces And (NoUsarLuces = False) Then
            LimpiarMapaDeLuces                                      'Limpio los datos de las luces en el mapa
            'If DLL_Luces.count And (NoUsarLuces = False) Then       'Hay luces?
            DLL_Luces.Actualizar                                    'Render de las luces
            'End If
            
            'Pasar a dll area de vision
                'Que al dll diga si se actualizo algo dentro del area
                'if si poner el flag para actualizar ESE area
                
            #If esMe = 1 Then
                If ME_MiniMap.miniMapaTipo And emmLuces Then MiniMapNeedToBeRedrawed = True
            #End If
            Cachear_Tiles = True
        End If

        Light_Update_Map = False
        
    Else
        AcumuladorTiempo = AcumuladorTiempo + timerElapsedTime
    End If

End Sub

Public Function InMapBounds(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If x < X_MINIMO_VISIBLE Or x > X_MAXIMO_VISIBLE Or y < Y_MINIMO_VISIBLE Or y > Y_MAXIMO_VISIBLE Then
        InMapBounds = False
        Exit Function
    End If
    
    InMapBounds = True
End Function



Private Sub LimpiarMapaDeLuces(): Call CopyMemory(ResultColorArray(1, 1), OriginalColorArray(1, 1), TILES_POR_MAPA * 4): End Sub
