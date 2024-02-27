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

Public Type fRGBA_COLOR
    b As Single
    g As Single
    r As Single
    a As Single
End Type

Public Enum TipoLuces
    Luz_Normal = 1
    Luz_Animada = 2
    Luz_Incandecente = 4
    Luz_Fuego = 8
End Enum

Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Sub ElevarTerreno Lib "MZEngine.dll" (ByVal XCord As Long, ByVal YCord As Long, ByRef altura As Single, ByVal map_x As Long, ByVal map_y As Long, ByVal theta As Single, ByVal range As Long, ByVal haltura As Byte)

Private Declare Sub CalcularNormal Lib "MZEngine.dll" (ByVal A0 As Single, ByVal A1 As Single, ByVal A2 As Single, ByVal A3 As Single, ByRef C1 As D3DVECTOR)
Private Declare Sub SetSunPos Lib "MZEngine.dll" (ByRef pos As D3DVECTOR)
Private Declare Sub CalcularSombraN Lib "MZEngine.dll" (ByRef n1 As D3DVECTOR, ByRef N2 As D3DVECTOR, ByRef C1 As Single, ByRef C2 As Single, ByRef C3 As Single)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Bytes As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef dest As Any, ByVal numbytes As Long)
Public Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long

'//void _stdcall CalcularSombra(DWORD* cOut, const NVEC* Normal, const D3DCOLORVALUE* Color_Terreno){
Private Declare Sub CalcularSombra Lib "MZEngine.dll" (ByRef cOut As Any, ByRef cIn As Any, ByRef normal As D3DVECTOR, ByRef sunpos As D3DVECTOR, ByRef altura As Integer)

Public Const Perspectiva As Single = 0.65

Public Type Alturas_udt
    plus(3) As Single
    alt As Single
End Type
    
Private Declare Sub Map_Compute_Lights Lib "MZEngine.dll" (ByRef DestArray As Long, ByRef Map_Color As Any, ByRef hlmd As Byte, ByRef color As BGRCOLOR_DLL)

'LUCES
    Public Intensidad_Del_Terreno(1 To 100, 1 To 100)   As Byte         'Guarda la intensidad de la luz de un vertice del mapa
    Public OriginalMapColor(1 To 100, 1 To 100)         As BGRACOLOR_DLL 'Colores precalculados en el mapeditor
    Public OriginalMapColorSombra(1 To 100, 1 To 100)   As Long         'OriginalMapColor * Sombra
    Public OriginalColorArray(1 To 100, 1 To 100)       As Long         'BACKUP DE ResultColorArray (OriginalMapColorSombra * AMBIENTE)
    Public ResultColorArray(1 To 100, 1 To 100)         As Long         'OriginalColorArray * LUCES DINÁMICAS
    Public ResultColorArraySinSombra(1 To 100, 1 To 100) As Long        'OriginalColorArray * LUCES DINÁMICAS

'/LUCES

'MONTAÑAS
    'Altura de cada vertice del mapa
    Public hMapData(1 To 100, 1 To 100)     As Alturas_udt
    ' Altura de donde pisa el pj, o donde flotan las cosas, o donde el árbol vuela. Sirve para ahcer escaleras.
    Public AlturaPie(1 To 100, 1 To 100)    As Integer
    'Es >0 si en la tile hay una altura distnta a cero.
    Public Alturas(1 To 100, 1 To 100)      As Integer
    ' Almacena el vector normalizado de los triángulos del mapa, para calcular la sombra Intensidad_sombra = DOT(NORMALIZED•SOL_POS)
    Public NormalData(1 To 100, 1 To 100)   As D3DVECTOR
    ' Altura del agua, no hay vuelta que darle.
    Public AlturaAgua                       As Integer
    Public TexturaAgua                      As Integer
    
    Public Sombra_Montañas(1 To 100, 1 To 100) As Byte
    
'/MONTAÑAS

'FLAGS PARA RECALCULAR
    ' Setear TRUE si se quiere recalcular el (color del mapa * ambiente) * luces
    Public Light_Update_Map                     As Boolean
    ' Setear TRUE si se quiere recalcular Color_y_ambiente_cache * luces
    Public Light_Update_Lights                  As Boolean
    Public Light_Update_Sombras                 As Boolean
'/FLAGS PARA RECALCULAR



'COLOR FINAL = LERP[BRILLO_TERRENO,COLOR_PRESET.rgb,(SOMBRA * AMBIENTE.rgb)] * LUCES


Private Declare Sub lerp2 Lib "MZEngine.dll" (ByRef cOut As Any, ByRef cIn1 As Any, ByRef cIn2 As Any, ByVal Mu As Byte)


Public DLL_Terreno As New Terreno
Public DLL_Luces As New LucesManager


Public Sub Init_Lights()
    Set DLL_Terreno = New Terreno
    Set DLL_Luces = New LucesManager
    DLL_Terreno.IniciarColores Intensidad_Del_Terreno(1, 1), VarPtr(OriginalMapColor(1, 1)), OriginalColorArray(1, 1), ResultColorArray(1, 1), OriginalMapColorSombra(1, 1)
    DLL_Luces.Iniciar 1000, VarPtr(ResultColorArray(1, 1))
End Sub

Public Sub Pre_Render_Lights()
'    'Funcion exclusiva del editor de mapas
'
'    'Recorre la lista de luces y las que no tengan ID las borra y cachea en el color del mapa
'
'    Dim i       As Integer
'
'    'ZeroMemory OriginalMapColor(1, 1), 40000
'
'    If light_count Then
'        For i = 1 To light_last
'            Pre_Render_Light i
'        Next i
'    End If

'TODO LUCES

End Sub

Public Sub Pre_Render_Light(ByVal i As Integer)
'
'    Dim color As RGBCOLOR
'
'    With light_list(i)
'        If .ID = 0 Then
'            If .active Then
'                color.r = .color.r
'                color.g = .color.g
'                color.b = .color.b
'                Pincelear_Mapa .map_x, .map_y, color, .range, .brillo, .pixel_pos_x, .pixel_pos_y
'                .active = 0
'            End If
'        End If
'    End With
'
'    Light_Update_Map = True

'TODO LUCES
End Sub


Private Sub Light_Render(ByVal light_index As Integer)
'
'Dim x           As Integer
'Dim Y           As Integer
'Dim Mu          As Single
'
'Dim tl As Long
'
'Dim tColor As BGRACOLOR_DLL
'
'Dim Cateto1!, Cateto2!, Hipotenusa!
'Dim max_x%, max_y%, min_x%, min_y%
'
'With light_list(light_index)
'    min_y = .map_y - .range - 1
'    min_x = .map_x - .range - 1
'    max_y = .map_y + .range + 1
'    max_x = .map_x + .range + 1
'
'    If (max_x > 99) Then
'        max_x = 99
'    Else
'        If (min_x < 2) Then min_x = 2
'    End If
'
'    If (min_y < 2) Then
'        min_y = 2
'    Else
'        If (max_y > 99) Then max_y = 99
'    End If
'
'    For Y = min_y To max_y
'        For x = min_x To max_x
'            Cateto1 = .pixel_pos_x - (x * 32)
'            Cateto2 = .pixel_pos_y - (Y * 32)
'            Hipotenusa = Sqr(Cateto1 * Cateto1 + Cateto2 * Cateto2) 'Obtengo la hipotenusa
'
'            If (Hipotenusa <= .rangoplus) Then
'                'COLOR
'                Mu = (Hipotenusa / .rangoplus) * (.brillo / 255)
'                DXCopyMemory tColor, ResultColorArray(x, Y), 4
'                tColor.r = Interp(.color.r, tColor.r, Mu)
'                tColor.g = Interp(.color.g, tColor.g, Mu)
'                tColor.b = Interp(.color.b, tColor.b, Mu)
'                DXCopyMemory ResultColorArray(x, Y), tColor, 4
'            End If
'        Next x
'    Next Y
'End With
End Sub


Public Sub Pincelear_Mapa(ByVal mapX As Integer, ByVal mapY As Integer, color As RGBCOLOR, ByVal radio As Integer, Optional ByVal brillo As Byte = 255, Optional ByVal PixelPosX As Integer = 0, Optional ByVal PixelPosY As Integer = 0, Optional ByVal Intensidad As Byte = 255)

Dim x           As Integer
Dim Y           As Integer
Dim Mu          As Single
Dim tIntensidad As Integer

Dim rRGB        As RGBCOLOR
Dim RadioPlus   As Integer

Dim Cateto1!, Cateto2!, Hipotenusa!
Dim max_x%, max_y%, min_x%, min_y%

    RadioPlus = radio * 32
    
    min_y = mapY - radio - 1
    min_x = mapX - radio - 1
    max_y = mapY + radio + 1
    max_x = mapX + radio + 1
    
    If (max_x > 99) Then
        max_x = 99
    Else
        If (min_x < 2) Then min_x = 2
    End If
    
    If (min_y < 2) Then
        min_y = 2
    Else
        If (max_y > 99) Then max_y = 99
    End If
    
    If PixelPosX = 0 Then
        PixelPosX = mapX * 32 + 16
        PixelPosY = mapY * 32 + 16
    End If
    For Y = min_y To max_y
        For x = min_x To max_x
        
            Cateto1 = PixelPosX - (x * 32)
            Cateto2 = PixelPosY - (Y * 32)
            Hipotenusa = Sqr(Cateto1 * Cateto1 + Cateto2 * Cateto2) 'Obtengo la hipotenusa
            
            If (Hipotenusa <= RadioPlus) Then
                'COLOR
                
                Mu = (Hipotenusa / RadioPlus) * (brillo / 255)
                
                rRGB.r = Interp(color.r, OriginalMapColor(x, Y).r, Mu)
                rRGB.g = Interp(color.g, OriginalMapColor(x, Y).g, Mu)
                rRGB.b = Interp(color.b, OriginalMapColor(x, Y).b, Mu)
                
                If OriginalMapColor(x, Y).r < rRGB.r Then _
                    OriginalMapColor(x, Y).r = rRGB.r
                    
                If OriginalMapColor(x, Y).g < rRGB.g Then _
                    OriginalMapColor(x, Y).g = rRGB.g
                    
                If OriginalMapColor(x, Y).b < rRGB.b Then _
                    OriginalMapColor(x, Y).b = rRGB.b
                    
                Mu = Hipotenusa / RadioPlus
                'BRILLO
                
                tIntensidad = Abs(Interp(Intensidad, Intensidad_Del_Terreno(x, Y), Mu))
                If tIntensidad > 255 Then tIntensidad = 255
                If tIntensidad > Intensidad_Del_Terreno(x, Y) Then _
                    Intensidad_Del_Terreno(x, Y) = tIntensidad
            End If
        Next x
    Next Y
Call CopyMemory(OriginalMapColorSombra(1, 1), OriginalMapColor(1, 1), 40000)

End Sub


Private Sub CalcularSombraVB(ByRef tv As D3DVECTOR)
Dim x   As Integer
Dim Y   As Integer


Dim tl As Long
Dim ts As Single

Dim d As Single


ts = Sqr(tv.x * tv.x + tv.Y * tv.Y + tv.z * tv.z)

If ts = 0 Then ts = 1


tv.x = tv.x / ts
tv.Y = tv.Y / ts
tv.z = tv.z / ts


D3DXVec3Normalize tv, tv

For x = 1 To 100
    For Y = 1 To 100
        ts = D3DXVec3Dot(NormalData(x, Y), tv)
        If ts < 0 Then
            #If COMO_A_MENDUZ_LE_GUSTA Then
            ts = 0 'ts * -128
            #Else
            ts = ts * -32
            #End If
        Else
            ts = ts * 128
        End If
        Sombra_Montañas(x, Y) = ts
    Next Y
Next x

For x = 2 To 99
    For Y = 2 To 99
        d = Sombra_Montañas(x, Y)
        d = d + Sombra_Montañas(x - 1, Y - 1)
        d = d + Sombra_Montañas(x - 1, Y)
        d = d + Sombra_Montañas(x - 1, Y + 1)
        d = d + Sombra_Montañas(x - 1, Y)
        d = d + Sombra_Montañas(x + 1, Y)
        
        d = d + Sombra_Montañas(x - 1, Y - 1)
        d = d + Sombra_Montañas(x + 1, Y - 1)
        d = d + Sombra_Montañas(x, Y + 1)
        
        d = d / 9
        Sombra_Montañas(x, Y) = d
    Next Y
Next x

End Sub

Public Sub Heightmap_Calculate(Optional ByVal sunpos_x As Single = 5, Optional ByVal sunpos_y As Single = 5, Optional ByVal sunpos_z As Single = 5)
    Dim sun As D3DVECTOR
    sun.x = sunpos_x
    sun.Y = sunpos_y
    sun.z = sunpos_z

'NUEVAAA

If Light_Update_Sombras Then
    CalcularSombraVB sun
    Light_Update_Sombras = False
End If

Dim x%, Y%, tl&
For x = 1 To 100
    For Y = 1 To 100
        tl = ResultColorArray(x, Y)
        If hMapData(x, Y).alt > 0 Then
            lerp2 ResultColorArray(x, Y), tl, CLng(&HFF000000), Abs(Sombra_Montañas(x, Y) / hMapData(x, Y).alt)
        Else
            lerp2 ResultColorArray(x, Y), tl, CLng(&HFF000000), Sombra_Montañas(x, Y)
        End If
    Next Y
Next x


'/NUEVAAA

    'CalcularSombra OriginalMapColorSombra(1, 1), OriginalMapColor(1, 1), NormalData(1, 1), sun, Alturas(1, 1)
    'CalcularSombra ResultColorArray(1, 1), OriginalColorArray(1, 1), NormalData(1, 1), sun, Alturas(1, 1)
    'Call CopyMemory(OriginalMapColorSombra(1, 1), OriginalMapColor(1, 1), 40000)
'    Light_Update_Map = True
End Sub

Public Sub map_render_light()
    Static LastLightUpdate As Long
    If LastLightUpdate + 32 < GetTickCount Then
        Lights_Update
        LastLightUpdate = GetTickCount
    
        Light_Update_Lights = Light_Update_Lights Or DLL_Luces.NeedUpdateLights
    
        If Light_Update_Lights Or Light_Update_Map Then
        
            If Light_Update_Map Then
            
                DLL_Terreno.Refresh color_mod_day_16.r, color_mod_day_16.g, color_mod_day_16.b
            
                Call CopyMemory(ResultColorArraySinSombra(1, 1), ResultColorArray(1, 1), 40000)
                Heightmap_Calculate frmMain.XX.Value, frmMain.YY.Value, sunposa.Y
                
                Call CopyMemory(OriginalColorArray(1, 1), ResultColorArray(1, 1), 40000)
                recalcular_opacidades_agua
                
                Light_Update_Lights = True
            End If
            
            If Light_Update_Lights Then
                Call CopyMemory(ResultColorArray(1, 1), OriginalColorArray(1, 1), 40000)
                
                If DLL_Luces.count() > 0 Then
                    DLL_Luces.Actualizar
                End If
                
                recalcular_colores_agua
            End If
    
            DLL_Luces.LightsUpdated
    
            Light_Update_Lights = False
            Light_Update_Map = False
            copy_tile_now = 128 '?FIXED
        End If
    End If
End Sub

Public Function Light_Remove(ByRef light_index As Integer) As Boolean
Call DLL_Luces.Quitar(light_index)
End Function

Public Function Light_Color_Value_Get(ByVal light_index As Long, ByRef color_value As RGBCOLOR) As Boolean
'Esto se usa? :S
'TODO LUCES
End Function
Public Function Light_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByVal r As Byte, ByVal g As Byte, ByVal b As Byte, Optional ByVal range As Byte = 1, Optional ByVal brillo As Byte = 255, Optional ByVal ID As Long, Optional ByVal tipo As TipoLuces = Luz_Normal) As Long
    Light_Create = DLL_Luces.Crear(map_x, map_y, r, g, b, range, brillo, ID, tipo)
End Function

Public Function Light_Move(ByVal light_index As Long, Optional ByVal map_x As Integer, Optional ByVal map_y As Integer, Optional ByVal PixelOffsetX As Integer, Optional ByVal PixelOffsetY As Integer, Optional ByVal PixelPosX As Integer, Optional ByVal PixelPosY As Integer) As Boolean
    Call DLL_Luces.Move(light_index, map_x, map_y, PixelOffsetX, PixelOffsetY)
End Function

Public Function Light_MovePixel(ByVal light_index As Long, ByVal PixelPosX As Integer, ByVal PixelPosY As Integer) As Boolean
    Call DLL_Luces.MovePixel(light_index, PixelPosX, PixelPosY)
End Function

Public Function Light_Toggle(ByVal light_index As Long, ByVal active As Byte) As Boolean
'    If Light_Check(light_index) Then
'        light_list(light_index).active = active
'        Light_Toggle = True
'    End If
'TODO LUCES
End Function

Public Function Light_Find(ByVal ID As Long) As Long
Light_Find = DLL_Luces.Find(ID)
End Function

Public Function Light_Remove_All() As Boolean
DLL_Luces.Remove_All
End Function



Public Sub Lights_Update()
'TODO LUCES
'    Dim Index As Long
'    Dim Tbyte As Integer
'    For Index = 1 To light_last
'        With light_list(Index)
'            If .ID = LUZ_TIPO_FUEGO Then
'
'                If MapData(.map_x, .map_y).luz = Index Then
'                    If MapData(.map_x, .map_y).ObjGrh.GrhIndex <> Fogata Then
'                        .active = False
'                    End If
'                End If
'
'                If .theta < 5 Then
'                    .theta = .theta + timerElapsedTime * 0.01
'                Else
'                    Tbyte = (.rf + timerElapsedTime) Mod 255
'                    .rf = Tbyte
'                    .theta = 5 + Abs(Seno(Tbyte * (360 / 255))) * 0.5
'                End If
'
'                .color.g = 200 - Rnd * 50
'
'                .color.r = 255 - Rnd * 40
'                .range = .theta
'                .rangoplus = .theta * 32
'
'                If Sqr((UserPos.X - .map_y) * (UserPos.X - .map_y) + (UserPos.Y - .map_x) * (UserPos.Y - .map_x)) < (15 + .theta) Then
'                    Light_Update_Lights = True
'                End If
'            End If
'        End With
'    Next Index
End Sub



