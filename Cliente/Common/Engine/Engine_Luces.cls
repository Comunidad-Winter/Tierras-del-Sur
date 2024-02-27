VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Engine_Luces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Compartido
' Luces encapsuladas en clase de VB para pasar métodos a C

#If USAR_ENGINE_COM = 0 Then

Option Explicit
'Implements ILucesManager 'INTERFACE
Private Type Light
    active          As Byte 'Do we ignore this light?
    
    map_x           As Integer 'Coordinates
    map_y           As Integer
    
    pixel_pos_x     As Integer
    pixel_pos_y     As Integer
    
    rangoplus       As Integer
    
    theta           As Single
    
    Progreso        As Single
    
    Color           As BGRACOLOR_DLL
    
    id              As Long
    start_tick      As Long
    
    brillo          As Byte
    
    range           As Byte
    tipo            As Byte
End Type

'Light list
Private light_list() As Light
Private light_count As Long
Private light_last As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal bytes As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef dest As Any, ByVal numbytes As Long)

Private pLight_Update_Map As Boolean
Private pLight_Update_Lights As Boolean

Private LucesTimer As clsPerformanceTimer

Private DeltaTime As Single

Private Iterador_Index As Integer

Public Sub Iterador_Iniciar()
    Iterador_Index = 0
End Sub

Public Function Iterar() As Integer

    Iterador_Index = Iterador_Index + 1 'Sumo uno al index
    
    If Iterador_Index > light_last Then 'Llegue al final?
        Iterador_Index = 0 'Reinicio el iterador
    Else
        Do Until light_list(Iterador_Index).active <> 0 'Es la luz iterable?
            If Iterador_Index = light_last Then
                Iterador_Index = 0 'Reinicio el iterador
                Exit Do
            End If
            Iterador_Index = Iterador_Index + 1
        Loop
    End If
    
    Iterar = Iterador_Index
End Function

Public Sub Actualizar()
    Dim Index As Integer
    For Index = 1 To light_last
        If Light_Check(Index) Then
            Light_Render Index
        End If
    Next Index
    pLight_Update_Lights = False
End Sub

Public Sub iniciar(cantidad As Long, DestArray As Long): End Sub

Private Sub Class_Initialize()
    ReDim Preserve light_list(1 To 1)
    Set LucesTimer = New clsPerformanceTimer
    LucesTimer.Time
    light_count = 0
    light_last = 0
End Sub

Private Sub Class_Terminate()
    'Marce On local error resume next
    Set LucesTimer = Nothing
    Erase light_list
End Sub

Public Function Count() As Integer
    Count = light_count
End Function


Public Function Find(ByVal id As Long) As Long
On Error GoTo ErrorHandler:
    Dim loopC As Long
    
    loopC = 1
    Do Until light_list(loopC).id = id
        If loopC = light_last Then
            Light_Find = 0
            Exit Function
        End If
        loopC = loopC + 1
    Loop
    
    Light_Find = loopC
Exit Function
ErrorHandler:
    Light_Find = 0
End Function

Public Function Remove_All() As Boolean
    Dim Index As Long

    For Index = 1 To light_last
        'Make sure it's a legal index
        If Light_Check(Index) Then
            Light_Destroy Index
        End If
    Next Index
    
    Remove_All = True
End Function

Public Function Quitar(ByRef light_index As Integer) As Boolean
    If Light_Check(light_index) Then
        Light_Destroy light_index
        pLight_Update_Lights = True
        Quitar = True
    End If
    light_index = 0
End Function

Public Function Get_Color(ByVal light_index As Long, ByRef r As Byte, ByRef g As Byte, ByRef b As Byte) As Boolean
    If Light_Check(light_index) Then
        r = light_list(light_index).Color.r
        g = light_list(light_index).Color.g
        b = light_list(light_index).Color.b
        Light_Color_Value_Get = True
    End If
End Function

Public Function Get_Light(ByVal light_index As Long, map_x As Byte, map_y As Byte, ByRef r As Byte, ByRef g As Byte, ByRef b As Byte, radio As Byte, brillo As Byte, id As Long, tipo As Integer) As Boolean
    If Light_Check(light_index) Then
        With light_list(light_index)
            r = .Color.r
            g = .Color.g
            b = .Color.b
            map_x = .map_x
            map_y = .map_y
            radio = .range
            brillo = .brillo
            id = .id
            tipo = .tipo
            Get_Light = True
        End With
    End If
End Function

Public Function Get_Light_Ext(ByVal light_index As Long, map_x As Byte, map_y As Byte, ByRef r As Byte, ByRef g As Byte, ByRef b As Byte, radio As Byte, brillo As Byte, id As Long, tipo As Integer, ByRef pixel_pos_x As Integer, ByRef pixel_pos_y As Integer) As Boolean
    If Light_Check(light_index) Then
        With light_list(light_index)
            r = .Color.r
            g = .Color.g
            b = .Color.b
            map_x = .map_x
            map_y = .map_y
            radio = .range
            brillo = .brillo
            id = .id
            tipo = .tipo
            pixel_pos_x = .pixel_pos_x
            pixel_pos_y = .pixel_pos_y
            Get_Light = True
        End With
    End If
End Function

Public Function crear(ByVal map_x As Integer, ByVal map_y As Integer, ByVal r As Byte, ByVal g As Byte, ByVal b As Byte, _
                            Optional ByVal range As Byte = 1, Optional ByVal brillo As Byte = 255, Optional ByVal id As Long, Optional ByVal tipo As TipoLuces = Luz_Normal) As Long
    If InMapBounds(map_x, map_y) Then
        crear = Light_Next_Open
        Light_Make crear, map_x, map_y, r, g, b, range, brillo, tipo, id
    End If
End Function

Public Function Move(ByVal light_index As Long, Optional ByVal map_x As Integer, Optional ByVal map_y As Integer, Optional ByVal PixelOffsetX As Integer, Optional ByVal PixelOffsetY As Integer) As Boolean
    If Light_Check(light_index) Then
        'If InMapBounds(map_x, map_y) Then
            With light_list(light_index)

                'Me da la ubicacion en coordenadas de mapa y un offset en pixeles
                If InMapBounds(map_x, map_y) Then
                    light_list(light_index).map_x = map_x
                    light_list(light_index).map_y = map_y
                    light_list(light_index).pixel_pos_x = map_x * 32 + PixelOffsetX + 16
                    light_list(light_index).pixel_pos_y = map_y * 32 + PixelOffsetY - 16
                End If
                
                pLight_Update_Lights = True
                
                Move = True
            End With
        'End If
    End If
End Function

Public Function MovePixel(ByVal light_index As Long, ByVal PixelPosX As Integer, ByVal PixelPosY As Integer) As Boolean
    If Light_Check(light_index) Then
        'If InMapBounds(map_x, map_y) Then
            With light_list(light_index)
                'Me da la ubicacion en pixeles
                light_list(light_index).map_x = PixelPosX / 32
                light_list(light_index).map_y = PixelPosY / 32
                light_list(light_index).pixel_pos_x = PixelPosX + 16
                light_list(light_index).pixel_pos_y = PixelPosY - 16

                
                pLight_Update_Lights = True
                
                MovePixel = True
            End With
        'End If
    End If
End Function

Public Property Get NeedUpdateLights() As Boolean
    NeedUpdateLights = pLight_Update_Lights
End Property

Public Property Get NeedUpdateMap() As Boolean
    NeedUpdateLights = pLight_Update_Map
End Property

Public Sub MapUpdated()
    pLight_Update_Map = False
End Sub

Public Sub LightsUpdated()
    pLight_Update_Lights = False
End Sub


'PRIVADOSSSSS


Private Sub Light_Make(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, ByVal r As Byte, ByVal g As Byte, ByVal b As Byte, _
                        ByVal range As Long, ByVal brillo As Byte, ByVal tipo As TipoLuces, Optional ByVal id As Long)
    If light_index > light_last Then
        light_last = light_index
        ReDim Preserve light_list(1 To light_last)
    End If
    
    'If light_index Then
        light_count = light_count + 1
    'End If
    
    With light_list(light_index)
        .active = 1

        .map_x = map_x
        .map_y = map_y
        .pixel_pos_x = map_x * 32 + 16
        .pixel_pos_y = map_y * 32 - 16

        .range = range
        
        .rangoplus = .range * 32
        
        .id = id
        .Color.r = r
        .Color.g = g
        .Color.b = b
        .Color.a = 255
        .brillo = brillo
        
        .tipo = tipo
        pLight_Update_Lights = True
    End With
    
End Sub

Private Function Light_Check(ByVal light_index As Long) As Boolean
    If light_index > 0 And light_index <= light_last Then
        If light_list(light_index).active Then
            Light_Check = True
        End If
    End If
End Function

Private Function Light_Next_Open() As Long
On Error GoTo ErrorHandler:
    Dim loopC As Long
    
    loopC = 1
    Do Until light_list(loopC).active = 0
        If loopC = light_last Then
            Light_Next_Open = light_last + 1
            Exit Function
        End If
        loopC = loopC + 1
    Loop
    
    Light_Next_Open = loopC
Exit Function
ErrorHandler:
    Light_Next_Open = 1
End Function

Private Sub Light_Destroy(ByVal light_index As Long)
    Dim temp As Light
    

    
    light_list(light_index) = temp
    
    'Update array size
    If light_index = light_last Then
        Do Until light_list(light_last).active
            light_last = light_last - 1
            If light_last = 0 Then
                light_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve light_list(1 To light_last + 1)
    End If
    light_count = light_count - 1
    pLight_Update_Lights = True
End Sub

Private Sub Light_Render(ByVal light_index As Integer)

Dim x           As Integer
Dim y           As Integer
Dim mu          As Single

Dim tl As Long

Dim tColor As BGRACOLOR_DLL

Dim Cateto1!, Cateto2!, Hipotenusa!
Dim max_x%, max_y%, min_x%, min_y%

With light_list(light_index)
    min_y = .map_y - .range - 1
    min_x = .map_x - .range - 1
    max_y = .map_y + .range + 1
    max_x = .map_x + .range + 1
    
    If (max_x > X_MAXIMO_VISIBLE) Then
        max_x = X_MAXIMO_VISIBLE
    ElseIf (min_x < X_MINIMO_VISIBLE) Then
        min_x = X_MINIMO_VISIBLE
    End If
    
    If (min_y < Y_MINIMO_VISIBLE) Then
        min_y = Y_MINIMO_VISIBLE
    ElseIf (max_y > Y_MAXIMO_VISIBLE) Then
        max_y = Y_MAXIMO_VISIBLE
    End If

    For y = min_y To max_y
        For x = min_x To max_x
            Cateto1 = .pixel_pos_x - (x * 32)
            Cateto2 = .pixel_pos_y - (y * 32)
            Hipotenusa = Sqr(Cateto1 * Cateto1 + Cateto2 * Cateto2) 'Obtengo la hipotenusa
            
            If (Hipotenusa <= .rangoplus) Then
                'COLOR
                mu = Interp((.brillo / 255), 1, (Hipotenusa / .rangoplus))
                DXCopyMemory tColor, ResultColorArray(x, y), 4
                tColor.r = Interp(.Color.r, tColor.r, mu)
                tColor.g = Interp(.Color.g, tColor.g, mu)
                tColor.b = Interp(.Color.b, tColor.b, mu)
                DXCopyMemory ResultColorArray(x, y), tColor, 4
            End If
        Next x
    Next y
End With
End Sub




#End If
