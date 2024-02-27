Attribute VB_Name = "Engine_Entidades"
Option Explicit

'Entidades:

    'Pueden ser INDEXADAS o DINÁMICAS. Ambas comparten la función de creación.
    'Solo que las indexadas se crean con otra que lee los parámetros y los pasa como argumentos

    'Tienen efecto cuando mueren, pueden crear otras entidades INDEXADAS, NO DINÁMICAS.
    
    'Puede programarse su muerte a cierto tick. Tienen también una vida maxima y actual asignada.
    
    'Pueden tardar en morir. (luces) para no desaparecer de golpe.
    
    'Son capaces de:
    '-Reproducir sonidos.
    '-Moverse de forma lineal en cualquier dirección.
    '-Crear/matar partículas.
    '-Crear/matar luces.
    '-Disponer de varios gráficos.
    
    'Con la función Entidad_Interactuar se reproduce el sonido correspondiente y se crea la partícula correspondiente
Public Enum Orientacion
    Vertical = 0
    Horizontal = 1
End Enum


Public Enum eTipoEntidadVida
    Nulo = 0
    puntos = 1
    tiempo = 2
End Enum

Public Type udtEntidades
    active          As Byte     '1 Si la entidad esta viva. 0 en caso contrario.
    id              As Long     'TODO Implementar un HASH?
       
    'Una entidad puede tener M cantidad de sonidos
    Grafico         As Grh ' El grafico actual que debe mostrar
    Graficos()      As Integer
    GraficosCount   As Byte ' Cantidad de graficos que tiene la entidades
    GraficoActual   As Byte ' Grafico actual
    GraficoAnterior As Byte ' El ultimo grafico mostrad
    
    'Una entidad puede tener M cantidad de sonidos
    Sonidos()       As Integer 'Lista de sonidos posibles
    SonidosCount    As Byte
    SonidoAnterior  As Byte
    
    streamSonido    As Long
    
    'Una entidad puede tener N cantidad de particulas
    Particulas()    As Integer
    ParticulasCreadas() As Engine_Particle_Group
    ParticulasCount As Byte
    ParticulaAnterior As Byte
    
    'Una entidad solo tiene una luz
    Brillo_radio    As Byte
    Luz_radio       As Byte
    Luz_color       As RGBCOLOR
    Luz_ID          As Integer  ' Para moverla y matarla
    Luz_tipo        As Integer
    Luz_inicio      As Byte
    Luz_fin         As Byte
    
    Velocidad       As Single   ' Desplazamiento 'TODO Se inicia pero no se usa.
    Progreso        As Single   ' 0...1
    
    NaceEnTick      As Long     ' TimeGetTime
    MuereEnTick     As Long     ' Getticckcount
    VidaTick        As Long     ' Cnatidad de milisegundos de vida en total que tiene la entidad
    
    destino         As D3DVECTOR2 ' Pixeles en mapa 'Posicion a donde debe llegar la entidad
    origen          As D3DVECTOR2 ' Pixeles en mapa. Posicion de origen desde donde se mueve
    MPPos           As D3DVECTOR2 ' Pixeles
    map_x           As Byte
    map_y           As Byte
    
    VidaTotal       As Long
    VidaActual      As Long
    VidaMu          As Single   ' == VidaActual / VidaTotal TODO. NO SE UTILIZA
    
    EntidadCuandoMuere  As Long ' Entidad que crea al morir en este mismo slot 'TODO Nunca se utiliza
    
    Char            As Integer  ' Está vinculado a algún char? (para seguir movimientos) (ver $proyectil)
        
    Orientacion     As Orientacion 'TODO No se esta usando
        
    offsetX         As Single 'TODO. Nunca se iniciliza / modifica
    offsetY         As Single 'TODO. Nunca se inicaliza / modifica
    
    Angulo          As Single 'El angulo que forma la entidad moviendose.
    
    Proyectil       As Byte 'Si es un proyectil la entidad se mueve desde $origen hasta $destino o $char
    
    Altura          As Integer 'TODO No se utiliza
    
    MoverLuz        As Byte 'Auxiliar que indica si la entidad se movio y por lo tanto hace falta mover la luz
                            'TODO Se podria utilizar una variable local en la funcion correspondiente
                            'Esta variable solo se utiliza en una funcion
                            
                            
    Next            As Integer 'Lista enlazada de entidades. Siguiente.
    prev            As Integer 'Lista doblemente enlazada de entidades. Anterior
    
    tieneQueMorir   As Boolean 'TODO No se utiliza.
    
    #If esMe = 1 Then
        accion As iAccionEditor
        numeroIndexadoEntidad As Integer
    #End If
End Type


Public Entidades()     As udtEntidades 'Slots para almacenar la informacion de las entidades que fueron agregadas al mapa
Private EntidadesMax    As Integer 'Maxima cantidad de entidades que puede haber en el mapa
Private EntidadesCount  As Integer 'Cantidad de entidades activas
Public EntidadesLast   As Integer 'Ultimo index de entidades que se utilizo


Public EntidadesMap(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)   As Integer ' Cada integer es el root de un Linked list.

'-------------------------------------------------------------------
'Private EntIndexadasCount   As Integer 'Cantidad de entidades indexadas TODO. No se usa mas que en el procedimiento correspondiente
' Private EntIndexadasLast   As Integer 'Ultimo index de entidades indexadas valido
'-------------------------------------------------------------------


Public Sub Entidades_Iniciar(ByVal maximo As Integer)
    EntidadesMax = maximo
    EntidadesCount = 0
    EntidadesLast = 0
    
    ReDim Entidades(1 To EntidadesMax)
End Sub


'Borra todas las entidades del mapa
Public Sub Entidades_ReIniciar()
    Dim i As Integer, j As Integer
    If EntidadesCount Then
        For i = 1 To EntidadesLast
            'Limpiar entidad. Para borrar luces y partículas que queden flotando
            With Entidades(i)
                If .Luz_ID Then DLL_Luces.Quitar .Luz_ID
                If .ParticulasCount Then
                    For j = 0 To .ParticulasCount
                        Set .ParticulasCreadas(j) = Nothing
                    Next j
                End If
                If .Char Then CharList(.Char).entidad = 0
            End With
        Next i
    End If
    
    For i = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
        For j = Y_MAXIMO_VISIBLE To Y_MAXIMO_VISIBLE
            EntidadesMap(i, j) = 0
        Next j
    Next i
    
    'Borro la info de las entidades
    ReDim Entidades(1 To EntidadesMax)

    EntidadesCount = 0
    EntidadesLast = 0
End Sub


' NumeroEntidad, identificador de la entidad indexada

Public Function Entidades_Crear_Indexada(ByVal map_x As Byte, ByVal map_y As Byte, ByVal id As Integer, entidadBase As tIndiceEntidad) As Integer
    Dim i As Integer
    
    Entidades_Crear_Indexada = -1
       
    Entidades_Crear_Indexada = Entidades_Agregar(id)
    
    If Entidades_Crear_Indexada = -1 Then Exit Function

    With Entidades(Entidades_Crear_Indexada)
            
            .map_x = map_x
            .map_y = map_y
            
            .Proyectil = entidadBase.Proyectil
            
            
            ' Graficos
            .GraficoAnterior = 0
    
            .GraficosCount = UBound(entidadBase.Graficos) + 1
                                   
            If Not (.GraficosCount = 1 And entidadBase.Graficos(0) = 0) Then
                ReDim .Graficos(.GraficosCount)
                DXCopyMemory .Graficos(0), entidadBase.Graficos(0), 2& * .GraficosCount
                 .GraficoActual = 0
                .GraficoAnterior = 0
            Else
                .GraficosCount = 0
            End If
            
            'Copiamos los arrays de sonidos
            .SonidosCount = UBound(entidadBase.Sonidos) + 1
            
            If Not (.SonidosCount = 1 And entidadBase.Sonidos(0) = 0) Then
                ReDim .Sonidos(.SonidosCount)
            
                If .SonidosCount Then
                    DXCopyMemory .Sonidos(0), entidadBase.Sonidos(0), 2& * .SonidosCount
                    .SonidoAnterior = 255

                    'TODO: si esta en el rango visible
                    'If .Sonidos(0) > 0 Then
                    '   Sonido_Play .Sonidos(0)
                    '   .streamSonido = 0
                    'Else
                    '   streamSonido = Sonido_PlayEX(-.Sonidos(0), True)
                    'end If
                End If
            Else
                .SonidosCount = 0
            End If
                       
            'Copiamos los arrays de partículas
            .ParticulasCount = UBound(entidadBase.Particulas) + 1
            .ParticulaAnterior = 0
            
            If Not (.ParticulasCount = 1 And entidadBase.Particulas(0) = 0) Then
                ReDim .Particulas(.ParticulasCount)
                ReDim .ParticulasCreadas(.ParticulasCount)
            
                If .ParticulasCount Then
                    DXCopyMemory .Particulas(0), entidadBase.Particulas(0), 2& * .ParticulasCount
                End If
            Else
                .ParticulasCount = 0
            End If
                    
            ' Inicializa los graficos
            If .GraficosCount Then InitGrh .Grafico, .Graficos(0)
                
            'Agrega al mapa
            If .map_x Then
                Entidad_AgregarAMapa Entidades_Crear_Indexada, .map_x, .map_y
                
                .MPPos.x = .map_x * 32
                .MPPos.y = .map_y * 32
                .origen = .MPPos
            End If
        
    
            'Inicializa
            .Angulo = 0
        
            .Luz_radio = entidadBase.luz.LuzRadio
            .Luz_color = entidadBase.luz.LuzColor
            .Brillo_radio = entidadBase.luz.LuzBrillo
            .Luz_tipo = entidadBase.luz.LuzTipo
            .Luz_inicio = entidadBase.luz.luzInicio
            .Luz_fin = entidadBase.luz.luzFin
            
            .Luz_ID = DLL_Luces.crear(map_x, map_y, .Luz_color.r, .Luz_color.g, .Luz_color.b, .Luz_radio, .Brillo_radio, .Luz_tipo, .Luz_inicio, .Luz_fin)
        
            .NaceEnTick = GetTimer
        
            If entidadBase.tipo = eTipoEntidadVida.tiempo Then
                .MuereEnTick = .NaceEnTick + entidadBase.Vida
            Else
                .VidaTotal = entidadBase.Vida
                .VidaActual = .VidaTotal
            End If
                
            'Crea las particullas
            If .ParticulasCount Then
                For i = 0 To .ParticulasCount - 1
                'i = 0
                    If .Particulas(i) Then
                
                        'Call Engine_Particles.Particle_Group_Create(.ParticulasCreadas(i), .Particulas(i))
                        'Engine_Particles.Particle_Group_Set_MPos .ParticulasCreadas(i), .map_x, .map_y
                        Set .ParticulasCreadas(i) = New Engine_Particle_Group
                        Engine_Particles_Storage.IniciarGrupoParticulas .Particulas(i), .ParticulasCreadas(i)
                        .ParticulaAnterior = 0
                    End If
                Next i
            End If
        
            .Progreso = 0
    End With

End Function

'Retorna un slot para establecer la entidad
Private Function Entidades_Agregar(ByVal id As Long) As Integer
    Entidades_Agregar = Entidades_ObtenerLibre
    
    If Entidades_Agregar <> -1 Then
        With Entidades(Entidades_Agregar)
            .active = 1
            .id = id
        End With
        If Entidades_Agregar > EntidadesLast Then EntidadesLast = Entidades_Agregar
        EntidadesCount = EntidadesCount + 1
    End If
End Function

'Obtiene un slot libre para guardar una entidad
Private Function Entidades_ObtenerLibre() As Integer
    Dim i As Integer
    
    If EntidadesCount < EntidadesMax Then 'nos aseguramos de que haya espacio
        For i = 1 To EntidadesMax
            If Entidades(i).active = 0 Then
                Entidades_ObtenerLibre = i
                Exit Function
            End If
        Next i
    End If
    
    Entidades_ObtenerLibre = -1
End Function

'Remueve la luz, las particulas, para el sonido
'TODO es una funcion, pero no retorna ningun valor
Private Function Entidad_Limpiar(ByVal Index As Integer)
    Dim i As Integer

    If EntidadesCount Then
        If Index <= EntidadesLast Then
            With Entidades(Index)
                If .Luz_ID Then DLL_Luces.Quitar .Luz_ID
                .Luz_ID = 0
                If .ParticulasCount Then
                    For i = 0 To .ParticulasCount - 1
                        If Not .ParticulasCreadas(i) Is Nothing Then
                            'Engine_Particles.Particle_Group_Erase .ParticulasCreadas(i)
                            Set .ParticulasCreadas(i) = Nothing
                        End If
                    Next i
                End If
                If .streamSonido Then BASS_ChannelStop .streamSonido
            End With
            
            Entidad_QuitarDeMapa Index
        End If
    End If
End Function

'Desactiva una entidad y la elimina
Private Function Entidades_Remover(ByRef Index As Integer) As Boolean 'True cuando se cambia el active de verdadero a falso
    Dim i As Integer
    Dim j As Integer
    
    If EntidadesCount Then
        If Index <= EntidadesLast Then
            Entidades_Remover = Entidades(Index).active
            Entidades(Index).active = 0
            
            Entidad_Limpiar Index
            EntidadesCount = EntidadesCount - 1

            If Index = EntidadesLast Then
                For i = EntidadesLast To 0 Step -1
                    If Entidades(i).active Then
                        EntidadesLast = i
                        Exit For
                    End If
                Next i
                If EntidadesLast = Index Then EntidadesLast = 0
            End If
        End If
    End If
End Function

Public Function Entidades_Buscar(ByVal id As Long) As Integer

    Dim i As Integer
    If EntidadesCount Then
        For i = 1 To EntidadesLast
            If Entidades(i).id = id Then
                Entidades_Buscar = i
                Exit Function
            End If
        Next i
    End If
    
    Entidades_Buscar = -1
End Function

'TODO Es una funcion pero no devuelve nada
Private Function Entidad_AgregarAMapa(ByVal entidad As Integer, ByVal MapX As Byte, ByVal MapY As Byte)
    Dim siguiente As Integer ' Actualizo el linked list. y lo agrego al EntidadesMap(xx,yy) para renderizsarloasd

    If EntidadesCount > 0 And entidad > 0 Then
        If entidad <= EntidadesLast Then
            siguiente = EntidadesMap(MapX, MapY)
            If siguiente Then
                Entidades(siguiente).prev = entidad
                Entidades(entidad).Next = siguiente
            End If
            EntidadesMap(MapX, MapY) = entidad
        End If
    End If

End Function

'TODO Es una funcion, pero no devuelve nada
Private Function Entidad_QuitarDeMapa(ByVal entidad As Integer) ' Quito un elemento del linked list.

    If EntidadesCount > 0 And entidad > 0 Then
        If entidad <= EntidadesLast Then
            With Entidades(entidad)
                If EntidadesMap(.map_x, .map_y) = entidad Then ' si es el root.
                    If .Next Then ' Tiene hijos ?
                        Entidades(.Next).prev = 0
                        EntidadesMap(.map_x, .map_y) = .Next
                        .Next = 0
                    Else
                        EntidadesMap(.map_x, .map_y) = 0
                    End If
                ElseIf .prev > 0 And .Next > 0 Then 'esta en el diome.
                    Entidades(.prev).Next = .Next
                    Entidades(.Next).prev = .prev
                    .prev = 0
                    .Next = 0
                ElseIf .Next = 0 And .prev > 0 Then ' está al final de la linked list.
                    Entidades(.prev).Next = 0
                End If
                
            End With
            EntidadesCount = EntidadesCount - 1
        End If
    End If
End Function

'TODO Es una funcion, pero no devuelve nada
Public Function Entidades_Actualizar()
    Dim i As Integer, j As Integer
    Dim tick As Long
    tick = GetTimer
    
    If EntidadesCount Then
        For i = 1 To EntidadesLast
            With Entidades(i)
                If .active Then
                    If .MuereEnTick Then
                        If .NaceEnTick < tick And .MuereEnTick > tick Then
                            .Progreso = (tick - .NaceEnTick) / (.MuereEnTick - .NaceEnTick)
                            .active = 1
                        ElseIf .MuereEnTick < tick Then
                            'Se murio.
                            If .NaceEnTick < tick Then
                                .active = 0
                                Entidad_Limpiar i
                            End If
                        End If
                    ElseIf .VidaTotal Then
                        If .VidaActual > 0 Then
                            .Progreso = (.VidaTotal - .VidaActual) / .VidaTotal
                            .active = 1
                        Else
                            'Se murio.
                            Entidad_Limpiar i
                            .active = 0
                        End If
                    End If
                    
                    If .Progreso > 1 Then
                        .Progreso = 1
                    End If
                    
                
                    'Sonidos
                    If .SonidosCount > 0 Then
                        Dim NuevoSonido As Integer
                        
                        NuevoSonido = (.Progreso * (.SonidosCount - 1)) Mod &HFF
                        'Reproducir sonido
                      
                        If .MPPos.x + offset_map.x > -64 And .MPPos.y + offset_map.y > -64 And .MPPos.x + offset_map.x < 550 And .MPPos.y + offset_map.y < 550 Then 'ESTOY EN RANGO
                            If NuevoSonido <> .SonidoAnterior And .Sonidos(NuevoSonido) Then
                                Debug.Print "ENT"; i; "> Play sound  "; NuevoSonido; Abs(.Sonidos(NuevoSonido))
                                
                                If .SonidoAnterior <> 255 Then
                                    If .streamSonido <> 0 And .Sonidos(NuevoSonido) <> .Sonidos(.SonidoAnterior) Then
                                        BASS_ChannelStop .streamSonido
                                        .streamSonido = 0
                                    End If
                                End If
                                
                                If .Sonidos(NuevoSonido) > 0 Then
                                    Sonido_Play .Sonidos(NuevoSonido)
                                    .streamSonido = 0
                                Else
                                    .streamSonido = Sonido_PlayEX(-.Sonidos(NuevoSonido), True)
                                    modBass.BASS_ChannelSlideAttribute .streamSonido, BASS_ATTRIB_VOL, 1, 200
                                End If
                            ElseIf .Sonidos(NuevoSonido) < 0 And .streamSonido = 0 Then
                                .streamSonido = Sonido_PlayEX(-.Sonidos(NuevoSonido), True)
                                modBass.BASS_ChannelSlideAttribute .streamSonido, BASS_ATTRIB_VOL, 1, 200
                            End If
                        Else
                            If .streamSonido <> 0 Then
                                BASS_ChannelStop .streamSonido
                                .streamSonido = 0
                            End If
                        End If
                        
                        .SonidoAnterior = NuevoSonido
                        
                    End If

                    'if Sigue viva?
                    If .active And (.NaceEnTick > 0 And .NaceEnTick < tick) Then
                        If .Proyectil Then
                            If .Char Then
                                
                                .MPPos.x = Interp(.origen.x, CharList(.Char).Pos.x * 32, .Progreso)
                                .MPPos.y = Interp(.origen.y, CharList(.Char).Pos.y * 32, .Progreso)
                                .Angulo = (Angulo(CharList(.Char).Pos.x * 32, CharList(.Char).Pos.y * 32, .MPPos.x, .MPPos.y) + 270) Mod 360
                            Else
                                
                                .MPPos.x = Interp(.origen.x, .destino.x, .Progreso)
                                .MPPos.y = Interp(.origen.y, .destino.y, .Progreso)
                                '.Angulo = Engine_GetAngle(.MPPos.X, .MPPos.Y, .destino.X, .destino.Y)
                                .Angulo = Angulo(.destino.x, .destino.y, .MPPos.x, .MPPos.y)
                            End If
                            
                            .MoverLuz = 1
                        ElseIf .destino.x Then
                            .MPPos.x = CosInterp(.origen.x, .destino.x, .Progreso)
                            .MPPos.y = CosInterp(.origen.y, .destino.y, .Progreso)
                            .MoverLuz = 1
                        End If
                        
                        'Particulas
                        If .ParticulasCount > 0 Then
                             'Engine_Particles.Particle_Group_Kill .ParticulasCreadas(.ParticulaAnterior), 0
                            Dim ParticulaNueva As Integer
                            
                            ParticulaNueva = (.Progreso * (.ParticulasCount - 1)) Mod &HFF

                            If ParticulaNueva <> .ParticulaAnterior And .Particulas(ParticulaNueva) Then
                                If Not .ParticulasCreadas(.ParticulaAnterior) Is Nothing Then
                                    .ParticulasCreadas(.ParticulaAnterior).Matar 1
                                End If
                                
                                Set .ParticulasCreadas(ParticulaNueva) = New Engine_Particle_Group
                                Engine_Particles_Storage.IniciarGrupoParticulas .Particulas(ParticulaNueva), .ParticulasCreadas(ParticulaNueva)
                            End If
                            
                            .ParticulaAnterior = ParticulaNueva
                        End If
                        
                        'Graficos
                        If .GraficosCount > 0 Then
                            .GraficoActual = (.Progreso * (.GraficosCount - 1)) Mod &HFF
                            If .GraficoActual <> .GraficoAnterior Then
                                .GraficoAnterior = .GraficoActual
                                InitGrh .Grafico, .Graficos(.GraficoActual)
                            End If
                        End If
                        
                        
                        'Luz
                        If .MoverLuz Then
                            If .ParticulasCount Then
                                For j = 0 To .ParticulasCount - 1
                                    If Not .ParticulasCreadas(j) Is Nothing Then
                                        'Call Engine_Particles.Particle_Group_Set_PPos(.ParticulasCreadas(j), .MPPos.X, .MPPos.Y)
                                        .ParticulasCreadas(j).SetPixelPos .MPPos.x, .MPPos.y
                                    End If
                                Next j
                            End If
                            
                            If .Luz_ID > 0 Then Call DLL_Luces.MovePixel(.Luz_ID, .MPPos.x, .MPPos.y)
                            .MoverLuz = 0
                        End If
                    End If
                    'End If sigue viva.
                End If
            End With
            'If Entidades(i).active = 0 Then Entidades_Remover i
        Next i
    End If
End Function

'TODO x! e y! no se usa.
'TODO Es una funcion pero no devuelve nada
Public Function Entidades_Render_Recursivo(ByVal x!, ByVal y!, ByVal MX As Integer, ByVal MY As Integer)
    Dim tick As Long
    Dim TmpEntidad As Integer
    
    tick = GetTimer
    
    If EntidadesCount > 0 Then
        TmpEntidad = EntidadesMap(MX, MY)
        If TmpEntidad > 0 Then
            If TmpEntidad <= EntidadesLast Then
                Do While TmpEntidad
                    Call Entidades_Render(TmpEntidad)
                    TmpEntidad = Entidades(TmpEntidad).Next
                Loop
            End If
        End If
    End If
End Function

'Es una funcion pero no devuelve nada
Public Function Entidades_ReproducirSonido(ByVal entidad As Integer, Optional ByVal Sonido As Integer = -1)
    If EntidadesCount > 0 And entidad > 0 Then
        If entidad <= EntidadesLast Then
            With Entidades(entidad)
                If .active Then
                    'If .NaceEnTick > tick Then
                        If .Progreso <= 1 Then
                            If Sonido = -1 Then
                                If .Sonidos(.SonidoAnterior) Then Sonido_Play .Sonidos(.SonidoAnterior)
                            Else
                                If .SonidosCount < Sonido Then
                                    If .Sonidos(Sonido) Then Sonido_Play .Sonidos(Sonido)
                                End If
                            End If
                        End If
                    'End If
                End If
            End With
        End If
    End If
End Function
'Es una funcion, pero no devuelve nada
Public Function Entidades_SetDestino(ByVal entidad As Integer, ByVal PixelPosX As Integer, PixelPosY As Integer)
    If EntidadesCount > 0 And entidad > 0 Then
        If entidad <= EntidadesLast Then
            With Entidades(entidad)
                If .active Then
                    .destino.x = PixelPosX
                    .destino.y = PixelPosY
                    .Char = 0
                End If
            End With
        End If
    End If
End Function
'TODO Es una funcion, pero no devuelve nada.
Public Function Entidades_SetCharDestino(ByVal entidad As Integer, ByVal Char As Integer)
    If EntidadesCount > 0 And entidad > 0 Then
        If entidad <= EntidadesLast Then
            With Entidades(entidad)
                If .active Then
                    .destino.x = 0
                    .destino.y = 0
                    .Char = Char
                End If
            End With
        End If
    End If
End Function
'TODO Es una funcion pero no devuelve nada
Public Function Entidades_SetVidaActual(ByVal entidad As Integer, ByVal Vida As Integer)
    If EntidadesCount > 0 And entidad > 0 Then
        If entidad <= EntidadesLast Then
            With Entidades(entidad)
                If .active Then
                    .VidaActual = Vida
                End If
            End With
        End If
    End If
End Function

'TODO Es una funcion pero no devuelve nada
Public Function Entidades_SetVidaMaxima(ByVal entidad As Integer, ByVal Vida As Integer)
    If EntidadesCount > 0 And entidad > 0 Then
        If entidad <= EntidadesLast Then
            With Entidades(entidad)
                If .active Then
                    .VidaTotal = Vida
                End If
            End With
        End If
    End If
End Function

'TODO Es una funcion pero no devuelve nada
Public Function Entidades_SetPPos(ByVal entidad As Integer, ByVal PixelPosX As Integer, ByVal PixelPosY As Integer)
    If EntidadesCount > 0 And entidad > 0 Then
        If entidad <= EntidadesLast Then
            With Entidades(entidad)
                If .active Then
                    .MPPos.x = PixelPosX
                    .MPPos.y = PixelPosY
                    .origen = .MPPos
                End If
            End With
        End If
    End If
End Function

'Es una funcion pero no devuelve nada
Public Function Entidades_SetMPos(ByVal entidad As Integer, ByVal MapX As Integer, ByVal MapY As Integer)
    If EntidadesCount > 0 And entidad > 0 Then
        If entidad <= EntidadesLast Then
            With Entidades(entidad)
                If .active Then
                    If MapX > 0 And MapY > 0 Then
                        Entidad_QuitarDeMapa entidad
                        
                        .MPPos.x = MapX * 32
                        .MPPos.y = MapY * 32
                        .origen = .MPPos
                        
                        .map_x = MapX
                        .map_y = MapY
                        
                        Entidad_AgregarAMapa entidad, MapX, MapY
                    End If
                End If
            End With
        End If
    End If
End Function

'Es una funcion pero no devuelve nada
Public Function Entidades_Render(ByVal entidad As Integer)
    Dim tick As Long
    Dim i As Byte
    Dim x As Single, y As Single
    
    tick = GetTimer
    
    If EntidadesCount > 0 And entidad > 0 Then
        If entidad <= EntidadesLast Then
            With Entidades(entidad)
                If .active Then
                    If .NaceEnTick < tick Then
                        If .Progreso <= 1 Then
                            
                            x = .MPPos.x + offset_map.x + minXOffset * 32
                            'y = .MPPos.y + offset_map.y
                            
                            If .map_x > 0 And .map_y > 0 Then
                                y = .MPPos.y + offset_map.y - AlturaPie(.map_x, .map_y) + minYOffset * 32
                            Else
                                y = .MPPos.y + offset_map.y + minYOffset * 32
                            End If
                            
                            
                            'y = y - .Altura
                            
                            If .Proyectil = 0 Then
                                If .GraficosCount Then
                                    Draw_Grh .Grafico, x + .offsetX, y + .offsetY, 1, .map_x, .map_y
                                End If
                            Else
                                Grh_Render_Rotated .Graficos(.GraficoActual), x, y, base_light, .Angulo ', , base_light, 180 - .Angulo
                            End If
                            
                                 
                            If .ParticulasCount Then
                                For i = 0 To .ParticulasCount 'Render particula.
                                    'If .ParticulasCreadas(i) Then Call Particle_Group_Render(.ParticulasCreadas(i))
                                    If Not .ParticulasCreadas(i) Is Nothing Then
                                        .ParticulasCreadas(i).SetPixelPos .MPPos.x, .MPPos.y
                                        If .ParticulasCreadas(i).Render = False Then
                                            Debug.Print "ENT"; entidad; "> Muere particula "; i
                                            Set .ParticulasCreadas(i) = Nothing
                                        End If
                                    End If
                                Next i
                            End If
                        End If
                    End If
                End If
            End With
        End If
    End If
End Function


'**** Funciones de Manipulacion para el map editor
#If esMe Then
    Public Sub Entidades_SetIDIndexada(ByVal entidad As Integer, ByVal Index As Integer)
        Entidades(entidad).numeroIndexadoEntidad = Index
    End Sub
    
    Public Sub Entidades_SetAccion(ByVal entidad As Integer, ByVal accion As iAccionEditor)
       Set Entidades(entidad).accion = accion
    End Sub
    
    'Elimina sin delay una entidad
    Public Sub eliminar(ByVal entidad As Integer)
            Entidad_Limpiar entidad
            Entidades(entidad).active = 0
    End Sub
#End If
