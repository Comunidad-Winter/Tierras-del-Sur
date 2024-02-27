Attribute VB_Name = "Engine_Sonido_Extend"
Option Explicit

Private Type udtLista_Sonidos
    active As Byte

    Sonido As Integer
    stream As Long
    
    tick_muerte As Long
End Type

Private Lista_Sonidos()     As udtLista_Sonidos
Private Lista_Sonidos_Max    As Integer
Private Lista_Sonidos_Count  As Integer
Private Lista_Sonidos_Last   As Integer

Public Sub Sonido_CambiarVolumen_Ambiente(volumen As Single)

    Dim i As Integer
    
    If Lista_Sonidos_Count < Lista_Sonidos_Max Then 'nos aseguramos de que haya espacio
        For i = 0 To Lista_Sonidos_Max
            If Not Lista_Sonidos(i).active = 0 Then
                modBass.BASS_ChannelSlideAttribute Lista_Sonidos(i).stream, BASS_ATTRIB_VOL, volumen, 0
            End If
        Next i
    End If
    
End Sub

Public Sub Sonido_Play_Ambiente(ByVal Sonido As Integer, Optional ByVal fadeinMS As Long = 0)
    Dim Index As Integer
    
    '¿Tiene activada la musica?
    If Not Musica Then Exit Sub
    
    Index = Sonido_Ambiental_Buscar(Sonido)
    
    If Index = -1 Then
        Index = Sonido_Ambiental_Agregar(Sonido)
    End If
    
    With Lista_Sonidos(Index)
        .active = .stream <> 0
        .Sonido = Sonido
        
        If .active = 0 Then
            .stream = Engine_Sonido.Sonido_PlayEX(Sonido, True, 0)
            .active = .stream <> 0
        
            If .active Then
                modBass.BASS_ChannelSlideAttribute .stream, BASS_ATTRIB_VOL, volumenMusica, fadeinMS
            End If
        End If
    End With
End Sub

Public Sub Sonido_Stop_Ambiente(ByVal Sonido As Integer, Optional ByVal fadeoutMS As Long = 0)

    Dim Index As Integer
    
    Index = Sonido_Ambiental_Buscar(Sonido)
    
    If Index <> -1 Then
        If fadeoutMS = 0 Then
            Engine_Sonido.Sonido_Stop Lista_Sonidos(Index).Sonido
            Sonido_Ambiental_Remover Index
        Else
            modBass.BASS_ChannelSlideAttribute Lista_Sonidos(Index).stream, BASS_ATTRIB_VOL, 0, fadeoutMS
            Lista_Sonidos(Index).tick_muerte = GetTimer + fadeoutMS
        End If
    End If
End Sub

Public Sub BatchSonidos()
    Static UltimoBatch As Long
    Dim TickActual As Long
    
    TickActual = GetTimer
    
    If TickActual - UltimoBatch > 100 Then
        UltimoBatch = TickActual
    
        If Lista_Sonidos_Count Then
            Dim i As Integer
            For i = 0 To Lista_Sonidos_Last
                If Lista_Sonidos(i).active Then
                    If Lista_Sonidos(i).tick_muerte < TickActual And Lista_Sonidos(i).tick_muerte > 0 Then
                        Engine_Sonido.Sonido_Stop Lista_Sonidos(i).Sonido
                        Sonido_Ambiental_Remover i
                    End If
                End If
            Next i
        End If
    End If

End Sub

Public Sub Sonido_Ambiental_Iniciar(ByVal maximo As Integer)
    Lista_Sonidos_Max = maximo
    Sonido_Ambiental_ReIniciar
End Sub

Public Sub Sonido_Ambiental_ReIniciar()
    Dim i As Integer
    If Lista_Sonidos_Count Then
        For i = 0 To Lista_Sonidos_Last
            If Lista_Sonidos(i).active Then
                Engine_Sonido.Sonido_Stop Lista_Sonidos(i).Sonido
                Lista_Sonidos(i).active = 0
            End If
        Next i
    End If

    ReDim Lista_Sonidos(Lista_Sonidos_Max)
    Lista_Sonidos_Count = 0
    Lista_Sonidos_Last = 0
End Sub

Private Function Sonido_Ambiental_Agregar(ByVal Sonido As Long) As Integer
    Dim i As Integer
    Sonido_Ambiental_Agregar = Sonido_Ambiental_ObtenerLibre
    
    If Sonido_Ambiental_Agregar <> -1 Then
        With Lista_Sonidos(Sonido_Ambiental_Agregar)
            .active = 1
            .Sonido = Sonido
        End With
        If Sonido_Ambiental_Agregar > Lista_Sonidos_Last Then Lista_Sonidos_Last = Sonido_Ambiental_Agregar
        Lista_Sonidos_Count = Lista_Sonidos_Count + 1
    End If
End Function

Private Function Sonido_Ambiental_ObtenerLibre() As Integer
    Dim i As Integer
    
    If Lista_Sonidos_Count < Lista_Sonidos_Max Then 'nos aseguramos de que haya espacio
        For i = 0 To Lista_Sonidos_Max
            If Lista_Sonidos(i).active = 0 Then
                Sonido_Ambiental_ObtenerLibre = i
                Exit Function
            End If
        Next i
    End If
    
    Sonido_Ambiental_ObtenerLibre = -1
End Function

Private Function Sonido_Ambiental_Remover(ByRef Index As Integer) As Boolean 'True cuando se cambia el active de verdadero a falso
    Dim i As Integer
    If Lista_Sonidos_Count Then
        If Index <= Lista_Sonidos_Last Then
            Sonido_Ambiental_Remover = Lista_Sonidos(Index).active
            Lista_Sonidos(Index).active = 0
            Lista_Sonidos(Index).stream = 0
            Lista_Sonidos_Count = Lista_Sonidos_Count - 1
            Lista_Sonidos(Index).tick_muerte = 0
            If Index = Lista_Sonidos_Last Then
                For i = Lista_Sonidos_Last To 0 Step -1
                    If Lista_Sonidos(i).active Then
                        Lista_Sonidos_Last = i
                        Exit For
                    End If
                Next i
                If Lista_Sonidos_Last = Index Then Lista_Sonidos_Last = 0
            End If
        End If
    End If
End Function

Private Function Sonido_Ambiental_Buscar(ByVal Sonido As Long) As Integer

    Dim i As Integer
    If Lista_Sonidos_Count Then
        For i = 0 To Lista_Sonidos_Last
            If Lista_Sonidos(i).Sonido = Sonido Then
                Sonido_Ambiental_Buscar = i
                Exit Function
            End If
        Next i
    End If
    
    Sonido_Ambiental_Buscar = -1
End Function
