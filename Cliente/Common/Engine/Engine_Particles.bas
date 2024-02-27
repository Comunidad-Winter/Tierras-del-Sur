Attribute VB_Name = "Engine_Particles"
' ESTE ARCHIVO ESTA COMPARTIDO

Option Explicit

Public Type PARTVERTEXc  'NO TOCAR POR NADA EN EL MUNDO
    v As D3DVECTOR
    rhw As Single       'NO TOCAR POR NADA EN EL MUNDO
    Tamaño As Single    'NO TOCAR POR NADA EN EL MUNDO
        alpha As Byte
        red As Byte
        green As Byte
        blue As Byte
    tu As Single        'NO TOCAR POR NADA EN EL MUNDO
    tv As Single        'NO TOCAR POR NADA EN EL MUNDO
End Type                'NO TOCAR POR NADA EN EL MUNDO


Private Type pa_gro
    PrtData()       As Particle
    PrtVertList()   As PARTVERTEX
    type            As Integer ' Particle_Stream
    
    progress        As Single
    dir             As Integer
    lifecounter     As Integer
    muere           As Byte
    stage(0 To 1)   As Integer

    emmisor         As mzVECTOR2

    killable        As Byte
End Type

Public Type tEtapa
    start As Integer
    end As Integer
End Type



Public Type Emisores_Combinados
    streams()   As pa_gro
    etapas()    As tEtapa
    etapa       As Integer
    estapas_num As Integer
    streams_num As Integer
    Target      As mzVECTOR2
    targeta     As Byte
    target_char As Integer
    id          As Integer
    killable    As Byte
    stages      As Byte
    X           As Byte
    Y           As Byte
End Type



Dim particle_group_list() As Emisores_Combinados 'pa_gro
Dim particle_group_count As Integer
Dim particle_group_last As Integer

Public mpz As Single

'SANGRE



Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public Function Particle_Group_Make(ByRef particle_group_index As Integer, ByVal map_x As Integer, ByVal map_y As Integer, ByVal Stream_Type As Integer, Optional ByVal Capa As Byte = 1) As Integer
''*****************************************************
''****** Coded by Menduz (lord.yo.wo@gmail.com) *******
''*****************************************************
'    If Stream_Type > emisores_particulas_count Then Exit Function
'    If map_x = 0 Then Exit Function
'    Dim a As Integer
'    Call Particle_Group_Create(MapData(map_x, map_y).Particles_groups(capa), Stream_Type)
'    particle_group_index = MapData(map_x, map_y).Particles_groups(capa)
'    With particle_group_list(particle_group_index)
'        'Map pos
'        .x = map_x
'        .y = map_y
'
'        For a = 0 To .streams_num
'            With .streams(a)
'                .emmisor.x = 32 * map_x + 16
'                .emmisor.y = 32 * map_y + 16
'            End With
'        Next a
'        .stages = capa
'    End With
'
'
'    Map_render_2array
'    Particle_Group_Make = particle_group_index
End Function

Public Sub Particle_Group_Set_MPos(ByVal particle_group_index As Integer, ByVal map_x As Integer, ByVal map_y As Integer)
''*****************************************************
''****** Coded by Menduz (lord.yo.wo@gmail.com) *******
''*****************************************************offset_map
'    If particle_group_index > particle_group_last Then Exit Sub
'    Dim a As Integer
'    For a = 0 To particle_group_list(particle_group_index).streams_num
'        With particle_group_list(particle_group_index).streams(a).emmisor
'            .x = 32 * map_x + 16
'            .y = 32 * map_y + 16
'        End With
'    Next a
'
End Sub

Public Sub Particle_Group_Set_PPos(ByVal particle_group_index%, ByVal map_x!, ByVal map_y!)
''*****************************************************
''****** Coded by Menduz (lord.yo.wo@gmail.com) *******
''*****************************************************
'    If particle_group_index > particle_group_last Then Exit Sub
'    Dim a As Integer
'    For a = 0 To particle_group_list(particle_group_index).streams_num
'        With particle_group_list(particle_group_index).streams(a).emmisor
'            .x = map_x
'            .y = map_y
'        End With
'    Next a
End Sub

Public Sub Particle_Group_Set_TMPos(ByVal particle_group_index As Integer, ByVal map_x As Integer, ByVal map_y As Integer)
'    If particle_group_index > particle_group_last Then Exit Sub
'    With particle_group_list(particle_group_index)
'        .Target.x = 32 * map_x + 16
'        .Target.y = 32 * map_y + 16
'        .targeta = 1
'    End With
End Sub

Public Sub Particle_Group_Set_TPPos(ByVal particle_group_index%, ByVal map_x!, ByVal map_y!)
'    If particle_group_index > particle_group_last Then Exit Sub
'    With particle_group_list(particle_group_index)
'        .Target.x = map_x
'        .Target.y = map_y
'        .targeta = 1
'    End With
End Sub

Public Sub Particle_Group_Set_TChar(ByVal particle_group_index As Integer, ByVal Char As Integer)
'    If particle_group_index <= particle_group_last Then Exit Sub
'    With particle_group_list(particle_group_index)
'        .Target.x = 0
'        .Target.y = 0
'        .target_char = Char
'    End With
End Sub

Public Sub Particle_Group_Kill(ByVal id%, ByVal times%)
''*****************************************************
''****** Coded by Menduz (lord.yo.wo@gmail.com) *******
''*****************************************************
'    If id > particle_group_last Then Exit Sub
'    Dim a As Integer
'    For a = 0 To particle_group_list(id).streams_num
'        With particle_group_list(id).streams(a)
'            .lifecounter = times * Particle_Stream(.type).NumOfParticles + 1
'            .muere = 1
'        End With
'    Next a

End Sub

Public Sub Particle_Group_Create(ByRef particle_group_index As Integer, ByVal Stream_Type As Integer)
''*****************************************************
''****** Coded by Menduz (lord.yo.wo@gmail.com) *******
''*****************************************************
'    Dim i%, a As Integer
'    If Stream_Type > emisores_particulas_count Or Stream_Type = 0 Then Exit Sub
'
'    Do
'        particle_group_index = particle_group_index + 1
'        'Update LastProjectile if we go over the size of the current array
'        If particle_group_index > particle_group_last Then
'            particle_group_last = particle_group_index
'            particle_group_count = particle_group_count + 1
'            ReDim Preserve particle_group_list(0 To particle_group_last)
''            Debug.Print "CPG>"; particle_group_index
'            Exit Do
'        End If
'    Loop While particle_group_list(particle_group_index).killable = 0
'
'    With particle_group_list(particle_group_index)
'        .streams_num = emisores_particulas(Stream_Type).streams_num
'        .estapas_num = emisores_particulas(Stream_Type).estapas_num
'        ReDim .streams(0 To .streams_num)
'        ReDim .etapas(0 To .estapas_num)
'
'        .etapa = 0
'
'        .targeta = 0
'        .target_char = 0
'        .killable = 0
'
'        For a = 0 To .streams_num
'            With .streams(a)
'                'Map pos
'                .type = emisores_particulas(Stream_Type).streams(a)
'                .dir = 1
'
'                .lifecounter = Particle_Stream(.type).vida * Particle_Stream(.type).NumOfParticles + 1
'                .muere = Particle_Stream(.type).muere
'
'                .progress = 0
'                .killable = 0
'
'                ReDim .PrtData(0 To Particle_Stream(.type).NumOfParticles)
'                ReDim .PrtVertList(0 To Particle_Stream(.type).NumOfParticles)
''                For i = 0 To Particle_Stream(.type).NumOfParticles
''                    .PrtData(i).viva = 0
''                Next i
'            End With
'        Next a
'        For a = 0 To .estapas_num
'            .etapas(a) = emisores_particulas(Stream_Type).etapas(a)
'        Next a
'    End With
End Sub

Public Sub Particle_Group_Erase(ByRef particle_group_index As Integer)
''*****************************************************
''****** Coded by Menduz (lord.yo.wo@gmail.com) *******
''*****************************************************
'    If UBound(particle_group_list) = particle_group_index Then
''        Erase particle_group_list(particle_group_index).PrtData
''        Erase particle_group_list(particle_group_index).PrtVertList
'        Erase particle_group_list(particle_group_index).streams
'        'Erase particle_group_list(particle_group_index).PrtVertList
'        With particle_group_list(particle_group_index)
'            If .x > 0 Then
'                MapData(.x, .y).Particles_groups(.stages) = 0
'            End If
'        End With
'        Debug.Print "EPG>"; particle_group_index
'        particle_group_last = particle_group_index - 1
'
'        ReDim Preserve particle_group_list(0 To particle_group_last)
'        particle_group_count = particle_group_count - 1
'    Else
'        particle_group_list(particle_group_index).killable = 1
'    End If
End Sub

Public Sub Particle_Group_Remove_All()
''*****************************************************
''****** Coded by Menduz (lord.yo.wo@gmail.com) *******
''*****************************************************
'    Dim y As Byte
'    Dim x As Byte
'
'    For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
'        For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
'            MapData(x, y).Particles_groups(0) = 0
'            MapData(x, y).Particles_groups(1) = 0
'            MapData(x, y).Particles_groups(2) = 0
'            MapData(x, y).Particles_groups_original(0) = 0
'            MapData(x, y).Particles_groups_original(1) = 0
'            MapData(x, y).Particles_groups_original(2) = 0
'        Next y
'    Next x
'    meteo_particle = 0
'    particle_group_count = 0
'    particle_group_last = 0
'    ReDim particle_group_list(0)
End Sub
Public Sub Particle_Group_Map_Remake_All()
''*****************************************************
''****** Coded by Menduz (lord.yo.wo@gmail.com) *******
''*****************************************************
'    Dim y As Byte
'    Dim x As Byte
'    Dim tmpp As Integer
'
'    For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
'        For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
'                If MapData(x, y).Particles_groups_original(0) Then Engine_Particles.Particle_Group_Make tmpp, x, y, MapData(x, y).Particles_groups_original(0), 0
'                If MapData(x, y).Particles_groups_original(1) Then Engine_Particles.Particle_Group_Make tmpp, x, y, MapData(x, y).Particles_groups_original(1), 1
'                If MapData(x, y).Particles_groups_original(2) Then Engine_Particles.Particle_Group_Make tmpp, x, y, MapData(x, y).Particles_groups_original(2), 2
'        Next y
'    Next x
End Sub

Public Sub Particle_Group_Render(ByRef ii%)
'    Dim i As Long, g&
'    Dim total As Long
'    Dim tmp As Byte
'    Dim tt As Byte
'    Dim tata As Byte
'
'
'
'    If particle_group_last >= ii And ii > 0 Then
'        Randomize
'        'If frmMain.ccc.value Then
'        Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_MODULATE)
''        Else
''        Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_SUBTRACT) 'D3DTOP_MODULATEINVALPHA_ADDCOLOR auras
''        End If
'        D3DDevice.SetVertexShader particleFVF
'
'        With particle_group_list(ii)
'
'            If .target_char > 0 Then
'                .Target.x = CharList(.target_char).MPPos.x
'                .Target.y = CharList(.target_char).MPPos.y
'            End If
'
'            If .x > 0 And .y > 0 Then
'                offset_map_part.y = offset_map_part.y - AlturaPie(.x, .y)
'            End If
'
'            For i = .etapas(.etapa).start To .etapas(.etapa).end
'                With .streams(i)
'                    If .killable = 0 Then
'                        #If MEDIR_PERFORMANCE = 1 Then
'                            timer_particles_performance.Time
'                        #End If
'                        g = UpdateParticles(.PrtData(0), .PrtVertList(0), 10, timerTicksPerFrame, .emmisor, particle_group_list(ii).Target, .progress, Particle_Stream(.type), offset_map_part, .muere, .lifecounter, Rnd * 20)
'                        #If MEDIR_PERFORMANCE = 1 Then
'                            particles_cant_render = particles_cant_render + g
'                            particles_time_calc = particles_time_calc + timer_particles_performance.TimeD
'                            timer_particles_performance.TimeD
'                        #End If
'                        If g = -1 Then
'                            .killable = 1
'                        Else
'
'                            total = total + 1
'                            If Particle_Stream(.type).Line = 1 Then
'                                Call GetTexture(0)
'                                D3DDevice.DrawPrimitiveUP D3DPT_LINESTRIP, g - 1, .PrtVertList(0), Part_size
'                            ElseIf Particle_Stream(.type).Line = 2 Then
'                                Call GetTexture(Particle_Stream(.type).texture)
'                                Set_Blend_Mode Particle_Stream(.type).blend_mode
'
'
'
'
'                            Else
'                                Call GetTexture(Particle_Stream(.type).texture)
'                                'tmp = SurfaceDB.GetTexturePNG(Particle_Stream(.type).texture)
'                                'If tmp = 0 Then D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
'                                'If Particle_Stream(.type).blend_mode Then
'                                Set_Blend_Mode Particle_Stream(.type).blend_mode
'                                'End If
'
'                                D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, g, .PrtVertList(0), Part_size
'                                'If tmp = 0 Then D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
'                            End If
'
'                            #If esME Then
'                                'part_totales = part_totales + G
'                                If frmMain.veremisor.Value Then
'                                    D3DDevice.SetVertexShader FVF
'                                    Engine.text_render_graphic Chr$(255) & "•" & Chr$(255) & " " & i & ">" & particle_group_list(ii).etapa & ">" & .progress, .emmisor.x - 4 + offset_map.x, .emmisor.y - 8 + offset_map.y, &H40FFFFFF
'                                    D3DDevice.SetVertexShader particleFVF
'                                End If
'                            #End If
'                        End If
'                        #If MEDIR_PERFORMANCE = 1 Then
'                            particles_time_render = particles_time_render + timer_particles_performance.TimeD
'                        #End If
'                    End If
'                End With
'                If .streams(i).progress = 1! Then
'                    tata = 255
'                    .streams(i).progress = 0
'                End If
'            Next i
'
'        End With
'        If tata Then tata = pasar_de_etapa(ii)
'        If total = 0 And tata = 0 Then
'            Particle_Group_Erase ii
'            ii = 0
'        Else
'            particle_group_list(ii).killable = 0
'        End If
'        D3DDevice.SetVertexShader FVF
'        Set_Blend_Mode
'        Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, lColorMod)
'    End If
End Sub

Function pasar_de_etapa(ByVal ii%) As Byte
'pasar_de_etapa = 0
'    With particle_group_list(ii)
'        If .estapas_num > .etapa Then
'            .etapa = .etapa + 1
'            pasar_de_etapa = 255
'            Debug.Print "ETAP"
'        End If
'    End With
End Function



