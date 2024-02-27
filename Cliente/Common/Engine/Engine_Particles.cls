VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Engine_Particles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type particle
    'pos
'        D3DVECTOR Pos;
'        float spdX, spdY, accX, accY, size,progress,progress_factor, alpha, alpha_factor;
'        D3DCOLORVALUE color;
'        unsigned char viva, j;
'        short i;
    v As D3DVECTOR
    spd As D3DVECTOR2
    acc As D3DVECTOR2
    
    size As Single

    progress As Single
    progress_factor As Single
    
    'color
    alpha As Single
    alpha_factor As Single
    color As D3DCOLORVALUE
    
    viva As Byte
    j As Byte
    i As Integer
End Type


Private Type pa_gro
    PrtData()       As particle
    PrtVertList()   As PARTVERTEX
    type            As Integer ' Particle_Stream
    
    progress        As Single
    dir             As Integer
    lifecounter     As Integer
    muere           As Byte
    stage(0 To 1)   As Integer

    emmisor         As D3DVECTOR2

    killable        As Byte
End Type


Private Type Stream
    'name As String
    
    varZ As Single

    accX As Single
    accY As Single
    spdX As Single
    spdY As Single

    alpha_factor As Single

    rnd_alpha_factor As Single
    rnd_x As Single
    rnd_y As Single

    trig_factor As Single
    
    proj_speed As Single

    mod_timer As Single
    'color
    AlphaInicial As Single
    RedInicial As Single
    GreenInicial As Single
    BlueInicial As Single
    RedFinal As Single
    GreenFinal As Single
    BlueFinal As Single
    size As Single
    
    offset_x As Single
    offset_y As Single
    
    progress_factor As Single
    friccion As Single
    
    StartColor As D3DCOLORVALUE
    EndColor As D3DCOLORVALUE
    
    NumOfParticles As Integer
    texture As Integer
    life As Integer
    vida As Integer

    tipo As Long
    
    alpha_coseno As Byte
    muere As Byte
    Line As Byte
    crece As Byte
    blend_mode As Byte
End Type

Private Type tEtapa
    start As Integer
    end As Integer
End Type

Private Type Emisor_Combinado
    streams()   As Integer
    etapas()    As tEtapa
    streams_num As Integer
    estapas_num As Integer
End Type

Private Type Emisores_Combinados
    streams()   As pa_gro
    etapas()    As tEtapa
    etapa       As Integer
    estapas_num As Integer
    streams_num As Integer
    target      As D3DVECTOR2
    targeta     As Byte
    target_char As Integer
    ID          As Integer
    killable    As Byte
    stages      As Byte
    x           As Byte
    Y           As Byte
End Type

Private Particle_Stream() As Stream
Private TotalStreams As Integer

Private emisores_particulas() As Emisor_Combinado
Private emisores_particulas_count As Integer

Dim particle_group_list() As Emisores_Combinados 'pa_gro
Dim particle_group_count As Integer
Dim particle_group_last As Integer

Public mpz As Single

'SANGRE

Private Const Gravedad As Single = 4
Private Const Sangre_grh As Integer = 14109 '263
Private Const Xvar As Single = 2 'Variación X cuando nace la partícula
Private Const Yvar As Single = 3.2 'Variación Y cuando nace la partícula
Private Const sangre_agrandamiento As Single = 0.15 'Variación Y cuando nace la partícula
Private Const Vida_gota As Integer = 900 'cada gota vive medio segundo antes de renacer

Private Type sangre_particula
    viva As Boolean
    MuereEnTick As Long
    NaceEnTick As Long
    ModificadorY As Integer 'altura en la que cae
    ModificadorY2 As Integer 'altura en la que cae
    x As Single
    Y As Single
    vX As Single
    vY As Single
    tamaño As Single
    alpha As Single
End Type

Private Type Sangres_fx
    viva As Boolean

    start As Long
    end As Long
    lenght As Long
    
    src_char As Integer
    
    src_x_map As Integer
    src_y_map As Integer

    fuente_x As Integer
    fuente_y As Integer

    cantidad As Integer
    particulas() As sangre_particula
    
    wind As Byte 'N=0 E=1 S=2 O=3
    
    last_act As Currency
    
    altura As Byte
    
    col As D3DCOLORVALUE
End Type

Private libre As Integer

Private Sangres() As Sangres_fx
Private ultima As Integer

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

'//char _stdcall Update_Particles(cParticle* pList, PARTVERTEX* vList, unsigned short cant, _
'    float elapsed_time, D3DVECTOR2 &emmisor,const D3DVECTOR2 &target, float &Progress, _
'    const cParticleStream* Strea, const D3DVECTOR2 &offset_map, unsigned char &muere, short &lifecounter){

Private Declare Function Update_Particles Lib "MZEngine.dll" (ByRef pList As particle, ByRef vList As PARTVERTEX, _
    ByVal count As Integer, ByVal elapsed_time As Single, ByRef emmisor As D3DVECTOR2, ByRef target As D3DVECTOR2, _
    ByRef progress As Single, ByRef pstream As Stream, ByRef offset_mapa As D3DVECTOR2, ByRef muere As Byte, ByRef life_counter As Integer, randoma As Long) As Long

Private Declare Sub Init_Particles Lib "MZEngine.dll" (ByRef tmanio As Single)

#If MEDIR_PERFORMANCE = 1 Then
    Private timer_particles_performance As New clsPerformanceTimer
    
    Private particles_time_calc As Double
    Private particles_time_render As Double
    Private particles_cant_render As Double
#End If

Private Function FactorDeAct(ID As Integer) As Single
    Dim start_time As Currency
    Static timer_freq As Currency

    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    Call QueryPerformanceCounter(start_time)
    
    FactorDeAct = ((start_time - Sangres(ID).last_act) / timer_freq * 1000) * 0.2
    
    Call QueryPerformanceCounter(Sangres(ID).last_act)
End Function

Private Sub Initialize_Particles()
ReDim Sangres(1) As Sangres_fx
ReDim Sangres(1).particulas(0) As sangre_particula
ReDim Sangres(0).particulas(0) As sangre_particula
ultima = 1
libre = 1

End Sub

Sub CargarParticle_Streams()
    '*****************************************************************
    'Menduz
    '*****************************************************************
    Dim streamfile As String
    Dim loopc As Integer
    Dim ct() As String
    Dim i As Integer
    Dim etapa As String
    streamfile = IniPath & "Particles.ini"

    TotalStreams = val(GetVar(streamfile, "INIT", "Total"))
    ReDim ct(0)
    ReDim Particle_Stream(1 To TotalStreams) As Stream
    
    For loopc = 1 To TotalStreams
        With Particle_Stream(loopc)
            '.name = GetVar(streamfile, CStr("GROUP" & loopc), "Name")
            .accX = GetVarSng(streamfile, CStr("GROUP" & loopc), "accX")
            .accY = GetVarSng(streamfile, CStr("GROUP" & loopc), "accY")
            .spdX = GetVarSng(streamfile, CStr("GROUP" & loopc), "spdX")
            .spdY = GetVarSng(streamfile, CStr("GROUP" & loopc), "spdY")
            .varZ = GetVarSng(streamfile, CStr("GROUP" & loopc), "VarZ")
            .AlphaInicial = GetVarSng(streamfile, CStr("GROUP" & loopc), "AlphaInicial")
            .RedInicial = GetVarSng(streamfile, CStr("GROUP" & loopc), "RedInicial")
            .GreenInicial = GetVarSng(streamfile, CStr("GROUP" & loopc), "GreenInicial")
            .BlueInicial = GetVarSng(streamfile, CStr("GROUP" & loopc), "BlueInicial")
            .alpha_factor = GetVarSng(streamfile, CStr("GROUP" & loopc), "AlphaFactor")
            .RedFinal = GetVarSng(streamfile, CStr("GROUP" & loopc), "RedFinal")
            .GreenFinal = GetVarSng(streamfile, CStr("GROUP" & loopc), "GreenFinal")
            .BlueFinal = GetVarSng(streamfile, CStr("GROUP" & loopc), "BlueFinal")
            .NumOfParticles = val(GetVar(streamfile, CStr("GROUP" & loopc), "NumOfParticles"))
            '.Gravity = GetVarSng(streamfile, CStr("GROUP" & loopc), "Gravity")
            .mod_timer = GetVarSng(streamfile, CStr("GROUP" & loopc), "Timer")
            .rnd_x = GetVarSng(streamfile, CStr("GROUP" & loopc), "RndX")
            .rnd_y = GetVarSng(streamfile, CStr("GROUP" & loopc), "RndY")
            .rnd_alpha_factor = GetVarSng(streamfile, CStr("GROUP" & loopc), "RndA")
            .texture = val(GetVar(streamfile, CStr("GROUP" & loopc), "texture"))
            .size = GetVarSng(streamfile, CStr("GROUP" & loopc), "Size")
            'Call DXCopyMemory(.Size_dword, .size, 4)
            .life = val(GetVar(streamfile, CStr("GROUP" & loopc), "Life"))
            .tipo = val(GetVar(streamfile, CStr("GROUP" & loopc), "Tipo"))
            If .tipo <> 1 And .NumOfParticles > 360 Then .NumOfParticles = 360
            .StartColor = CreateColorVal(.AlphaInicial, .RedInicial, .GreenInicial, .BlueInicial)
            .EndColor = CreateColorVal(.AlphaInicial, .RedFinal, .GreenFinal, .BlueFinal)
            .vida = val(GetVar(streamfile, CStr("GROUP" & loopc), "Vida"))
            .Line = val(GetVar(streamfile, CStr("GROUP" & loopc), "Line"))
            .offset_x = val(GetVar(streamfile, CStr("GROUP" & loopc), "OffX"))
            .offset_y = val(GetVar(streamfile, CStr("GROUP" & loopc), "OffY"))
            .crece = val(GetVar(streamfile, CStr("GROUP" & loopc), "Crece")) > 0
            .friccion = GetVarSng(streamfile, CStr("GROUP" & loopc), "Fricc")
            .muere = 0
            .proj_speed = GetVarSng(streamfile, CStr("GROUP" & loopc), "PSpd")
            If .proj_speed = 0 Then .proj_speed = 2
            .progress_factor = GetVarSng(streamfile, CStr("GROUP" & loopc), "Pf")
            .blend_mode = CByte(val(GetVar(streamfile, CStr("GROUP" & loopc), "Blend")))
            If .vida > 0 Then .muere = 1
            If .AlphaInicial = -1 Then .alpha_coseno = 1
            If .NumOfParticles > 0 Then .trig_factor = 360 / .NumOfParticles
            #If esME Then
            frmMain.lstParticulas.AddItem loopc
            #End If
        End With
    Next loopc
    
    emisores_particulas_count = val(GetVar(streamfile, "INIT", "Emi"))
    
    ReDim emisores_particulas(1 To emisores_particulas_count)
    
    For loopc = 1 To emisores_particulas_count
        With emisores_particulas(loopc)
            ct = Split(GetVar(streamfile, CStr("E" & loopc), "P"), ",")
            .streams_num = UBound(ct())
            ReDim .streams(0 To .streams_num)
            For i = 0 To .streams_num
                .streams(i) = Int(val(ct(i)))
            Next i
            #If esME Then
            frmMain.lstEmisores.AddItem loopc
            #End If
            etapa = GetVar(streamfile, CStr("E" & loopc), "Etapas")
            If Len(etapa) Then
                ct = Split(etapa, "-")
                .estapas_num = UBound(ct())
                ReDim .etapas(0 To .estapas_num)
                For i = 0 To .estapas_num
                    .etapas(i).start = Int(val(ReadField(1, ct(i), 32)))
                    .etapas(i).end = Int(val(ReadField(2, ct(i), 32)))
                Next i
            Else
                .estapas_num = 0
                ReDim .etapas(0)
                .etapas(0).start = 0
                .etapas(0).end = .streams_num
            End If
        End With
    Next loopc
    
    Initialize_Particles
End Sub

Public Sub Crear_Sangre(char As Integer, cantidad As Integer, Duracion As Long, altura As Byte)
Dim tick As Long
Dim actual As Integer
tick = GetTickCount

    actual = Buscar_Siguiente
    
    With Sangres(actual)
        .viva = True
        .src_char = char
        .src_x_map = CharList(char).pos.x
        .src_y_map = CharList(char).pos.Y
        .start = tick
        .end = tick + Duracion
        .lenght = Duracion
        .cantidad = cantidad
        .altura = altura
        .fuente_x = CharList(char).pos.x - 16
        .fuente_y = CharList(char).pos.Y
        If .src_x_map > 0 Then
            Call Long2RGB(ResultColorArray(.src_x_map, .src_y_map), .col.r, .col.g, .col.b)
        Else
            .col.r = 1
            .col.g = 1
            .col.b = 1
        End If
        
        If UBound(.particulas()) <> cantidad Then
            ReDim .particulas(cantidad) As sangre_particula
        End If
 
        'MapData(.fuente_x, .fuente_y).sangre_fx = actual
        Call Init_sangresita(actual)
    End With
End Sub

Private Sub limpiar_sangre(ID As Integer)
    If ultima >= ID Then
        If Sangres(ID).viva = False Then 'Sangres(ID).end < GetTickCount Then
        libre = ID
            With Sangres(ID)
                .viva = False
                
                .src_char = 0
                
                'If .src_x_map > 0 Then
                '    MapData(.src_x_map, .src_y_map).sangre_fx = 0
                '    Sangres(ID).src_x_map = 0
                '    Sangres(ID).src_y_map = 0
                'End If
            End With
        End If
    End If
End Sub

Private Function Buscar_Siguiente() As Integer
    Dim i As Integer
    
    'libre = -1
    'For i = 1 To ultima 'nunca va a ser un numero grande, en la pantalla no entra tanta sangre mueejjej
    '    If Sangres(i).viva = False Then
    '        libre = i
    '        Exit For
    '    End If
    'Next i
    
    If libre <> -1 Then
        If Sangres(libre).viva = False Then Buscar_Siguiente = libre
    End If
    If Buscar_Siguiente = 0 Then
        ReDim Preserve Sangres(ultima + 1)
        ultima = ultima + 1
        ReDim Sangres(ultima).particulas(0) As sangre_particula
        Buscar_Siguiente = ultima
    End If
    'Debug.Print "SE ABRIO"; Buscar_Siguiente; "PARA SANGRE"
End Function

Public Sub render_blood(ByVal offsetX As Single, ByVal offsetY As Single)
    Dim i As Integer
    Dim conteo As Integer
    For i = 1 To ultima 'nunca va a ser un numero grande, en la pantalla no entra tanta sangre mueejjej
        If Sangres(i).viva = True Then
            Render_ID i, offsetX, offsetY
            conteo = conteo + 1
        End If
    Next i
    If conteo = 0 Then
        ReDim Sangres(1) As Sangres_fx
        ReDim Sangres(1).particulas(0) As sangre_particula
        ReDim Sangres(0).particulas(0) As sangre_particula
        ultima = 1
        libre = 1
    End If
End Sub

Private Sub Render_ID(ID As Integer, ByVal offsetX As Single, ByVal offsetY As Single)
    Dim tick As Long
    Dim factor As Single
    Dim i As Integer
    Dim my As Integer
    Dim matar As Byte
    Dim cant As Integer
    Dim mox As Integer
    Dim moy As Integer
    Dim TLlist() As PARTVERTEX
    mox = offset_map.x '(Sangres(ID).src_x_map - MinX - TileBufferSizeX + 1) * 32 + OffsetX
    moy = offset_map.Y '(Sangres(ID).src_y_map - MinY - TileBufferSizeY + 1) * 32 + OffsetY
    
    If ultima >= ID And ID > 0 Then
        If Sangres(ID).viva = True Then
        
            tick = GetTickCount
            
            If Sangres(ID).end >= tick Then
            
                factor = FactorDeAct(ID)
                
'                Select Case (CharList(Sangres(ID).src_char).Heading - 1) 'Sangres(ID).wind
'                    Case 0
'                        my = 1
'                    Case 1
'                        my = 0
'                    Case 2
'                        my = 1
'                    Case 3
'                        my = 0
'                End Select
'                my = 1
                matar = 0
                
                ReDim TLlist(0 To Sangres(ID).cantidad)
                For i = 0 To Sangres(ID).cantidad
                    With Sangres(ID).particulas(i)
                        If tick > .MuereEnTick Then
                        .viva = False
                        End If
                        If .viva Then
                            If tick > .NaceEnTick Then
                                If .Y < .ModificadorY2 Then 'CharList(Sangres(ID).src_char).mppos.y - .ModificadorY Then
                                    .x = .x + (factor * .vX)
                                    .vY = .vY + Gravedad * (-my + 1) * factor * 0.01
                                    .Y = .Y + (factor * .vY) '+ (Gravedad * (-my + 1)) * factor * 0.2
                                

                                If .tamaño < 64 Then _
                                    .tamaño = .tamaño + factor * sangre_agrandamiento * 2
                                End If
                                TLlist(cant).tamaño = .tamaño
                                .alpha = .alpha - factor * 0.003
                                If .alpha < 0 Then
                                    .viva = False
                                    .alpha = 0
                                End If
                                
                                TLlist(cant).tamaño = .tamaño
                                TLlist(cant).v.x = .x + mox
                                TLlist(cant).v.Y = .Y + moy + 16
                                
                                TLlist(cant).color = D3DColorMake(Sangres(ID).col.r, Sangres(ID).col.g, Sangres(ID).col.b, .alpha)
                                TLlist(cant).rhw = 1
                                cant = cant + 1
                            Else
                                .x = CharList(Sangres(ID).src_char).mppos.x + 16 + Rnd * 5 - Rnd * 5
                                .Y = CharList(Sangres(ID).src_char).mppos.Y - Sangres(ID).altura + Rnd * 5 - Rnd * 5
                                .ModificadorY2 = CharList(Sangres(ID).src_char).mppos.Y - .ModificadorY
                            End If
                            matar = 1

                        End If
                        
                    End With
                Next i

                If matar = 0 Then
                    Sangres(ID).viva = False
                    libre = ID
                    limpiar_sangre ID
                Else
                    If cant > 0 Then
                        ReDim Preserve tmpTLlist(0 To cant - 1)
                        'D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
                        'Debug.Print cant
                        Call GetTexture(Sangre_grh)
                        Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_MODULATE)
                        D3DDevice.SetVertexShader particleFVF
                        D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, cant, TLlist(0), Part_size
                        Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, lColorMod)
                        D3DDevice.SetVertexShader FVF
                        'D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
                        'Engine.Grh_Render_Simple_box Sangre_grh, 0, 0, -1&, 32
                    Else
                    Debug.Print "a"
                    End If
                End If
            Else
                Sangres(ID).viva = False
                limpiar_sangre ID
                libre = ID
            End If
        Else
            limpiar_sangre ID
            libre = ID
        End If
    End If
End Sub

Private Sub Init_sangresita(ID As Integer)
    Dim tick As Long
    tick = GetTickCount
    Dim i As Integer
    Dim mx As Integer
    Dim my As Integer
    Dim MMX As Single
    Dim MMY As Single
    MMX = 0
    MMY = 0
    Call FactorDeAct(ID)
    
    Select Case (CharList(Sangres(ID).src_char).Heading - 1) 'Sangres(ID).wind
        Case 0
            my = 1
            mx = 0
            MMX = Rnd
            MMY = -1
        Case 1
            my = 0
            mx = 1
            MMY = Rnd
        Case 2
            my = 1
            mx = 0
            MMX = Rnd
        Case 3
            my = 0
            mx = -1
            MMY = Rnd
    End Select
                
    For i = 0 To Sangres(ID).cantidad
        With Sangres(ID).particulas(i)
            Randomize
            .vX = (Xvar * 0.5 * mx) - (Rnd * mx * Xvar * 0.5) + ((Rnd * MMX) - (Rnd * MMX))
            .vY = 0 '(Yvar * my) - (Rnd * my * Yvar) - ((Rnd * MMY) + (Rnd * MMY))
            '.x = mx * 5 - 16
            '.y = -Sangres(ID).altura
            .x = CharList(Sangres(ID).src_char).mppos.x
            .Y = CharList(Sangres(ID).src_char).mppos.Y - Sangres(ID).altura
            .ModificadorY = -16 + (Rnd * 5) - (Rnd * 5) '(Rnd * 5) - (my * 10)
            .ModificadorY2 = CharList(Sangres(ID).src_char).mppos.Y + .ModificadorY
            Dim tmp As Integer
            tmp = (Rnd * Vida_gota) * 0.5
            .NaceEnTick = tick + tmp
            .MuereEnTick = tick + Vida_gota + tmp + (Rnd * Vida_gota) + (Rnd * Vida_gota) + (Rnd * Vida_gota)
            If .MuereEnTick > Sangres(ID).end Then
            Sangres(ID).end = .MuereEnTick
            End If
            .viva = True
            .tamaño = 1
            .alpha = 1
        End With
    Next i
    
End Sub


Public Function Particle_Group_Make(ByRef particle_group_index As Integer, ByVal map_x As Integer, ByVal map_y As Integer, ByVal Stream_Type As Integer, Optional ByVal capa As Byte = 1) As Integer
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    If Stream_Type > emisores_particulas_count Then Exit Function
    If map_x = 0 Then Exit Function
    Dim a As Integer
    Call Particle_Group_Create(MapData(map_x, map_y).Particles_groups(capa), Stream_Type)
    particle_group_index = MapData(map_x, map_y).Particles_groups(capa)
    With particle_group_list(particle_group_index)
        'Map pos
        .x = map_x
        .Y = map_y
        
        For a = 0 To .streams_num
            With .streams(a)
                .emmisor.x = 32 * map_x + 16
                .emmisor.Y = 32 * map_y + 16
            End With
        Next a
        .stages = capa
    End With

    
    Map_render_2array
    Particle_Group_Make = particle_group_index
End Function

Public Sub Particle_Group_Set_MPos(ByVal particle_group_index As Integer, ByVal map_x As Integer, ByVal map_y As Integer)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************offset_map
    If particle_group_index > particle_group_last Then Exit Sub
    Dim a As Integer
    For a = 0 To particle_group_list(particle_group_index).streams_num
        With particle_group_list(particle_group_index).streams(a).emmisor
            .x = 32 * map_x + 16
            .Y = 32 * map_y + 16
        End With
    Next a
    
End Sub

Public Sub Particle_Group_Set_PPos(ByVal particle_group_index%, ByVal map_x!, ByVal map_y!)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    If particle_group_index > particle_group_last Then Exit Sub
    Dim a As Integer
    For a = 0 To particle_group_list(particle_group_index).streams_num
        With particle_group_list(particle_group_index).streams(a).emmisor
            .x = map_x
            .Y = map_y
        End With
    Next a
End Sub

Public Sub Particle_Group_Set_TMPos(ByVal particle_group_index As Integer, ByVal map_x As Integer, ByVal map_y As Integer)
    If particle_group_index > particle_group_last Then Exit Sub
    With particle_group_list(particle_group_index)
        .target.x = 32 * map_x + 16
        .target.Y = 32 * map_y + 16
        .targeta = 1
    End With
End Sub

Public Sub Particle_Group_Set_TPPos(ByVal particle_group_index%, ByVal map_x!, ByVal map_y!)
    If particle_group_index > particle_group_last Then Exit Sub
    With particle_group_list(particle_group_index)
        .target.x = map_x
        .target.Y = map_y
        .targeta = 1
    End With
End Sub

Public Sub Particle_Group_Set_TChar(ByVal particle_group_index As Integer, ByVal char As Integer)
    If particle_group_index <= particle_group_last Then Exit Sub
    With particle_group_list(particle_group_index)
        .target.x = 0
        .target.Y = 0
        .target_char = char
    End With
End Sub

Public Sub Particle_Group_Kill(ByVal ID%, ByVal times%)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    If ID > particle_group_last Then Exit Sub
    Dim a As Integer
    For a = 0 To particle_group_list(ID).streams_num
        With particle_group_list(ID).streams(a)
            .lifecounter = times * Particle_Stream(.type).NumOfParticles + 1
            .muere = 1
        End With
    Next a

End Sub

Public Sub Particle_Group_Create(ByRef particle_group_index As Integer, ByVal Stream_Type As Integer)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    Dim i%, a As Integer
    If Stream_Type > emisores_particulas_count Or Stream_Type = 0 Then Exit Sub
    
    Do
        particle_group_index = particle_group_index + 1
        'Update LastProjectile if we go over the size of the current array
        If particle_group_index > particle_group_last Then
            particle_group_last = particle_group_index
            particle_group_count = particle_group_count + 1
            ReDim Preserve particle_group_list(0 To particle_group_last)
'            Debug.Print "CPG>"; particle_group_index
            Exit Do
        End If
    Loop While particle_group_list(particle_group_index).killable = 0

    With particle_group_list(particle_group_index)
        .streams_num = emisores_particulas(Stream_Type).streams_num
        .estapas_num = emisores_particulas(Stream_Type).estapas_num
        ReDim .streams(0 To .streams_num)
        ReDim .etapas(0 To .estapas_num)
        
        .etapa = 0
        
        .targeta = 0
        .target_char = 0
        .killable = 0
        
        For a = 0 To .streams_num
            With .streams(a)
                'Map pos
                .type = emisores_particulas(Stream_Type).streams(a)
                .dir = 1
                
                .lifecounter = Particle_Stream(.type).vida * Particle_Stream(.type).NumOfParticles + 1
                .muere = Particle_Stream(.type).muere
                
                .progress = 0
                .killable = 0

                ReDim .PrtData(0 To Particle_Stream(.type).NumOfParticles)
                ReDim .PrtVertList(0 To Particle_Stream(.type).NumOfParticles)
'                For i = 0 To Particle_Stream(.type).NumOfParticles
'                    .PrtData(i).viva = 0
'                Next i
            End With
        Next a
        For a = 0 To .estapas_num
            .etapas(a) = emisores_particulas(Stream_Type).etapas(a)
        Next a
    End With
End Sub

Public Sub Particle_Group_Erase(ByRef particle_group_index As Integer)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    If UBound(particle_group_list) = particle_group_index Then
'        Erase particle_group_list(particle_group_index).PrtData
'        Erase particle_group_list(particle_group_index).PrtVertList
        Erase particle_group_list(particle_group_index).streams
        'Erase particle_group_list(particle_group_index).PrtVertList
        With particle_group_list(particle_group_index)
            If .x > 0 Then
                MapData(.x, .Y).Particles_groups(.stages) = 0
            End If
        End With
        Debug.Print "EPG>"; particle_group_index
        particle_group_last = particle_group_index - 1

        ReDim Preserve particle_group_list(0 To particle_group_last)
        particle_group_count = particle_group_count - 1
    Else
        particle_group_list(particle_group_index).killable = 1
    End If
End Sub

Public Sub Particle_Group_Remove_All()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    Dim Y As Byte
    Dim x As Byte
    For x = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
            MapData(x, Y).Particles_groups(0) = 0
            MapData(x, Y).Particles_groups(1) = 0
            MapData(x, Y).Particles_groups(2) = 0
            MapData(x, Y).Particles_groups_original(0) = 0
            MapData(x, Y).Particles_groups_original(1) = 0
            MapData(x, Y).Particles_groups_original(2) = 0
        Next Y
    Next x
    meteo_particle = 0
    particle_group_count = 0
    particle_group_last = 0
    ReDim particle_group_list(0)
End Sub


Public Sub Particle_Group_Render(ByRef ii%)
    Dim i As Long, g&
    Dim total As Long
    Dim tmp As Byte
    Dim tt As Byte
    Dim tata As Byte
    
    Dim offset_map_part As D3DVECTOR2
    
    If particle_group_last >= ii And ii > 0 Then
        Randomize
        'If frmMain.ccc.value Then
        Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_MODULATE)
'        Else
'        Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_SUBTRACT) 'D3DTOP_MODULATEINVALPHA_ADDCOLOR auras
'        End If
        D3DDevice.SetVertexShader particleFVF
        
        With particle_group_list(ii)
        
            If .target_char > 0 Then
                .target.x = CharList(.target_char).mppos.x
                .target.Y = CharList(.target_char).mppos.Y
            End If
    
            'tt = (.target.X <> 0) And .target_char = 0
            offset_map_part.x = offset_map.x
            
            If .x > 0 And .Y > 0 Then
                offset_map_part.Y = offset_map.Y - AlturaPie(.x, .Y)
            Else
                offset_map_part.Y = offset_map.Y
            End If
            
            For i = .etapas(.etapa).start To .etapas(.etapa).end
                With .streams(i)
                    If .killable = 0 Then
                        #If MEDIR_PERFORMANCE = 1 Then
                            timer_particles_performance.Time
                        #End If
                        g = Update_Particles(.PrtData(0), .PrtVertList(0), 10, timerTicksPerFrame, .emmisor, particle_group_list(ii).target, .progress, Particle_Stream(.type), offset_map_part, .muere, .lifecounter, Rnd * 20)
                        #If MEDIR_PERFORMANCE = 1 Then
                            particles_cant_render = particles_cant_render + g
                            particles_time_calc = particles_time_calc + timer_particles_performance.TimeD
                            timer_particles_performance.TimeD
                        #End If
                        If g = -1 Then
                            .killable = 1
                        Else

                            total = total + 1
                            If Particle_Stream(.type).Line = 1 Then
                                Call GetTexture(0)
                                D3DDevice.DrawPrimitiveUP D3DPT_LINESTRIP, g - 1, .PrtVertList(0), Part_size
                            ElseIf Particle_Stream(.type).Line = 2 Then
                                Call GetTexture(Particle_Stream(.type).texture)
                                Set_Blend_Mode Particle_Stream(.type).blend_mode
                                
                                
                                
                                
                            Else
                                Call GetTexture(Particle_Stream(.type).texture)
                                'tmp = SurfaceDB.GetTexturePNG(Particle_Stream(.type).texture)
                                'If tmp = 0 Then D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
                                'If Particle_Stream(.type).blend_mode Then
                                Set_Blend_Mode Particle_Stream(.type).blend_mode
                                'End If

                                D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, g, .PrtVertList(0), Part_size
                                'If tmp = 0 Then D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
                            End If
                            
                            #If esME Then
                                part_totales = part_totales + g
                                If frmMain.veremisor.Value Then
                                    D3DDevice.SetVertexShader FVF
                                    Engine.text_render_graphic Chr$(255) & "•" & Chr$(255) & " " & i & ">" & particle_group_list(ii).etapa & ">" & .progress, .emmisor.x - 4 + offset_map.x, .emmisor.Y - 8 + offset_map.Y, &H40FFFFFF
                                    D3DDevice.SetVertexShader particleFVF
                                End If
                            #End If
                        End If
                        #If MEDIR_PERFORMANCE = 1 Then
                            particles_time_render = particles_time_render + timer_particles_performance.TimeD
                        #End If
                    End If
                End With
                If .streams(i).progress = 1! Then
                    tata = 255
                    .streams(i).progress = 0
                End If
            Next i
            
        End With
        If tata Then tata = pasar_de_etapa(ii)
        If total = 0 And tata = 0 Then
            Particle_Group_Erase ii
            ii = 0
        Else
            particle_group_list(ii).killable = 0
        End If
        D3DDevice.SetVertexShader FVF
        Set_Blend_Mode
        Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, lColorMod)
    End If
End Sub

Private Function pasar_de_etapa(ByVal ii%) As Byte
    pasar_de_etapa = 0
    With particle_group_list(ii)
        If .estapas_num > .etapa Then
            .etapa = .etapa + 1
            pasar_de_etapa = 255
            Debug.Print "ETAP"
        End If
    End With
End Function


'        mod_particletimer = timerTicksPerFrame * Particle_Stream(.type).mod_timer
'        .progress = .progress + mod_particletimer * 36
'        If .progress > 360 Then .progress = .progress - 360
'        If .progress > 360 Then .progress = .progress - 360
'
'        If .targeta Then
'            Ang = Engine_GetAngle(.emmisor.x, .emmisor.y, .target.x, .target.y) + 180
'            .emmisor.x = .emmisor.x - Sin(Ang * DegreeToRadian) * Particle_Stream(.type).proj_speed * mod_particletimer
'            .emmisor.y = .emmisor.y + Cos(Ang * DegreeToRadian) * Particle_Stream(.type).proj_speed * mod_particletimer
'        ElseIf .target_char > 0 Then
'            Ang = Engine_GetAngle(.emmisor.x, .emmisor.y, charlist(.target_char).mppos.x, charlist(.target_char).mppos.y) + 180
'            .emmisor.x = .emmisor.x - Sin(Ang * DegreeToRadian) * Particle_Stream(.type).proj_speed * mod_particletimer
'            .emmisor.y = .emmisor.y + Cos(Ang * DegreeToRadian) * Particle_Stream(.type).proj_speed * mod_particletimer
'        End If
'
'
'
'        For i = 0 To Particle_Stream(.type).NumOfParticles
'            If .PrtData(i).viva = 1 Then
'                ''.PrtData(i).alpha = .PrtData(i).alpha - .PrtData(i).alpha_factor * mod_particletimer
'                .PrtData(i).progress = .PrtData(i).progress + timerTicksPerFrame * .PrtData(i).progress_factor
'
'                If .PrtData(i).progress > 1 Then
'                    .PrtData(i).viva = 0
'                    .PrtVertList(i).color = &H0
'                    '.PrtData(i).progress = 0
'                    i = i - 1
'                Else
'                    If Particle_Stream(.type).tipo = 5 Then
'                        .PrtVertList(i).tamaño = .PrtVertList(i).tamaño + mod_particletimer * 5
'                    End If
'                    If .PrtVertList(i).tamaño > mpz Then
'                        .PrtData(i).viva = 0
'                        .PrtVertList(i).color = &H0
'                        '.PrtData(i).progress = 0
'                        i = i - 1
'                    Else
'                        If Particle_Stream(.type).alpha_coseno = 1 Then
'                            'tmp = (nowa - .PrtData(i).created) / Particle_Stream(.type).life
'                            tmp = .PrtData(i).progress
'                            .PrtData(i).alpha = -4 * (tmp - 1) * tmp * .PrtData(i).alpha_factor '-(tmp * 2 - 1) ^ 2 + 1 ' (-Cos(tmp * Pi2) + 1) / 2
'                            tmp = .PrtData(i).alpha * Particle_Stream(.type).rnd_alpha_factor
'                        Else
'                            .PrtData(i).alpha = .PrtData(i).alpha - .PrtData(i).alpha_factor * mod_particletimer
'
'                        End If
'                        If .PrtData(i).alpha > 1 Then .PrtData(i).alpha = 1
'
'                        .PrtData(i).v.x = .PrtData(i).v.x + .PrtData(i).spd.x * mod_particletimer * tmp
'                        .PrtData(i).v.y = .PrtData(i).v.y + .PrtData(i).spd.y * mod_particletimer * tmp
'                        .PrtData(i).spd.x = .PrtData(i).spd.x + .PrtData(i).acc.x * mod_particletimer
'                        .PrtData(i).spd.y = .PrtData(i).spd.y + .PrtData(i).acc.y * mod_particletimer
'
'                        .PrtVertList(i).v.x = .PrtData(i).v.x + offset_map.x
'                        .PrtVertList(i).v.y = .PrtData(i).v.y + offset_map.y
'                    'color
'                        Call D3DXColorLerp(.PrtData(i).color, Particle_Stream(.type).StartColor, Particle_Stream(.type).EndColor, .PrtData(i).progress) '(nowa - .PrtData(i).created) / Particle_Stream(.type).life)
'
'                        .PrtVertList(i).color = D3DColorMake(.PrtData(i).color.r, .PrtData(i).color.g, .PrtData(i).color.b, .PrtData(i).alpha)
'
'                        c = 1
'                    End If
'                End If
'            Else
'                If i = 1 Then If .muere = 1 Then .lifecounter = .lifecounter - 1
'                If .lifecounter >= 0 Then
'
'                        .PrtData(i).viva = 1
'                        .PrtData(i).created = nowa
'                        .PrtData(i).progress = 0
'                        '.PrtVertList(i).tamaño = Particle_Stream(.type).size
'
'                        Select Case Particle_Stream(.type).tipo
'                            Case 1, 5
'                                    .PrtData(i).acc.x = Particle_Stream(.type).accX
'                                .PrtData(i).acc.y = Particle_Stream(.type).accY
'                                .PrtData(i).v.x = .emmisor.x + RandomNumber(-Particle_Stream(.type).VarZ, Particle_Stream(.type).VarZ)
'                                .PrtData(i).v.y = .emmisor.y
'                                .PrtData(i).spd.x = Particle_Stream(.type).spdX + Particle_Stream(.type).rnd_x * Rnd - Particle_Stream(.type).rnd_x * Rnd
'                                .PrtData(i).spd.y = Particle_Stream(.type).spdY + Particle_Stream(.type).rnd_y * Rnd - Particle_Stream(.type).rnd_y * Rnd
'                            Case 2
'                                .PrtData(i).v.x = .emmisor.x + Coseno(i * Particle_Stream(.type).trig_factor) * Particle_Stream(.type).VarZ
'                                .PrtData(i).v.y = .emmisor.y + Seno(i * Particle_Stream(.type).trig_factor) * Particle_Stream(.type).VarZ
'                                .PrtData(i).spd.x = Coseno(i * Particle_Stream(.type).trig_factor) * (Particle_Stream(.type).rnd_x + 1)
'                                .PrtData(i).spd.y = Seno(i * Particle_Stream(.type).trig_factor) * (Particle_Stream(.type).rnd_y + 1)
'                            Case 3
'                                .PrtData(i).v.y = .emmisor.y - Particle_Stream(.type).VarZ * Coseno(i * Particle_Stream(.type).trig_factor) * Particle_Stream(.type).rnd_y
'                                .PrtData(i).v.x = .emmisor.x + Particle_Stream(.type).VarZ * Seno(i * Particle_Stream(.type).trig_factor) * Particle_Stream(.type).rnd_x
'                                .PrtData(i).spd.y = Particle_Stream(.type).spdY '- Seno(i * Particle_Stream(.type).trig_factor) * Particle_Stream(.type).VarZ
'                                .PrtData(i).spd.x = Particle_Stream(.type).spdX '- Coseno(i * Particle_Stream(.type).trig_factor) * Particle_Stream(.type).VarZ
'                            Case 4
'                                .PrtData(i).v.y = .emmisor.y - Particle_Stream(.type).VarZ * Coseno(i * Particle_Stream(.type).trig_factor) * Particle_Stream(.type).rnd_y
'                                .PrtData(i).v.x = .emmisor.x + Particle_Stream(.type).VarZ * Seno(i * Particle_Stream(.type).trig_factor) * Particle_Stream(.type).rnd_x
'                                .PrtData(i).spd.x = Particle_Stream(.type).spdX + Particle_Stream(.type).rnd_x * Rnd
'                                .PrtData(i).spd.y = Particle_Stream(.type).spdY + Particle_Stream(.type).rnd_y * Rnd
'                                .PrtData(i).acc.x = Particle_Stream(.type).accX - Seno(i * Particle_Stream(.type).trig_factor)
'                                .PrtData(i).acc.y = Particle_Stream(.type).accY + Coseno(i * Particle_Stream(.type).trig_factor)
'                            Case 6
'                                .PrtData(i).v.x = .emmisor.x - Coseno(i * Particle_Stream(.type).trig_factor) * Coseno(.progress) * Particle_Stream(.type).rnd_x
'                                .PrtData(i).v.y = .emmisor.y + Seno(i * Particle_Stream(.type).trig_factor) * Coseno(.progress) * Particle_Stream(.type).rnd_y
'                                .PrtData(i).spd.x = 0 '(Rnd - Rnd) * Particle_Stream(.type).VarZ
'                                .PrtData(i).spd.y = 0 '(Rnd - Rnd) * Particle_Stream(.type).VarZ
'                                .PrtData(i).acc.x = Particle_Stream(.type).accX - Seno(i * Particle_Stream(.type).trig_factor)
'                                .PrtData(i).acc.y = Particle_Stream(.type).accY + Coseno(i * Particle_Stream(.type).trig_factor)
'                            Case 7 'mediaesfera
'                                .PrtData(i).v.y = .emmisor.y - Particle_Stream(.type).VarZ * Coseno(i * Particle_Stream(.type).trig_factor) * Particle_Stream(.type).rnd_y
'                                .PrtData(i).v.x = .emmisor.x + Particle_Stream(.type).VarZ * Seno(i * Particle_Stream(.type).trig_factor) * Particle_Stream(.type).rnd_x
'                                .PrtData(i).spd.y = Particle_Stream(.type).spdY - Seno(i * Particle_Stream(.type).trig_factor) * Particle_Stream(.type).VarZ
'                                .PrtData(i).spd.x = Particle_Stream(.type).spdX - Coseno(i * Particle_Stream(.type).trig_factor) * Particle_Stream(.type).VarZ
'                                .PrtData(i).acc.x = Particle_Stream(.type).accX - Seno(i * Particle_Stream(.type).trig_factor)
'                                .PrtData(i).acc.y = Particle_Stream(.type).accY - Coseno(i * Particle_Stream(.type).trig_factor)
'
'                        End Select
'
'                        If Particle_Stream(.type).tipo = 5 Then
'                            .PrtVertList(i).tamaño = 16
'                        Else
'                            .PrtVertList(i).tamaño = Particle_Stream(.type).size * Rnd
'                        End If
'
'                        If Particle_Stream(.type).alpha_coseno Then
'                            .PrtData(i).alpha = 0
'                            .PrtData(i).alpha_factor = Particle_Stream(.type).alpha_factor - (Rnd * Particle_Stream(.type).rnd_alpha_factor - Rnd * Particle_Stream(.type).rnd_alpha_factor) / 2
'                        Else
'                            .PrtData(i).alpha = Particle_Stream(.type).AlphaInicial
'                            .PrtData(i).alpha_factor = Particle_Stream(.type).alpha_factor + Rnd * Particle_Stream(.type).rnd_alpha_factor
'                        End If
'
'                        .PrtData(i).progress_factor = .PrtData(i).alpha_factor * 0.15
'
'                        .PrtData(i).color = Particle_Stream(.type).StartColor
'
'                        .PrtVertList(i).v.x = .PrtData(i).v.x + offset_map.x
'                        .PrtVertList(i).v.y = .PrtData(i).v.y + offset_map.y
'
'                        .PrtVertList(i).color = D3DColorMake(.PrtData(i).color.r, .PrtData(i).color.g, .PrtData(i).color.b, .PrtData(i).alpha)
'                        c = 1
'                Else
'                    c = 0
'                    .muere = 1
'                End If
'            End If
'        Next i
'
'        If .muere = 1 Then
'            If c = 0 Then
'                ii = 0
'                .killable = 1
'            End If
'        End If


Private Sub Long2RGB(LongCol As Long, r As Single, g As Single, b As Single)
b = (LongCol And &HFF) / 255 'ESTO NO VA ACÁ, MOVER A MODCOMMONMATH
g = ((LongCol And &HFF00) \ &H100) / 255
r = ((LongCol And &HFF0000) \ &H10000) / 255
End Sub

