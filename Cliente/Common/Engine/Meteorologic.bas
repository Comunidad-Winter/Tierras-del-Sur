Attribute VB_Name = "Engine_Meteorologic"
'ESTE ARCHIVO ESTÁ COMPARTIDO!

Option Explicit

Public world_wind   As D3DVECTOR2
Public wind_act     As Boolean

Public raining      As Boolean

Public Enum Tipos_Clima
    climalluvia = 2
    ClimaNeblina = 1
    ClimaNiebla = 4
    ClimaTormenta_de_arena = 8
    ClimaNublado = 16
    ClimaNieve = 32
    ClimaRayos_de_luz = 64
End Enum

Public base_light   As Long
Public base_light_techo   As Long
Public day_r_old    As Byte
Public day_g_old    As Byte
Public day_b_old    As Byte

Public ambient      As D3DCOLORVALUE

Public estado_time  As Byte

Private trueno      As Boolean
Private trueno_tick As Byte
Private backup_color As Long

Type luzxhora
    r As Integer
    g As Integer
    b As Integer
End Type

Public luz_dia(0 To 24) As luzxhora '¬¬ la hora 24 dura 1 minuto entre las 24 y las 0

Public modR As Single, modG As Single, modB As Single


Public tamColor(3)      As Long, outVecSol As D3DVECTOR

Public color_mod_day    As D3DCOLORVALUE
Public color_mod_day_16 As BGRCOLOR_DLL
Public color_mod_c      As D3DCOLORVALUE
Public base_color       As RGBCOLOR
Public color_mod_day_argb As ARGB_COLOR

Public lights2act As Boolean

Public meteo_particle As Engine_Particle_Group

Private MultiploColorNublado As New clsAlpha

Public Forzar_Dia As Boolean

Public AlphaNiebla As New clsAlpha
Public AlphaArena As New clsAlpha

Public HoraDelDia As Single
Public fraccionDelDia As Byte




'DWORD skyBtmColors[] = {0xFF303E57, 0xFFAC7963, 0xFFCAD7DB};
'
'int seq[]={0, 0, 1, 2, 2, 2, 1, 0, 0};


Public Function update_day_light() As D3DCOLORVALUE
   Dim seq_id%
    Dim seq_residue!
    Static timea As Single
    Dim timex As Single
    Dim col1 As D3DCOLORVALUE, col2 As D3DCOLORVALUE
    Dim cOut As D3DCOLORVALUE
    
    Dim seq()
    seq() = Array(1, 1, 1, 2, 2, 2, 1, 1, 1, 1)
    
    Dim skyBtmColors(2) As D3DCOLORVALUE
    skyBtmColors(0).a = 255
    skyBtmColors(1).a = 255
    skyBtmColors(2).a = 255
    
    

    skyBtmColors(0).r = 50
    skyBtmColors(0).g = 50
    skyBtmColors(0).b = 50

    skyBtmColors(1).r = 100
    skyBtmColors(1).g = 100
    skyBtmColors(1).b = 100

    skyBtmColors(2).r = 155
    skyBtmColors(2).g = 155
    skyBtmColors(2).b = 155
    
    Dim PosicionesSol(2) As D3DVECTOR
    
    With PosicionesSol(0)
        .Y = -40
        .X = -40
        .z = 1
    End With

    With PosicionesSol(1)
        .Y = 0
        .X = -40
        .z = 1
    End With

    With PosicionesSol(2)
        .Y = 40
        .X = -30
        .z = 1
    End With
    
    
    'timex = Hour(Time) + Minute(Time) / 60 + Second(Time) / 6000
'#If esCLIENTE = 1 Then
'    timex = (1 - 1) + Minutos / 60 '+ Second(Time) / 6000
'#Else
'    #If esMe Then
     timex = Abs((fraccionDelDia - 1) / 96) * 24
'    #Else
        'timea = timea + timerTicksPerFrame * 0.005
        'If timea > 24 Then timea = timea Mod 24
    
'        timex = timea 'Hour(Time)
'    #End If

'#End If
'    timea = timea + timerTicksPerFrame * 0.05
'    If timea > 24 Then timea = timea Mod 24
'
'    timex = timea 'Hour(Time)


    HoraDelDia = timex

    Call DLL_Luces.SetFraccionDia(((timex / 24) * 96) + 1)
    
    seq_id = timex \ 3
    seq_residue = timex / 3 - seq_id

    col1 = skyBtmColors(seq(seq_id))
    col2 = skyBtmColors(seq(seq_id + 1))
    
    D3DXColorLerp update_day_light, col1, col2, seq_residue
    D3DXVec3Lerp outVecSol, PosicionesSol(seq(seq_id)), PosicionesSol(seq(seq_id + 1)), seq_residue
    
    With update_day_light
        If .r > 255 Then .r = 255
        If .g > 255 Then .g = 255
        If .b > 255 Then .b = 255
        If .r < 0 Then .r = 0
        If .g < 0 Then .g = 0
        If .b < 0 Then .b = 0
        
        If Forzar_Dia = True Then
            .r = 180
            .g = 180
            .b = 180
        End If
    End With
End Function

Public Sub Cambiar_estado_climatico(ByVal Nuevo_estado As Tipos_Clima)
    
    If Not estado_time = Nuevo_estado Then
        If Not meteo_particle Is Nothing Then
            meteo_particle.Matar 1
        End If
    End If

    estado_time = Nuevo_estado

    If (estado_time And Tipos_Clima.climalluvia) And meteo_particle Is Nothing Then
        Set meteo_particle = New Engine_Particle_Group
        meteo_particle = PARTICULAS_LLUVIA
    ElseIf (estado_time And Tipos_Clima.ClimaNieve) And meteo_particle Is Nothing Then
        Set meteo_particle = New Engine_Particle_Group
        meteo_particle = PARTICULAS_NIEVE
    Else
        If Not meteo_particle Is Nothing Then
            meteo_particle.Matar 1
        End If
    End If

    'If (estado_time And Tipos_Clima.ClimaNublado) Or (estado_time And Tipos_Clima.ClimaLluvia) Then
   '     MultiploColorNublado.value = 200
    'Else
    '    MultiploColorNublado.value = 255
   ' End If
End Sub

Public Sub cron_fxs()
If trueno = True Then
    If trueno_tick >= 2 Then
        trueno = False
        base_light = backup_color
    Else
        init_trueno trueno_tick + 1
    End If
End If
End Sub

Public Function cron_tiempo() As Boolean
    Static LastUpdate As Long
    If LastUpdate + 32 < GetTimer Then
        LastUpdate = GetTimer
        
            Dim Hora As Byte
            Dim hacer As Boolean
            Hora = Hour(Time)
            
            hacer = change_day_effect()
            If hacer = True Then
                Light_Update_Map = True
                NormalesMapaNececitanActualizar = True
            End If
            
            #If esMe = 1 Then
                If frmMain.chkAnimarDia.value = vbChecked Then
                    NormalesMapaNececitanActualizar = True
                End If
            #End If
            
            If (estado_time And Tipos_Clima.climalluvia) Then
                If meteo_particle Is Nothing Then
                    Set meteo_particle = New Engine_Particle_Group
                    meteo_particle = PARTICULAS_LLUVIA
                Else
                    If meteo_particle.PGID <> PARTICULAS_LLUVIA Then
                        meteo_particle.Matar 1
                    End If
                End If
            ElseIf (estado_time And Tipos_Clima.ClimaNieve) Then
                If meteo_particle Is Nothing Then
                    Set meteo_particle = New Engine_Particle_Group
                    meteo_particle = PARTICULAS_NIEVE
                Else
                    If meteo_particle.PGID <> PARTICULAS_NIEVE Then
                        meteo_particle.Matar 1
                    End If
                End If
            Else
                If Not CBool((estado_time And Tipos_Clima.ClimaNieve) Or (estado_time And Tipos_Clima.climalluvia)) Then
                    If Not meteo_particle Is Nothing Then
                        meteo_particle.Matar 1
                    End If
                End If
            End If
            
            If (estado_time And Tipos_Clima.ClimaNublado) Or (estado_time And Tipos_Clima.climalluvia) Then
                MultiploColorNublado.value = 200
            Else
                MultiploColorNublado.value = 255
            End If
            
        'TODO: degradé de fogs
            
            If (estado_time And Tipos_Clima.ClimaNiebla) Or (estado_time And Tipos_Clima.ClimaNeblina) Then
                If estado_time And Tipos_Clima.ClimaNublado Then
                    AlphaNiebla.value = 128
                Else
                    AlphaNiebla.value = 64
                End If
            Else
                AlphaNiebla.value = 0
            End If
            
            If estado_time And Tipos_Clima.ClimaTormenta_de_arena Then
                AlphaArena.value = 200
            Else
                AlphaArena.value = 0
            End If
            
            If estado_time And Tipos_Clima.ClimaRayos_de_luz Then
                Lightbeam_do = 10
            End If
            
            If Not meteo_particle Is Nothing Then
                meteo_particle.SetPos UserPos.X, UserPos.Y
            End If
            
            cron_tiempo = hacer
    End If
End Function

Public Sub init_trueno(Optional ByVal tick As Byte = 0)
    'TODO:Nuevos truenos
End Sub

Private Function change_day_effect() As Boolean
    'On Error GoTo ehand
    Static lastca As Single, ll!, l!
    Dim r%, g%, b%, oc As D3DCOLORVALUE, tmpFloat!
    
    If mapinfo.ColorPropio Then
        update_day_light
        oc.r = mapinfo.BaseColor.r
        oc.g = mapinfo.BaseColor.g
        oc.b = mapinfo.BaseColor.b
    Else
        oc = update_day_light
    End If
    
    If mapinfo.puede_nublado Then
        tmpFloat = Round(MultiploColorNublado.value) / 255
        r = oc.r * tmpFloat
        g = oc.g * tmpFloat
        b = oc.b * tmpFloat
    Else
        r = oc.r
        g = oc.g
        b = oc.b
    End If

    change_day_effect = Not (r = color_mod_c.r And g = color_mod_c.g And b = color_mod_c.b)
    
    If change_day_effect = True Then
        color_mod_day.r = r / 255
        color_mod_day.g = g / 255
        color_mod_day.b = b / 255
        base_color.r = day_r_old * color_mod_day.r
        base_color.g = day_g_old * color_mod_day.g
        base_color.b = day_b_old * color_mod_day.b
        color_mod_c.r = r
        color_mod_c.g = g
        color_mod_c.b = b
        color_mod_day_16.r = r
        color_mod_day_16.g = g
        color_mod_day_16.b = b
        color_mod_day_argb.r = r
        color_mod_day_argb.g = g
        color_mod_day_argb.b = b
        color_mod_day_argb.a = 255
        
        MoverSol outVecSol.X, outVecSol.Y, outVecSol.z
        
        Light_Update_Map = True
        base_light = D3DColorXRGB(r, g, b)
        base_light_techo = D3DColorXRGB(r, g, b) And &HFFFFFF
    End If
    
    Exit Function
ehand:
    LogError "Error en change_day_effect t="
End Function

Public Sub setup_ambient()
    AlphaNiebla.Speed = 30000
    AlphaArena.value = 15000
    MultiploColorNublado.Speed = 256000
    MultiploColorNublado.InitialValue = 255
End Sub

