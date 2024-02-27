Attribute VB_Name = "Engine_Sonido"
' ESTE ARCHIVO ESTA COMPARTIDO POR TODOS LOS PROGRAMAS.

Option Explicit

Private Type SOUND_DB_ENTRY
    active          As Byte
    FileName        As Integer
    UltimoAcceso    As Long
    Sample          As Long
    Channel         As Long
    size            As Long 'Memoria
End Type

Private Tabla() As Integer
Private TablaMax As Integer

Private DB_SoundMem() As SOUND_DB_ENTRY

Private pSoundCant      As Integer
Private pSoundMax       As Integer
Private pSoundLast      As Integer
Private pSoundMaxMemory As Long

'Private Declare Function GetTickCount Lib "kernel32" () As Long

Private hBass           As Boolean

Public pakSonidos       As clsEnpaquetado

Private Type SonidoAmbiente
    active As Byte

    Sonido As Integer
    stream As Long
    
    tick1 As Long
    vol1 As Single
    tick2 As Long
    vol2 As Long
    
    Matar As Byte 'cuando tick_actual > tick2 mata al sonido y a la estructura
End Type


Public Sub Sonido_DeInit()
'Marce On error resume next
    Dim i As Long
    
    For i = 1 To pSoundLast
        If DB_SoundMem(i).Sample Then
            Call BASS_SampleFree(DB_SoundMem(i).Sample)
        End If
    Next i
    
    BASS_Free
    
    Erase DB_SoundMem
    
End Sub

Public Function Sonido_Init(ByVal MaxMemory As Long, ByVal MaxEntries As Long, PakPath As String) As Boolean
'Marce On error resume next
    pSoundMax = MaxEntries
    
    If pSoundMax < 1 Then 'por lo menos 1 gráfico
        Exit Function
    End If
    
    ReDim DB_SoundMem(pSoundMax)
    pSoundLast = 0
    
    pSoundMaxMemory = MaxMemory
    
    Set pakSonidos = New clsEnpaquetado
    pakSonidos.Cargar PakPath

    ' check the correct BASS was loaded
    If (HiWord(BASS_GetVersion) = BASSVERSION) Then
    
    Else
        Call MsgBox("Version incorrecta de BASS.dll", vbCritical)
        Call LogError("Version incorrecta de BASS.dll")
        Exit Function
    End If

    If BASS_Init(-1, 44100, BASS_DEVICE_3D, frmMain.hWnd, 0) Then
        hBass = True
    Else
        Call LogError("No se puede iniciar el dispositivo de sonido.")
        hBass = False
        Exit Function

    End If
    
    Call BASS_Set3DFactors(1, 1, 1)
    
    ReDim Tabla(600)
    TablaMax = 600
    
    Sonido_Init = True

End Function

Public Sub Sonido_BorrarTodo()
'Marce On error resume next
    Dim i As Long
    
    For i = 1 To pSoundLast
        If DB_SoundMem(i).Sample Then
            Call BASS_SampleFree(DB_SoundMem(i).Sample)
        End If
        DB_SoundMem(i).active = 0
    Next i
    
    ReDim Tabla(600)
    TablaMax = 600
    ReDim DB_SoundMem(0)
    pSoundLast = 0
End Sub

Public Function Sonido_PlayEX(ByVal Sonido As Long, Optional ByVal looping As Boolean, Optional ByVal volume As Single = 1, Optional ByVal pan As Single = 0) As Long

Dim chn     As Long
Dim Sample  As Long
Dim Index   As Integer
Dim info    As BASS_SAMPLE

If hBass Then

    If TablaMax >= Sonido Then
        If Tabla(Sonido) Then
            Index = Tabla(Sonido)
        Else
           If Sonido_Load(Sonido, Index) = False Then Exit Function
        End If
    Else
        If Sonido_Load(Sonido, Index) = False Then Exit Function
    End If
    If Index Then
        With DB_SoundMem(Index)
            If .Sample And .active = 1 Then
                BASS_SampleGetInfo .Sample, info
                                
                .Channel = BASS_SampleGetChannel(.Sample, BASSFALSE)

                Call BASS_ChannelSetAttribute(.Channel, BASS_ATTRIB_VOL, volume)
                Call BASS_ChannelSetAttribute(.Channel, BASS_ATTRIB_PAN, pan)

                If looping Then
                    Call BASS_ChannelFlags(.Channel, BASS_SAMPLE_LOOP, BASS_SAMPLE_LOOP)
                Else
                    Call BASS_ChannelFlags(.Channel, 0, BASS_SAMPLE_LOOP)
                End If
                
                Call BASS_ChannelPlay(.Channel, BASSTRUE)
                
                Sonido_PlayEX = .Channel
                
            End If
        End With
    End If
End If
End Function

Public Function Sonido_Play(ByVal Sonido As Long) As Long

Dim chn     As Long
Dim Sample  As Long
Dim Index   As Integer
Dim info    As BASS_SAMPLE

If Not EfectosSonidoActivados Then Exit Function

If hBass Then

    If TablaMax >= Sonido Then
        If Tabla(Sonido) Then
            Index = Tabla(Sonido)
        Else
            If Sonido_Load(Sonido, Index) = False Then Exit Function
        End If
    Else
        If Sonido_Load(Sonido, Index) = False Then Exit Function
    End If
    
    If Index > 0 Then
        With DB_SoundMem(Index)
            If .Sample And .active = 1 Then
                .Channel = BASS_SampleGetChannel(.Sample, BASSFALSE)

                Call BASS_ChannelSetAttribute(.Channel, BASS_ATTRIB_VOL, VolumenF)

                Call BASS_ChannelPlay(.Channel, BASSFALSE)
                Sonido_Play = .Channel
            End If
        End With
    End If
End If
End Function

Public Sub Sonido_Stop(ByVal Sonido As Long)
Dim chn     As Long
Dim Sample  As Long
Dim Index   As Integer
If hBass Then
    If TablaMax >= Sonido Then
        If Tabla(Sonido) Then
            Index = Tabla(Sonido)
            With DB_SoundMem(Index)
                If .Sample Then
                    If .Channel Then Call BASS_ChannelStop(.Channel)
                    .Channel = BASS_SampleGetChannel(.Sample, BASSFALSE)
                    If .Channel Then Call BASS_ChannelStop(.Channel)
                End If
            End With
        End If
    End If
End If
End Sub

Private Function Sonido_BorrarMenosUsado() As Integer
    Dim valor As Long
    Dim i As Long
    'Inicializamos todo
    valor = GetTimer() 'DB_SoundMem(1).UltimoAcceso
    Sonido_BorrarMenosUsado = 1
    
    'Buscamos cual es el que lleva más tiempo sin ser utilizado
    For i = 1 To pSoundLast
        If DB_SoundMem(i).UltimoAcceso < valor Then
            valor = DB_SoundMem(i).UltimoAcceso
            Sonido_BorrarMenosUsado = i
        End If
    Next i
    
    'Disminuimos el contador
    pSoundCant = pSoundCant - 1
    
    'Borramos

    If DB_SoundMem(i).Sample Then
        Call BASS_SampleFree(DB_SoundMem(i).Sample)
        DB_SoundMem(i).active = 0
    End If

    Tabla(DB_SoundMem(Sonido_BorrarMenosUsado).FileName) = 0
    pSoundMaxMemory = pSoundMaxMemory - DB_SoundMem(Sonido_BorrarMenosUsado).size
End Function

Private Function Sonido_BuscarVacio() As Integer
    Dim i As Integer
    
    If pSoundCant < pSoundMax Then 'nos aseguramos de que haya espacio
        For i = 1 To pSoundMax
            If DB_SoundMem(i).active = 0 Then
                Sonido_BuscarVacio = i
                Exit Function
            End If
        Next i
    End If
    
    Sonido_BuscarVacio = -1
End Function

Private Function Sonido_Load(ByVal numero As Integer, ByRef db_index As Integer) As Boolean
    Dim Data()  As Byte
    Dim Ptr     As Long
    Dim size    As Long
    Dim Offset  As Long
    
    Dim hs      As Long
    
    Dim i       As Integer
    
    
    'obtener data() o Ptr y size
    If pakSonidos.Leer(numero, Data) Then
        Ptr = VarPtr(Data(0))
        
        If TablaMax >= numero Then
            If Tabla(numero) Then
                If DB_SoundMem(Tabla(numero)).Sample Then
                    Call BASS_SampleFree(DB_SoundMem(Tabla(numero)).Sample)
                End If
                DB_SoundMem(Tabla(numero)).active = 0
            End If
        Else
            TablaMax = TablaMax + 20
            ReDim Preserve Tabla(TablaMax)
        End If
        
        If pSoundCant = pSoundMax Then
            i = Sonido_BorrarMenosUsado
        Else
            i = Sonido_BuscarVacio
            If i = 0 Or i = -1 Then
                MsgBox "Error. No se encontraron espacios vacios para el sonido.", vbCritical, "CATASTROFE!"
                Exit Function
            End If
        End If
        
        hs = BASS_SampleLoad(BASSTRUE, Ptr, 0, pakSonidos.Cabezal_GetFileSize(numero), 3, BASS_SAMPLE_OVER_POS)
        
        If hs Then
            With DB_SoundMem(i)
                .active = 1
                .Sample = hs
                .FileName = numero
                .UltimoAcceso = GetTimer
                .size = size
            End With
            db_index = i
            Tabla(numero) = i
            Sonido_Load = True
        Else
            LogError "Engine_Sonido::Load Error: code: " & BASS_ErrorGetCode()
        End If
    End If
End Function

Public Sub Sonido_Set_Volumen(ByVal Sonido As Long, ByVal vol As Single)
    Call modBass.BASS_ChannelSetAttribute(Sonido, BASS_ATTRIB_VOL, vol)
End Sub
