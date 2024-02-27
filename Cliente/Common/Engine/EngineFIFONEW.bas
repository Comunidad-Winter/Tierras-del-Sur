Attribute VB_Name = "Engine_FIFO"
'                  ____________________________________________
'                 /_____/  http://www.arduz.com.ar/ao/   \_____\
'                //            ____   ____   _    _ _____      \\
'               //       /\   |  __ \|  __ \| |  | |___  /      \\
'              //       /  \  | |__) | |  | | |  | |  / /        \\
'             //       / /\ \ |  _  /| |  | | |  | | / /   II     \\
'            //       / ____ \| | \ \| |__| | |__| |/ /__          \\
'           / \_____ /_/    \_\_|  \_\_____/ \____//_____|_________/ \
'           \________________________________________________________/

Option Explicit


Public Type map_corners
    left As Integer
    top As Integer
    right As Integer
    bottom As Integer
End Type
Public Corners As map_corners

Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef dest As Any, ByVal numbytes As Long)


'Private Declare Sub CalcularNormal Lib "MZEngine.dll" (ByVal A0 As Single, ByVal A1 As Single, ByVal A2 As Single, ByVal A3 As Single, ByRef C1 As D3DVECTOR, ByRef C2 As D3DVECTOR)

Public pakMapas As clsEnpaquetado

Public Sub CargarPakMapas(Path As String)
    Set pakMapas = New clsEnpaquetado

    If Not FileExist(Path) Then
    '    pakMapas.CrearVacio Path, 2000
        MsgBox "No se pudo cargar el archivo de mapas. Por favor reinstale el juego descargandolo desde www.tierrasdelsur.cc"
        End
    Else
        pakMapas.Cargar Path
    End If
   
End Sub

Sub SwitchMap(ByVal map As Integer)
Static timerxx As New clsPerformanceTimer
Dim IH As INFOHEADER

timerxx.Time

Call LimpiarMapa

'If objMapManager.Getstored_map(map, CurMap) = False Then

   ' If pakMapas.Cabezal_GetFileSize(map) Then
   '     pakMapas.IH_Get map, IH
        Cargar_Mapa_CLI app.Path & "/Recursos/Mapas/" & map & ".clientmap", 0
       ' Cargar_Mapa_CLI pakMapas.Path_res, IH.EmpiezaByte
   ' End If
    
    CurMap = map
    Debug.Print "DESDE DISCO DURO TARDO EN CARGAR MAPA:"; timerxx.Time
'Else
'    Debug.Print "DESDE MEMORIA TARDO EN CARGAR MAPA:"; timerxx.Time
'End If

End Sub

Public Function encode_decode_text(text As String, ByVal off As Integer, Optional ByVal cript As Byte, Optional ByVal Encode As Byte) As String
    Dim i As Integer, l As String
    If Encode Then off = 256 - off
    Dim ba() As Byte, bo() As Byte
    Dim lenn%
    ba = StrConv(text, vbFromUnicode)
    lenn = UBound(ba)
    ReDim bo(0 To lenn)
    For i = 0 To lenn
       bo(i) = ((ba(i) Xor cript) + off) Mod 256 Xor cript
    Next i
    encode_decode_text = StrConv(bo, vbUnicode)
End Function

Sub LimpiarMapa()
On Error GoTo 0
    Dim Y As Long
    Dim X As Long

    Dim tt%
    
    Dim ResizeBackBufferX As Integer
    Dim ResizeBackBufferY As Integer

    DLL_Luces.Remove_All

    FX_Projectile_Erase_All
    FX_Hit_Erase_All
    
    For Y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
        For X = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE

            Set mapdata(X, Y).Particles_groups(0) = Nothing
            Set mapdata(X, Y).Particles_groups(1) = Nothing
            Set mapdata(X, Y).Particles_groups(2) = Nothing

            If CharMap(X, Y) > 0 Then
                tt = CharMap(X, Y)
                Call EraseChar(tt)
                CharMap(X, Y) = 0
            End If
            mapdata(X, Y).ObjGrh.GrhIndex = 0
        Next X
    Next Y
End Sub





Sub jojoparticulas()
CargarParticle_Streams
Dim i As Long
Dim X As Integer
Dim Y As Integer

For i = 0 To 10
    X = UserPos.X - Rnd * 15 + Rnd * 15
    Y = UserPos.Y - Rnd * 10 + Rnd * 10
    'Engine_Landscape.Light_Create X, Y, 255, 200, 60, 5, 3
    'Engine_Particles.Particle_Group_Make 0, X, Y, 16
Next i
    'Engine_Particles.Particle_Group_Make 1, UserPos.x, UserPos.y, 12, 0

End Sub

