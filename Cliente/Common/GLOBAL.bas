Attribute VB_Name = "modGlobal"
'____________________________________________
'                 /_____/  http://www.arduz.com.ar/ao/   \_____\
'                //            ____   ____   _    _ _____      \\
'               //       /\   |  __ \|  __ \| |  | |___  /      \\
'              //       /  \  | |__) | |  | | |  | |  / /        \\
'             //       / /\ \ |  _  /| |  | | |  | | / /   II     \\
'            //       / ____ \| | \ \| |__| | |__| |/ /__          \\
'           / \_____ /_/    \_\_|  \_\_____/ \____//_____|_________/ \
'           \________________________________________________________/


#Const Debuging = 0
#Const TimerPerformance = 1

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Declare Function QueryPerformanceCounter Lib "kernel32" (X As Currency) As Boolean
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (X As Currency) As Boolean

Private Declare Function GetTickCount Lib "kernel32" () As Long


#If LOCALHOST = 0 Then
     Public Const WEB_API = "https://api.tierrasdelsur.cc"
#Else
    Public Const WEB_API = "localhost:8080"
#End If
        

Public Sub checkTimers()
    If GetTickCount < 0 Or timeGetTime < 0 Then
        MsgBox "Tu pc está prendida hace mucho tiempo (más de " & Round((4294967296# + GetTickCount) / (3600000), 1) & " horas). Por favor reinicia la pc. Si tenes una notebook, apagala. Asegurate que no quede en modo hibernación."
        End
    End If
    GetTimer
End Sub

Public Function GetTimer(Optional ByVal usarCacheado As Boolean = False) As Double
'Importantes Para windows XP y AMD, deberiamos detectar si es xp y si es amd y usar otro timer:
'-http://support.microsoft.com/kb/895980
'-http://support.microsoft.com/kb/909944
'Para volver atras la movida de las FPS. volver a version 426
Static TimerCacheado As Double

If TimerCacheado = 0 Or usarCacheado = False Then
    #If TimerPerformance = 1 Then
        Static Inicio_De_Los_Tiempos As Currency
        Static freq As Currency
        
        If freq = 0 Then
            QueryPerformanceFrequency freq
            QueryPerformanceCounter Inicio_De_Los_Tiempos
            freq = freq / 1000
        Else
            Dim Tmp As Currency
            
            QueryPerformanceCounter Tmp
            TimerCacheado = (Tmp - Inicio_De_Los_Tiempos) / freq
        End If
    #ElseIf TimerPerformance = 2 Then
        Static Inicio_De_Los_Tiempos As Long
        If Inicio_De_Los_Tiempos = 0 Then
            Inicio_De_Los_Tiempos = timeGetTime And &H7FFFFFFF
            TimerCacheado = 0
        Else
            TimerCacheado = (timeGetTime And &H7FFFFFFF) - Inicio_De_Los_Tiempos
        End If
    #Else
        Static Inicio_De_Los_Tiempos As Long
        If Inicio_De_Los_Tiempos = 0 Then
            Inicio_De_Los_Tiempos = GetTickCount And &H7FFFFFFF
            TimerCacheado = 0
        Else
            TimerCacheado = (GetTickCount And &H7FFFFFFF) - Inicio_De_Los_Tiempos
        End If
    #End If
End If

GetTimer = TimerCacheado

End Function

Public Function GetCfg(app$, master$, key$, Optional default$) As String
'Marce On error resume next
    GetCfg = Xor_String_Cfg(GetSetting(app, master, key, default), 109)
End Function

Public Function SaveCfg(app$, master$, key$, value$) As String
'Marce On error resume next
    Call SaveSetting(app, master, key, Xor_String_Cfg(value, 109))
End Function

Private Function Xor_String_Cfg(ByRef T As String, ByVal Code As Byte) As String
    Dim bytes() As Byte
    bytes = StrConv(T, vbFromUnicode)
    Call Xor_Bytes_Cfg(bytes, Code)
    Xor_String_Cfg = StrConv(bytes, vbUnicode)
End Function

Private Sub Xor_Bytes_Cfg(ByRef ByteArray() As Byte, ByVal Code As Byte)
    Dim i As Integer

    For i = 0 To UBound(ByteArray)
        ByteArray(i) = Code Xor (ByteArray(i) Xor CryptKey)
    Next
End Sub
