Attribute VB_Name = "Admin"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Public Type tMotd
    texto As String
    Formato As String
End Type

Public MaxLines As Integer
Public MOTD() As tMotd

Public Type tAPuestas
    Ganancias As Long
    Perdidas As Long
    Jugadas As Long
End Type
Public Apuestas As tAPuestas

Public NPCs As Long
Public DebugSocket As Boolean

Public Horas As Long
Public Dias As Long
Public MinsRunning As Long

Public ReiniciarServer As Long

Public tInicioServer As Long
Public EstadisticasWeb As New clsEstadisticasIPC

Public SanaIntervaloSinDescansar As Integer
Public StaminaIntervaloSinDescansar As Integer
Public SanaIntervaloDescansar As Integer
Public StaminaIntervaloDescansar As Integer
Public IntervaloSed As Integer
Public IntervaloHambre As Integer
Public IntervaloVeneno As Integer
Public IntervaloParalizado As Integer
Public IntervaloInvisible As Integer
Public IntervaloFrio As Integer
Public IntervaloWavFx As Integer
Public IntervaloLanzaHechizo As Integer
Public IntervaloNPCPuedeAtacar As Integer
Public IntervaloNPCAI As Integer
Public IntervaloInvocacion As Integer
Public IntervaloUserPuedeAtacar As Long
Public IntervaloUserPuedeCastear As Long
Public IntervaloUserPuedeTrabajar As Long
Public IntervaloParaConexion As Long
Public IntervaloCerrarConexion As Long '[Gonzalo]
Public IntervaloUserPuedeUsar As Long
Public IntervaloAutoReiniciar As Long   'segundos


Public MinutosWs As Long
Public Puerto As Integer

Public MAXPASOS As Long

Public BootDelBackUp As Byte
Public Lloviendo As Boolean
Public DeNoche As Boolean

Public IpList As New Collection
Public ClientsCommandsQueue As Byte

Public Type TCPESStats
    BytesEnviados As Double
    BytesRecibidos As Double
    BytesEnviadosXSEG As Long
    BytesRecibidosXSEG As Long
    BytesEnviadosXSEGMax As Long
    BytesRecibidosXSEGMax As Long
    BytesEnviadosXSEGCuando As Date
    BytesRecibidosXSEGCuando As Date
End Type

Public TCPESStats As TCPESStats

'Public ResetThread As New clsThreading

Function VersionOK(ByVal Ver As String) As Boolean
VersionOK = (Ver = ULTIMAVERSION)
End Function

Public Function VersionesActuales(ByVal v1 As Integer, ByVal v2 As Integer, ByVal v3 As Integer, ByVal v4 As Integer, ByVal v5 As Integer, ByVal v6 As Integer, ByVal v7 As Integer) As Boolean


End Function


Public Function ValidarLoginMSG(ByVal N As Integer) As Integer
On Error Resume Next
Dim AuxInteger As Integer
Dim AuxInteger2 As Integer
AuxInteger = SD(N)
AuxInteger2 = SDM(N)
ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function


Sub ReSpawnOrigPosNpcs()
On Error Resume Next

Dim i As Integer
Dim MiNPC As npc
   
For i = 1 To LastNPC
   'OJO
   If Npclist(i).flags.NPCActive Then
        
        If InMapBounds(Npclist(i).Orig.Map, Npclist(i).Orig.X, Npclist(i).Orig.Y) And Npclist(i).Numero = Guardias Then
                MiNPC = Npclist(i)
                Call QuitarNPC(i)
                Call ReSpawnNpc(MiNPC)
        End If
        
        'tildada por sugerencia de yind
        'If Npclist(i).Contadores.TiempoExistencia > 0 Then
        '        Call MuereNpc(i, 0)
        'End If
   End If
   
Next i

End Sub

Sub WorldSave()
On Error Resume Next
'Call LogTarea("Sub WorldSave")

Dim loopX As Integer
Dim Porc As Long

Call SendData(ToAll, 0, 0, "Y30")

Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales

Dim j As Integer, k As Integer

For j = 1 To NumMaps
    If MapInfo(j).BackUp = 1 Then k = k + 1
Next j

FrmStat.ProgressBar1.Min = 0
FrmStat.ProgressBar1.max = k
FrmStat.ProgressBar1.Value = 0

For loopX = 1 To NumMaps
    'DoEvents
    
    If MapInfo(loopX).BackUp = 1 Then
    
            Call SaveMapData(loopX)
            FrmStat.ProgressBar1.Value = FrmStat.ProgressBar1.Value + 1
    End If

Next loopX

FrmStat.Visible = False

If FileExist(DatPath & "\bkNpc.dat", vbNormal) Then Kill (DatPath & "bkNpc.dat")
If FileExist(DatPath & "\bkNPCs-HOSTILES.dat", vbNormal) Then Kill (DatPath & "bkNPCs-HOSTILES.dat")

For loopX = 1 To LastNPC
    If Npclist(loopX).flags.BackUp = 1 Then
            Call BackUPnPc(loopX)
    End If
Next

Call SendData(ToAll, 0, 0, "Y31")

End Sub

Public Sub PurgarPenas()
Dim i As Integer
For i = 1 To LastUser
    If UserList(i).flags.UserLogged Then
    
        If UserList(i).Counters.Pena > 0 Then
                
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    Call WarpUserChar(i, Libertad.Map, Libertad.X, Libertad.Y, True)
                    Call SendData(ToIndex, i, 0, "Y32")
                End If
                
        End If
        
    End If
Next i
End Sub


Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal Minutos As Long, Optional ByVal GmName As String = "")
        
        UserList(UserIndex).Counters.Pena = Minutos
       
        
        Call WarpUserChar(UserIndex, Prision.Map, Prision.X, Prision.Y, True)
        
        If GmName = "" Then
            Call SendData(ToIndex, UserIndex, 0, "||Has sido encarcelado, deberas permanecer en la carcel " & Minutos & " minutos." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||" & GmName & " te ha encarcelado, deberas permanecer en la carcel " & Minutos & " minutos." & FONTTYPE_INFO)
        End If
        
End Sub


Public Sub BorrarUsuario(ByVal UserName As String)
On Error Resume Next
If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
    Kill CharPath & UCase$(UserName) & ".chr"
End If
End Sub

Public Function BANCheck(ByVal Name As String) As Boolean

BANCheck = (val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban")) = 1) 'Or _
(val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "AdminBan")) = 1)

End Function

Public Function PersonajeExiste(ByVal Name As String) As Boolean

PersonajeExiste = FileExist(CharPath & UCase$(Name) & ".chr", vbNormal)

End Function

Public Function UnBan(ByVal Name As String) As Boolean
'Unban the character
Call WriteVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban", "0")

'Remove it from the banned people database
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "BannedBy", "NOBODY")
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Reason", "NO REASON")
End Function

Public Function MD5ok(ByVal md5formateado As String) As Boolean
Dim i As Integer

If MD5ClientesActivado = 1 Then
    For i = 0 To UBound(MD5s)
        If (md5formateado = MD5s(i)) Then
            MD5ok = True
            Exit Function
        End If
    Next i
    MD5ok = False
Else
    MD5ok = True
End If

End Function

Public Sub MD5sCarga()
Dim LoopC As Integer

MD5ClientesActivado = val(GetVar(IniPath & "Server.ini", "MD5Hush", "Activado"))

If MD5ClientesActivado = 1 Then
    ReDim MD5s(val(GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptados")))
    For LoopC = 0 To UBound(MD5s)
        MD5s(LoopC) = GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptado" & (LoopC + 1))
        MD5s(LoopC) = txtOffset(hexMd52Asc(MD5s(LoopC)), 53)
    Next LoopC
End If

End Sub

Public Sub BanIpAgrega(ByVal ip As String)
BanIps.Add ip

Call BanIpGuardar
End Sub

Public Function BanIpBuscar(ByVal ip As String) As Long
Dim Dale As Boolean
Dim LoopC As Long

Dale = True
LoopC = 1
Do While LoopC <= BanIps.Count And Dale
    Dale = (BanIps.Item(LoopC) <> ip)
    LoopC = LoopC + 1
Loop

If Dale Then
    BanIpBuscar = 0
Else
    BanIpBuscar = LoopC - 1
End If
End Function

Public Function BanIpQuita(ByVal ip As String) As Boolean

On Error Resume Next

Dim N As Long

N = BanIpBuscar(ip)
If N > 0 Then
    BanIps.Remove N
    BanIpGuardar
    BanIpQuita = True
Else
    BanIpQuita = False
End If

End Function

Public Sub BanIpGuardar()
Dim ArchivoBanIp As String
Dim ArchN As Long
Dim LoopC As Long

ArchivoBanIp = App.Path & "\Dat\BanIps.dat"

ArchN = FreeFile()
Open ArchivoBanIp For Output As #ArchN

For LoopC = 1 To BanIps.Count
    Print #ArchN, BanIps.Item(LoopC)
Next LoopC

Close #ArchN

End Sub

Public Sub BanIpCargar()
Dim ArchN As Long
Dim Tmp As String
Dim ArchivoBanIp As String

ArchivoBanIp = App.Path & "\Dat\BanIps.dat"

Do While BanIps.Count > 0
    BanIps.Remove 1
Loop

ArchN = FreeFile()
Open ArchivoBanIp For Input As #ArchN

Do While Not EOF(ArchN)
    Line Input #ArchN, Tmp
    BanIps.Add Tmp
Loop

Close #ArchN

End Sub

Public Sub ActualizaEstadisticasWeb()

Static Andando As Boolean
Static Contador As Long
Dim Tmp As Boolean

Contador = Contador + 1

If Contador >= 10 Then
    Contador = 0
    Tmp = EstadisticasWeb.EstadisticasAndando()
    
    If Andando = False And Tmp = True Then
        Call InicializaEstadisticas
    End If
    
    Andando = Tmp
End If

End Sub

Public Sub ActualizaStatsES()

Static TUlt As Single
Dim Transcurrido As Single

Transcurrido = Timer - TUlt

If Transcurrido >= 5 Then
    TUlt = Timer
    With TCPESStats
        .BytesEnviadosXSEG = CLng(.BytesEnviados / Transcurrido)
        .BytesRecibidosXSEG = CLng(.BytesRecibidos / Transcurrido)
        .BytesEnviados = 0
        .BytesRecibidos = 0
        
        If .BytesEnviadosXSEG > .BytesEnviadosXSEGMax Then
            .BytesEnviadosXSEGMax = .BytesEnviadosXSEG
            .BytesEnviadosXSEGCuando = CDate(Now)
        End If
        
        If .BytesRecibidosXSEG > .BytesRecibidosXSEGMax Then
            .BytesRecibidosXSEGMax = .BytesRecibidosXSEG
            .BytesRecibidosXSEGCuando = CDate(Now)
        End If
        
        If frmEstadisticas.Visible Then
            Call frmEstadisticas.ActualizaStats
        End If
    End With
End If

End Sub


Public Function UserDarPrivilegioLevel(ByVal Name As String) As Long
If EsDios(Name) Then
    UserDarPrivilegioLevel = 3
ElseIf EsSemiDios(Name) Then
    UserDarPrivilegioLevel = 2
ElseIf EsConsejero(Name) Then
    UserDarPrivilegioLevel = 1
Else
    UserDarPrivilegioLevel = 0
End If
End Function

