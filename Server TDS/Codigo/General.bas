Attribute VB_Name = "General"
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

'Global ANpc As Long
'Global Anpc_host As Long

Global LeerNPCs As New clsLeerInis
Global LeerNPCsHostiles As New clsLeerInis
'********************Misery_Ezequiel 28/05/05********************'
Public conn As ADODB.Connection

Public rs As ADODB.Recordset
Public sql As String
Public db_name As String
Public db_server As String
Public db_port As Integer
Public db_user As String
Public db_pass As String
Public constr As String
Option Explicit

Sub DarCuerpoDesnudo(ByVal UserIndex As Integer, Optional ByVal Mimetizado As Boolean = False)
'[Wizard] Si navega damos el Body de la barca:)
If Not UserList(UserIndex).flags.Navegando Then
Select Case UCase$(UserList(UserIndex).Raza)
    
    Case "HUMANO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 21
                    Else
                        UserList(UserIndex).Char.Body = 21
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 39
                    Else
                        UserList(UserIndex).Char.Body = 39
                    End If
      End Select
    Case "ELFO OSCURO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 32
                    Else
                        UserList(UserIndex).Char.Body = 32
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 40
                    Else
                        UserList(UserIndex).Char.Body = 40
                    End If
      End Select
    Case "ENANO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 53
                    Else
                        UserList(UserIndex).Char.Body = 53
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 60
                    Else
                        UserList(UserIndex).Char.Body = 60
                    End If
      End Select
    Case "GNOMO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 53
                    Else
                        UserList(UserIndex).Char.Body = 53
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 60
                    Else
                        UserList(UserIndex).Char.Body = 60
                    End If
      End Select
    Case Else
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 21
                    Else
                        UserList(UserIndex).Char.Body = 21
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 39
                    Else
                        UserList(UserIndex).Char.Body = 39
                    End If
      End Select
    
End Select
Else
    UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.BarcoObjIndex).Ropaje
End If

UserList(UserIndex).flags.Desnudo = 1
End Sub

Sub Bloquear(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Map As Integer, ByVal X As Integer, ByVal Y As Integer, b As Byte)
'b=1 bloquea el tile en (x,y)
'b=0 desbloquea el tile indicado
Call Senddata(sndRoute, sndIndex, sndMap, "BQ" & X & "," & Y & "," & b)
End Sub

Function HayAgua(Map As Integer, X As Integer, Y As Integer) As Boolean
If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
    If MapData(Map, X, Y).Graphic(1) >= 1505 And _
       MapData(Map, X, Y).Graphic(1) <= 1520 And _
       MapData(Map, X, Y).Graphic(2) = 0 Then
            HayAgua = True
    Else
            HayAgua = False
    End If
Else
  HayAgua = False
End If
End Function

Sub LimpiarMundo()
On Error Resume Next
Dim i As Integer
For i = 1 To TrashCollector.Count
    Dim d As cGarbage
    Set d = TrashCollector(1)
    Call EraseObj(ToMap, 0, d.Map, 1, d.Map, d.X, d.Y)
    Call TrashCollector.Remove(1)
    Set d = Nothing
Next i
End Sub

Sub EnviarSpawnList(ByVal UserIndex As Integer)
Dim k As Integer, SD As String
SD = "SPL" & UBound(SpawnList) & ","

For k = 1 To UBound(SpawnList)
    SD = SD & SpawnList(k).NpcName & ","
Next k

Call Senddata(ToIndex, UserIndex, 0, SD)
End Sub

Sub ConfigListeningSocket(ByRef Obj As Object, ByVal Port As Integer)
#If UsarQueSocket = 0 Then
Obj.AddressFamily = AF_INET
Obj.protocol = IPPROTO_IP
Obj.SocketType = SOCK_STREAM
Obj.Binary = False
Obj.Blocking = False
Obj.BufferSize = 1024
Obj.LocalPort = Port
Obj.backlog = 5
Obj.listen
#End If
End Sub

Sub Main()
On Error Resume Next

If App.PrevInstance Then
    MsgBox "Este programa ya est� corriendo.", vbInformation, "Tirras Del Sur"
End
End If
ULVerUp = GetVar(App.Path & "/Updater.ini", "Ultima", "Fecha")

ChDir App.Path
ChDrive App.Path

Call LoadMotd
Call BanIpCargar

Prision.Map = 66
Libertad.Map = 66

Prision.X = 75
Prision.Y = 47
Libertad.X = 75
Libertad.Y = 65

LastBackup = Format(Now, "Short Time")
Minutos = Format(Now, "Short Time")

ReDim Npclist(1 To MAXNPCS) As npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer
ReDim Parties(1 To MAX_PARTIES) As clsParty

IniPath = App.Path & "\"
DatPath = App.Path & "\Dat\"

LevelSkill(1).LevelValue = 3
LevelSkill(2).LevelValue = 5
LevelSkill(3).LevelValue = 7
LevelSkill(4).LevelValue = 10
LevelSkill(5).LevelValue = 13
LevelSkill(6).LevelValue = 15
LevelSkill(7).LevelValue = 17
LevelSkill(8).LevelValue = 20
LevelSkill(9).LevelValue = 23
LevelSkill(10).LevelValue = 25
LevelSkill(11).LevelValue = 27
LevelSkill(12).LevelValue = 30
LevelSkill(13).LevelValue = 33
LevelSkill(14).LevelValue = 35
LevelSkill(15).LevelValue = 37
LevelSkill(16).LevelValue = 40
LevelSkill(17).LevelValue = 43
LevelSkill(18).LevelValue = 45
LevelSkill(19).LevelValue = 47
LevelSkill(20).LevelValue = 50
LevelSkill(21).LevelValue = 53
LevelSkill(22).LevelValue = 55
LevelSkill(23).LevelValue = 57
LevelSkill(24).LevelValue = 60
LevelSkill(25).LevelValue = 63
LevelSkill(26).LevelValue = 65
LevelSkill(27).LevelValue = 67
LevelSkill(28).LevelValue = 70
LevelSkill(29).LevelValue = 73
LevelSkill(30).LevelValue = 75
LevelSkill(31).LevelValue = 77
LevelSkill(32).LevelValue = 80
LevelSkill(33).LevelValue = 83
LevelSkill(34).LevelValue = 85
LevelSkill(35).LevelValue = 87
LevelSkill(36).LevelValue = 90
LevelSkill(37).LevelValue = 93
LevelSkill(38).LevelValue = 95
LevelSkill(39).LevelValue = 97
LevelSkill(40).LevelValue = 100
LevelSkill(41).LevelValue = 100
LevelSkill(42).LevelValue = 100
LevelSkill(43).LevelValue = 100
LevelSkill(44).LevelValue = 100
LevelSkill(45).LevelValue = 100
LevelSkill(46).LevelValue = 100
LevelSkill(47).LevelValue = 100
LevelSkill(48).LevelValue = 100
LevelSkill(49).LevelValue = 100
LevelSkill(50).LevelValue = 100

ReDim ListaRazas(1 To NUMRAZAS) As String
ListaRazas(1) = "Humano"
ListaRazas(2) = "Elfo"
ListaRazas(3) = "Elfo Oscuro"
ListaRazas(4) = "Gnomo"
ListaRazas(5) = "Enano"

ReDim ListaClases(1 To NUMCLASES) As String
ListaClases(1) = "Mago"
ListaClases(2) = "Clerigo"
ListaClases(3) = "Guerrero"
ListaClases(4) = "Asesino"
ListaClases(5) = "Ladron"
ListaClases(6) = "Bardo"
ListaClases(7) = "Druida"
ListaClases(8) = "Paladin"
ListaClases(9) = "Cazador"
ListaClases(10) = "Pescador"
ListaClases(11) = "Herrero"
ListaClases(12) = "Le�ador"
ListaClases(13) = "Minero"
ListaClases(14) = "Carpintero"
ListaClases(15) = "Sastre"
ListaClases(16) = "Pirata"

ReDim SkillsNames(1 To NUMSKILLS) As String
SkillsNames(1) = "Resistencia M�gica"
SkillsNames(2) = "Magia"
SkillsNames(3) = "Robar"
SkillsNames(4) = "Tacticas de combate"
SkillsNames(5) = "Combate con armas"
SkillsNames(6) = "Meditar"
SkillsNames(7) = "Apu�alar"
SkillsNames(8) = "Ocultarse"
SkillsNames(9) = "Supervivencia"
SkillsNames(10) = "Talar arboles"
SkillsNames(11) = "Comercio"
SkillsNames(12) = "Defensa con escudos"
SkillsNames(13) = "Pesca"
SkillsNames(14) = "Mineria"
SkillsNames(15) = "Carpinteria"
SkillsNames(16) = "Herreria"
SkillsNames(17) = "Liderazgo"
SkillsNames(18) = "Domar animales"
SkillsNames(19) = "Armas de proyectiles"
SkillsNames(20) = "Wresterling"
SkillsNames(21) = "Navegacion"

ReDim UserSkills(1 To NUMSKILLS) As Integer
ReDim UserAtributos(1 To NUMATRIBUTOS) As Integer
ReDim AtributosNames(1 To NUMATRIBUTOS) As String
AtributosNames(1) = "Fuerza"
AtributosNames(2) = "Agilidad"
AtributosNames(3) = "Inteligencia"
AtributosNames(4) = "Carisma"
AtributosNames(5) = "Constitucion"

frmCargando.Show

'Call PlayWaveAPI(App.Path & "\wav\harp3.wav")
frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
ENDL = Chr(13) & Chr(10)
ENDC = Chr(1)
IniPath = App.Path & "\"
CharPath = App.Path & "\Charfile\"
'Bordes del mapa
MinXBorder = XMinMapSize + (XWindow \ 2)
MaxXBorder = XMaxMapSize - (XWindow \ 2)
MinYBorder = YMinMapSize + (YWindow \ 2)
MaxYBorder = YMaxMapSize - (YWindow \ 2)
DoEvents
frmCargando.Label1(2).Caption = "Iniciando Arrays..."

Call LoadGuildsDB
Call CargarSpawnList
Call CargarForbidenWords
'�?�?�?�?�?�?�?� CARGAMOS DATOS DESDE ARCHIVOS �??�?�?�?�?�?�?�
frmCargando.Label1(2).Caption = "Cargando Server.ini"

MaxUsers = 0
Call LoadSini
Call CargaApuestas
'*************************************************
Call CargaNpcsDat
'*************************************************
frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
'Call LoadOBJData
Call LoadOBJData

frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
Call CargarHechizos

Call LoadArmasHerreria
Call LoadArmadurasHerreria
Call LoadObjCarpintero

If BootDelBackUp Then
    frmCargando.Label1(2).Caption = "Cargando BackUp"
    Call CargarBackUp_Nuevo2
Else
    frmCargando.Label1(2).Caption = "Cargando Mapas"
    Call LoadMapData
    'Call LoadMapData_Nuevo
End If
'Comentado porque hay worldsave en ese mapa!
'Call CrearClanPretoriano(MAPA_PRETORIANO, ALCOBA2_X, ALCOBA2_Y)
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�
Dim LoopC As Integer
'Resetea las conexiones de los usuarios
For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
    UserList(LoopC).ConnIDValida = False
Next LoopC
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�
With frmMain
    .AutoSave.Enabled = True
    .tLluvia.Enabled = True
   ' .tPiqueteC.Enabled = True Marche
    .Timer1.Enabled = True
    If ClientsCommandsQueue <> 0 Then
        .CmdExec.Enabled = True
    Else
        .CmdExec.Enabled = False
    End If
    .GameTimer.Enabled = True
    .Auditoria.Enabled = True
    .TIMER_AI.Enabled = True
    .npcataca.Enabled = True
End With
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'Configuracion de los sockets
#If UsarQueSocket = 1 Then
Call IniciaWsApi(frmMain.hwnd)
SockListen = ListenForConnect(Puerto, hWndMsg, "")
#ElseIf UsarQueSocket = 0 Then
frmCargando.Label1(2).Caption = "Configurando Sockets"
frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Binary = False
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048
Call ConfigListeningSocket(frmMain.Socket1, Puerto)
#ElseIf UsarQueSocket = 2 Then
frmMain.Serv.Iniciar Puerto
#ElseIf UsarQueSocket = 3 Then
frmMain.TCPServ.Encolar True
frmMain.TCPServ.IniciarTabla 1009
frmMain.TCPServ.SetQueueLim 51200
frmMain.TCPServ.Iniciar Puerto
#End If

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�


db_name = "tds_mysql"
db_server = "localhost"
db_port = "3306"    'default port is 3306"
db_user = "root"
db_pass = "E7JC"
constr = "Provider=MSDASQL.1;Password=;Persist Security Info=True;User ID=;Extended Properties=" & Chr$(34) & "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & db_name & ";SERVER=" & db_server & ";UID=" & db_user & ";PASSWORD=" & db_pass & ";PORT=" & db_port & ";OPTION=16387;STMT=;" & Chr$(34)
Set conn = New ADODB.Connection
conn.Open constr
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseServer

Unload frmCargando
'Log
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
Close #N
'Ocultar
If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If
tInicioServer = GetTickCount() And &H7FFFFFFF
Call InicializaEstadisticas
Randomize Timer
'ResetThread.CreateNewThread AddressOf ThreadResetActions, tpNormal
'Call MainThread
End Sub

Function FileExist(file As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
On Error Resume Next
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************
If Dir(file, FileType) = "" Then
    FileExist = False
Else
    FileExist = True
End If
End Function

Public Function ReadField(Pos As Integer, Text As String, SepASCII As Integer) As String
Dim delimiter As String
delimiter = Chr(SepASCII)
Dim i As Long
    Dim LastPos As Long
    Dim CurrentPos As Long
    
    For i = 1 To Pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = Mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        ReadField = Mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)
    End If
End Function

Function MapaValido(ByVal Map As Integer) As Boolean
MapaValido = Map >= 1 And Map <= NumMaps
End Function

Sub MostrarNumUsers()
frmMain.CantUsuarios.Caption = "Numero de usuarios jugando: " & NumUsers
End Sub

Public Sub LogCriticEvent(Desc As String)
On Error GoTo errhandler
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile
Exit Sub
errhandler:
End Sub

Public Sub LogEjercitoReal(Desc As String)
On Error GoTo errhandler
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
Print #nfile, Desc
Close #nfile
Exit Sub
errhandler:
End Sub

Public Sub LogEjercitoCaos(Desc As String)
On Error GoTo errhandler
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
Print #nfile, Desc
Close #nfile
Exit Sub
errhandler:
End Sub

Public Sub LogIndex(ByVal Index As Integer, ByVal Desc As String)
On Error GoTo errhandler
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\" & Index & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile
Exit Sub
errhandler:
End Sub

Public Sub LogError(Desc As String)

On Error GoTo errhandler
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\errores.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile
Exit Sub
errhandler:
End Sub

Public Sub LogStatic(Desc As String)
On Error GoTo errhandler
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile
Exit Sub
errhandler:
End Sub

Public Sub LogTarea(Desc As String)
'On Error GoTo errhandler
'Dim nfile As Integer
'nfile = FreeFile(1) ' obtenemos un canal
'Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
'Print #nfile, Date & " " & Time & " " & Desc
'Close #nfile
'Exit Sub
'errhandler:
End Sub

Public Sub LogDesarrollo(ByVal str As String)
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\desarrollo.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile
End Sub

Public Sub LogGM(Nombre As String, texto As String, Consejero As Boolean)
On Error GoTo errhandler
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
'[Wizard 03/09/05]-> Cambie la localizacion de los logs de GMs y consejeros.
If Consejero Then
    Open App.Path & "\logs\Gms\consejeros\" & Nombre & ".log" For Append Shared As #nfile
Else
    Open App.Path & "\logs\Gms\" & Nombre & ".log" For Append Shared As #nfile
End If
'[/Wizard]
Print #nfile, Date & " " & Time & " " & texto
Close #nfile
Exit Sub
errhandler:
End Sub

Public Sub SaveDayStats()
''On Error GoTo errhandler
''
''Dim nfile As Integer
''nfile = FreeFile ' obtenemos un canal
''Open App.Path & "\logs\" & Replace(Date, "/", "-") & ".log" For Append Shared As #nfile
''
''Print #nfile, "<stats>"
''Print #nfile, "<ao>"
''Print #nfile, "<dia>" & Date & "</dia>"
''Print #nfile, "<hora>" & Time & "</hora>"
''Print #nfile, "<segundos_total>" & DayStats.Segundos & "</segundos_total>"
''Print #nfile, "<max_user>" & DayStats.MaxUsuarios & "</max_user>"
''Print #nfile, "</ao>"
''Print #nfile, "</stats>"
''
''
''Close #nfile
Exit Sub
errhandler:
End Sub

Public Sub LogAsesinato(texto As String)
On Error GoTo errhandler
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile
Exit Sub
errhandler:
End Sub

Public Sub logVentaCasa(ByVal texto As String)
On Error GoTo errhandler
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile
Exit Sub
errhandler:
End Sub

Public Sub LogHackAttemp(texto As String)
On Error GoTo errhandler
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile
Exit Sub
errhandler:
End Sub

Public Sub LogCheating(texto As String)
On Error GoTo errhandler
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CH.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile
Exit Sub
errhandler:
End Sub

Public Sub LogCriticalHackAttemp(texto As String)
On Error GoTo errhandler
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile
Exit Sub
errhandler:
End Sub

Public Sub LogAntiCheat(texto As String)
On Error GoTo errhandler
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, ""
Close #nfile
Exit Sub
errhandler:
End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
Dim Arg As String
Dim i As Integer

For i = 1 To 33
Arg = ReadField(i, cad, 44)
If Arg = "" Then Exit Function
Next i
ValidInputNP = True
End Function

Sub Restart()
'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next
If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."
Dim LoopC As Integer
#If UsarQueSocket = 0 Then
    frmMain.Socket1.Cleanup
    frmMain.Socket1.Startup
    frmMain.Socket2(0).Cleanup
    frmMain.Socket2(0).Startup
#ElseIf UsarQueSocket = 1 Then
    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, "")
#ElseIf UsarQueSocket = 2 Then
#End If

For LoopC = 1 To MaxUsers
    Call CloseSocket(LoopC)
Next

ReDim UserList(1 To MaxUsers)

For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
    UserList(LoopC).ConnIDValida = False
Next LoopC

LastUser = 0
NumUsers = 0

ReDim Npclist(1 To MAXNPCS) As npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer

Call LoadSini
Call LoadOBJData
Call LoadMapData
Call CargarHechizos

#If UsarQueSocket = 0 Then
'*****************Setup socket
frmMain.Socket1.AddressFamily = AF_INET
frmMain.Socket1.protocol = IPPROTO_IP
frmMain.Socket1.SocketType = SOCK_STREAM
frmMain.Socket1.Binary = False
frmMain.Socket1.Blocking = False
frmMain.Socket1.BufferSize = 1024
frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048
'Escucha
frmMain.Socket1.LocalPort = val(Puerto)
frmMain.Socket1.listen

#ElseIf UsarQueSocket = 1 Then
#ElseIf UsarQueSocket = 2 Then
#End If

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
'Log it
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " servidor reiniciado."
Close #N
'Ocultar
If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If
End Sub

Public Function Intemperie(ByVal UserIndex As Integer) As Boolean
    If MapInfo(UserList(UserIndex).Pos.Map).Zona <> "DUNGEON" Then
        If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 1 And _
           MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 2 And _
           MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 4 Then Intemperie = True
    Else
        Intemperie = False
    End If
End Function

Public Sub EfectoLluvia(ByVal UserIndex As Integer)
On Error GoTo errhandler
Dim massta As Integer
If UserList(UserIndex).flags.UserLogged Then
    If Intemperie(UserIndex) Then
        '[Misery_Ezequiel 09/07/05]
        massta = CInt(RandomNumber(1, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5)))
        UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta + massta
        If UserList(UserIndex).Stats.MinSta > UserList(UserIndex).Stats.MaxSta Then UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta
        Call SendUserStatsBox(UserIndex)
        '[\]Misery_Ezequiel 09/07/05]
    End If
End If
Exit Sub
errhandler:
 LogError ("Error en EfectoLluvia")
End Sub


Public Sub EfectoNieve(ByVal UserIndex As Integer)
On Error GoTo errhandler
Dim modifi As Integer
If MapInfo(UserList(UserIndex).Pos.Map).Terreno = Nieve Then
    Call Senddata(ToIndex, UserIndex, 0, "Y10")
    modifi = Porcentaje(UserList(UserIndex).Stats.MaxHP, 40)
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - modifi
    If UserList(UserIndex).Stats.MinHP < 1 Then
            Call Senddata(ToIndex, UserIndex, 0, "Y22")
            UserList(UserIndex).Stats.MinHP = 0
            Call UserDie(UserIndex)
    End If
End If

Exit Sub
errhandler:
 LogError ("Error en EfectoNieve")
End Sub


Public Sub TiempoInvocacion(ByVal UserIndex As Integer)
Dim i As Integer
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
           Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = _
           Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia - 1
           If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
        End If
    End If
Next i
End Sub
Public Sub EfectoMimetismo(ByVal UserIndex As Integer)

If UserList(UserIndex).Counters.Mimetismo < IntervaloInvisible Then
    UserList(UserIndex).Counters.Mimetismo = UserList(UserIndex).Counters.Mimetismo + 1
Else
    'restore old char
    Call Senddata(ToIndex, UserIndex, 0, "||Recuperas tu apariencia normal." & FONTTYPE_INFO)
    
    UserList(UserIndex).Char.Body = UserList(UserIndex).CharMimetizado.Body
    UserList(UserIndex).Char.Head = UserList(UserIndex).CharMimetizado.Head
    UserList(UserIndex).Char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
    UserList(UserIndex).Char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
    UserList(UserIndex).Char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
        
    
    UserList(UserIndex).Counters.Mimetismo = 0
    UserList(UserIndex).flags.Mimetizado = 0
    Call ChangeUserChar(ToMap, UserIndex, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
End If
            
End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer)
'Reescrito para que vuelva a quitar vida si esta desnudo en un mapa frio[Wizard]
Dim modifi As Integer

If UserList(UserIndex).flags.Privilegios > 0 Then Exit Sub

If UserList(UserIndex).Counters.Frio < IntervaloFrio * 10 Then
  UserList(UserIndex).Counters.Frio = UserList(UserIndex).Counters.Frio + 1
Else
   If MapInfo(UserList(UserIndex).Pos.Map).Frio = 1 Then
        Call Senddata(ToIndex, UserIndex, 0, "||��Estas muriendo de frio, abrigate o moriras!!." & FONTTYPE_INFO)
        modifi = Porcentaje(UserList(UserIndex).Stats.MaxHP, 5)
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - modifi
        Call SendUserVida(UserIndex)
        If UserList(UserIndex).Stats.MinHP < 1 Then
            Call Senddata(ToIndex, UserIndex, 0, "||��Has muerto de frio!!." & FONTTYPE_INFO)
            UserList(UserIndex).Stats.MinHP = 0
            Call UserDie(UserIndex)
        End If
   Else
        modifi = Porcentaje(UserList(UserIndex).Stats.MaxSta, 5)
        Call QuitarSta(UserIndex, modifi)
        'Call SendData(ToIndex, UserIndex, 0, "||��Has perdido stamina, si no te abrigas rapido perderas toda!!." & FONTTYPE_INFO)
        Call SendUserEsta(UserIndex)
        UserList(UserIndex).Counters.Frio = 0
   End If
End If
  
  


End Sub

'[Misery_Ezequiel 11/07/05]
Public Sub EfectoFrioPowa(ByVal UserIndex As Integer)
Dim modifi As Integer
'[Misery_Ezequiel 27/05/05]
If UserList(UserIndex).flags.Privilegios >= 1 Then Exit Sub
'[\]Misery_Ezequiel 27/05/05]
If UserList(UserIndex).Counters.Frio < IntervaloFrioDeNieve Then
  UserList(UserIndex).Counters.Frio = UserList(UserIndex).Counters.Frio + 1
Else
  If MapInfo(UserList(UserIndex).Pos.Map).Terreno = Nieve Then
    Call Senddata(ToIndex, UserIndex, 0, "Y10")
    modifi = Porcentaje(UserList(UserIndex).Stats.MaxHP, 5)
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - modifi
    If UserList(UserIndex).Stats.MinHP < 1 Then
            Call Senddata(ToIndex, UserIndex, 0, "Y22")
            UserList(UserIndex).Stats.MinHP = 0
            Call UserDie(UserIndex)
    End If
  Else
    modifi = Porcentaje(UserList(UserIndex).Stats.MaxSta, 5)
    Call QuitarSta(UserIndex, modifi)
    'Call SendData(ToIndex, UserIndex, 0, "||��Has perdido stamina, si no te abrigas rapido perderas toda!!." & FONTTYPE_INFO)
  End If
  
  UserList(UserIndex).Counters.Frio = 0
  Call SendUserStatsBox(UserIndex)
End If
End Sub
'[\]Misery_Ezequiel 11/07/05]

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)
If UserList(UserIndex).Counters.Invisibilidad < IntervaloInvisible Then
  'cazador con armadura de cazador oculto no se hace visible
  'mersada de inmediata direccion pero no me importa pq esta
  'version ya fue :D
  '[Misery_Ezequiel 29/06/05]
  If UCase$(UserList(UserIndex).Clase) = "CAZADOR" And UserList(UserIndex).flags.Oculto > 0 And UserList(UserIndex).Stats.UserSkills(Ocultarse) > 90 Then
    If UserList(UserIndex).Invent.ArmourEqpObjIndex = 612 Or UserList(UserIndex).Invent.ArmourEqpObjIndex = 360 Then
        Exit Sub
    End If
  End If
  '[\]Misery_Ezequiel 29/06/05]
  UserList(UserIndex).Counters.Invisibilidad = UserList(UserIndex).Counters.Invisibilidad + 1
Else
  Call Senddata(ToIndex, UserIndex, 0, "Y23")
  UserList(UserIndex).Counters.Invisibilidad = 0
  UserList(UserIndex).flags.Invisible = 0
  UserList(UserIndex).flags.Oculto = 0
'   no ecripto los noverx,0
'  If EncriptarProtocolosCriticos Then
'    Call SendCryptedData(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.charindex & ",0")
'  Else
    Call Senddata(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.charindex & ",0")
'  End If
End If
End Sub

Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)
If Npclist(NpcIndex).Contadores.Paralisis > 0 Then
    Npclist(NpcIndex).Contadores.Paralisis = Npclist(NpcIndex).Contadores.Paralisis - 1
Else
    Npclist(NpcIndex).flags.Paralizado = 0
        Npclist(NpcIndex).flags.Inmovilizado = 0
End If
End Sub

Public Sub EfectoCegueEstu(ByVal UserIndex As Integer)
If UserList(UserIndex).Counters.Ceguera > 0 Then
    UserList(UserIndex).Counters.Ceguera = UserList(UserIndex).Counters.Ceguera - 1
Else
    If UserList(UserIndex).flags.Ceguera = 1 Then
        UserList(UserIndex).flags.Ceguera = 0
        Call Senddata(ToIndex, UserIndex, 0, "NSEGUE")
    End If
    If UserList(UserIndex).flags.Estupidez = 1 Then
        UserList(UserIndex).flags.Estupidez = 0
        Call Senddata(ToIndex, UserIndex, 0, "NESTUP")
    End If

End If
End Sub

Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)
If UserList(UserIndex).Counters.Paralisis > 0 Then
    UserList(UserIndex).Counters.Paralisis = UserList(UserIndex).Counters.Paralisis - 1
Else
    UserList(UserIndex).flags.Paralizado = 0
    'UserList(UserIndex).Flags.AdministrativeParalisis = 0
    Call Senddata(ToIndex, UserIndex, 0, "PARADOK")
End If
End Sub

Public Sub RecStamina(UserIndex As Integer, EnviarStats As Boolean, Intervalo As Integer)
'[Wizard 02/09/05]
If UserList(UserIndex).flags.Desnudo = 1 Then Exit Sub
'[/Wizard]
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 2 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 4 Then Exit Sub

If UserList(UserIndex).Stats.MinHam = 0 Or UserList(UserIndex).Stats.MinAGU = 0 Then Exit Sub

Dim massta As Integer
If UserList(UserIndex).Stats.MinSta < UserList(UserIndex).Stats.MaxSta Then
   If UserList(UserIndex).Counters.STACounter < Intervalo Then
       UserList(UserIndex).Counters.STACounter = UserList(UserIndex).Counters.STACounter + 1
   Else
       UserList(UserIndex).Counters.STACounter = 0
       massta = CInt(RandomNumber(1, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5)))
       UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta + massta
       If UserList(UserIndex).Stats.MinSta > UserList(UserIndex).Stats.MaxSta Then UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta
           EnviarStats = True
       End If
End If
End Sub

Public Sub EfectoVeneno(UserIndex As Integer, EnviarStats As Boolean)
Dim N As Integer

If UserList(UserIndex).Counters.Veneno < IntervaloVeneno Then
  UserList(UserIndex).Counters.Veneno = UserList(UserIndex).Counters.Veneno + 1
Else
  Call Senddata(ToIndex, UserIndex, 0, "Y39")
  UserList(UserIndex).Counters.Veneno = 0
  N = RandomNumber(1, 5)
  UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - N
  If UserList(UserIndex).Stats.MinHP < 1 Then Call UserDie(UserIndex)
  EnviarStats = True
End If
End Sub

Public Sub DuracionPociones(UserIndex As Integer)
'Controla la duracion de las pociones
If UserList(UserIndex).flags.DuracionEfecto > 0 Then
   UserList(UserIndex).flags.DuracionEfecto = UserList(UserIndex).flags.DuracionEfecto - 1
   If UserList(UserIndex).flags.DuracionEfecto = 0 Then
        UserList(UserIndex).flags.TomoPocion = False
        UserList(UserIndex).flags.TipoPocion = 0
        'volvemos los atributos al estado normal
        Dim loopX As Integer
        For loopX = 1 To NUMATRIBUTOS
              UserList(UserIndex).Stats.UserAtributos(loopX) = UserList(UserIndex).Stats.UserAtributosBackUP(loopX)
        Next
   End If
End If
End Sub

Public Sub HambreYSed(UserIndex As Integer, fenviarAyS As Boolean)
'[Misery_Ezequiel 27/05/05]
If UserList(UserIndex).flags.Privilegios >= 2 Then Exit Sub
'[\]Misery_Ezequiel 27/05/05]
'Sed
If UserList(UserIndex).Stats.MinAGU > 0 Then
    If UserList(UserIndex).Counters.AGUACounter < IntervaloSed Then
          UserList(UserIndex).Counters.AGUACounter = UserList(UserIndex).Counters.AGUACounter + 1
    Else
          UserList(UserIndex).Counters.AGUACounter = 0
          UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU - 10
                            
          If UserList(UserIndex).Stats.MinAGU <= 0 Then
               UserList(UserIndex).Stats.MinAGU = 0
               UserList(UserIndex).flags.Sed = 1
          End If
          fenviarAyS = True
                            
    End If
End If
'hambre
If UserList(UserIndex).Stats.MinHam > 0 Then
   If UserList(UserIndex).Counters.COMCounter < IntervaloHambre Then
        UserList(UserIndex).Counters.COMCounter = UserList(UserIndex).Counters.COMCounter + 1
   Else
        UserList(UserIndex).Counters.COMCounter = 0
        UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam - 10
        If UserList(UserIndex).Stats.MinHam < 0 Then
               UserList(UserIndex).Stats.MinHam = 0
               UserList(UserIndex).flags.Hambre = 1
        End If
        fenviarAyS = True
    End If
End If
End Sub

Public Sub Sanar(UserIndex As Integer, EnviarStats As Boolean, Intervalo As Integer)
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 2 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 4 Then Exit Sub

Dim mashit As Integer
'con el paso del tiempo va sanando....pero muy lentamente ;-)
If UserList(UserIndex).Stats.MinHP < UserList(UserIndex).Stats.MaxHP Then
   If UserList(UserIndex).Counters.HPCounter < Intervalo Then
      UserList(UserIndex).Counters.HPCounter = UserList(UserIndex).Counters.HPCounter + 1
   Else
      mashit = CInt(RandomNumber(2, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5)))
                           
      UserList(UserIndex).Counters.HPCounter = 0
      UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + mashit
      If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
         Call Senddata(ToIndex, UserIndex, 0, "Y40")
         EnviarStats = True
      End If
End If
End Sub

Public Sub CargaNpcsDat()
'Dim NpcFile As String
'
'NpcFile = DatPath & "NPCs.dat"
'ANpc = INICarga(NpcFile)
'Call INIConf(ANpc, 0, "", 0)
'
'NpcFile = DatPath & "NPCs-HOSTILES.dat"
'Anpc_host = INICarga(NpcFile)
'Call INIConf(Anpc_host, 0, "", 0)
Dim npcfile As String

npcfile = DatPath & "NPCs.dat"
LeerNPCs.Abrir npcfile
npcfile = DatPath & "NPCs-HOSTILES.dat"
LeerNPCsHostiles.Abrir npcfile
End Sub

Public Sub DescargaNpcsDat()
'If ANpc <> 0 Then Call INIDescarga(ANpc)
'If Anpc_host <> 0 Then Call INIDescarga(Anpc_host)
End Sub

Sub PasarSegundo()
    Dim i As Integer
    For i = 1 To LastUser
        'Cerrar usuario
        If UserList(i).Counters.Saliendo Then
            UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
            If UserList(i).Counters.Salir <= 0 Then
                Call Senddata(ToIndex, i, 0, "Y41")
                Call Senddata(ToIndex, i, 0, "FINOK")
                Call CloseSocket(i)
                Exit Sub
            End If
        End If
      Next i
    
    
    If Minutossinlluvia = 864 Or Minutossinlluvia = 597 Or Minutossinlluvia = 100 Then
        Call Senddata(ToAll, 0, 0, "CHE")
        Call LimpiarMundo
    End If
    
' manejador de lluvia lo saque del Intervalo que hacia
If Lloviendo = False Then
     If Minutossinlluvia > 1200 Then
        If RandomNumber(1, 500) <= 7 Then
        Call Senddata(ToAll, 0, 0, "LLU")
        Call Senddata(ToAll, 0, 0, "NIE")
        Lloviendo = True
        Nevando = True
        Minutossinlluvia = 0
        End If
     Else
     Minutossinlluvia = Minutossinlluvia + 1
     End If
Else
    If Minutoslloviendo > 180 Then 'Minutos en realidad son segundos
    Call Senddata(ToAll, 0, 0, "LLU")
    Call Senddata(ToAll, 0, 0, "NIE")
    Minutoslloviendo = 0
    Lloviendo = False
    Nevando = False
    Else
    Minutoslloviendo = Minutoslloviendo + 1
    End If
End If
End Sub
 
Sub GuardarUsuarios()
    haciendoBK = True
    Call Senddata(ToAll, 0, 0, "BKW")
    Call Senddata(ToAll, 0, 0, "Y43")
    Dim i As Integer
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            Call SaveUser(i, UserList(i).Name)
        End If
    Next i
    Call Senddata(ToAll, 0, 0, "Y44")
    Call Senddata(ToAll, 0, 0, "BKW")
    haciendoBK = False
End Sub

Sub InicializaEstadisticas()
Dim Ta As Long
Ta = GetTickCount() And &H7FFFFFFF
Call EstadisticasWeb.Inicializa(frmMain.hwnd)
Call EstadisticasWeb.Informar(CANTIDAD_MAPAS, NumMaps)
Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
Call EstadisticasWeb.Informar(UPTIME_SERVER, (Ta - tInicioServer) / 1000)
Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)
End Sub

Public Sub LogDebug(Desc As String)
On Error GoTo errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\debug.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile
Exit Sub
errhandler:
End Sub
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
    'WorldSave
    Call DoBackUp
    ' experiencias
    Call mdParty.ActualizaExperiencias
    'Guardar Pjs
    Call GuardarUsuarios
    'Guilds
    Call SaveGuildsDB
    If EjecutarLauncher Then Shell (App.Path & "\ao.exe")
    'Chauuu
    Unload frmMain
End Sub

Public Sub LogRetos(ByVal str1 As String)
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\retos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str1
Close #nfile
End Sub

Public Sub LogCSu(ByVal str2 As String)
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\logcsu.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str2
Close #nfile
End Sub
'********************Misery_Ezequiel 28/05/05********************'

Function CharExist(file As String) As Boolean
    sql = "select * from usuarios where NickB = '" & file & "'"
   rs.Open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF = False Then
        CharExist = True
    Else
        CharExist = False
       rs.Close
    End If
    
End Function

Public Sub EfectoNevando(ByVal UserIndex As Integer)
On Error GoTo errhandler
Dim modifi As Integer
If MapInfo(UserList(UserIndex).Pos.Map).Terreno = Nieve Then
    Call Senddata(ToIndex, UserIndex, 0, "Y10")
    modifi = Porcentaje(UserList(UserIndex).Stats.MaxHP, 60)
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - modifi
    If UserList(UserIndex).Stats.MinHP < 1 Then
            Call Senddata(ToIndex, UserIndex, 0, "Y22")
            UserList(UserIndex).Stats.MinHP = 0
            Call UserDie(UserIndex)
    End If
End If

Exit Sub
errhandler:
 LogError ("Error en EfectoNieve")
End Sub

Public Sub LogCheats(texto As String)
On Error GoTo errhandler
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\cheats.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile
Exit Sub
errhandler:
End Sub
