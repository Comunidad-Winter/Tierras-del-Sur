Attribute VB_Name = "modConfig"
Option Explicit

Public BootDelBackUp As Boolean             ' Los mapas se cargan desde el backup?
Public Puerto As Integer
Public EnTesting As Boolean
Public HideMe As Byte
Public ServerSoloGMs As Byte

Public SanaIntervaloSinDescansar As Long
Public StaminaIntervalo As Long
Public SanaIntervaloDescansar As Long

Public IntervaloSed As Long
Public IntervaloHambre As Long
Public IntervaloParalizado As Long
Public IntervaloParalizadoGuerrero As Long
Public IntervaloParalizadoCazador As Long
Public IntervaloParalizadoNPC As Long
Public IntervaloInvisible As Long
Public IntervaloMimetizado As Long
Public IntervaloFrio As Long
Public IntervaloCalor As Long

'Variables cambidas
Public IntervaloInvocacion As Long
Public IntervaloInvocacionTierra As Long
Public IntervaloInvocacionAgua As Long
Public IntervaloInvocacionFuego As Long
Public IntervaloDuracionPociones As Long

'Variables cambiadas
Public IntervaloUserPuedeAtacar As Long
Public IntervaloCerrarConexion As Long
Public IntervaloUserPuedeUsar As Long

'General
Public IntervaloGolpe As Single
Public IntervaloMagia As Single
Public IntervaloFlecha As Single
Public IntervaloTotal As String
Public IntervaloU As Single
Public IntervaloClick As Single

'Guerrero
Public IntervaloGolpeG As Single
Public IntervaloMagiaG As Single
Public IntervaloFlechaG As Single
Public IntervaloUG As Single
Public IntervaloClickG As Single
Public IntervaloTotalG As String

'Cazador
Public IntervaloTotalC As String
Public IntervaloGolpeC As Single
Public IntervaloMagiaC As Single
Public IntervaloFlechaC As Single
Public IntervalouC As Single
Public IntervaloClickC As Single

Public IntervalosAntiLanzarAutomatico As String

Public INTERVALO_ATAQUE As Long

Public MinutosWs As Long

Public Nix As WorldPos
Public Ullathorpe As WorldPos
Public Banderbill As WorldPos
Public CiudadOscura As WorldPos


Public Type carcel
    posiciones(1 To 3) As WorldPos
    salida As WorldPos
End Type
    
Public NixCarcel As carcel


Sub LoadSini()

Dim Temporal As Long
Dim Temporal1 As Long

BootDelBackUp = val(GetVar(IniPath & "Server.ini", "INIT", "IniciarDesdeBackUp")) = 1

ConexionWeb.SERVIDOR_WEB_IP = GetVar(IniPath & "Server.ini", "SERVERWEB", "IP")
ConexionWeb.SERVIDOR_WEB_PUERTO = CInt(GetVar(IniPath & "Server.ini", "SERVERWEB", "PUERTO"))

'Misc
LastSockListen = val(GetVar(IniPath & "Server.ini", "INIT", "LastSockListen"))
Puerto = val(GetVar(IniPath & "Server.ini", "INIT", "StartPort"))
HideMe = val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))

ServerSoloGMs = val(GetVar(IniPath & "server.ini", "init", "ServerSoloGMs"))

EnTesting = val(GetVar(IniPath & "server.ini", "INIT", "Testing"))

'Intervalos
SanaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar"))

StaminaIntervalo = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloStamina"))

SanaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar"))

IntervaloSed = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed"))

IntervaloHambre = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre"))

IntervaloParalizado = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado"))
IntervaloParalizadoGuerrero = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizadoGuerrero"))
IntervaloParalizadoCazador = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizadoCazador"))

IntervaloParalizadoNPC = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizadoNPC"))

IntervaloInvisible = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible"))

IntervaloMimetizado = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloMimetizado"))

IntervaloFrio = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio"))
IntervaloCalor = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCalor"))


'''''''''''''''
IntervaloInvocacion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion"))
IntervaloInvocacionTierra = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacionTierra"))
IntervaloInvocacionAgua = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacionAgua"))
IntervaloInvocacionFuego = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacionFuego"))
IntervaloDuracionPociones = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloDuracionPociones"))

'***************GOlPE***********************
IntervaloGolpe = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "Golpe"))
IntervaloGolpeG = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "GolpeG"))
IntervaloGolpeC = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "GolpeC"))
'*******************************************

IntervaloTotal = IntervaloGolpe
IntervaloTotalG = IntervaloGolpeG
IntervaloTotalC = IntervaloGolpeC

'**************Magia**************************
IntervaloMagia = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "Magia"))
IntervaloMagiaG = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "MagiaG"))
IntervaloMagiaC = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "MagiaC"))
'**********************************************

IntervaloTotal = IntervaloTotal & "-" & IntervaloMagia
IntervaloTotalG = IntervaloTotalG & "-" & IntervaloMagiaG
IntervaloTotalC = IntervaloTotalC & "-" & IntervaloMagiaC


'****************Flecha***************************
IntervaloFlecha = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "Flecha"))
IntervaloFlechaG = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "FlechaG"))
IntervaloFlechaC = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "FlechaC"))
'*************************************************

IntervaloTotal = IntervaloTotal & "-" & IntervaloFlecha
IntervaloTotalG = IntervaloTotalG & "-" & IntervaloFlechaG
IntervaloTotalC = IntervaloTotalC & "-" & IntervaloFlechaC

'***************U************************************
IntervaloU = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "U"))
IntervaloUG = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "UG"))
IntervalouC = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "UC"))
'******************************************************

IntervaloTotal = IntervaloTotal & "-" & IntervaloU
IntervaloTotalG = IntervaloTotalG & "-" & IntervaloUG
IntervaloTotalC = IntervaloTotalC & "-" & IntervalouC

'*********************CLICK***************************
IntervaloClick = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "Click"))
IntervaloClickG = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "ClickG"))
IntervaloClickC = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "ClickC"))
'********************************************************

IntervaloTotal = IntervaloTotal & "-" & IntervaloClick
IntervaloTotalG = IntervaloTotalG & "-" & IntervaloClickG
IntervaloTotalC = IntervaloTotalC & "-" & IntervaloClickC

'*****************************************************************************
'************************** ANTICHEAT ****************************************

IntervalosAntiLanzarAutomatico = GetVar(IniPath & "Server.ini", "ANTICHEAT", "IntervaloSolapaLanzar") & "-"
IntervalosAntiLanzarAutomatico = IntervalosAntiLanzarAutomatico & GetVar(IniPath & "Server.ini", "ANTICHEAT", "IntervaloSolapaLanzarSuper") & "-"
IntervalosAntiLanzarAutomatico = IntervalosAntiLanzarAutomatico & GetVar(IniPath & "Server.ini", "ANTICHEAT", "IntervaloHechizoLanzar") & "-"
IntervalosAntiLanzarAutomatico = IntervalosAntiLanzarAutomatico & GetVar(IniPath & "Server.ini", "ANTICHEAT", "IntervaloHechizoLanzarSuper") & "-"
IntervalosAntiLanzarAutomatico = IntervalosAntiLanzarAutomatico & GetVar(IniPath & "Server.ini", "ANTICHEAT", "UmbralAlerta")

'&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&
frmMain.TIMER_AI.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcAI"))

INTERVALO_ATAQUE = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcPuedeAtacar"))

IntervaloUserPuedeAtacar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar"))

MinutosWs = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWS"))

UltimaFechaProcesada = GetVar(IniPath & "Server.ini", "INTERVALOS", "UltimaFechaProcesada")

If MinutosWs < 60 Then MinutosWs = 180

IntervaloCerrarConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCerrarConexion"))
IntervaloUserPuedeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeUsar"))
  
RecordUsuarios = val(GetVar(IniPath & "Server.ini", "INIT", "Record"))
  
'Max users
Temporal = val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))
If MaxUsers = 0 Then
    MaxUsers = Temporal
    ReDim UserList(1 To MaxUsers) As User
End If

Nix.map = GetVar(DatPath & "Ciudades.dat", "NIX", "Mapa")
Nix.x = GetVar(DatPath & "Ciudades.dat", "NIX", "X")
Nix.y = GetVar(DatPath & "Ciudades.dat", "NIX", "Y")

Ullathorpe.map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
Ullathorpe.x = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
Ullathorpe.y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")

Banderbill.map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
Banderbill.x = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
Banderbill.y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")

CiudadOscura.map = GetVar(DatPath & "Ciudades.dat", "CiudadOscura", "Mapa")
CiudadOscura.x = GetVar(DatPath & "Ciudades.dat", "CiudadOscura", "X")
CiudadOscura.y = GetVar(DatPath & "Ciudades.dat", "CiudadOscura", "Y")

NixCarcel = getCarcel("CarcelNix")

DayStats.MinUsuarios = MaxUsers

End Sub

Private Function getCarcel(nombre As String) As carcel
Dim carcel As carcel
Dim mapa As Integer
Dim loopCarcel As Integer

mapa = val(GetVar(DatPath & "Ciudades.dat", nombre, "Mapa"))

For loopCarcel = 1 To 3
    carcel.posiciones(loopCarcel).map = mapa
    carcel.posiciones(loopCarcel).x = GetVar(DatPath & "Ciudades.dat", nombre, "X" & loopCarcel)
    carcel.posiciones(loopCarcel).y = GetVar(DatPath & "Ciudades.dat", nombre, "Y" & loopCarcel)
Next

carcel.salida.map = mapa
carcel.salida.x = val(GetVar(DatPath & "Ciudades.dat", nombre, "SalidaX"))
carcel.salida.y = val(GetVar(DatPath & "Ciudades.dat", nombre, "SalidaY"))

getCarcel = carcel

End Function
