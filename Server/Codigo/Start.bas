Attribute VB_Name = "Start"
Option Explicit

Public Sub Main()

If App.PrevInstance Then
    MsgBox "Este programa ya está corriendo.", vbInformation, "Tirras Del Sur"
    End
End If
            
ChDir App.Path
ChDrive App.Path

IniPath = App.Path & "\"
DatPath = App.Path & "\Dat\"
MapPath = App.Path & "\Maps\"

'Inicio el Manager
If Not API_Manager.iniciarManager Then
    LogError "No se pudo conectar con el Manager"
End If

' Configuraciones rápidas
servidorAtacado = False                     ' AntiDDos
denunciarActivado = True                    ' Se puede denunciar
charlageneral = True                        ' Chat global
ProfilePaquetes = False                     ' Profile paquetes que se reciben.
fraccionDelDia = 62
Randomize Timer

#If SERVER_PRUEBAS = 1 Then
    frmMain.Caption = "MAPAS_TEST"
#Else
    frmMain.Caption = "INNOVA"
#End If

frmCargando.Show

Call LoadSini                               ' Cargamos la configuracion

Call General.iniciarEstructuras             ' Npcs, gms, etc

Call modPersonajes.iniciarEstructuras

Call modMySql.iniciarConexionBaseDeDatos    ' Me conecto a la base de datos

Call Admin.actualizarOnlinesDB(True)        ' Actualizo los online que hay en el juego

Call CryptoInit                             ' Inicia codigos de encriptacion

Call Constantes_Generales.inicializarConstantes   ' Constantes relacionadas al juego

Call modClases.inicializarClases

Call inicializarRazas

Call LoadMotd                               ' El mensaje de bienvenida

Call CargarRequisitos                       ' Requisitos para entrar a la armada

Call InitTimeGetTime                        ' Intervalos

Call NPCs.iniciarEstructurasNpcs

Call modEventos.iniciarEstructuraEventos    ' Lista de Eventos

Call modDescansos.iniciarZonasDescanso      ' Zonas usadas por los eventos

Call modRings.iniciarRings                  ' Carga de rings

Call modRetos.iniciar                       ' Iniciar Sistema de Retos

Call modCapturarPantalla.iniciar            ' Sistema de captura de pantalla del usuario

Call Anticheat_MemCheck.iniciarEstructuras  ' Sistema anticheat para chequear la edición de la memoria del usuario

Call CargarHechizos                         ' Hechizos

Call CargaNpcsDat                           ' Criaturas

Call LoadOBJData                            ' Objecitos

Call LoadCofres                             ' Cofres

Call LoadPergaminos                         ' Pergaminos

Call LoadPasajes                            ' Viajes entre ciudades

Call LoadArmasHerreria

Call LoadArmadurasHerreria

Call LoadObjCarpintero

Call CargaApuestas                          ' Sistema de apuestas

Call CargarSpawnList                        ' Sistema de entrenameinto

Call mdClanes.iniciar                       ' Sistema de clanes

Call modResucitar.iniciar

Call SV_Mundo.cargarMundo  'Cargo el Mundo del Juego

DoEvents

'Resetea las conexiones de los usuarios
Dim loopC As Integer

For loopC = 1 To MaxUsers
    UserList(loopC).ConnID = INVALID_SOCKET
    UserList(loopC).InicioConexion = 0
    UserList(loopC).ConfirmacionConexion = 0
    
    UserList(loopC).PacketNumber = 1
    UserList(loopC).MinPacketNumber = 1
    UserList(loopC).CryptOffset = 0
    
    UserList(loopC).UserIndex = loopC
Next loopC

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
With frmMain
    .AutoSave.Enabled = True
    .GameTimer.Enabled = True
    .Auditoria.Enabled = True
    .TIMER_AI.Enabled = True
    .timerTrabajo.Enabled = True
End With
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Call iniciarSockets

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

Unload frmCargando

'Log
Call Logs.LogMain("Server iniciado " & App.Major & "." & App.Minor & "." & App.Revision)

'Ocultar
If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

End Sub

Private Sub iniciarSockets()
    
    If App.LogMode = 0 Then
        If LastSockListen >= 0 Then Call apiclosesocket(LastSockListen) 'Cierra el socket de escucha
    End If
    
    'Configuracion de los sockets
    Call IniciaWsApi(frmMain.hWnd)
    LogDesarrollo ("Abriendo puerto " & Puerto)
    SockListen = ListenForConnect(Puerto, hWndMsg, "")

    If SockListen <> -1 Then
        Call WriteVar(IniPath & "Server.ini", "INIT", "LastSockListen", SockListen) ' Guarda el socket escuchando
    Else
        MsgBox "Ha ocurrido un error al iniciar el socket del Servidor.", vbCritical + vbOKOnly
        End
    End If
End Sub
