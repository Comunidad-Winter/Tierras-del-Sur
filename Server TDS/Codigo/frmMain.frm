VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Argentum Online"
   ClientHeight    =   4920
   ClientLeft      =   1950
   ClientTop       =   1815
   ClientWidth     =   6300
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4920
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Timer Trabajo 
      Interval        =   850
      Left            =   960
      Top             =   3120
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   2400
      Top             =   1680
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   2160
   End
   Begin VB.CheckBox SUPERLOG 
      Caption         =   "log"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton CMDDUMP 
      Caption         =   "dump"
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   120
      Top             =   2160
   End
   Begin VB.Timer CmdExec 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   1680
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   120
      Top             =   2640
   End
   Begin VB.Timer tLluvia 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1560
      Top             =   1680
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   600
      Top             =   1680
   End
   Begin VB.Timer npcataca 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1080
      Top             =   2160
   End
   Begin VB.Timer TIMER_AI 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1680
   End
   Begin VB.Frame Frame1 
      Caption         =   "BroadCast"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4935
      Begin VB.CommandButton Command2 
         Caption         =   "Broadcast consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Broadcast clientes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox BroadMsg 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tierras del Sur v 9.9F"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   480
      TabIndex        =   10
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label Escuch 
      Caption         =   "Label2"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label CantUsuarios 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de usuarios:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1725
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   45
   End
   Begin VB.Menu mnuControles 
      Caption         =   "Argentum"
      Begin VB.Menu mnuServidor 
         Caption         =   "Configuracion"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "Systray Servidor"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar Servidor"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modificado por marche ulitma vez el 28/4/05
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'********************Misery_Ezequiel 28/05/05********************'
Option Explicit
Public ESCUCHADAS As Long
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Const NIM_ADD = 0
Const NIM_MODIFY = 1
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4
Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hwnd As Long, ID As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA
    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hwnd = hwnd
    nidTemp.uID = ID
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)
    setNOTIFYICONDATA = nidTemp
End Function

Sub CheckIdleUser()
Dim iUserIndex As Integer
For iUserIndex = 1 To MaxUsers
   'Conexion activa? y es un usuario loggeado?
   If UserList(iUserIndex).ConnID <> -1 And UserList(iUserIndex).flags.UserLogged Then
        'Actualiza el contador de inactividad
        UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
        If UserList(iUserIndex).Counters.IdleCount >= 5 Then
           ' No desconectar usuarios privilegiados por inactividad (dioses, semi, consejeros...) - 2005-03-25 byGorlok
            If UserList(iUserIndex).flags.Privilegios = 0 And UserList(iUserIndex).flags.Inmovilizado = 0 And UserList(iUserIndex).flags.Paralizado = 0 Then
                Call Senddata(ToIndex, iUserIndex, 0, "!!Demasiado tiempo inactivo. Has sido desconectado..")
                'mato los comercios seguros
                If UserList(iUserIndex).ComUsu.DestUsu > 0 Then
                    If UserList(UserList(iUserIndex).ComUsu.DestUsu).flags.UserLogged Then
                        If UserList(UserList(iUserIndex).ComUsu.DestUsu).ComUsu.DestUsu = iUserIndex Then
                            Call Senddata(ToIndex, UserList(iUserIndex).ComUsu.DestUsu, 0, "Y129")
                            Call FinComerciarUsu(UserList(iUserIndex).ComUsu.DestUsu)
                        End If
                    End If
                    Call FinComerciarUsu(iUserIndex)
                End If
                Call Cerrar_Usuario(iUserIndex)
            Else
                UserList(iUserIndex).Counters.IdleCount = 0
            End If
        End If
  End If
Next iUserIndex
End Sub

Private Sub Auditoria_Timer()
On Error GoTo errhand
'Dim k As Integer
'For k = 1 To LastUser
'    If UserList(k).ConnID <> -1 Then
'        DayStats.Segundos = DayStats.Segundos + 1
'    End If
'Next k
Call PasarSegundo 'sistema de desconexion de 10 segs
'Call ActualizaEstadisticasWeb
Call ActualizaStats
Exit Sub
errhand:
Call LogError("Error en Timer Auditoria. Err: " & Err.Description & " - " & Err.Number)
End Sub

Private Sub AutoSave_Timer()
On Error GoTo errhandler
'fired every minute
Static Minutos As Long
Static MinutosLatsClean As Long
Static MinsSocketReset As Long
Static MinsPjesSave As Long
Static MinutosNumUsersCheck As Long
Dim i As Integer
Dim num As Long
'If ReiniciarServer = 666 Then
'#If True = False Then
'    Call SendData(ToAll, 0, 0, "||Servidor> Reiniciando..." & FONTTYPE_SERVER)
'
'    'WorldSave
'    Call DoBackUp
'
'    'Guardar Pjs
'#If UsarQueSocket = 1 Then
'    Call apiclosesocket(SockListen)
'#ElseIf UsarQueSocket = 0 Then
'
'#End If
'
'    For i = 1 To MaxUsers
'        If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
'            Call CloseSocket(i)
'        End If
'    Next i
'
'    'Guilds
'    Call SaveGuildsDB
'
'    ChDrive App.Path
'    ChDir App.Path
'
''    If FileExist(App.Path & "\" & App.EXEName, vbNormal) Then
'        Call Shell(App.Path & "\" & App.EXEName, vbNormalNoFocus)
''    End If
'
'    'Chau
'    Unload frmMain
'
'    Exit Sub
'#End If
'End If
MinsRunning = MinsRunning + 1
If MinsRunning = 60 Then
    Horas = Horas + 1
    If Horas = 24 Then
        Call SaveDayStats
        DayStats.MaxUsuarios = 0
        DayStats.Segundos = 0
        DayStats.Promedio = 0
        Call DayElapsed
        'Dias = Dias + 1
        Horas = 0
'        If AutoReiniciar = 1 Then
'            Call SendData(ToAll, 0, 0, "||Servidor> El servidor se reiniciará en 1 minuto." & FONTTYPE_SERVER)
'            ReiniciarServer = 666
'        End If
    End If
    MinsRunning = 0
End If
Minutos = Minutos + 1
#If UsarQueSocket = 1 Then
' ok la cosa es asi, este cacho de codigo es para
' evitar los problemas de socket. a menos que estes
' seguro de lo que estas haciendo, te recomiendo
' que lo dejes tal cual está.
' alejo.
MinsSocketReset = MinsSocketReset + 1
' cada 1 minutos hacer el checkeo
If MinsSocketReset >= 5 Then
    MinsSocketReset = 0
    For i = 1 To MaxUsers
        If UserList(i).ConnID <> -1 And Not UserList(i).flags.UserLogged Then
            If UserList(i).Counters.IdleCount > ((IntervaloCerrarConexion * 2) / 3) Then
                Call CloseSocket(i)
            End If
        End If
    Next i
    'Call ReloadSokcet
    Call LogCriticEvent("NumUsers: " & NumUsers & " WSAPISock2Usr: " & WSAPISock2Usr.Count)
End If
#End If
MinutosNumUsersCheck = MinutosNumUsersCheck + 1
If MinutosNumUsersCheck >= 2 Then
    MinutosNumUsersCheck = 0
    num = 0
    For i = 1 To MaxUsers
        If UserList(i).ConnID <> -1 And UserList(i).flags.UserLogged Then
            num = num + 1
        End If
    Next i
    If num <> NumUsers Then
        NumUsers = num
        'Call SendData(ToAdmins, 0, 0, "Servidor> Error en NumUsers. Contactar a algun Programador." & FONTTYPE_SERVER)
        Call LogCriticEvent("Num <> NumUsers")
    End If
End If
If Minutos >= MinutosWs Then
    Call DoBackUp
    Call aClon.VaciarColeccion
    Minutos = 0
End If
If MinutosLatsClean >= 15 Then
        MinutosLatsClean = 0
        Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
        Call LimpiarMundo
Else
        MinutosLatsClean = MinutosLatsClean + 1
End If

'[Consejeros]
'If MinsPjesSave >= 30 Then
'    MinsPjesSave = 0
'    Call GuardarUsuarios
'Else
'    MinsPjesSave = MinsPjesSave + 1
'End If
Call PurgarPenas
Call CheckIdleUser
'<<<<<-------- Log the number of users online ------>>>
Dim N As Integer
N = FreeFile(1)
Open App.Path & "\logs\numusers.log" For Output Shared As N
Print #N, NumUsers
Close #N
'<<<<<-------- Log the number of users online ------>>>
Exit Sub
errhandler:
    Call LogError("Error en TimerAutoSave " & Err.Number & ": " & Err.Description)
End Sub

Private Sub CMDDUMP_Click()
On Error Resume Next
Dim i As Integer
For i = 1 To MaxUsers
    Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).Name & " UserLogged: " & UserList(i).flags.UserLogged)
Next i
Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)
End Sub

Private Sub CmdExec_Timer()
Dim i As Integer
Static N As Long
On Error Resume Next ':(((
N = N + 1
For i = 1 To MaxUsers
    If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
        If Not UserList(i).CommandsBuffer.IsEmpty Then
            Call HandleData(i, UserList(i).CommandsBuffer.Pop)
        End If
        If N >= 10 Then
            If UserList(i).ColaSalida.Count > 0 Then ' And UserList(i).SockPuedoEnviar Then
    #If UsarQueSocket = 1 Then
                Call IntentarEnviarDatosEncolados(i)
    '#ElseIf UsarQueSocket = 0 Then
    '            Call WrchIntentarEnviarDatosEncolados(i)
    '#ElseIf UsarQueSocket = 2 Then
    '            Call ServIntentarEnviarDatosEncolados(i)
    #ElseIf UsarQueSocket = 3 Then
        'NADA, el control deberia ocuparse de esto!!!
        'si la cola se llena, dispara un on close
    #End If
            End If
        End If
    End If
Next i
If N >= 10 Then
    N = 0
End If
Exit Sub
hayerror:
End Sub

Private Sub Command1_Click()
'[Misery_Ezequiel 10/06/05]
If Len(BroadMsg.Text) > 0 Then
Call Senddata(ToAll, 0, 0, "!!" & BroadMsg.Text & ENDC)
Else
Exit Sub
End If
'[\]Misery_Ezequiel 10/06/05]
End Sub

Public Sub InitMain(ByVal f As Byte)
If f = 1 Then
    Call mnuSystray_Click
Else
    frmMain.Show
End If
End Sub

Private Sub Command2_Click()
'[Misery_Ezequiel 10/06/05]
If Len(BroadMsg.Text) > 0 Then
Call Senddata(ToAll, 0, 0, "||Servidor> " & BroadMsg.Text & FONTTYPE_SERVER)
Else
Exit Sub
End If
'[\]Misery_Ezequiel 10/06/05]
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   If Not Visible Then
        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hwnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
End Sub

Private Sub QuitarIconoSystray()
On Error Resume Next
'Borramos el icono del systray
Dim i As Integer
Dim nid As NOTIFYICONDATA
nid = setNOTIFYICONDATA(frmMain.hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")
i = Shell_NotifyIconA(NIM_DELETE, nid)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Call QuitarIconoSystray
#If UsarQueSocket = 1 Then
Call LimpiaWsApi(frmMain.hwnd)
#ElseIf UsarQueSocket = 0 Then
Socket1.Cleanup
#ElseIf UsarQueSocket = 2 Then
Serv.Detener
#End If
Call DescargaNpcsDat

Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
Next

'Log
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " server cerrado."
Close #N
End
End Sub



Private Sub GameTimer_Timer()
Dim iUserIndex As Integer
Dim bEnviarStats As Boolean
Dim bEnviarAyS As Boolean
Dim iNpcIndex As Integer
Static lTirarBasura As Long
Static lPermiteAtacar As Long
Static lPermiteCast As Long
Static lPermiteTrabajar As Long
'[Alejo]
If lPermiteAtacar < IntervaloUserPuedeAtacar Then
    lPermiteAtacar = lPermiteAtacar + 1
End If
If lPermiteCast < IntervaloUserPuedeCastear Then
    lPermiteCast = lPermiteCast + 1
End If
If lPermiteTrabajar < IntervaloUserPuedeTrabajar Then
     lPermiteTrabajar = lPermiteTrabajar + 1
End If
'[/Alejo]
On Error GoTo hayerror
 '<<<<<< Procesa eventos de los usuarios >>>>>>
 
For iUserIndex = 1 To LastUser
   'Conexion activa?
   If UserList(iUserIndex).ConnID <> -1 Then
      '¿User valido?
      If UserList(iUserIndex).ConnIDValida And UserList(iUserIndex).flags.UserLogged Then
         '[Alejo-18-5]
         bEnviarStats = False
         bEnviarAyS = False
         UserList(iUserIndex).NumeroPaquetesPorMiliSec = 0

         Call DoTileEvents(iUserIndex, UserList(iUserIndex).Pos.Map, UserList(iUserIndex).Pos.X, UserList(iUserIndex).Pos.Y)
         
        If UserList(iUserIndex).flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
        If UserList(iUserIndex).flags.Ceguera = 1 Or _
            UserList(iUserIndex).flags.Estupidez Then Call EfectoCegueEstu(iUserIndex)
        If UserList(iUserIndex).flags.Muerto = 0 Then

        If UserList(iUserIndex).flags.Desnudo And UserList(iUserIndex).flags.Privilegios = 0 Then Call EfectoFrio(iUserIndex)
        If UserList(iUserIndex).flags.Meditando Then Call DoMeditar(iUserIndex)
        If UserList(iUserIndex).flags.Envenenado = 1 And UserList(iUserIndex).flags.Privilegios = 0 Then Call EfectoVeneno(iUserIndex, bEnviarStats)
        If UserList(iUserIndex).flags.AdminInvisible <> 1 And UserList(iUserIndex).flags.Invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
        If UserList(iUserIndex).flags.Mimetizado = 1 Then Call EfectoMimetismo(iUserIndex)
        Call DuracionPociones(iUserIndex)
        Call HambreYSed(iUserIndex, bEnviarAyS)

'que feo q es esto :S
If Lloviendo Then
            If Not Intemperie(iUserIndex) Then
                    If Not UserList(iUserIndex).flags.Descansar And (UserList(iUserIndex).flags.Hambre = 0 And UserList(iUserIndex).flags.Sed = 0) Then
                        'No esta descansando
                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                        ElseIf UserList(iUserIndex).flags.Descansar Then
                        'esta descansando
                            Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                            Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                                          'termina de descansar automaticamente
                                          If UserList(iUserIndex).Stats.MaxHP = UserList(iUserIndex).Stats.MinHP And _
                                             UserList(iUserIndex).Stats.MaxSta = UserList(iUserIndex).Stats.MinSta Then
                                                    Call Senddata(ToIndex, iUserIndex, 0, "DOK")
                                                    Call Senddata(ToIndex, iUserIndex, 0, "Y130")
                                                    UserList(iUserIndex).flags.Descansar = False
                                          End If
                                 End If 'Not UserList(UserIndex).Flags.Descansar And (UserList(UserIndex).Flags.Hambre = 0 And UserList(UserIndex).Flags.Sed = 0)
                    End If
            Else
            
                    If Not UserList(iUserIndex).flags.Descansar And (UserList(iUserIndex).flags.Hambre = 0 And UserList(iUserIndex).flags.Sed = 0) And UserList(iUserIndex).flags.Desnudo = 0 Then
                    'No esta descansando
                             Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                             Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                    ElseIf UserList(iUserIndex).flags.Descansar Then
                    'esta descansando
                             Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                             Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                             'termina de descansar automaticamente
                             If UserList(iUserIndex).Stats.MaxHP = UserList(iUserIndex).Stats.MinHP And _
                                UserList(iUserIndex).Stats.MaxSta = UserList(iUserIndex).Stats.MinSta Then
                                     Call Senddata(ToIndex, iUserIndex, 0, "DOK")
                                     Call Senddata(ToIndex, iUserIndex, 0, "Y130")
                                     UserList(iUserIndex).flags.Descansar = False
                             End If
                    End If 'Not UserList(UserIndex).Flags.Descansar And (UserList(UserIndex).Flags.Hambre = 0 And UserList(UserIndex).Flags.Sed = 0)
               End If
            End If
               
               If bEnviarStats Then Call SendUserStatsBox(iUserIndex)
               If bEnviarAyS Then Call EnviarHambreYsed(iUserIndex)
               If UserList(iUserIndex).NroMacotas > 0 Then Call TiempoInvocacion(iUserIndex)
       End If 'Muerto
       
       
     Else 'no esta logeado?
        UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
        If UserList(iUserIndex).Counters.IdleCount > IntervaloParaConexion Then
              UserList(iUserIndex).Counters.IdleCount = 0
              Call CloseSocket(iUserIndex)
        End If
End If 'UserLogged

Next iUserIndex

If Not lPermiteAtacar < IntervaloUserPuedeAtacar Then
    lPermiteAtacar = 0
End If
If Not lPermiteCast < IntervaloUserPuedeCastear Then
    lPermiteCast = 0
End If
If Not lPermiteTrabajar < IntervaloUserPuedeTrabajar Then
     lPermiteTrabajar = 0
End If
Exit Sub
hayerror:
'[/Alejo]
  'DoEvents
End Sub

Private Sub mnuCerrar_Click()
Call SaveGuildsDB
If MsgBox("¡¡Atencion!! Si cierra el servidor puede provocar la perdida de datos. ¿Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
    Dim f
    For Each f In Forms
        Unload f
    Next
End If
End Sub

Private Sub mnusalir_Click()
    Call mnuCerrar_Click
End Sub

Public Sub mnuMostrar_Click()
On Error Resume Next
    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0
End Sub

Private Sub mnuServidor_Click()
frmServidor.Visible = True
End Sub

Private Sub mnuSystray_Click()
Dim i As Integer
Dim s As String
Dim nid As NOTIFYICONDATA

s = "ARGENTUM-ONLINE"
nid = setNOTIFYICONDATA(frmMain.hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, s)
i = Shell_NotifyIconA(NIM_ADD, nid)
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False
End Sub

Private Sub npcataca_Timer()
On Error Resume Next
Dim npc As Integer

For npc = 1 To LastNPC
    Npclist(npc).CanAttack = 1
Next npc
End Sub

'#If UsarQueSocket = 2 Then
'
'
'Private Sub Serv_Close(ByVal ID As Long)
'#If UsarQueSocket = 2 Then
'
'Dim UserIndex As Integer
'
'UserIndex = CInt(Serv.GetDato(ID))
'
'If UserIndex > 0 Then
'    If UserList(UserIndex).flags.UserLogged Then
'        Call CloseSocketSL(UserIndex)
'        Call Cerrar_Usuario(UserIndex)
'    Else
'        Call CloseSocket(UserIndex)
'    End If
'End If
'
'#End If
'End Sub
'
'Private Sub Serv_Eror(ByVal Numero As Long, ByVal Descripcion As String)
'#If UsarQueSocket = 2 Then
'Call LogCriticEvent("Serv_Eror " & Numero & ": " & Descripcion)
'#End If
'End Sub
'
'Private Sub Serv_NuevaConn(ByVal ID As Long)
'#If UsarQueSocket = 2 Then
''==========================================================
'
'If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "Pedido de conexion SocketID:" & ID & vbCrLf
'
'On Error Resume Next
'
'    Dim NewIndex As Integer
'    Dim Ret As Long
'    Dim i As Long
'    Dim tStr As String
'
'    If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "NextOpenUser" & vbCrLf
'
'    NewIndex = NextOpenUser ' Nuevo indice
'    If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "UserIndex asignado " & NewIndex & vbCrLf
'
'    If NewIndex <= MaxUsers Then
'        If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "Cargando Socket " & NewIndex & vbCrLf
'
'
'        UserList(NewIndex).ip = Serv.GetIP(ID)
'
'        'Busca si esta banneada la ip
'        For i = 1 To BanIps.Count
'            If BanIps.Item(i) = UserList(NewIndex).ip Then
'                Call Serv.CerrarSocket(ID)
'                Exit Sub
'            End If
'        Next i
'
'
'        '=============================================
'        If aDos.MaxConexiones(UserList(NewIndex).ip) Then
'            UserList(NewIndex).ConnID = -1
'            If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "User slot reseteado " & NewIndex & vbCrLf
'            If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "Socket unloaded" & NewIndex & vbCrLf
'            'Call LogCriticEvent(UserList(NewIndex).ip & " intento crear mas de 3 conexiones.")
'            Call aDos.RestarConexion(UserList(NewIndex).ip)
'            Call Serv.CerrarSocket(ID)
'            'Exit Sub
'        End If
'
'        Call Serv.SetDato(ID, NewIndex)
'
'        UserList(NewIndex).SockPuedoEnviar = True
'        UserList(NewIndex).ConnID = ID
'        UserList(NewIndex).ConnIDValida = True
'        Set UserList(NewIndex).CommandsBuffer = New CColaArray
'        Set UserList(NewIndex).ColaSalida = New Collection
'
'        If NewIndex > LastUser Then LastUser = NewIndex
'
''        Debug.Print "Conexion desde " & UserList(NewIndex).ip
'
'        If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & UserList(NewIndex).ip & " logged." & vbCrLf
'    Else
'        Call LogCriticEvent("No acepte conexion porque no tenia slots")
'
'        tStr = "ERRServer lleno" & ENDC
'        Call Serv.Enviar(ID, tStr, Len(tStr))
'        Call Serv.CerrarSocket(ID)
'    End If
'
'#End If
'End Sub
'
'Private Sub Serv_Read(ByVal ID As Long, ByVal Datos As String, ByVal Cantidad As Long)
'#If UsarQueSocket = 2 Then
'
'Dim t() As String
'Dim LoopC As Long
'Dim UserIndex As Integer
'
'UserIndex = CInt(Serv.GetDato(ID))
'
'If UserIndex > 0 Then
'    TCPESStats.BytesRecibidos = TCPESStats.BytesRecibidos + Len(Datos)
'
'    UserList(UserIndex).RDBuffer = UserList(UserIndex).RDBuffer & Datos
'
'    'If InStr(1, UserList(Slot).RDBuffer, Chr(2)) > 0 Then
'    '    UserList(Slot).RDBuffer = "CLIENTEVIEJO" & ENDC
'    '    Debug.Print "CLIENTEVIEJO"
'    'End If
'
'    t = Split(UserList(UserIndex).RDBuffer, ENDC)
'    If UBound(t) > 0 Then
'        UserList(UserIndex).RDBuffer = t(UBound(t))
'
'        For LoopC = 0 To UBound(t) - 1
'            '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'            '%%% SI ESTA OPCION SE ACTIVA SOLUCIONA %%%
'            '%%% EL PROBLEMA DEL SPEEDHACK          %%%
'            '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'            If ClientsCommandsQueue = 1 Then
'                If t(LoopC) <> "" Then If Not UserList(UserIndex).CommandsBuffer.Push(t(LoopC)) Then Call Cerrar_Usuario(UserIndex)
'
'            Else ' SH tiebe efecto
'                  If UserList(UserIndex).ConnID <> -1 Then
'                    Call HandleData(UserIndex, t(LoopC))
'                  Else
'                    Exit Sub
'                  End If
'            End If
'        Next LoopC
'    End If
'End If
'
'#End If
'End Sub
'
'Private Sub Serv_Write(ByVal ID As Long)
'#If UsarQueSocket = 2 Then
'
'#End If
'End Sub
'
'#End If
'
'Private Sub Socket1_Accept(SocketId As Integer)
'#If UsarQueSocket = 0 Then
'
''=========================================================
''USO DEL CONTROL SOCKET WRENCH
''=============================
'
'If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "Pedido de conexion SocketID:" & SocketId & vbCrLf
'
'On Error Resume Next
'
'    Dim NewIndex As Integer
'
'
'    If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "NextOpenUser" & vbCrLf
'
'    NewIndex = NextOpenUser ' Nuevo indice
'    If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "UserIndex asignado " & NewIndex & vbCrLf
'
'    If NewIndex >= 1 And NewIndex <= MaxUsers Then
'            If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "Cargando Socket " & NewIndex & vbCrLf
'
'            Unload Socket2(NewIndex)
'            Load Socket2(NewIndex)
'
'            Socket2(NewIndex).AddressFamily = AF_INET
'            Socket2(NewIndex).protocol = IPPROTO_IP
'            Socket2(NewIndex).SocketType = SOCK_STREAM
'            Socket2(NewIndex).Binary = False
'            Socket2(NewIndex).BufferSize = SOCKET_BUFFER_SIZE
'            Socket2(NewIndex).Blocking = False
'            Socket2(NewIndex).Linger = 1
'
'            Socket2(NewIndex).accept = SocketId
'
'            UserList(NewIndex).ip = Socket2(NewIndex).PeerAddress
'            If BanIpBuscar(UserList(NewIndex).ip) > 0 Then
'                Call CloseSocket(NewIndex)
'                Exit Sub
'            End If
'
'
'            If aDos.MaxConexiones(Socket2(NewIndex).PeerAddress) Then
'
'                UserList(NewIndex).ConnID = -1
'                If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "User slot reseteado " & NewIndex & vbCrLf
'
'
'
'                If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "Socket unloaded" & NewIndex & vbCrLf
'
'                'Call LogCriticEvent(Socket2(NewIndex).PeerAddress & " intento crear mas de 3 conexiones.")
'                Call aDos.RestarConexion(Socket2(NewIndex).PeerAddress)
'                'Socket2(NewIndex).Disconnect
'                Unload frmMain.Socket2(NewIndex)
'
'                Exit Sub
'            End If
'
'            Set UserList(NewIndex).CommandsBuffer = New CColaArray
'            Set UserList(NewIndex).ColaSalida = New Collection
'            UserList(NewIndex).ConnIDValida = True
'            UserList(NewIndex).ConnID = SocketId
'            UserList(NewIndex).SockPuedoEnviar = True
'
'            If NewIndex > LastUser Then
'                LastUser = NewIndex
'                If LastUser > MaxUsers Then
'                    LastUser = MaxUsers
'                    Call CloseSocket(NewIndex)
'                End If
'            End If
'
'            If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & Socket2(NewIndex).PeerAddress & " logged." & vbCrLf
'    Else
'        Call LogCriticEvent("No acepte conexion porque no tenia slots")
'    End If
'
'Exit Sub
'
'#End If
'End Sub
'
'Private Sub Socket1_Blocking(Status As Integer, Cancel As Integer)
'Cancel = True
'End Sub
'
'Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
'
'If DebugSocket Then frmDebugSocket.Text2.Text = frmDebugSocket.Text2.Text & Time & " " & ErrorString & vbCrLf
'
'frmDebugSocket.Label3.Caption = Socket1.State
'End Sub
'
'Private Sub Socket1_Write()
''
'
'End Sub
'
'Private Sub Socket2_Blocking(Index As Integer, Status As Integer, Cancel As Integer)
''Cancel = True
'End Sub
'
'Private Sub Socket2_Connect(Index As Integer)
''If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "Conectado" & vbCrLf
'
'On Error Resume Next
'
'If Index >= LBound(UserList) And Index <= UBound(UserList) Then
'    Set UserList(Index).CommandsBuffer = New CColaArray
'End If
'
'End Sub
'
'Private Sub Socket2_Disconnect(Index As Integer)
'On Error GoTo hayerror
'
'    If UserList(Index).flags.UserLogged And _
'        UserList(Index).Counters.Saliendo = False Then
'        Call Cerrar_Usuario(Index)
'    ElseIf Not UserList(Index).flags.UserLogged Then
'        Call CloseSocket(Index)
'    Else
'        Call CloseSocketSL(Index)
'    End If
'
'Exit Sub
'hayerror:
'
'End Sub
'
''Private Sub Socket2_LastError(Index As Integer, ErrorCode As Integer, ErrorString As String, Response As Integer)
'''24004   WSAEINTR    Blocking function was canceled
'''24009   WSAEBADF    Invalid socket descriptor passed to function
'''24013   WSAEACCES   Access denied
'''24014   WSAEFAULT   Invalid address passed to function
'''24022   WSAEINVAL   Invalid socket function call
'''24024   WSAEMFILE   No socket descriptors are available
'''24035   WSAEWOULDBLOCK  Socket would block on this operation
'''24036   WSAEINPROGRESS  Blocking function in progress
'''24037   WSAEALREADY Function being canceled has already completed
'''24038   WSAENOTSOCK Invalid socket descriptor passed to function
'''24039   WSAEDESTADDRREQ Destination address is required
'''24040   WSAEMSGSIZE Datagram was too large to fit in specified buffer
'''24041   WSAEPROTOTYPE   Specified protocol is the wrong type for this socket
'''24042   WSAENOPROTOOPT  Socket option is unknown or unsupported
'''24043   WSAEPROTONOSUPPORT  Specified protocol is not supported
'''24044   WSAESOCKTNOSUPPORT  Specified socket type is not supported in this address family
'''24045   WSAEOPNOTSUPP   Socket operation is not supported
'''24046   WSAEPFNOSUPPORT Specified protocol family is not supported
'''24047   WSAEAFNOSUPPORT Specified address family is not supported by this protocol
'''24048   WSAEADDRINUSE   Specified address is already in use
'''24049   WSAEADDRNOTAVAIL    Specified address is not available
'''24050   WSAENETDOWN Network subsystem has failed
'''24051   WSAENETUNREACH  Network cannot be reached from this host
'''24052   WSAENETRESET    Network dropped connection on reset
'''24053   WSAECONNABORTED Connection was aborted due to timeout or other failure
'''24054   WSAECONNRESET   Connection was reset by remote network
'''24055   WSAENOBUFS  No buffer space is available
'''24056   WSAEISCONN  Socket is already connected
'''24057   WSAENOTCONN Socket Is Not Connected
'''24058   WSAESHUTDOWN    Socket connection has been shut down
'''24060   WSAETIMEDOUT    Operation timed out before completion
'''24061   WSAECONNREFUSED Connection refused by remote network
'''24064   WSAEHOSTDOWN    Remote host is down
'''24065   WSAEHOSTUNREACH Remote host is unreachable
'''24091   WSASYSNOTREADY  Network subsystem is not ready for communication
'''24092   WSAVERNOTSUPPORTED  Requested version is not available
'''24093   WSANOTINITIALIZED   Windows sockets library not initialized
'''25001   WSAHOST_NOT_FOUND   Authoritative Answer Host not found
'''25002   WSATRY_AGAIN    Non-authoritative Answer Host not found
'''25003   WSANO_RECOVERY  Non-recoverable error
'''25004   WSANO_DATA  No data record of requested type
'''Response = SOCKET_ERRIGNORE
''If ErrorCode = 24053 Then Call CloseSocket(Index)
''End Sub
'
'Private Sub Socket2_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)
'#If UsarQueSocket = 0 Then
'
'On Error GoTo ErrorHandler
'
''*********************************************
''Separamos las lineas con ENDC y las enviamos a HandleData()
''*********************************************
'Dim LoopC As Integer
'Dim RD As String
'Dim rBuffer(1 To COMMAND_BUFFER_SIZE) As String
'Dim CR As Integer
'Dim tChar As String
'Dim sChar As Integer
'Dim eChar As Integer
'Dim aux$
'Dim OrigCad As String
'
'Dim LenRD As Long
'
''<<<<<<<<<<<<<<<<<< Evitamos DoS >>>>>>>>>>>>>>>>>>>>>>>>>>>
''Call AddtoVar(UserList(Index).NumeroPaquetesPorMiliSec, 1, 1000)
''
''If UserList(Index).NumeroPaquetesPorMiliSec > 700 Then
''   'UserList(Index).Flags.AdministrativeBan = 1
''   Call LogCriticalHackAttemp(UserList(Index).Name & " " & frmMain.Socket2(Index).PeerAddress & " alcanzo el max paquetes por iteracion.")
''   Call SendData(ToIndex, Index, 0, "ERRSe ha perdido la conexion, por favor vuelva a conectarse.")
''   Call CloseSocket(Index)
''   Exit Sub
''End If
'
'Call Socket2(Index).Read(RD, DataLength)
'
'OrigCad = RD
'LenRD = Len(RD)
'
''Call AddtoVar(UserList(Index).BytesTransmitidosUser, LenB(RD), 100000)
'
''[¡¡BUCLE INFINITO!!]'
'If LenRD = 0 Then
'    UserList(Index).AntiCuelgue = UserList(Index).AntiCuelgue + 1
'    If UserList(Index).AntiCuelgue >= 150 Then
'        UserList(Index).AntiCuelgue = 0
'        Call LogError("!!!! Detectado bucle infinito de eventos socket2_read. cerrando indice " & Index)
'        Unload Socket2(Index)
'        Call CloseSocket(Index)
'        Exit Sub
'    End If
'Else
'    UserList(Index).AntiCuelgue = 0
'End If
''[¡¡BUCLE INFINITO!!]'
'
''Verificamos por una comando roto y le agregamos el resto
'If UserList(Index).RDBuffer <> "" Then
'    RD = UserList(Index).RDBuffer & RD
'    UserList(Index).RDBuffer = ""
'End If
'
''Verifica por mas de una linea
'sChar = 1
'For LoopC = 1 To LenRD
'
'    tChar = Mid$(RD, LoopC, 1)
'
'    If tChar = ENDC Then
'        CR = CR + 1
'        eChar = LoopC - sChar
'        rBuffer(CR) = Mid$(RD, sChar, eChar)
'        sChar = LoopC + 1
'    End If
'
'Next LoopC
'
''Verifica una linea rota y guarda
'If Len(RD) - (sChar - 1) <> 0 Then
'    UserList(Index).RDBuffer = Mid$(RD, sChar, Len(RD))
'End If
'
''Enviamos el buffer al manejador
'For LoopC = 1 To CR
'
'    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'    '%%% SI ESTA OPCION SE ACTIVA SOLUCIONA %%%
'    '%%% EL PROBLEMA DEL SPEEDHACK          %%%
'    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'    If ClientsCommandsQueue = 1 Then
'        If rBuffer(LoopC) <> "" Then If Not UserList(Index).CommandsBuffer.Push(rBuffer(LoopC)) Then Call Cerrar_Usuario(Index)
'
'    Else ' SH tiebe efecto
'          If UserList(Index).ConnID <> -1 Then
'            Call HandleData(Index, rBuffer(LoopC))
'          Else
'            Exit Sub
'          End If
'    End If
'
'Next LoopC
'
'Exit Sub
'
'
'ErrorHandler:
'    Call LogError("Error en Socket read." & Err.Description & " Numero paquetes:" & UserList(Index).NumeroPaquetesPorMiliSec & " . Rdata:" & OrigCad)
'
'#End If
'End Sub
'
'
'
'Private Sub Socket2_Write(Index As Integer)
''On Error Resume Next
''
''If Index >= LBound(UserList) And Index <= UBound(UserList) Then
''    UserList(Index).SockPuedoEnviar = True
''End If
''
'End Sub

Private Sub TIMER_AI_Timer()
On Error GoTo ErrorHandler
Dim NpcIndex As Integer
Dim X As Integer
Dim Y As Integer
Dim UseAI As Integer
Dim mapa As Integer
Dim e_p As Integer
'Barrin 29/9/03
If Not haciendoBK And Not EnPausa Then
    'Update NPCs
    For NpcIndex = 1 To LastNPC
        If Npclist(NpcIndex).flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
            e_p = esPretoriano(NpcIndex)
            If e_p > 0 Then
                If Npclist(NpcIndex).flags.Paralizado = 1 Then Call EfectoParalisisNpc(NpcIndex)
                Select Case e_p
                    Case 1  ''clerigo
                        Call PRCLER_AI(NpcIndex)
                    Case 2  ''mago
                        Call PRMAGO_AI(NpcIndex)
                    Case 3  ''cazador
                        Call PRCAZA_AI(NpcIndex)
                    Case 4  ''rey
                        Call PRREY_AI(NpcIndex)
                    Case 5  ''guerre
                        Call PRGUER_AI(NpcIndex)
                End Select
            Else
                ''ia comun
                If Npclist(NpcIndex).flags.Paralizado = 1 Then
                      Call EfectoParalisisNpc(NpcIndex)
                Else
                '[Misery_Ezequiel 12/06/05]
                     'Usamos AI si hay algun user en el mapa
                     If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
                        Call EfectoParalisisNpc(NpcIndex)
                     End If
                '[\]Misery_Ezequiel 12/06/05]
                     mapa = Npclist(NpcIndex).Pos.Map
                     If mapa > 0 Then
                          If MapInfo(mapa).NumUsers > 0 Then
                                  If Npclist(NpcIndex).Movement <> ESTATICO Then
                                        Call NPCAI(NpcIndex)
                                  End If
                          End If
                     End If
                     
                End If
            End If
        End If
    Next NpcIndex
End If
Exit Sub
ErrorHandler:
 
 If Npclist(NpcIndex).MaestroUser > 0 Then Exit Sub
 Call MuereNpc(NpcIndex, 0)
 Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).Name & " mapa:" & Npclist(NpcIndex).Pos.Map)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim i As Integer
For i = 1 To LastUser
    If UserList(i).flags.UserLogged Then _
        If UserList(i).flags.Oculto = 1 Then Call DoPermanecerOculto(i)
Next i
End Sub

Private Sub Timer2_Timer()
Dim Msgs As String
Msgs = Conteo
If Conteo = 0 Then Msgs = "YA"
Call Senddata(ToAll, 0, 0, "||Conteo> " & Msgs & FONTTYPE_SERVER)
If Conteo <> 0 Then
Conteo = val(Conteo) - 1
End If
If Msgs = "YA" Then Timer2.Enabled = False
End Sub



Private Sub Timer4_Timer()

End Sub

Private Sub tLluvia_Timer()
'Aca me parece que "Mysery_Ezequiel", por que voy a poner que cuando recupere
' se fije si esta lllvoeidno y si llueve que recupere mas lento
'por que asi lo dijo ar en balance
'aca solo queda para la nieve
' asi este timer casi ni existe
'Solo nieva en los mapas 169, 170, 171

'Esto es la nieve enrealidad
On Error GoTo errhandler

Dim iCount As Integer

For iCount = 1 To LastUser
If UserList(iCount).flags.Muerto = 0 And UserList(iCount).flags.Privilegios = 0 Then
    If UserList(iCount).Pos.Map = 169 Or UserList(iCount).Pos.Map = 170 Or UserList(iCount).Pos.Map = 171 Then
        If UserList(iCount).Invent.ArmourEqpObjIndex = 665 Or UserList(iCount).Invent.ArmourEqpObjIndex = 666 Or UserList(iCount).Invent.ArmourEqpObjIndex = 667 Then
        Else
            If Nevando Then
            Call EfectoNevando(iCount)
            Call SendUserVida(iCount)
            Else
            Call EfectoNieve(iCount)
            Call SendUserVida(iCount)
            End If
        End If
        
    End If
    
    If Lloviendo Then
        '[Wizard]
        If UserList(iCount).flags.Desnudo Or UserList(iCount).Stats.MinAGU = 0 Or UserList(iCount).Stats.MinHam = 0 Then GoTo 1
        UserList(iCount).Stats.MinSta = UserList(iCount).Stats.MinSta + CInt(RandomNumber(1, Porcentaje(UserList(iCount).Stats.MaxSta, 5)))
        If UserList(iCount).Stats.MinSta > UserList(iCount).Stats.MaxSta Then UserList(iCount).Stats.MinSta = UserList(iCount).Stats.MaxSta
        Call SendUserEsta(iCount)
    End If
End If
1
Next iCount

Exit Sub
errhandler:
Call LogError("tLluvia " & Err.Number & ": " & Err.Description)
End Sub


'Private Sub tPiqueteC_Timer()
'Esto lo borro por que soy tan powa que lo hize en cliente
''''''''''''''''''Echo por Marche'''''''''''''''''''''''''
'On Error GoTo errhandler
'Static Segundos As Integer
'Segundos = Segundos + 6
'Dim i As Integer
'For i = 1 To LastUser
'    If UserList(i).flags.UserLogged Then
 '           If MapData(UserList(i).Pos.Map, UserList(i).Pos.X, UserList(i).Pos.Y).trigger = TRIGGER_ANTIPIQUETE Then
  '                  UserList(i).Counters.PiqueteC = UserList(i).Counters.PiqueteC + 1
   '                 Call Senddata(ToIndex, i, 0, "Y131")
       '             If UserList(i).Counters.PiqueteC > 23 Then
    '                        UserList(i).Counters.PiqueteC = 0
     '                       Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)
      '              End If
        '    Else
         '           If UserList(i).Counters.PiqueteC > 0 Then UserList(i).Counters.PiqueteC = 0
          '  End If
            
           ' If Segundos >= 18 Then
'                Dim nfile As Integer
'                nfile = FreeFile ' obtenemos un canal
'                Open App.Path & "\logs\maxpasos.log" For Append Shared As #nfile
'                Print #nfile, UserList(i).Counters.Pasos
'                Close #nfile
            '    If Segundos >= 18 Then UserList(i).Counters.Pasos = 0
           ' End If
            
    'End If
'Next i
'If Segundos >= 18 Then Segundos = 0
'Exit Sub
'errhandler:
'    Call LogError("Error en tPiqueteC_Timer " & Err.Number & ": " & Err.Description)
'End Sub

Private Sub tTraficStat_Timer()
'Dim i As Integer
'
'If frmTrafic.Visible Then frmTrafic.lstTrafico.Clear
'
'For i = 1 To LastUser
'    If UserList(i).Flags.UserLogged Then
'        If frmTrafic.Visible Then
'            frmTrafic.lstTrafico.AddItem UserList(i).Name & " " & UserList(i).BytesTransmitidosUser + UserList(i).BytesTransmitidosSvr & " bytes per second"
'        End If
'        UserList(i).BytesTransmitidosUser = 0
'        UserList(i).BytesTransmitidosSvr = 0
'    End If
'Next i
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''USO DEL CONTROL TCPSERV'''''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

#If UsarQueSocket = 3 Then

Private Sub TCPServ_Eror(ByVal Numero As Long, ByVal Descripcion As String)
    Call LogError("TCPSERVER SOCKET ERROR: " & Numero & "/" & Descripcion)
End Sub

Private Sub TCPServ_NuevaConn(ByVal ID As Long)
On Error GoTo errorHandlerNC
    ESCUCHADAS = ESCUCHADAS + 1
    Escuch.Caption = ESCUCHADAS
    Dim i As Integer
    Dim NewIndex As Integer
    NewIndex = NextOpenUser
    
    If NewIndex <= MaxUsers Then
        'call logindex(NewIndex, "******> Accept. ConnId: " & ID)
        TCPServ.SetDato ID, NewIndex
        UserList(NewIndex).CryptOffset = 0 'Gorlok
        If aDos.MaxConexiones(TCPServ.GetIP(ID)) Then
            Call aDos.RestarConexion(TCPServ.GetIP(ID))
            Call ResetUserSlot(NewIndex)
            Exit Sub
        End If
        If NewIndex > LastUser Then LastUser = NewIndex
        UserList(NewIndex).ConnID = ID
        UserList(NewIndex).ip = TCPServ.GetIP(ID)
        UserList(NewIndex).ConnIDValida = True
        Set UserList(NewIndex).CommandsBuffer = New CColaArray
        For i = 1 To BanIps.Count
            If BanIps.Item(i) = TCPServ.GetIP(ID) Then
                Call ResetUserSlot(NewIndex)
                Exit Sub
            End If
        Next i
    Else
        Call CloseSocket(NewIndex, True)
        LogCriticEvent ("NEWINDEX > MAXUSERS. IMPOSIBLE ALOCATEAR SOCKETS")
    End If
Exit Sub
errorHandlerNC:
Call LogError("TCPServer::NuevaConexion " & Err.Description)
End Sub

Private Sub TCPServ_Close(ByVal ID As Long, ByVal MiDato As Long)
    On Error GoTo eh
    '' No cierro yo el socket. El on_close lo cierra por mi.
    'call logindex(MiDato, "******> Remote Close. ConnId: " & ID & " Midato: " & MiDato)
    Call CloseSocket(MiDato, False)
Exit Sub
eh:
    Call LogError("Ocurrio un error en el evento TCPServ_Close. ID/miDato:" & ID & "/" & MiDato)
End Sub

Private Sub TCPServ_Read(ByVal ID As Long, Datos As Variant, ByVal Cantidad As Long, ByVal MiDato As Long)
Dim t() As String
Dim LoopC As Long
Dim RD As String
On Error GoTo errorh
If UserList(MiDato).ConnID <> UserList(MiDato).ConnID Then
    Call LogError("Recibi un read de un usuario con ConnId alterada")
    Exit Sub
End If
RD = StrConv(Datos, vbUnicode)
'call logindex(MiDato, "Read. ConnId: " & ID & " Midato: " & MiDato & " Dato: " & RD)
UserList(MiDato).RDBuffer = UserList(MiDato).RDBuffer & RD
t = Split(UserList(MiDato).RDBuffer, ENDC)
If UBound(t) > 0 Then
    UserList(MiDato).RDBuffer = t(UBound(t))
    For LoopC = 0 To UBound(t) - 1
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        '%%% SI ESTA OPCION SE ACTIVA SOLUCIONA %%%
        '%%% EL PROBLEMA DEL SPEEDHACK          %%%
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        If ClientsCommandsQueue = 1 Then
            If t(LoopC) <> "" Then
                If Not UserList(MiDato).CommandsBuffer.Push(t(LoopC)) Then
                    Call LogError("Cerramos por no encolar. Userindex:" & MiDato)
                    Call CloseSocket(MiDato)
                End If
            End If
        Else ' no encolamos los comandos (MUY VIEJO)
              If UserList(MiDato).ConnID <> -1 Then
                Call HandleData(MiDato, t(LoopC))
              Else
                Exit Sub
              End If
        End If
    Next LoopC
End If
Exit Sub
errorh:
Call LogError("Error socket read: " & MiDato & " dato:" & RD & " userlogged: " & UserList(MiDato).flags.UserLogged & " connid:" & UserList(MiDato).ConnID & " ID Parametro" & ID & " error:" & Err.Description)
End Sub
#End If

Private Sub Trabajo_Timer()
Dim i As Integer
'1.pescar
'2.talar
'3.minar
'4.lingotear

    For i = 1 To LastUser  'Agregue un grupo de trabajadores. Es una locura que se fije que esta haciendo cada usuario
         If UserList(i).ConnIDValida Then
         If UserList(i).TyTrabajo > 0 And UserList(i).flags.Trabajando Then
            If UserList(i).TyTrabajo = 1 Then
                If UserList(i).TyTrabajoMod = OBJTYPE_CAÑA Then
                    Call DoPescar(i)
                Else
                    Call DoPescarRed(i)
                End If
                Exit Sub
            ElseIf UserList(i).TyTrabajo = 2 Then
                Call DoTalar(i, UserList(i).TyTrabajoMod)
                Exit Sub
            ElseIf UserList(i).TyTrabajo = 3 Then
                Call DoMineria(i)
            ElseIf UserList(i).TyTrabajo = 4 Then
                Call FundirMineral(i, UserList(i).Suerte)
            End If
         Else
         If UserList(i).TyTrabajo > 0 Then Call DejarDeTrabajar(i)
         End If
         End If
    Next i
    
    
End Sub
