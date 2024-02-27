VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Argentum Online"
   ClientHeight    =   5415
   ClientLeft      =   1950
   ClientTop       =   1815
   ClientWidth     =   7245
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
   ScaleHeight     =   5415
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.CheckBox chkForzarDia 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Forzar Dia"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2520
      TabIndex        =   27
      Top             =   4200
      Width           =   1215
   End
   Begin VB.HScrollBar scrollfraccionDelDia 
      Height          =   255
      Left            =   2640
      Max             =   96
      Min             =   1
      TabIndex        =   25
      Top             =   4560
      Value           =   1
      Width           =   2175
   End
   Begin VB.TextBox txtModificaroTrabajo 
      Height          =   315
      Left            =   5400
      TabIndex        =   24
      Text            =   "1"
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdActualizarTrabajo 
      Caption         =   "Actualizar trabajo"
      Height          =   360
      Left            =   5400
      TabIndex        =   23
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox Sensibilidad 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      TabIndex        =   18
      Text            =   "3"
      Top             =   3120
      Width           =   615
   End
   Begin MSWinsockLib.Winsock WinsockWeb 
      Index           =   0
      Left            =   6480
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Conexiones originadas por la web"
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   4815
      Begin VB.Label cantidadConexionesWeb 
         Caption         =   "0"
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.Label labelCantidadConexionesWeb 
         Caption         =   "Cantidad:"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Timer timerTrabajo 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6480
      Top             =   2040
   End
   Begin VB.Timer AntiMacrosCen 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6360
      Top             =   1320
   End
   Begin VB.Timer AntiMacros 
      Interval        =   60000
      Left            =   6720
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5400
      Top             =   1320
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5400
      Top             =   2760
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5400
      Top             =   480
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5400
      Top             =   2040
   End
   Begin VB.Timer TIMER_AI 
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   6480
      Top             =   480
   End
   Begin VB.Frame Frame1 
      Caption         =   "BroadCast"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   360
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
   Begin VB.Label lblNoche 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hora del dia:"
      Height          =   210
      Left            =   3000
      TabIndex        =   26
      Top             =   4920
      Width           =   1005
   End
   Begin VB.Label lblMensajeError 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje Error"
      Height          =   570
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   6930
   End
   Begin VB.Label estadoServidor 
      BackStyle       =   0  'Transparent
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   2040
      TabIndex        =   21
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label nuevoSocket 
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   20
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Sensibilidad"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "PasarSegundo"
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   14
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajo"
      Height          =   255
      Left            =   6360
      TabIndex        =   13
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Auco"
      Height          =   255
      Left            =   5400
      TabIndex        =   12
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "GameTimer"
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   11
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "PasarMinuto"
      Height          =   255
      Left            =   5160
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Npcs"
      Height          =   255
      Left            =   6480
      TabIndex        =   9
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Centinelas"
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "¡ TDS 2020 !"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1200
      TabIndex        =   7
      Top             =   1560
      Width           =   3375
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
      Top             =   120
      Width           =   1725
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   1200
      TabIndex        =   0
      Top             =   4320
      Width           =   765
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
Option Explicit

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Const NIM_ADD = 0
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4
Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONUP = &H205
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private LastIndex As Integer


Private Function setNOTIFYICONDATA(hWnd As Long, id As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA
    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uID = id
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)
    setNOTIFYICONDATA = nidTemp
End Function

'---------------------------------------------------------------------------------------
' Procedure : CheckIdleUser
' DateTime  : 18/02/2007 21:26
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub CheckIdleUser()
Dim iUserIndex As Integer

For iUserIndex = 1 To LastUser

   'Conexion activa? y es un usuario loggeado?
   If Not UserList(iUserIndex).ConnID = INVALID_SOCKET Then
   
        'Actualiza el contador de inactividad
        UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
   
        ' ¿Esta jugando?
        If UserList(iUserIndex).flags.UserLogged Then
            ' El tiempo no avanza para personajes inmovializados
            If Not (UserList(iUserIndex).flags.Privilegios = 0 And _
                UserList(iUserIndex).flags.Inmovilizado = 0 And _
                UserList(iUserIndex).flags.Paralizado = 0) Then
                
                UserList(iUserIndex).Counters.IdleCount = 0
            End If
    
            If UserList(iUserIndex).Counters.IdleCount > 5 Then
                EnviarPaquete Paquetes.MensajeFight, "Demasiado tiempo inactivo...", iUserIndex
                Call Cerrar_Usuario(UserList(iUserIndex))
            End If
            
        Else ' Creando personaje
    
            If UserList(iUserIndex).Counters.IdleCount > 10 Then
                If Not CloseSocket(iUserIndex) Then LogError ("Check Idle Crear")
            End If

        End If
  End If
  
Next iUserIndex

End Sub

Private Sub AntiMacros_Timer()
Call modCentinelas.AntiMacrosL 'EL YIND
End Sub

Private Sub AntiMacrosCen_Timer() 'EL YIND
Call modCentinelas.procesarCentinelas
End Sub

Private Sub Auditoria_Timer()
    Call PasarSegundo 'sistema de desconexion de 10 segs
End Sub

Private Sub calcularWorldSave()
    Static minutosSinWorldSave As Long

    minutosSinWorldSave = minutosSinWorldSave + 1

    'Es momento de hacer worldSave?. Definido en el archivo de configuración
    If minutosSinWorldSave >= modConfig.MinutosWs Then
        Call DoBackUp
        If Admin.servidorAtacado Then Admin.servidorTerminaAtaque
        minutosSinWorldSave = 0
    ElseIf minutosSinWorldSave + 3 = MinutosWs Then '3 minutos antes de que se haga, aviso.
        EnviarPaquete Paquetes.MensajeServer, "En 3 minutos se realizará el WorldSave.", 0, ToAll
    ElseIf minutosSinWorldSave + 1 = MinutosWs Then  '1 minuto antes de que se haga, aviso.
        EnviarPaquete Paquetes.MensajeServer, "En un minuto se realizará el WorldSave.", 0, ToAll
        EnviarPaquete Paquetes.MensajeServer, "En un minuto se realizará el WorldSave.", 0, ToAll
    End If
    
End Sub

Private Sub revisarChatGlobal()

    Static minutos As Long

    minutos = minutos + 1

    ' Chat global. Se revisa cada 50 minutos
    If minutos > 10 Then
        If NumUsers < 100 Then
            'Si hay menos de 100 usuarios y el chat no esta activado. Lo activo
            If modChatGlobal.charlageneral = False Then
               Call modChatGlobal.activarChatGlobal
            End If
        ElseIf NumUsers > 110 Then
            'Si hay mas de  110 usuarios y el chat esta activado. Lo desactivo.
            If modChatGlobal.charlageneral = True Then
                Call modChatGlobal.desactivarChatGlobal
            End If
        End If
    End If

End Sub
Private Sub revisarPosicionCriaturas()

    Static MinutosLatsClean As Long
    
    If MinutosLatsClean >= 45 Then
        MinutosLatsClean = 0
        Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
    Else
        MinutosLatsClean = MinutosLatsClean + 1
    End If

End Sub
'Este timer se ejecuta una vez por minuto.
Private Sub AutoSave_Timer()

    
    Call modClima.calcularClima             ' Calculamos el clima
    
    Call revisarPosicionCriaturas           ' Las criaturas vuelven a su posicion inciial. Una vez cada 45 minutos
    
    Call PurgarPenas                        ' Penas de carcel de los personajes.
    
    Call CheckIdleUser                      ' Inactividad de personajes. Una vez por minuto
    
    Call actualizarOnlinesDB                ' Actualización en la base de datos de usuarios online. Una vez por minuto
    
    Call modEventos.procesarTimeOutMinuto   ' Eventos. Una vez por minuto
    
    Call modEstadisticasTCP.ActualizaStats  ' Estadisticas de consumo de ancho de banda. Una vez por minuto
    
    Call modCapturarPantalla.eliminarCorruptas  ' Fotos solicitadas a los usuarios
    
    Call Anticheat_MemCheck.hook_pasar_Minuto   ' Sistema de chequeo de Edicion de Memoria
    
    Call revisarChatGlobal                      ' Chat Global
    
    Call mdClanes.revisarEstadoClanes           ' Elecciones / Infracciones
    
    
    Call modMySql.enviarPingBaseDeDatos         ' Por las dudas, hacemos una consulta para que no se caiga
    
    ' Cada 30 minutos.
    If Minute(Now) = 0 Or Minute(Now) = 30 Then
        ' Estadisticas de Online
        Call modEstadisticasTCP.GuardarEstadisticasFraccion(Date$)
        ' Limpiamos el mundo
        Call modMundo.LimpiarMundo
    End If
    
    ' Cada 5 minutos
    If Minute(Now) Mod 5 = 0 Then
        fraccionDelDia = fraccionDelDia + 1
        
        If ForzarDia = 0 Then ' Bug. esta al revez en cl elcinete,
            ' Es de dia!
            If fraccionDelDia > 96 Then
                fraccionDelDia = 1
                ForzarDia = 1
            End If
        Else
            ' No foramos el dia
            If fraccionDelDia > 67 Then
                fraccionDelDia = 36
                ForzarDia = 0
            End If
        End If
        
        Call enviarEstadoNoche
    End If
    
    Call calcularWorldSave
End Sub


Private Sub Command1_Click()

    If Len(BroadMsg.Text) > 0 Then
        EnviarPaquete Paquetes.mBox2, BroadMsg.Text, 0, ToAll
    End If

End Sub

Public Sub InitMain(ByVal f As Byte)
    Me.WinsockWeb(0).LocalPort = SERVIDOR_WEB_PUERTO
    Me.WinsockWeb(0).listen
    LastIndex = 0
    
    Dim i As Integer

    For i = 1 To 20
      Load WinsockWeb(i)
    Next i
    
    'NUEVO SOCKET
    'IniciarComponente
If f = 1 Then
    Call mnuSystray_Click
Else
    frmMain.Show
End If
End Sub

Private Sub Command2_Click()
If Len(BroadMsg.Text) > 0 Then
    EnviarPaquete Paquetes.MensajeServer, BroadMsg.Text, 0, ToAll
    Exit Sub
End If
End Sub

Private Sub Form_Load()
#If TDSFacil Then
    Me.Label2 = "TDS FÁCIL"
#Else
    Me.Label2 = "TDS"
#End If

Me.scrollfraccionDelDia.value = fraccionDelDia
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Not Visible Then
        Select Case x \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
End Sub

'CSEH: Nada
Private Sub QuitarIconoSystray()
On Error Resume Next
'Borramos el icono del systray
Dim i As Integer
Dim nid As NOTIFYICONDATA
nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")
i = Shell_NotifyIconA(NIM_DELETE, nid)
End Sub

'CSEH: Nada
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Dim loopC As Integer

For loopC = 1 To MaxUsers
    If UserList(loopC).flags.UserLogged Then
        Call CloseSocket(loopC)
    End If
Next

Call QuitarIconoSystray
Call LimpiaWsApi(frmMain.hWnd)

'Log
Call LogMain("Server cerrado.")

End

End Sub



Private Sub GameTimer_Timer()

Dim iUserIndex As Integer
Dim bEnviarStats As Boolean
Dim bEnviarAyS As Boolean
Dim bEnviarEnergias As Boolean
Dim afectaLluvia As Boolean
Dim tiempo As Long
Dim ahora As Long

Static UltimoLoop As Long

 '<<<<<< Procesa eventos de los usuarios >>>>>>
If UltimoLoop = 0 Then UltimoLoop = 30000

ahora = GetTickCount
tiempo = ahora - UltimoLoop

For iUserIndex = 1 To LastUser
   'Conexion activa?j

With UserList(iUserIndex)

' Los Timers son validos solo para personajes activos
If Not .ConnID = INVALID_SOCKET Then

      '¿User valido?
      If .flags.UserLogged Then

         bEnviarStats = False
         bEnviarAyS = False
         bEnviarEnergias = False

        If .flags.Muerto = 0 Then 'And UserList(iUserIndex).flags.Privilegios = 0 Then
        
                If .flags.Paralizado = 1 Then Call EfectoParalisisUser(UserList(iUserIndex), tiempo)
                
                If .flags.Meditando Then Call DoMeditar(UserList(iUserIndex), tiempo)
                
                If .flags.Envenenado = 1 Then Call EfectoVeneno(UserList(iUserIndex), bEnviarStats, tiempo)
                
                If .flags.Invisible = 1 And .flags.Oculto = 0 Then Call EfectoInvisibilidad(UserList(iUserIndex), tiempo)
                
                If .flags.Mimetizado = 1 Then Call EfectoMimetismo(UserList(iUserIndex), tiempo)
                
                If .flags.DuracionEfecto > 0 Then Call DuracionPociones(iUserIndex, tiempo)
                
                If .flags.Oculto = 1 Then Call DoPermanecerOculto(UserList(iUserIndex), tiempo)
                
                If .NroMacotas > 0 Then Call TiempoInvocacion(iUserIndex, tiempo)

                If .controlCheat.VecesAtack > Me.Sensibilidad Then
                    EnviarPaquete Paquetes.mensajeinfo, .Name & " posible speed para pegar/magia. Gravedad: " & .controlCheat.VecesAtack, 0, ToAdmins
                    .controlCheat.VecesAtack = 0
                Else
                    .controlCheat.VecesAtack = 0
                End If
                        
                If MapInfo(.pos.map).Frio = 1 Then Call EfectoFrio(iUserIndex, tiempo, bEnviarStats)
                
                If MapInfo(.pos.map).Calor = 1 Then
                    Call EfectoCalor(UserList(iUserIndex), tiempo, bEnviarStats)
                End If
                
                Call HambreYSed(iUserIndex, bEnviarAyS, tiempo)
                
                ' ¿Esta lloviendo?
                ' Si esta lloviendo...
                '   Si el personaje NO está a la intemperie, o sea se esta mojando.
                '       Si no está descansando... recupera tipo 4
                '       Si el personaje está descansnsado... recupera tipo 3
                '   Si  está a la intemperie
                '       Si no está desnudo... recupera 4
                '       Si está desnudo o tiene hambre... recupera tipo 5
                ' Si NO está lloviendo
                '   Si NO está descansando, no tiene hambre/sed y no está desnudo.. recupera 2
                '   Si está descansando... recupera 1
                '   Si está desnudo... recupera 5
                
                If Lloviendo Then
                    '   ¿Le pega el agua?
                    afectaLluvia = modPersonaje.estaALaIntemperie(UserList(iUserIndex))
                Else
                    afectaLluvia = False
                End If
                    
                
                If afectaLluvia Then
                
                    If .flags.Desnudo Then
                        ' Pierde energia
                        Call RecStamina(UserList(iUserIndex), 5, bEnviarEnergias)
                    ElseIf .flags.Hambre = 0 And .flags.Sed = 0 Then ' No tiene ni hambre ni sed
                        ' Gana vida
                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                        ' Gana energia
                        Call RecStamina(UserList(iUserIndex), 4, bEnviarEnergias)
                    End If
                    
                Else
                
                    ' ¿Esta desnudo?
                    If .flags.Desnudo = 1 Then
                        ' Pierde energia
                        Call RecStamina(UserList(iUserIndex), 5, bEnviarEnergias)
                    ElseIf .flags.Descansar = True And (.flags.Hambre = 0 And .flags.Sed = 0) Then
                        ' Recupera vida
                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                        ' Recupera energia
                        Call RecStamina(UserList(iUserIndex), 1, bEnviarEnergias)
                    ElseIf .flags.Hambre = 0 And .flags.Sed = 0 Then
                        ' Recupera energia
                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                        Call RecStamina(UserList(iUserIndex), 2, bEnviarEnergias)
                    End If
                    
                    ' Termina de descansar automaticamente
                    If .flags.Descansar Then
                        If .Stats.MaxHP = .Stats.minHP And .Stats.MaxSta = .Stats.MinSta Then
                            .flags.Descansar = False
                            EnviarPaquete Paquetes.MDescansar, "", iUserIndex
                        End If
                    End If
                    
                End If
            
                If bEnviarStats Then Call SendUserStatsBoxBasicas(iUserIndex)
                If Not bEnviarStats And (.flags.Trabajando Or bEnviarEnergias) Then EnviarPaquete Paquetes.EnviarST, Codify(.Stats.MinSta), iUserIndex, ToIndex
                If bEnviarAyS Then Call EnviarHambreYsed(iUserIndex)
        
        End If ' Cierra user muerto

       End If 'con idd
       
End If

End With

Next iUserIndex

UltimoLoop = GetTickCount

End Sub

Private Sub mnuCerrar_Click()
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
nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, s)
i = Shell_NotifyIconA(NIM_ADD, nid)
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False
End Sub

Private Sub enviarEstadoNoche()

Dim iUserIndex As Integer

For iUserIndex = 1 To LastUser
    ' Los Timers son validos solo para personajes activos
    If Not UserList(iUserIndex).ConnID = INVALID_SOCKET Then
    
          '¿User valido?
          If UserList(iUserIndex).flags.UserLogged Then
            Call enviarNoche(UserList(iUserIndex))
          End If
          
    End If
Next

LogDesarrollo ("Se actualiza la situación de la nohe. Fracción del día " & fraccionDelDia & ". Forzar dia" & ForzarDia)
End Sub

Private Sub scrollfraccionDelDia_Change()
    fraccionDelDia = Me.scrollfraccionDelDia.value
    ForzarDia = Me.chkForzarDia.value = 1
    
    Me.lblNoche.Caption = "Hora: " & obtener_hora_fraccion(fraccionDelDia) & "hs."
    
    enviarEstadoNoche
End Sub

Private Sub TIMER_AI_Timer()
    Call NPCs.procesarNpcs
End Sub

Private Sub Timer2_Timer()
Dim Msgs As String
Msgs = Conteo

If Conteo = 0 Then Msgs = "YA"
EnviarPaquete Paquetes.MensajeSpell, "Conteo > " & Msgs, 0, ToAll
If Conteo <> 0 Then
Conteo = val(Conteo) - 1
End If
If Msgs = "YA" Then Timer2.Enabled = False
End Sub


Private Sub timerTrabajo_Timer()

Dim UserIndex As Integer

TrabajadoresGroup.itIniciar

Do While (TrabajadoresGroup.ithasNext)

    UserIndex = TrabajadoresGroup.itnext

    If UserList(UserIndex).Trabajo.tipo > 0 Then
    
        ' Desactivamos la inactividad
        UserList(UserIndex).Counters.IdleCount = 0
        
        ' ¿Que quiere hacer?
        Select Case UserList(UserIndex).Trabajo.tipo
            Case eTrabajos.Pesca
                Call DoPescar(UserList(UserIndex))
            Case eTrabajos.Tala
                Call DoTalar(UserList(UserIndex))
            Case eTrabajos.Mineria
                Call DoMineria(UserList(UserIndex))
            Case eTrabajos.Fundicion
                Call DoFundirMineral(UserList(UserIndex))
            Case eTrabajos.Herreria
                Call DoHerreria(UserList(UserIndex))
            Case eTrabajos.Carpinteria
                Call DoCarpinteria(UserList(UserIndex))
        End Select
    Else
        Call DejarDeTrabajar(UserList(UserIndex))
    End If
    
Loop

End Sub


Private Sub WinsockWeb_Close(index As Integer)
    WinsockWeb(index).Close
    Call LogDesarrollo("Se cierra conexion de web index " & index)
End Sub

Private Sub WinsockWeb_ConnectionRequest(index As Integer, ByVal requestID As Long)
    
    If index = 0 Then
        
        Me.cantidadConexionesWeb = Int(Me.cantidadConexionesWeb) + 1
        
        ' Preparar socket para el cliente
        LastIndex = LastIndex + 1
        
        If LastIndex > 20 Then
            LastIndex = 1
        End If
       
        Call LogDesarrollo("Se abre conexion de web index " & LastIndex & " " & WinsockWeb(index).RemoteHostIP)
        
        WinsockWeb(LastIndex).LocalPort = SERVIDOR_WEB_PUERTO
        WinsockWeb(LastIndex).accept requestID
    
    End If

End Sub

Private Sub WinsockWeb_DataArrival(index As Integer, ByVal bytesTotal As Long)

Dim datos As String
Dim numAccion As Byte
Dim argumento As String

    ' Obtengo la informacion que me llego en el socket
    Me.WinsockWeb(index).GetData datos
    ' Obtengo el numero de accion y el argumento correspondiente
    numAccion = Asc(Left(datos, 1))
    argumento = mid(datos, 2, Len(datos) - 1)
    
    ' proceso
    Call procesarHandleWeb(index, numAccion, argumento)
    ' guardo en el log
    Call LogDesarrollo("Se recibe datos conexion de web index " & index & datos)
    ' cierro el socket
    Call WinsockWeb_Close(index)
    
End Sub
Private Sub procesarHandleWeb(index As Integer, numAccion As Byte, argumento As String)
    Dim TempInt As Integer
    Dim TempVar As Variant
    Dim tempbyte As Byte
    Dim tempstr As String
    Dim tempLong As Long
    
    
    Select Case numAccion
        
        Case 1  ' Desloguear personaje
            TempInt = NameIndex(argumento)
             
            If TempInt <= 0 Then
               ' Me.WinsockWeb(Index).Senddata "1"
            Else
                If Not CloseSocket(TempInt) Then LogError ("procesar handle web 1")
               ' Me.WinsockWeb(Index).Senddata "0"
            End If
        Case 2 ' Encarcelar
        
            ' Los argumentos estan conpuestos por
            ' Nick a encarcelar-Tiempo en la carcel-Gm-Razon
            TempVar = Split(argumento, "-", 4)
            
            'Obtengo el userindex del usuario
            TempInt = NameIndex(Trim(TempVar(0)))
            ' Obtengo el tiempo a encarcelar
            tempbyte = Int(val(TempVar(1)))
            ' Obtengo la cadena explicatoria
            tempstr = Trim(TempVar(3))
            
            If TempInt > 0 Then
            
                Encarcelar UserList(TempInt), tempbyte
                
                If LenB(UserList(TempInt).flags.Penasas) = 0 Then
                    UserList(TempInt).flags.Penasas = tempstr
                Else
                    UserList(TempInt).flags.Penasas = UserList(TempInt).flags.Penasas & vbCrLf & tempstr
                End If
                
                LogAccionesWeb ("Carcel para " & UserList(TempInt).Name & "por " & TempVar(3))
            Else
                LogAccionesWeb ("No se pudo encarcelar a un usuario por que esta offline. " & TempVar(0))
            End If
            
            'Guardamos para obtener consistencia y que se actualice en la web también
            'para no ser penado dos veces por lo mismo
           Call SaveUser(TempInt, 1)
     Case 3 ' Banear
        ' Los argumentos estan conpuestos por
        ' Nick a banear-Tiempo de ban-Gm-Razon
        TempVar = Split(argumento, "-", 4)
        ' Obtengo la cadena explicatoria
            
        Call BanearUsuario(Trim(TempVar(2)), Trim(TempVar(0)), Trim(TempVar(3)), Int(val(TempVar(1))), False)
    Case 4 ' Mensaje por consola a todos los usuarios
        
        tempstr = argumento
        'Envo el mensaje a todos
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(14) & tempstr, 0, ToAll
        
    Case 5 ' Otorga oro a un usuario
    
        tempstr = argumento
        TempVar = Split(argumento, "-")
        ' ID del personaje - Cantidad de oro a dar - Motivo - ID DE PAGO
    
        TempInt = IDIndex(CLng(TempVar(0)))
        tempLong = CLng(TempVar(1))
        
        '¿Esta online?
        If TempInt > 0 Then
        
            'Le doy el oro
            UserList(TempInt).Stats.Banco = UserList(TempInt).Stats.Banco + tempLong
            
            'Le aviso
            EnviarPaquete Paquetes.MensajeTalk, TempVar(2) & ". Se te han sumado " & FormatNumber(val(tempLong), 0, vbTrue, vbFalse, vbTrue) & " monedas de oro en tu cuenta bancaria.", TempInt, ToIndex
            
            'Agrego el log
            sql = "INSERT DELAYED " & DB_NAME_PRINCIPAL & ".web_pagos(IDPAGO, IDPERSONAJE, MONTO) VALUES(" & CLng(TempVar(3)) & "," & UserList(TempInt).id & "," & tempLong & ")"
            
            conn.Execute sql, , adExecuteNoRecords
        End If
    
    End Select
End Sub

Private Sub WinsockWeb_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call LogDesarrollo("Error en la conexion  " & index & "desc " & Description)
End Sub
