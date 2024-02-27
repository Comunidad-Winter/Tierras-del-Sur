Attribute VB_Name = "TCP"



Option Explicit
Public fogataaaa As Integer

Rem////////////////////////////TCP//////////////////////////
Rem/                    Modulo desarrollado por Wizard
Rem/        Agiliza, Prescribe y Ahorra el envio y la
Rem/        recepcion de paquetes.
Rem/////////////////////////////////////////////////////////

Public Enum Paquetes
        Comandos = 1
        Hablar = 2
        MirarNorte = 3
        DeclararAlly = 4
        GuildInfo = 5
        ComandosSemi = 6
        CHerrero = 7
        ComandosDios = 8
        SkillSetDomar = 9
        Susurrar = 10
        MSurM = 11
        PeaceProp = 12
        MWest = 13
        MirarSur = 14
        FEST = 15
        EnviarGuildComen = 16
        Lachiteo = 17
        ClickSkill = 18
        EcharGuild = 19
        RetoAccpt = 20
        EFotoDenuncia = 21
        RetoCncl = 22
        CallForFama = 23
        CallForSkill = 24
        Encarcelame = 25
        RechazarComUsu = 26
        FinReto = 27
        Vender = 28
        Usar = 29
        Depositar = 30
        CallForAtributos = 31
        ComandosAdmin = 32
        MSouth = 33
        preConnect = 34
        MemberInfo = 35
        DeclararWar = 36
        MEast = 37
        ChangeItemsSlot = 38
        OfrecerComUsu = 39
        MOesteM = 40
        ComandosConse = 41
        GuildDetail = 42
        Expulsarparty = 43
        ' RevivirAutomaticamente = 44
        SosDone = 45
        MNorth = 46
        DejadeLaburar = 47
        CCarpintero = 48
        SeguroClan = 49
        MEsteM = 50
        PostearForo = 51
        ComUsuOk = 52
        CrearReto = 53
        Gritar = 54
        Moverhechi = 55
        CrearRetoD = 56
        BancoOk = 57
        ChangeItemsSlotboveda = 58
        UNLAG = 59
        SOSViewList = 60
        Pong2 = 61
        LaChiteo2 = 62
        Agarrar = 63
        comprar = 64
        MNorteM = 65
        ClickAccion = 66
        entrenador = 67
        DIClick = 68
        Drag = 69
        Retirar = 70
        Aprobaringresoparty = 71
        EnvPeaceOffer = 72
        ComOk = 73
        Salirparty = 74
        Seguro = 75
        SkillMod = 76
        AceptarGuild = 77
        PEACEDET = 78
        FaccionMsg = 79
        Tirar = 80
        ActualizarGNews = 81
        SOSAddItem = 82
        FinComUsu = 83
        SkillSetRobar = 84
        ClickIzquierdo = 85
        PeaceAccpt = 86
        GuildCode = 87
        Equipar = 88
        MTrabajar = 89
        LanzarHechizo = 90
        Crearparty = 91
        InfoHechizo = 92
        ArrojarDados = 93
        RechazarGuild = 94
        iParty = 95
        ccParty = 96
        MCombate = 97
        ConnectPj = 98
        MirarOeste = 99
        GuildSol = 100
        MirarEste = 101
        Spawn = 102
        GuildDSend = 103
        CreatePj = 104
        URLChange = 105
        SkillSetOcultar = 106
        ExitOk = 107
        Pegar = 108
    
        obtClanMiembros = 109
        obtClanSolicitudes = 110
        obtClanNews = 111
        
        infoTransferencia = 112
        respuesta = 113
End Enum

Public Enum sPaquetes
    pMensajeSimple = 1
    pMensajeCompuesto = 2
    PrenderFogata = 3
    MostrarCartel = 4
    MensajeForo = 5
    MensajeForo2 = 6
    WavSnd = 7
    pNpcInventory = 8
    TransOK = 9
    EnPausa = 10
    pIniciarComercioNpc = 11
    Loguea = 12 'Modificar
    VeUser = 13
    VeObjeto = 14
    VeNpc = 15
    DescNpc = 16
    DescNpc2 = 17
    ModCeguera = 18
    ModEstupidez = 19
    BloquearTile = 20
    pEnviarSpawnList = 21
    EquiparItem = 22
    DesequiparItem = 23
    BorrarObj = 24
    CrearObjeto = 25
    ApuntarProyectil = 26
    ApuntarTrb = 27
    EnviarArmasConstruibles = 28
    EnviarObjConstruibles = 29
    EnviarArmadurasConstruibles = 30
    ShowCarp = 31
    InitComUsu = 32
    ComUsuInv = 33
    FinComUsuOk = 34
    InitBanco = 35
    EnviarBancoObj = 36
    BancoOk = 37
    PeaceSolRequest = 38
    EnviarPeaceProp = 39
    PeticionClan = 40
    EnviarCharInfo = 41
    EnviarLeaderInfo = 42
    EnviarGuildsList = 43
    EnviarGuildDetails = 44
    HechizoFX = 45
    MensajeTalk = 46
    MensajeSpell = 47
    MensajeFight = 48
    MensajeInfo = 49
    CambiarHechizo = 50
    pCrearNPC = 51
    ChangeNpc = 52
    BorrarNpc = 53
    MoveChar = 54
    EnviarNpclst = 55
    COMBRechEsc = 56
    COMBNpcHIT = 57
    COMBMuereUser = 58
    COMBNpcFalla = 59
    COMBUserFalla = 60
    COMBEnemEscu = 61
    SangraUser = 62
    COMBUserImpcNpc = 63
    COMBEnemFalla = 64
    COMBEnemHitUs = 65
    COMBUserHITUser = 66
    Navega = 67
    AuraFx = 68
    Meditando = 69
    NoParalizado = 70
    Paralizado2 = 71
    NoParalizado2 = 72
    invisible = 73
    Visible = 74
    pChangeUserChar = 75
    LevelUP = 76
    SendSkills = 77
    SendFama = 78
    SendAtributos = 79
    MiniEst = 80
    BorrarUser = 81
    crearChar = 82
    EnviarPos = 83
    EnviarStat = 85
    EnviarF = 86
    EnviarA = 87
    EnviarOro = 88
    EnviarHP = 89
    EnviarMP = 90
    EnviarST = 91
    EnviarEXP = 92
    EnviarSYM = 93
    EnviarSYH = 94
    EnviarFA = 95
    EnviarHYS = 96
    QDL = 97
    MDescansar = 98
    ChangeMap = 99
    ChangeMusic = 100
    QTDL = 101
    IndiceChar = 102
    mBox = 103
    Lluvia = 104
    SOSAddItem = 105
    SOSViewList = 106
    MensajeServer = 107
    MensajeGMSG = 108
    UserTalk = 109
    UserShout = 110
    UserWhisper = 111
    TurnToNorth = 112
    TurnToSouth = 113
    TurnToWest = 114
    TurnToEast = 115
    FinComOk = 116
    FinBanOk = 117
    SndDados = 118
    ShowHerreriaForm = 119
    'EnviarUI = 120             no se usa
    EnviarGuildNews = 121
    InvRefresh = 122
    InitGuildFundation = 123
    MensajeClan1 = 124
    MensajeClan2 = 125
    SaidMagicWords = 126
    MoveNpc = 127
    pEnviarNpcInvBySlot = 128
    mTransError = 129
    CrearObjetoInicio = 130 'Para resolver el error de colgar en los mapas donde hay muchos items. Tendria que solucionar con este nuevo tcp pero por las dudas... Marche
    pMensajeSimple2 = 131
    noche = 132
    SegOFF = 133
    SegOn = 134
    Nieva = 135
    DejaDeTrabajar = 136
    TXA = 137 'El yind, efectos. Agregado por marche
    mBox2 = 138
    FXH = 139
    FundoParty = 140
    PNI = 141 'Pide ingreso a la pary
    Integranteparty = 142
    OnParty = 143
    Mest = 144
    AnimGolpe = 145 'El Yind - Animaciones de golpe y esucdo
    AnimEscu = 146
    CFXH = 147
    MensajeGuild = 148
    ClickObjeto = 149
    LISTUSU = 150 ' lista de usuarios para los gms
    Traba = 151 'Lista de usuarios trabajando
    UserTalkDead = 152 'Hablarle alos muertos
    TiempoRetos = 153 ' 3 2 1 ya a retar
    Pang = 154 ' Lo contrario al ping
    TalkQuest = 155
    pChangeUserCharCasco = 156 'Cambia solo el casco
    pChangeUserCharEscudo = 157 'Cambia solo el escudo
    pChangeUserCharArmadura = 158 ' Cambia solo la armadura
    pChangeUserCharArma = 159 'Cambia solo el arma
    EnCentinela = 160
    TXAII = 161
    EnviarStatsBasicas = 162
    MensajeArmadas = 163
    MensajeCaos = 164
    EmpiezaTrabajo = 165
    MensajeGlobal = 166
    PartyAcomodarS = 167
    PPI = 168
    PPE = 169
    Sefuedeparty = 170
    MensajeBoveda = 171
    IniciarAutoUpdater = 172
    EstaEnvenenado = 173
    Actualizarestado = 174
    MoverMuerto = 175
    ocultar = 176
    Desocultar = 177
    pNpcActualizarPrecios = 178
    ActualizaNick = 179
    ActualizaCantidadItem = 180
    ActualizarAreaUser = 181
    ActualizarAreanpc = 182
    CambiarHeadingNpc = 183
    BorrarArea = 184
    MoverWest = 185
    MoverEast = 186
    MoverNorth = 187
    MoverSouth = 188
    accion = 189
    Pong = 190
    SonidoTomarPociones = 191
    infoLogin = 192
    
    EnviarLeaderInfoSolicitudes = 193
    EnviarLeaderInfoMiembros = 194
    EnviarLeaderInfoNovedades = 195
    
    InfoAdminEventos = 196
    MensajeAdminEventos = 197
    InfoEventoAdminEventos = 198
    
    transferenciaIniciar = 199      ' Se solicita que comience la transferencia
    transferenciaOK = 200           ' Paso de la transferencia se hizo ok
    infoClan = 201                  ' Le envia informacion del clan del usuario
    checkMem = 202                  ' Chequea que no se este editando la memoria
    angulonpc = 204
End Enum



Rem ////////////////¡¡COMANDOS!!////////////////////
Public Enum Simple 'Comandos sin texto
    online = 1 'Todos los online
    Comerciar = 2
    boveda = 3
    Meditar = 4
    ENLISTAR = 5
    informacion = 6
    DESINVOCAR = 7
    Resucitar = 8
    Curar = 9
    Descansar = 10
    entrenar = 11
    ACOMPAÑAR = 12
    QUIETO = 13
    Balance = 14
    SALIRCLAN = 15
    FUNDARCLAN = 16
    ONLINECLAN = 17
    DONDECLAN = 18
    Salir = 19
    Ayuda = 20
    EST = 21
    RECOMPENSA = 22
    MOTD = 23
    GM = 24
    MOVER = 25
    Retirartodo = 102
    DepositarTodo = 103
    Abandonar = 104
    '105
    PartyPorcentaje = 106 'Party. Le pide al server lso porctenajes y los pjs actuales de la party
    Penas = 107
    pmsg = 108
    DisolverClan = 109
    '110 en dobles
    '111 en dobles
    '112 en dobles
    '113 en dobles
    '114 en dobles
    '115 en dobles
    minutoEnTorneo = 115
    '116
    tiempo = 117
End Enum

Public Enum Complejo
    'ONLINEP = 26 'Busca online 1 personaje en particular
    Retirar = 27
    Depositar = 28
    PASARORO = 29
    APOSTAR = 30
    VOTO = 31
    PASSWD = 32
    UDESC = 33
    Bug = 34
    CMSG = 35
    
    Denunciar = 95
    Centinela = 26
    Cheque = 97
    Activar = 98
    PING = 99
    RetarS = 100
    AcomodarPorcentajesDeParty = 101
    '102 En simples
    '103 En simples
    '104 en simples
    Fianza = 105
    '106 en simples
    '107 en simples
    '108 en simples
    '109 en simples
    reanudarclan = 110
    Participar = 111
    Aceptar = 112
    eventoInfo = 113
    decirEnTorneo = 114
    '115 en simples
    rechazar = 116
    '117
End Enum

Public Enum Conse1
    Hora = 36
    TELEPLOC = 37
    SHOW_SOS = 38
    CINVISIBLE = 39
    PANELGM = 40
    TRABAJANDO = 41
End Enum

Public Enum Conse2
    crem = 42
    TELEP = 43
    NENE = 44
    donde = 45
    IRA = 46
    
    CARCEL = 54
    
    Penas = 106
    
    ONLINEMAP = 53
    
    RMSG = 107
End Enum

Public Enum SemiDios2
    info = 47
    INV = 48
    BOV = 49
    CSKILLS = 50
    Revivir = 51
    IP2NICK = 52
    
    
    PERDON = 55
    CECHAR = 56
    ban = 57
    UNBAN = 58
    CSUM = 59
    RESETINV = 60
    NICK2IP = 62
    GMSG = 95
    ejecutar = 96
    qtalk = 97
    AUCO = 98
    
    '107
    Amapa = 109
    ROSG = 113
    
    MaxLevelMap = 114
    MinLevelMap = 115
    
    Spawn = 118
    
    LimiteUserMap = 122
    CName = 123
    '124
    '125
    CrearEvento = 126
    '127
    ObtenerInfoEvento = 128
    CancelarEvento = 129
    '130
    '131
     CT = 132
    
    publicarEvento = 133
    inscribirEvento = 134
End Enum

Public Enum SemiDios1
    ONLINEGM = 63
    cc = 64
    limpiar = 65
    SEGUIR = 66
    LagarCentinelas = 107
    NoseCaemap = 108
    Antinw = 111
    MapaFrio = 116
    HabilitarRobo = 117
    InfoMap = 119
    OnlyCiuda = 120
    OnlyCrimi = 121
    OnlyCaos = 124
    OnlyArmada = 125
    '126
    ObtenerEventos = 127
    '128
    '129
    dt = 130
    BLOQ = 131
    '132
    '133
    '134
End Enum

Public Enum Dios1
    MASSDEST = 67
    'BANIPLIST = 68
    'BANIPRELOAD = 69
    PASSDAY = 70
    '71
    dest = 72
    '73
    MATA = 74
    MASSKILL = 75
    MOTDCAMBIA = 76
    'ACC = 77
    NAVE = 78
    'APAGAR = 79
    'backup = 80
    GRABAR = 81
    BORRAR_SOS = 82
    SHOW_INT = 83
    Lluvia = 84
End Enum

Public Enum Dios2
    ACC = 77
   '85
    CTRIGGER = 86
    'BANIP = 87
    'UNBANIP = 88
    LASTIP = 89
    RACC = 90
    'CONDEN = 91
    'RAJAR = 92
    RAJARCLAN = 93
    'CMod = 94
    'CI = 95
    ZONEST = 96
    RETEST = 97
    CHATEST = 98
    'TCPEST = 99
    ECHARTODOSPJS = 100
    NickMac = 101
    BanMac = 102
    UnbanMac = 103
    'ReloadServer = 104
    'CName = 105
    Rettings = 106
    Habilitar = 107
    AceptarConsejo = 108
    ExpulsarConsejo = 109
    ModoRol = 110
    'CTE = 112
    CheqCli = 113
    VerPongs = 114
    CNameClan = 115
    EcharClan = 116
    
    sCapturarPantalla = 117
    consultarMem = 118
End Enum

Public Enum cmdAdmin
    GRABAR = 1
    APAGAR = 2
    backup = 3
    CONDEN = 4
    RAJAR = 5
    CMod = 6
    CI = 7
    TCPEST = 8
    ReloadServer = 9
    BANIPRELOAD = 10
    
End Enum

'////////////////Variables/////////////////////////////
Private TempByte2 As Byte
Private TempByte As Byte
Private TempInt As Integer
Private tempLong As Long
Private Tempvar As Variant
Private TempStr As String

Public Type tPaq
    TC As Long
    Rdata As String
End Type
Public TempPaq() As tPaq

Public StartTC As Long
Public PS As Long

Public recibiPaquete As Boolean

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long


'//////////////////////////////////////////////////////

Public Sub CrearAccion(Rdata As String)
        PS = PS + 1
        ReDim Preserve TempPaq(PS)
        TempPaq(PS).TC = GetTickCount - StartTC
        TempPaq(PS).Rdata = Rdata
End Sub
    
Sub ProcesarComando(ByVal comando As String)
On Error GoTo haybug
    Dim Main As String
    'Borramos la barrita
    comando = right$(comando, Len(comando) - 1)
    
    'Tomamos el Main del comando
    'Main = ReadField(1, Comando, Asc(" "))
    'Comandos siples de usuario.
    Select Case UCase$(comando)
        Case "ONLINE"
            Call sSendData(Paquetes.Comandos, Simple.online)
        Exit Sub
        Case "COMERCIAR"
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.Comerciar)
        Exit Sub
        Case "BOVEDA"
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.boveda)
        Exit Sub
        Case "MEDITAR"
            If UserMaxMAN = 0 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "No puedes meditar.", 65, 190, 156, 0): Exit Sub
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.Meditar)
        Exit Sub
        Case "ENLISTAR"
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.ENLISTAR)
        Exit Sub
        Case "INFORMACION"
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.informacion)
        Exit Sub
        Case "DESINVOCAR"
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.DESINVOCAR)
        Exit Sub
        Case "RESUCITAR"
            If UserStats(SlotStats).UserEstado = 0 Then Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.Resucitar)
        Exit Sub
        Case "CURAR"
        If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
        If UserStats(SlotStats).UserMinHP = UserMaxHP Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡Ya estás curado!", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.Curar)
        Exit Sub
        Case "DESCANSAR"
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            If UserStats(SlotStats).UserMinSTA = UserMaxSTA Then AddtoRichTextBox frmConsola.ConsolaFlotante, "No estás cansado.", 65, 190, 156, False, False, False: Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.Descansar)
        Exit Sub
        Case "ENTRENAR"
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.entrenar)
        Exit Sub
        Case "ACOMPAÑAR"
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.ACOMPAÑAR)
        Exit Sub
        Case "QUIETO"
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.QUIETO)
        Exit Sub
        Case "BALANCE"
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.Balance)
        Exit Sub
        Case "SALIRCLAN"
            If Not isTengoClan() Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡No perteneces a ningún clan!", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.SALIRCLAN)
        Exit Sub
        Case "FUNDARCLAN"
            If isTengoClan() Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡No puedes fundar un clan si ya eres de uno!", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.FUNDARCLAN)
        Exit Sub
        Case "ONLINECLAN"
            If Not isTengoClan() Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡No perteneces a ningún clan!", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.ONLINECLAN)
        Exit Sub
        Case "DONDECLAN"
            If Not isTengoClan() Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡No perteneces a ningún clan!", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.DONDECLAN)
        Exit Sub
        Case "SALIR"
            If UserStats(SlotStats).UserParalizado Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "No puedes salir estando paralizado.", 60, 190, 156, 0, 0): Exit Sub
            If UserStats(SlotStats).UserCentinela Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "No puedes salir si estas con un centinela.", 60, 190, 156, 0, 0): Exit Sub
            If Istrabajando Then modMiPersonaje.DejarDeTrabajar
            Call sSendData(Paquetes.Comandos, Simple.Salir)
        Exit Sub
        Case "AYUDA"
            Call sSendData(Paquetes.Comandos, Simple.Ayuda)
        Exit Sub
        Case "EST"
            Call sSendData(Paquetes.Comandos, Simple.EST)
        Exit Sub
        Case "RECOMPENSA"
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.RECOMPENSA)
        Exit Sub
        Case "MOTD"
            Call sSendData(Paquetes.Comandos, Simple.MOTD)
        Exit Sub
        Case "MOVER"
          '  If UserStats(SlotStats).userestado = 1 Then Call AddtoRichTextBox(frmMain.RecTxt, "Estás muerto.", 65, 190, 156, 0): Exit Sub
          '  Call sSendData(Paquetes.Comandos, Simple.MOVER)
        Exit Sub
        Case "ACTIVAR"
            Call sSendData(Paquetes.Comandos, Complejo.Activar)
        Exit Sub
        Case "PING"
            Call sSendData(1, Complejo.PING)
            PingPerformanceTimer.Time
            frmMain.PING = "Cargando"
            'PingTime = GetTickCount
        Exit Sub
        Case "RETIRARTODO"
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.Retirartodo)
        Exit Sub
        Case "DEPOSITARTODO"
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            If UserGLD = 0 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "No tienes oro para depositar.", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.DepositarTodo)
        Exit Sub
        Case "ABANDONAR"
            Call sSendData(Paquetes.Comandos, Simple.Abandonar)
            Exit Sub
        Case "PENAS"
            Call sSendData(Paquetes.Comandos, Simple.Penas)
            Exit Sub
        Case "TIEMPO"
            Call sSendData(Paquetes.Comandos, Simple.tiempo)
            Exit Sub
        Case "SEG"
            EnviarPaquete Paquetes.Seguro
            If UserSeguro Then
            UserSeguro = False
            frmMain.IconoSeg = "X"
            Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "SEGURO DESACTIVADO", 255, 0, 0, True, False, False)
            Else
            frmMain.IconoSeg = ""
            Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "SEGURO ACTIVADO", 0, 255, 0, True, False, False)
            UserSeguro = True
            End If
            Exit Sub
        Case "DRAG"
            EnviarPaquete Paquetes.Drag
            If frmMain.IconoDyd = "" Then
            frmMain.IconoDyd = "X"
            Else
            frmMain.IconoDyd = ""
            End If
            Exit Sub
        Case "RETIRAR"
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Complejo.Retirar)
            Exit Sub
        Case "DISOLVERCLAN"
            If Not isTengoClan() Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡No puedes disolver un clan al cual no perteneces!", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.DisolverClan)
        Exit Sub
        Case "MINUTO"
            Call sSendData(Paquetes.Comandos, Simple.minutoEnTorneo)
        Exit Sub
    End Select
    If UserPrivilegios > 0 Then
        Select Case UCase$(comando)
            Case "HORA"
                Call sSendData(Paquetes.ComandosConse, Conse1.Hora)
            Exit Sub
            Case "TELEPLOC"
                Call sSendData(Paquetes.ComandosConse, Conse1.TELEPLOC)
            Exit Sub
            Case "INVISIBLE"
                Call sSendData(Paquetes.ComandosConse, Conse1.CINVISIBLE)
            Exit Sub
      '      Case "SHOW SOS"
       '         Call sSendData(Paquetes.ComandosConse, Conse1.SHOW_SOS)
        '    Exit Sub
            Case "PANELGM"
                frmPanelGm.Show , frmMain
            Exit Sub
            Case "TRABAJANDO"
                Call sSendData(Paquetes.ComandosConse, Conse1.TRABAJANDO)
            Exit Sub
        End Select
    End If
    If UserPrivilegios > 1 Then
        Select Case UCase$(comando)
            Case "SEGUIR"
                Call sSendData(Paquetes.ComandosSemi, SemiDios1.SEGUIR)
            Exit Sub
            Case "CC"
                Call sSendData(Paquetes.ComandosSemi, SemiDios1.cc)
            Exit Sub
            Case "LIMPIAR"
                Call sSendData(Paquetes.ComandosSemi, SemiDios1.limpiar)
            Exit Sub
            Case "ONLINEGM"
                Call sSendData(Paquetes.ComandosSemi, SemiDios1.ONLINEGM)
            Exit Sub
            Case "CENTINELAS"
                 Call sSendData(Paquetes.ComandosSemi, SemiDios1.LagarCentinelas)
            Exit Sub
            Case "NOCAEMAP"
                Call sSendData(Paquetes.ComandosSemi, SemiDios1.NoseCaemap, "")
            Exit Sub
            Case "ANTINW"
                Call sSendData(Paquetes.ComandosSemi, SemiDios1.Antinw, "")
            Exit Sub
            Case "FRIO"
                Call sSendData(Paquetes.ComandosSemi, SemiDios1.MapaFrio, "")
            Exit Sub
            Case "HABILITARROBO"
                Call sSendData(Paquetes.ComandosSemi, SemiDios1.HabilitarRobo, "")
            Exit Sub
            Case "INFOMAP"
                Call sSendData(Paquetes.ComandosSemi, SemiDios1.InfoMap, "")
            Exit Sub
            Case "ONLYCIUDA"
                Call sSendData(Paquetes.ComandosSemi, SemiDios1.OnlyCiuda, "")
            Exit Sub
            Case "ONLYCRIMI"
                Call sSendData(Paquetes.ComandosSemi, SemiDios1.OnlyCrimi, "")
            Exit Sub
            Case "ONLYCAOS"
                Call sSendData(Paquetes.ComandosSemi, SemiDios1.OnlyCaos, "")
            Exit Sub
            Case "ONLYARMADA"
                Call sSendData(Paquetes.ComandosSemi, SemiDios1.OnlyArmada, "")
            Exit Sub
            Case "CREAREVENTO"
                frmAdminEventos.Show
            Exit Sub
            Case "BLOQ"
                Call sSendData(Paquetes.ComandosSemi, SemiDios1.BLOQ)
            Exit Sub
            Case "DT"
                Call sSendData(Paquetes.ComandosSemi, SemiDios1.dt)
            Exit Sub
            End Select
    End If
    If UserPrivilegios > 2 Then
        Select Case UCase$(comando)
            Case "MASSDEST" 'REVISAR!
                Call sSendData(Paquetes.ComandosDios, Dios1.MASSDEST)
            Exit Sub
           ' Case "BANIPLIST"
            '    Call sSendData(Paquetes.ComandosDios, Dios1.BANIPLIST)
            'Exit Sub
            Case "DEST"
                Call sSendData(Paquetes.ComandosDios, Dios1.dest)
            Exit Sub
            Case "MATA"
                Call sSendData(Paquetes.ComandosDios, Dios1.MATA)
            Exit Sub
            Case "MASSKILL"
                Call sSendData(Paquetes.ComandosDios, Dios1.MASSKILL)
            Exit Sub
            Case "BORRAR SOS"
                Call sSendData(Paquetes.ComandosDios, Dios1.BORRAR_SOS)
            Exit Sub
            Case "LLUVIA"
                Call sSendData(Paquetes.ComandosDios, Dios1.Lluvia)
            Exit Sub
            Case "ZONEST"
                Call sSendData(Paquetes.ComandosDios, Dios2.ZONEST, "")
            Exit Sub
            Case "RETEST"
                Call sSendData(Paquetes.ComandosDios, Dios2.RETEST, "")
            Exit Sub
            Case "CHATEST"
                Call sSendData(Paquetes.ComandosDios, Dios2.CHATEST, "")
            Exit Sub
            Case "ECHARTODOSPJS"
                Call sSendData(Paquetes.ComandosDios, Dios2.ECHARTODOSPJS, "")
            Exit Sub
            Case "HABILITAR"
                Call sSendData(Paquetes.ComandosDios, Dios2.Habilitar)
            Exit Sub
            Case "ESTRINGS"
                Call sSendData(Paquetes.ComandosDios, Dios2.Rettings)
            Exit Sub
            Case "CHEQ"
                Call sSendData(Paquetes.ComandosDios, Dios2.CheqCli, "")
            Exit Sub
            Case "PONGS"
                Call sSendData(Paquetes.ComandosDios, Dios2.VerPongs, "")
            Exit Sub
        End Select
    End If
    
    'Adminsitradores
    If UserPrivilegios > 3 Then
        Select Case UCase$(comando)
            Case "APAGAR"
                Call sSendData(Paquetes.ComandosAdmin, cmdAdmin.APAGAR)
            Exit Sub
            Case "DOBACKUP"
                Call sSendData(Paquetes.ComandosAdmin, cmdAdmin.backup)
            Exit Sub
            Case "GRABAR"
                Call sSendData(Paquetes.ComandosAdmin, cmdAdmin.GRABAR)
            Exit Sub
            Case "RECARGARSERVER"
                Call sSendData(Paquetes.ComandosAdmin, cmdAdmin.ReloadServer, "")
            Exit Sub
            Case "TCPEST"
                Call sSendData(Paquetes.ComandosAdmin, cmdAdmin.TCPEST, "")
            Exit Sub
            Case "BANIPRELOAD"
                Call sSendData(Paquetes.ComandosAdmin, cmdAdmin.BANIPRELOAD)
            Exit Sub
        End Select
    End If
'***************************************************************************************
'               COMANDOS QUE TIENEN ARGUMENTOS.
'***************************************************************************************
    Main = ReadField(1, comando, Asc(" "))
    'El comando deberia tener argumentos sino fue encontrado ya.!
    Select Case UCase$(Main)
        Case "RETIRAR"
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            If LenB(right$(comando, Len(comando) - (Len(Main) + 1))) = 0 Then Exit Sub
            If val(right$(comando, Len(comando) - (Len(Main) + 1))) < 0 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "No puedes retirar cantidades menores a 0.", 255, 255, 255, False, False): Exit Sub
            Call sSendData(Paquetes.Comandos, Complejo.Retirar, right$(comando, Len(comando) - (Len(Main) + 1)))
        Exit Sub
        Case "DEPOSITAR"
            If LenB(comando) = LenB(Main) Then AddtoRichTextBox frmConsola.ConsolaFlotante, "Debes escribir: /DEPOSITAR 'CANTIDAD'.", 65, 190, 156: Exit Sub
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            If LenB(right$(comando, Len(comando) - (Len(Main) + 1))) = 0 Then Exit Sub
            If val(right$(comando, Len(comando) - (Len(Main) + 1))) < 0 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "No puedes depositar cantidades menores a 0.", 255, 255, 255, False, False): Exit Sub
            Call sSendData(Paquetes.Comandos, Complejo.Depositar, right$(comando, Len(comando) - (Len(Main) + 1)))
        Exit Sub
        Case "PASARORO"
            If LenB(comando) = LenB(Main) Then AddtoRichTextBox frmConsola.ConsolaFlotante, "Debes escribir: /PASARORO 'CANTIDAD'.", 65, 190, 156: Exit Sub
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            If LenB(right$(comando, Len(comando) - (Len(Main) + 1))) = 0 Then Exit Sub
            If val(right$(comando, Len(comando) - (Len(Main) + 1))) < 0 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "No puedes pasar cantidades menores a 0.", 255, 255, 255, False, False): Exit Sub
            Call sSendData(Paquetes.Comandos, Complejo.PASARORO, right$(comando, Len(comando) - (Len(Main) + 1)))
        Exit Sub
        Case "APOSTAR"
            If LenB(comando) = LenB(Main) Then AddtoRichTextBox frmConsola.ConsolaFlotante, "Debes escribir: /APOSTAR 'CANTIDAD' ", 65, 190, 156: Exit Sub
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            If LenB(right$(comando, Len(comando) - (Len(Main) + 1))) = 0 Then Exit Sub
            If val(right$(comando, Len(comando) - (Len(Main) + 1))) < 0 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "No puedes apostar cantidades menores a 0.", 255, 255, 255, False, False): Exit Sub
            Call sSendData(Paquetes.Comandos, Complejo.APOSTAR, right$(comando, Len(comando) - (Len(Main) + 1)))
        Exit Sub
        Case "VOTO"
            If LenB(comando) = LenB(Main) Then AddtoRichTextBox frmConsola.ConsolaFlotante, "Debes escribir: /VOTO 'NICK' ", 65, 190, 156: Exit Sub
            Call sSendData(Paquetes.Comandos, Complejo.VOTO, right$(comando, Len(comando) - (Len(Main) + 1)))
        Exit Sub
        Case "PASSWD"
            frmMain.Enabled = False
            cpassword.Show
        Exit Sub
        Case "DESC"
            If LenB(comando) = LenB(Main) Then AddtoRichTextBox frmConsola.ConsolaFlotante, "Debes escribir: /DESC 'NUEVA DESC'.", 65, 190, 156: Exit Sub
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "No puedes cambiar tu descripcion mientras estás muerto.", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Complejo.UDESC, right$(comando, Len(comando) - (Len(Main) + 1)))
        Exit Sub
        Case "BUG"
            If LenB(comando) = LenB(Main) Then AddtoRichTextBox frmConsola.ConsolaFlotante, "Debes escribir: /BUG 'BUG a INFORMAR'.", 65, 190, 156: Exit Sub
            Call sSendData(Paquetes.Comandos, Complejo.Bug, right$(comando, Len(comando) - (Len(Main) + 1)))
        Exit Sub
        Case "CMSG"
            If LenB(comando) = LenB(Main) Then Exit Sub
            If Not isTengoClan() Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡No perteneces a ningún clan!", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Complejo.CMSG, right$(comando, Len(comando) - (Len(Main) + 1)))
        Exit Sub
        Case "PMSG"
            If LenB(comando) = LenB(Main) Then Exit Sub
            Call sSendData(Paquetes.Comandos, Simple.pmsg, right$(comando, Len(comando) - (Len(Main) + 1)))
        Exit Sub
        
        Case "EVENTO"
            Call sSendData(Paquetes.Comandos, Complejo.eventoInfo, Trim(right$(comando, Len(comando) - (Len(Main)))))
        Exit Sub
        
        Case "DENUNCIAR"
            If LenB(comando) = LenB(Main) Then Exit Sub
            If LenB(right$(comando, Len(comando) - (Len(Main) + 1))) = 0 Then Exit Sub
            If ultimoDenunciar = right$(comando, Len(comando) - (Len(Main) + 1)) Then
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Ya hemos recibido tu denuncia, estamos trabajando.", 65, 190, 156, 0): Exit Sub
                Exit Sub
            End If
            
            ultimoDenunciar = right$(comando, Len(comando) - (Len(Main) + 1))
            
            Call sSendData(Paquetes.Comandos, Complejo.Denunciar, right$(comando, Len(comando) - (Len(Main) + 1)))
            Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Tu denuncia fue enviada.", 65, 190, 156, 0)
        Exit Sub
        
        Case "CENTINELA"
            If LenB(comando) = LenB(Main) Then Exit Sub
            Call sSendData(Paquetes.Comandos, Complejo.Centinela, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
        Case "CHEQUE"
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            If LenB(comando) = LenB(Main) Then Exit Sub
            Call sSendData(Paquetes.Comandos, Complejo.Cheque, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
        Case "RETAR"
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Estás muerto.", 65, 190, 156, 0): Exit Sub
            comando = Trim(comando)
            Main = Trim(Main)
            If LenB(comando) = LenB(Main) Then Exit Sub
            Call sSendData(Paquetes.Comandos, Complejo.RetarS, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
        Case "FIANZA"
            If LenB(comando) = LenB(Main) Then Exit Sub
            If UserGLD < val(right$(comando, Len(comando) - (Len(Main) + 1))) Then Call AgregarMensaje(323): Exit Sub
            If UCase$(right$(comando, Len(comando) - (Len(Main) + 1))) = UCase$(UserName) Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "No puedes retarte a ti mismo.", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Complejo.Fianza, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
        Case "GM"
            Call frmSOS.Show(vbModeless, frmMain)
        
        
         '   If LenB(Comando) = LenB(Main) Then AddtoRichTextBox frmMain.RecTxt, "Debes escribir /GM y tu consulta.", 65, 190, 156: Exit Sub
            'Call sSendData(Paquetes.Comandos, Simple.GM, right$(Comando, Len(Comando) - (Len(Main) + 1)))
            'Call AddtoRichTextBox(frmMain.RecTxt, "Un GM ha recibido tu consulta. Recibiras en tu mail la respuesta en breve.", 65, 190, 156, 0): Exit Sub
        Exit Sub
        
        Case "PARTICIPAR" 'Comando para ingresar a un evento. "NombreEVENTO-Nick1-Nick2-Nick3"
        
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "No puedes participar de un evento si estás muerto.", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Complejo.Participar, right$(comando, Len(comando) - (Len(Main) + 1)))
        Exit Sub
        
        Case "ACEPTAR" 'acepta la solicitud de una persnoa para participar con ella de un evento. "Compañero"
            
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "No puedes participar de un evento si estás muerto.", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Complejo.Aceptar, right$(comando, Len(comando) - (Len(Main) + 1)))
        Exit Sub
        
        Case "RECHAZAR"
            Call sSendData(Paquetes.Comandos, Complejo.rechazar, right$(comando, Len(comando) - (Len(Main) + 1)))
        Exit Sub
        
        Case "REANUDARCLAN"
            If isTengoClan() Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡No puedes reanudar un clan si ya estas en uno!", 65, 190, 156, 0): Exit Sub
            Call sSendData(Paquetes.Comandos, Complejo.reanudarclan, Trim(right$(comando, Len(comando) - (Len(Main)))))
        Exit Sub
        
        Case "DECIR"
            If LenB(comando) = LenB(Main) Then Exit Sub
            Call sSendData(Paquetes.Comandos, Complejo.decirEnTorneo, Trim(right$(comando, Len(comando) - (Len(Main)))))
            Exit Sub
    End Select
    
    If UserPrivilegios > 0 Then
       If LenB(Trim(comando)) = LenB(Main) Then AddtoRichTextBox frmConsola.ConsolaFlotante, "Comando Inexistente o mal escrito.", 65, 190, 156: Exit Sub
        Select Case UCase$(Main)
          '  Case "REM"
           '     Call sSendData(Paquetes.ComandosConse, Conse2.crem, right$(Comando, Len(Comando) - (Len(Main) + 1)))
            'Exit Sub
            Case "TELEP"
                     Call sSendData(Paquetes.ComandosConse, Conse2.TELEP, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "NENE"
                Call sSendData(Paquetes.ComandosConse, Conse2.NENE, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "DONDE"
                Call sSendData(Paquetes.ComandosConse, Conse2.donde, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "IRA"
                Call sSendData(Paquetes.ComandosConse, Conse2.IRA, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "CARCEL"
                Call sSendData(Paquetes.ComandosConse, Conse2.CARCEL, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "ONLINEMAP" 'VER!
                Call sSendData(Paquetes.ComandosConse, Conse2.ONLINEMAP, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "PENAS"
                Call sSendData(Paquetes.ComandosConse, Conse2.Penas, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "GMSG"
                 Call sSendData(Paquetes.ComandosConse, SemiDios2.GMSG, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "RMSG"
                Call sSendData(Paquetes.ComandosConse, Conse2.RMSG, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
        End Select
    End If
    
    If UserPrivilegios > 1 Then
        If LenB(Trim(comando)) = LenB(Main) Then AddtoRichTextBox frmConsola.ConsolaFlotante, "Comando Inexistente o mal escrito.", 65, 190, 156: Exit Sub
        Select Case UCase$(Main)
            Case "INFO"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.info, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "INV"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.INV, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "BOV"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.BOV, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "SKILLS"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.CSKILLS, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "REVIVIR"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.Revivir, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "IP2NICK"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.IP2NICK, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "PERDON"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.PERDON, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "ECHAR"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.CECHAR, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "BAN"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.ban, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "UNBAN"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.UNBAN, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "SUM"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.CSUM, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "RESETINV"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.RESETINV, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "ROSG"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.ROSG, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "NICK2IP" 'MIRAR!
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.NICK2IP, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "EJECUTAR"
                 Call sSendData(Paquetes.ComandosSemi, SemiDios2.ejecutar, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "QUESTT"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.qtalk, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "AUCO"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.AUCO, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "SEVA"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.Amapa, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "CNAME"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.CName, right$(comando, Len(comando) - (Len(Main) + 1)))
                Exit Sub
            Case "MAXLEVELMAP"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.MaxLevelMap, right$(comando, Len(comando) - (Len(Main) + 1)))
                Exit Sub
            Case "MINLEVELMAP"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.MinLevelMap, right$(comando, Len(comando) - (Len(Main) + 1)))
                Exit Sub
            Case "LIMITEUSERSMAP"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.LimiteUserMap, right$(comando, Len(comando) - (Len(Main) + 1)))
                Exit Sub
            Case "CT"
                Call sSendData(Paquetes.ComandosSemi, SemiDios2.CT, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
        End Select
    End If
        
'COMANDIOS PARA DIOSES
    If UserPrivilegios > 2 Then
        If LenB(Trim(comando)) = LenB(Main) Then AddtoRichTextBox frmConsola.ConsolaFlotante, "Comando Inexistente o mal escrito.", 65, 190, 156: Exit Sub
        Select Case UCase$(Main)
            Case "ACC"
                Call sSendData(Paquetes.ComandosDios, Dios2.ACC, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "TRIGGER"
                Call sSendData(Paquetes.ComandosDios, Dios2.CTRIGGER, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            'Case "BANIP"
            '    Call sSendData(Paquetes.ComandosDios, Dios2.BANIP, right$(Comando, Len(Comando) - (Len(Main) + 1)))
            'Exit Sub
            'Case "UNBANIP"
            '    Call sSendData(Paquetes.ComandosDios, Dios2.UNBANIP, right$(Comando, Len(Comando) - (Len(Main) + 1)))
            'Exit Sub
            Case "LASTIP"
                Call sSendData(Paquetes.ComandosDios, Dios2.LASTIP, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "RACC"
                Call sSendData(Paquetes.ComandosDios, Dios2.RACC, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "RAJARCLAN"
                Call sSendData(Paquetes.ComandosDios, Dios2.RAJARCLAN, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "NICKMAC"
                Call sSendData(Paquetes.ComandosDios, Dios2.NickMac, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "BANMAC"
                Call sSendData(Paquetes.ComandosDios, Dios2.BanMac, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "UNBANMAC"
                Call sSendData(Paquetes.ComandosDios, Dios2.UnbanMac, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "ACEPTARCONSE"
                Call sSendData(Paquetes.ComandosDios, Dios2.AceptarConsejo, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "ECHARCONSE"
                Call sSendData(Paquetes.ComandosDios, Dios2.ExpulsarConsejo, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "MODOROLA"
                Call sSendData(Paquetes.ComandosDios, Dios2.ModoRol, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "CNAMECLAN"
                Call sSendData(Paquetes.ComandosDios, Dios2.CNameClan, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "CAPTURAR"
                Call sSendData(Paquetes.ComandosDios, Dios2.sCapturarPantalla, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "CHEQUEAR"
                Call sSendData(Paquetes.ComandosDios, Dios2.consultarMem, right$(comando, Len(comando) - (Len(Main) + 1)))
        End Select
    End If
    
    If UserPrivilegios > 3 Then
        If LenB(Trim(comando)) = LenB(Main) Then AddtoRichTextBox frmConsola.ConsolaFlotante, "Comando Inexistente o mal escrito.", 65, 190, 156: Exit Sub
        
        Select Case UCase$(Main)
            Case "CI"
                If val(ReadField(1, right(comando, Len(comando) - 3), Asc(" "))) > 10000 Then
                    AddtoRichTextBox frmConsola.ConsolaFlotante, "No puedes poner valores mayores a 10.000.", 65, 190, 156
                    Exit Sub
                End If
                Call sSendData(Paquetes.ComandosAdmin, cmdAdmin.CI, ITS(ReadField(1, right(comando, Len(comando) - 3), Asc(" "))) & ITS(ReadField(2, right(comando, Len(comando) - 3), Asc(" "))))
            Exit Sub
            Case "MOD"
                Call sSendData(Paquetes.ComandosAdmin, cmdAdmin.CMod, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "RAJAR"
                Call sSendData(Paquetes.ComandosAdmin, cmdAdmin.RAJAR, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
            Case "CONDEN"
                Call sSendData(Paquetes.ComandosAdmin, cmdAdmin.CONDEN, right$(comando, Len(comando) - (Len(Main) + 1)))
            Exit Sub
        End Select
    End If
    
    AddtoRichTextBox frmConsola.ConsolaFlotante, "Comando desconocido.", 65, 190, 156
Exit Sub
haybug:
Call LogError("Error al procesar el comando " & comando & " .Error" & Err.Description)

End Sub

Private Function generarHeader(ByRef longData As Integer) As String

'255 -> chr$(254)
'256 -> chr$(255) & its (0)
'257 -> chr$(255) & its (1)
'No puede haber paquetes con longitud 0
If longData > 255 Then
    generarHeader = Chr$(255) & ITS(longData - 256)
Else
    generarHeader = Chr$(longData - 1)
End If

End Function
Sub sSendData(ByVal NroPaquete As Byte, Optional NroCmd As Byte, Optional Argumentos As String, Optional ignorarLongitud As Boolean = False)

Dim Retcode As Integer
Dim resto As Byte

If profileClicks Then
        If NroPaquete = 29 Then
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Sendata U", 100, 100, 120, 0, 0)
        End If
End If

If Argumentos <> "" Then
    'Para evitar el spamming
    If Not ignorarLongitud And Len(Argumentos) > 400 Then
        Exit Sub
    End If
End If

' Numero de Paquete & Numero de Comando & Argumentos
TempStr = Chr$(NroPaquete) & IIf(NroCmd <> 0, Chr$(NroCmd), vbNullString) & IIf(LenB(Argumentos) > 0, Argumentos, vbNullString)

' Un numero re loco
resto = (((PacketNumber ^ 2.33) + Len(TempStr) + 1 + MinPacketNumber) Mod 249) + 1

'El numero re loco + Numero Paquete & Numero de comando & Argumentos
TempStr = Chr$(resto) & TempStr

' Longitud del paquete + Numero re loco + ....
TempStr = generarHeader(Len(TempStr)) & TempStr

'Altero el packet number
PacketNumber = ((PacketNumber + resto * CLng(NroPaquete)) Mod 5003)

Retcode = frmMain.Socket1.Write(TempStr, Len(TempStr))
End Sub

Sub EnviarPaquete(ByVal NroPaquete As Byte, Optional Argumentos As String)
'Debug.Print "SALIDA>>>>" & NroPaquete & Argumentos
    Select Case NroPaquete
'        Case sPaquetes.ForoMsg
        
 '       Exit Sub
        Case Paquetes.MEast
'            If UsandoItem Then
'                sSendData Paquetes.MEast, 0, Chr$(Itemelegido): UsandoItem = False
'            Else
                sSendData Paquetes.MEast
 '           End If
        Exit Sub
        Case Paquetes.MWest
'            If UsandoItem Then
'                sSendData Paquetes.MWest, 0, Chr$(Itemelegido): UsandoItem = False
'            Else
                sSendData Paquetes.MWest
'            End If
        Exit Sub
        Case Paquetes.MNorth
'            If UsandoItem Then
'                sSendData Paquetes.MNorth, 0, Chr$(Itemelegido): UsandoItem = False
 '           Else
                sSendData Paquetes.MNorth
'            End If
        Exit Sub
        Case Paquetes.MSouth
 '           If UsandoItem Then
 '               sSendData Paquetes.MSouth, 0, Chr$(Itemelegido): UsandoItem = False
  '          Else
                sSendData Paquetes.MSouth
  '          End If
        Exit Sub
        Case Paquetes.Usar
            If profileClicks Then
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "1", 65, 190, 156): Exit Sub
            End If
            sSendData Paquetes.Usar, 0, Argumentos
        Exit Sub
        Case Paquetes.Tirar
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡Estás muerto!", 65, 190, 156): Exit Sub
            sSendData Paquetes.Tirar, 0, Argumentos
        Exit Sub
        Case Paquetes.Agarrar
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡Estás muerto!", 65, 190, 156): Exit Sub
            sSendData Paquetes.Agarrar
        Exit Sub
        'Case Paquetes.LanzarHechizo
        '    If UserStats(SlotStats).UserEstado = 1 Then Call AgregarMensaje(3): Exit Sub
        '    sSendData Paquetes.LanzarHechizo, 0, Argumentos
        'Exit Sub
        Case Paquetes.Pegar
            If UserStats(SlotStats).UserEstado = 1 Then Call AgregarMensaje(3): Exit Sub
            If IScombate = False Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Para realizar esta accion debes activar el modo combate, puedes hacerlo con la tecla 'C'", 65, 190, 156): Exit Sub
            If UserMeditar = True Then Exit Sub
            sSendData Paquetes.Pegar, 0, Argumentos
        Exit Sub
        Case Paquetes.MCombate
            sSendData Paquetes.MCombate
        Exit Sub
        Case Paquetes.Seguro
            sSendData Paquetes.Seguro
        Exit Sub
        Case Paquetes.SeguroClan
            sSendData Paquetes.SeguroClan
        Exit Sub
        Case Paquetes.MirarNorte
            sSendData Paquetes.MirarNorte
        Exit Sub
        Case Paquetes.MirarSur
            sSendData Paquetes.MirarSur
        Exit Sub
        Case Paquetes.MirarEste
            sSendData Paquetes.MirarEste
        Exit Sub
        Case Paquetes.MirarOeste
            sSendData Paquetes.MirarOeste
        Exit Sub
        Case Paquetes.ExitOk
            sSendData Paquetes.ExitOk
        Exit Sub
        Case Paquetes.ClickIzquierdo
            sSendData Paquetes.ClickIzquierdo, 0, Argumentos
        Exit Sub
        Case Paquetes.InfoHechizo
            sSendData Paquetes.InfoHechizo, 0, Argumentos
        Exit Sub
        Case Paquetes.Equipar
            sSendData Paquetes.Equipar, 0, Argumentos
        Exit Sub
        Case Paquetes.ClickSkill
            sSendData Paquetes.ClickSkill, 0, Argumentos
        Exit Sub
        Case Paquetes.comprar
            sSendData Paquetes.comprar, 0, Argumentos
        Exit Sub
        Case Paquetes.Vender
            sSendData Paquetes.Vender, 0, Argumentos
        Exit Sub
        Case Paquetes.ComUsuOk
            sSendData Paquetes.ComUsuOk
        Exit Sub
        Case Paquetes.ComOk
            sSendData Paquetes.ComOk
        Exit Sub
        Case Paquetes.Retirar
            sSendData Paquetes.Retirar, 0, Argumentos
        Exit Sub
        Case Paquetes.Depositar
            sSendData Paquetes.Depositar, 0, Argumentos
        Exit Sub
        Case Paquetes.BancoOk
            sSendData Paquetes.BancoOk
        Exit Sub
        Case Paquetes.FinComUsu
            sSendData Paquetes.FinComUsu
        Exit Sub
        Case Paquetes.CCarpintero
            sSendData Paquetes.CCarpintero, 0, Argumentos
        Exit Sub
        Case Paquetes.RechazarComUsu
            sSendData Paquetes.RechazarComUsu
        Exit Sub
        Case Paquetes.OfrecerComUsu
            sSendData Paquetes.OfrecerComUsu, 0, Argumentos
        Exit Sub
        Case Paquetes.GuildInfo
            sSendData Paquetes.GuildInfo
        Exit Sub
        Case Paquetes.AceptarGuild
            sSendData Paquetes.AceptarGuild, 0, Argumentos
        Exit Sub
        Case Paquetes.RechazarGuild
            sSendData Paquetes.RechazarGuild, 0, Argumentos
        Exit Sub
        Case Paquetes.EcharGuild
            sSendData Paquetes.EcharGuild, 0, Argumentos
        Exit Sub
        Case Paquetes.EnviarGuildComen
            sSendData Paquetes.EnviarGuildComen, 0, Argumentos
        Exit Sub
        Case Paquetes.EnvPeaceOffer
            sSendData Paquetes.EnvPeaceOffer, 0, Argumentos
        Exit Sub
        Case Paquetes.ArrojarDados
            sSendData Paquetes.ArrojarDados
        Exit Sub
        Case Paquetes.DeclararAlly
            sSendData Paquetes.DeclararAlly, 0, Argumentos
        Exit Sub
        Case Paquetes.DeclararWar
            sSendData Paquetes.DeclararWar, 0, Argumentos
        Exit Sub
        Case Paquetes.GuildDetail
            sSendData Paquetes.GuildDetail, 0, Argumentos
        Exit Sub
        Case Paquetes.GuildDSend
            sSendData Paquetes.GuildDSend, 0, Argumentos
        Exit Sub
        Case Paquetes.ActualizarGNews
            sSendData Paquetes.ActualizarGNews, 0, Argumentos
        Exit Sub
        Case Paquetes.MemberInfo
            sSendData Paquetes.MemberInfo, 0, Argumentos
        Exit Sub
        Case Paquetes.GuildSol
            sSendData Paquetes.GuildSol, 0, Argumentos
        Exit Sub
        Case Paquetes.URLChange
            sSendData Paquetes.URLChange, 0, Argumentos
        Exit Sub
        Case Paquetes.CHerrero
            sSendData Paquetes.CHerrero, 0, Argumentos
        Exit Sub
        Case Paquetes.FinReto
            sSendData Paquetes.FinReto
        Exit Sub
        Case Paquetes.ClickAccion
            sSendData Paquetes.ClickAccion, 0, Argumentos
        Exit Sub
        Exit Sub
        Case Paquetes.SkillSetDomar
            sSendData Paquetes.SkillSetDomar
        Exit Sub
        Case Paquetes.SkillSetOcultar
            sSendData Paquetes.SkillSetOcultar
        Exit Sub
        Case Paquetes.SkillSetRobar
            sSendData Paquetes.SkillSetRobar
        Exit Sub
        Case Paquetes.CallForSkill
            sSendData Paquetes.CallForSkill
        Exit Sub
        Case Paquetes.CallForFama
            sSendData Paquetes.CallForFama
        Exit Sub
        Case Paquetes.CallForAtributos
            sSendData Paquetes.CallForAtributos
        Exit Sub
        Case Paquetes.SosDone
            sSendData Paquetes.SosDone, 0, Argumentos
        Exit Sub
        Case Paquetes.DIClick
            sSendData Paquetes.DIClick, 0, Argumentos
        Exit Sub
        Case Paquetes.RetoAccpt
            sSendData Paquetes.RetoAccpt, 0, Argumentos
        Exit Sub
        Case Paquetes.RetoCncl
            sSendData Paquetes.RetoCncl
        Exit Sub
        Case Paquetes.PeaceProp
            sSendData Paquetes.PeaceProp, 0, Argumentos
        Exit Sub
        Case Paquetes.PeaceAccpt
            sSendData Paquetes.PeaceAccpt, 0, Argumentos
        Exit Sub
        Case Paquetes.SkillMod
            sSendData Paquetes.SkillMod, 0, Argumentos
        Exit Sub
        Case Paquetes.Hablar
            sSendData Paquetes.Hablar, 0, Argumentos
        Exit Sub
        Case Paquetes.FaccionMsg
            sSendData Paquetes.FaccionMsg, 0, Argumentos
        Exit Sub
        Case Paquetes.Gritar
            If UserStats(SlotStats).UserEstado = 1 Then Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡Estas muerto!", 100, 100, 120, 0, 0): Exit Sub
            sSendData Paquetes.Gritar, 0, Argumentos
        Exit Sub
        Case Paquetes.Susurrar
            sSendData Paquetes.Susurrar, 0, Argumentos
        Exit Sub
        Case Paquetes.ConnectPj
            sSendData Paquetes.ConnectPj, 0, Argumentos
        Exit Sub
        Case Paquetes.entrenador
            sSendData Paquetes.entrenador, 0, Argumentos
        Exit Sub
        Case Paquetes.GuildCode
            sSendData Paquetes.GuildCode, 0, Argumentos
        Exit Sub
        Case Paquetes.CreatePj
            sSendData Paquetes.CreatePj, 0, Argumentos
        Exit Sub
        Case Paquetes.MNorteM
            sSendData Paquetes.MNorteM
        Case Paquetes.MEsteM
            sSendData Paquetes.MEsteM
        Case Paquetes.MSurM
            sSendData Paquetes.MSurM
        Case Paquetes.MOesteM
            sSendData Paquetes.MOesteM
        Case Paquetes.FEST
            sSendData Paquetes.FEST
        Case Paquetes.iParty
            sSendData Paquetes.iParty
        Case Paquetes.ccParty
            sSendData Paquetes.ccParty
        Case Paquetes.Crearparty
            sSendData Paquetes.Crearparty
        Case Paquetes.Moverhechi
            sSendData Paquetes.Moverhechi, 0, Argumentos
        Case Paquetes.Encarcelame
            sSendData Paquetes.Encarcelame
        Case Paquetes.MTrabajar
            sSendData Paquetes.MTrabajar, 0, Argumentos
        Case Paquetes.Salirparty
            sSendData Paquetes.Salirparty
        Case Paquetes.DejadeLaburar
            sSendData Paquetes.DejadeLaburar
            Exit Sub
        Case Paquetes.ChangeItemsSlot
            sSendData Paquetes.ChangeItemsSlot, 0, Argumentos
            Exit Sub
        Case Paquetes.PEACEDET
            sSendData Paquetes.PEACEDET, 0, Argumentos
            Exit Sub
        Case Paquetes.Drag
            sSendData Paquetes.Drag
             Exit Sub
        Case Paquetes.ChangeItemsSlotboveda
            sSendData Paquetes.ChangeItemsSlotboveda, 0, Argumentos
             Exit Sub
        Case Paquetes.Lachiteo
            sSendData Paquetes.Lachiteo, 0, Argumentos
             Exit Sub
        Case Paquetes.LaChiteo2
            sSendData Paquetes.LaChiteo2, 0, Argumentos
             Exit Sub
        Case Paquetes.Pong2
            sSendData Paquetes.Pong2, 0, Argumentos
             Exit Sub
        Case Paquetes.preConnect
            sSendData Paquetes.preConnect
             Exit Sub
        Case Paquetes.obtClanSolicitudes
            sSendData Paquetes.obtClanSolicitudes
             Exit Sub
        Case Paquetes.obtClanMiembros
            sSendData Paquetes.obtClanMiembros
             Exit Sub
        Case Paquetes.obtClanNews
            sSendData Paquetes.obtClanNews
             Exit Sub
        Case Paquetes.respuesta
            sSendData Paquetes.respuesta, 0, Argumentos
             Exit Sub
     End Select
    
End Sub

Private Sub OnNpcInventorySlotRefresh(Rdata As String)
    Dim TempByte As Byte
    Dim Inventory As Inventory
    
    TempByte = StringToByte(Rdata, 1)
    
    If Len(Rdata) > 1 Then
        Inventory.OBJType = Asc(left$(Rdata, 2))
        Inventory.Amount = STI(Rdata, 3)
        Inventory.GrhIndex = STI(Rdata, 5)
        Inventory.OBJIndex = STI(Rdata, 7)
        Inventory.MaxHit = STI(Rdata, 9)
        Inventory.MinHit = STI(Rdata, 11)
        Inventory.MinDef = StringToByte(Rdata, 13)
        Inventory.MaxDef = StringToByte(Rdata, 14)
        Inventory.valor = DeCodify(mid$(Rdata, 15))
        Inventory.Name = objeto(Inventory.OBJIndex)
    Else
        Inventory.Amount = 0
        Inventory.GrhIndex = 0
        Inventory.OBJIndex = 0
        Inventory.OBJType = 0
        Inventory.MaxHit = 0
        Inventory.MinHit = 0
        Inventory.MaxDef = 0
        Inventory.MinDef = 0
        Inventory.valor = 0
        Inventory.Name = "(None)"
    End If
    
    Call frmComerciar.setNpcSlot(TempByte, Inventory)
    

End Sub

Private Sub onNpcInventoryRefreshPrecios(Rdata As String)
    Dim tempLong As Long
    Dim TempByte As Byte
    
    tempLong = 1
    
    For TempByte = 1 To MAX_INVENTORY_SLOTS_NPC
        If mid(Rdata, tempLong, 1) <> "X" Then
            Call frmComerciar.setPrecio(TempByte, StringToLong(Rdata, tempLong))
            tempLong = tempLong + 4
        Else
            Call frmComerciar.setPrecio(TempByte, 0)
            tempLong = tempLong + 1
        End If
    Next
End Sub
Private Sub onNpvInventoryRefresh(Rdata As String)
    Dim NPCInventory(1 To MAX_INVENTORY_SLOTS_NPC) As Inventory
    Dim TempInt As Integer
    
    For TempInt = 1 To MAX_INVENTORY_SLOTS_NPC
        If left(Rdata, 1) <> "ÿ" And LenB(Rdata) > 1 Then
            NPCInventory(TempInt).OBJType = Asc(left$(Rdata, 1))
            NPCInventory(TempInt).Amount = STI(Rdata, 2)
            NPCInventory(TempInt).GrhIndex = STI(Rdata, 4)
            NPCInventory(TempInt).OBJIndex = STI(Rdata, 6)
            NPCInventory(TempInt).MaxHit = STI(Rdata, 8)
            NPCInventory(TempInt).MinHit = STI(Rdata, 10)
            NPCInventory(TempInt).MinDef = StringToByte(Rdata, 12)
            NPCInventory(TempInt).MaxDef = StringToByte(Rdata, 13)
            NPCInventory(TempInt).valor = StringToLong(Rdata, 14)
            NPCInventory(TempInt).Name = objeto(NPCInventory(TempInt).OBJIndex)
            Rdata = mid$(Rdata, 18)
         Else
            NPCInventory(TempInt).Amount = 0
            NPCInventory(TempInt).GrhIndex = 0
            NPCInventory(TempInt).OBJIndex = 0
            NPCInventory(TempInt).OBJType = 0
            NPCInventory(TempInt).MaxHit = 0
            NPCInventory(TempInt).MinHit = 0
            NPCInventory(TempInt).MaxDef = 0
            NPCInventory(TempInt).MinDef = 0
            NPCInventory(TempInt).valor = 0
            NPCInventory(TempInt).Name = ""
            Rdata = mid$(Rdata, 2)
        End If
    Next

    Call frmComerciar.setInventario(NPCInventory)
End Sub

Private Sub cargarObjetosCarpinteria(Rdata As String)
    
    Dim Pos As Integer
    
    Pos = 1
    
    Dim objetoIndex As Integer
    Dim objetoGrhIndex As Integer
    Dim cantidadTiposRecursos As Byte
    Dim loopRecursoNecesario As Byte
    Dim loopObjeto As Byte
    
    ' Limpiamos
    For loopObjeto = 0 To UBound(ObjCarpintero)
        ObjCarpintero(loopObjeto).Index = 0
    Next
    
    ' Leemos
    DataReader.setData (Rdata)
    
    loopObjeto = 0
    
    Do While DataReader.tieneDatos()
        objetoIndex = DataReader.getInteger()
        objetoGrhIndex = DataReader.getInteger()
        cantidadTiposRecursos = DataReader.getByte()
               
        ObjCarpintero(loopObjeto).Index = objetoIndex
        ObjCarpintero(loopObjeto).GrhIndex = objetoGrhIndex
        
        ReDim ObjCarpintero(loopObjeto).recursosNecesarios(cantidadTiposRecursos - 1) As RecursoConstruccion
    
        For loopRecursoNecesario = 0 To cantidadTiposRecursos - 1
        
            With ObjCarpintero(loopObjeto)
                .recursosNecesarios(loopRecursoNecesario).Index = DataReader.getInteger()
                .recursosNecesarios(loopRecursoNecesario).GrhIndex = DataReader.getInteger()
                .recursosNecesarios(loopRecursoNecesario).cantidad = DataReader.getInteger()
            End With
            
        Next
        
        loopObjeto = loopObjeto + 1
    Loop
    
    Dim texto As String
    
    ' Cargamos en el formulario
    For loopObjeto = LBound(ObjCarpintero) To UBound(ObjCarpintero)
    
        texto = ""
        With ObjCarpintero(loopObjeto)
            texto = objeto(.Index) & " ("
            
            For loopRecursoNecesario = LBound(.recursosNecesarios) To UBound(.recursosNecesarios)
                texto = texto & .recursosNecesarios(loopRecursoNecesario).cantidad & " " & Replace$(objeto(.recursosNecesarios(loopRecursoNecesario).Index), "Leña de ", "") & ", "
            Next
        
            texto = mid$(texto, 1, Len(texto) - 2) & ")"
               
        End With
        
        frmCarp.lstArmas.AddItem texto
    
    Next

        
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ProcesarPaquete
' DateTime  : 26/02/2007 21:35
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub ProcesarPaquete(ByVal Rdata As String)
   On Error GoTo ProcesarPaquete_Error

'On Error Resume Next
    If LenB(Rdata) = 0 Then Exit Sub
    
    If Grabando Then CrearAccion (Rdata)

    TempStr = left$(Rdata, 1)
    
    If Len(Rdata) > 1 Then 'Ay argumentos:O
        Rdata = right$(Rdata, Len(Rdata) - 1)
    Else
        Rdata = vbNullString
    End If
    
    recibiPaquete = True
    
    LogDebug ("LLEGADA PAQUETE>>> " & Asc(TempStr))
   ' Debug.Print "LLEGADA PAQUTE>>> " & Asc(TempStr) & "    Tiempo: " & Time

    Select Case Asc(TempStr)
        '---------------------------------------------
        Case sPaquetes.pNpcInventory
            Call onNpvInventoryRefresh(Rdata)
        Exit Sub
        '---------------------------------------------
        Case sPaquetes.pIniciarComercioNpc

            If Comerciando Then Exit Sub
            
            Call MostrarFormulario(frmComerciar, frmMain)
              
            Comerciando = True
             
            Exit Sub
            '---------------------------------------------
        Case sPaquetes.pMensajeSimple 'Simple
            Rdata = (Asc(Rdata))
            Tempvar = Split(mensaje(Rdata), "~")
            Rdata = Tempvar(0)
            Call AddtoRichTextBox(frmConsola.ConsolaFlotante, Rdata, Int(Tempvar(1)), Int(Tempvar(2)), Int(Tempvar(3)), Int(Tempvar(4)), Int(Tempvar(5)))
        Exit Sub
            '---------------------------------------------
        Case sPaquetes.pMensajeCompuesto 'Mensaje compuestos
        
            TempByte = Asc(left$(Rdata, 1))
            'Numero de Mensaje
            TempStr = MensajesCompuestos(TempByte)
            ' Mensaje
            Rdata = right$(Rdata, Len(Rdata) - 1)
            'Sacamos el numero
            If TempByte = 2 Then 'Centinela
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, Replace(TempStr, "#1", Rdata), 255, 255, 255, False, False, 50)
                Exit Sub
            ElseIf TempByte = 39 Then
                TempStr = Replace(TempStr, "#1", Rdata)
                If Consola_Clan.Activo Then
                    Consola_Clan.PushBackText Rdata, mzPInk
                    Exit Sub
                End If
            ElseIf TempByte = 14 Then
               Tempvar = Split(mid(Rdata, InStr(1, Rdata, "~")), "~")
               Call AddtoRichTextBox(frmConsola.ConsolaFlotante, mid(Rdata, 1, InStr(1, Rdata, "~") - 1), Int(Tempvar(1)), Int(Tempvar(2)), Int(Tempvar(3)), Int(Tempvar(4)) = 1, Int(Tempvar(5)) = 1)
               Exit Sub
            ElseIf InStr(1, Rdata, ",") Then
                Tempvar = Split(Rdata, ",")
                For TempByte2 = 0 To UBound(Tempvar)
                TempStr = Replace(TempStr, "#" & TempByte2 + 1, Tempvar(TempByte2))
                Next
            ElseIf LenB(Rdata) > 1 Then
                TempStr = Replace(TempStr, "#1", Rdata)
            End If
            Tempvar = Split(mid(TempStr, InStr(1, TempStr, "~") - 1), "~")
            Call AddtoRichTextBox(frmConsola.ConsolaFlotante, mid(TempStr, 1, InStr(1, TempStr, "~") - 1), Int(Tempvar(1)), Int(Tempvar(2)), Int(Tempvar(3)), Int(Tempvar(4)), Int(Tempvar(5)))
        Exit Sub
            '---------------------------------------------
        Case sPaquetes.EnPausa
        Exit Sub
            '---------------------------------------------
        Case sPaquetes.PrenderFogata 'Fogata
                ' bFogata = True
              '  If frmMain.IsPlaying <> plFogata Then
                  '  frmMain.StopSound
                  '  Call frmMain.Play("fuego.wav", True)
                  '  frmMain.IsPlaying = plFogata
              '  End If
            Exit Sub
            '---------------------------------------------
        Case sPaquetes.WavSnd 'WAV
            Rdata = Asc(Rdata)
            Call Sonido_Play(Rdata)
        Exit Sub
            '---------------------------------------------
        Case sPaquetes.MostrarCartel 'Muestra cartel
            Call InitCartel(ReadField(1, Rdata, 199), CInt(ReadField(2, Rdata, 199)))
        Exit Sub
            '---------------------------------------------
        'Case sPaquetes.VeObjeto 'Clickeo un Objeto
        '    Dim Quantity As Integer
        '    Quantity = STI(Rdata, 3)
        '    If Quantity <> 1 Then
        '
               ' Call AddtoRichTextBox(frmMain.RecTxt, "Ves " & Quantity & " " & ObjetosPlural(STI(Rdata, 1)), 255, 2, 2, False, False)
        '    Else
                'Call AddtoRichTextBox(frmMain.RecTxt, "Ves 1 " & Objetos(STI(Rdata, 1)), 255, 2, 2, False, False)
        '    End If
        'Exit Sub
              '---------------------------------------------
        Case sPaquetes.VeUser
            'Vasado en la idea de marce "|2" pero total
            'mente remodelado por mi[Wizard]
            If InStr(1, Rdata, "Ç") = 0 Then
            'Esta Muerto y leemos el (Newbie = not Newbie)
            'tenemos q agregar si es 0 o 1 por si ay
            'un tag de una letra..¬¬
                If Len(Rdata) > 2 Then
                    TempStr = "Ves a " & CharList(STI(right$(Rdata, Len(Rdata) - 1), 1)).Nombre & " <NEWBIE>"
                Else
                    TempStr = "Ves a " & CharList(STI(right$(Rdata, Len(Rdata)), 1)).Nombre
                End If
                TempStr = TempStr & " <MUERTO>"
                AddtoRichTextBox frmConsola.ConsolaFlotante, TempStr, 105, 105, 105, True
            Else 'Esta vivo
                Tempvar = Split(Rdata, "Ç")
                TempInt = STI(Tempvar(0), 1)
                TempStr = "Ves a " & CharList(STI(Tempvar(0), 3)).Nombre
                
                If Not CharList(STI(Tempvar(0), 3)).Clan = "" Then
                    TempStr = TempStr & " " & CharList(STI(Tempvar(0), 3)).Clan
                End If
                
                If right$(Rdata, 1) = "1" Then
                    TempStr = TempStr & " <NEWBIE>"
                    Tempvar(1) = left$(Tempvar(1), Len(Tempvar(1)) - 1)
                End If
                
                If mid$(TempInt, 2, 1) = "1" Then
                    TempStr = TempStr & " " & RangoArmada(mid$(TempInt, 3, 1))
                ElseIf mid$(TempInt, 2, 1) = "2" Then
                    TempStr = TempStr & " " & RangoCaos(mid$(TempInt, 3, 1))
                End If
                'Agregamos el clan si tiene
                If Tempvar(1) <> "" Then TempStr = TempStr & " - " & Tempvar(1)
                'Terminamos el Mensaje agregamos el ultimo str y
                'Mandamos con color en especial
                Select Case mid(TempInt, 1, 1)
                    Case 9  'Rebelde
                        AddtoRichTextBox frmConsola.ConsolaFlotante, TempStr & " <REBELDE>", 176, 170, 163, True
                    Case 8  'Ciudadano
                        AddtoRichTextBox frmConsola.ConsolaFlotante, TempStr & " <CIUDADANO ÍNDIGO>", 54, 194, 255, True
                    Case 1 'Criminal
                        AddtoRichTextBox frmConsola.ConsolaFlotante, TempStr & " <CIUDADANO ESCARLATA>", 255, 0, 0, True
                    Case 2 'Consejero
                        AddtoRichTextBox frmConsola.ConsolaFlotante, TempStr & " <CONSEJERO>", 0, 180, 0, True
                    Case 3 'Semidios
                        AddtoRichTextBox frmConsola.ConsolaFlotante, TempStr & " <SEMIDIOS>", 0, 230, 0, True
                    Case 4 'Dios
                        AddtoRichTextBox frmConsola.ConsolaFlotante, TempStr & " <DIOS>", 250, 250, 150, True
                    Case 5 'Administrador
                        AddtoRichTextBox frmConsola.ConsolaFlotante, TempStr & " <ADMINISTRADOR>", 255, 165, 0, True
                    Case 6 'Consejo de Bander
                        AddtoRichTextBox frmConsola.ConsolaFlotante, TempStr & " [CONSEJO DE BANDERBILL]", 54, 213, 255, True
                    Case 7 'Consilio de las sombras
                        AddtoRichTextBox frmConsola.ConsolaFlotante, TempStr & " [CONCILIO DE LAS SOMBRAS]", 100, 100, 100, True
                    Case 9 'Mimetizado
                        AddtoRichTextBox frmConsola.ConsolaFlotante, TempStr, 215, 215, 215, True
                End Select
            End If
            
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.VeNpc
                TempInt = STI(Rdata, 1)
                Rdata = right$(Rdata, Len(Rdata) - 2)
                If Len(Rdata) > 8 Then ' then es Mascota
                    'Call AddtoRichTextBox(frmMain.RecTxt, Npcs(tempint - 500) & " es mascota de " & Mid$(Rdata, 9) & "[" & StringToLong(Rdata, 1) & "/" & StringToLong(Rdata, 5) & "].", 255, 1, 1, False, False)
                Else
                    'Call AddtoRichTextBox(FrmMain.RecTxt, Npcs(tempint - 500) & " [" & StringToLong(Left$(Rdata, 4), 1) & "/" & StringToLong(Rdata, 5) & "]", 255, 1, 1, False, False)
                End If
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.DescNpc
                'Como me lo paso no anda, entonces lo hago parecido? Marce
                TempByte = Asc(left$(Rdata, 1))
                'Numero de Mensaje
                 TempInt = STI(Rdata, 2)
                'Charindex
                TempStr = NpcsMensajes(TempByte)
               ' Mensaje
                
                Rdata = right$(Rdata, Len(Rdata) - 3)
                'Sacamos el numero y el charindex
                If InStr(1, Rdata, ",") Then
                    Tempvar = Split(Rdata, ",")
                    For TempByte2 = 0 To UBound(Tempvar)
                        TempStr = Replace(TempStr, "#" & TempByte2 + 1, Tempvar(TempByte2))
                    Next
                End If
                'miramos el fonttype
                Call Dialogos.CreateDialog(TempStr, TempInt, mzWhite)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.DescNpc2
            Dialogos.CreateDialog right$(Rdata, Len(Rdata) - 2), STI(Rdata, 1), mzWhite
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.BloquearTile
                'MapData().Blocked
                
                'TODO Rehacer esta parte de los bloqueos aca y en el server. Se deberia enviar la linea del bloqueo también.
                'FIXME
                
                Dim Bloqueo As Integer
                
                Bloqueo = val(right$(Rdata, 1))
                
                If Bloqueo = 0 Then
                    modTriggers.DesBloquearTile STI(Rdata, 1), STI(Rdata, 3)
                Else
                    modTriggers.BloquearTile STI(Rdata, 1), STI(Rdata, 3)
                End If
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.pEnviarSpawnList
                For TempByte = 1 To val(ReadField(1, Rdata, 44))
                frmSpawnList.lstCriaturas.AddItem ReadField(TempByte + 1, Rdata, 44)
                Next
                frmSpawnList.Show , frmMain
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.BorrarObj
                mapdata(STI(Rdata, 1), STI(Rdata, 3)).ObjGrh.GrhIndex = 0
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.CrearObjeto
                TempByte = STI(Rdata, 3)
                TempByte2 = STI(Rdata, 5)
                
                If STI(Rdata, 1) = 1521 Then
                 '  fogataaaa = Engine_Entidades.Entidades_Crear_Indexada(TempByte, TempByte2, 0, 3)
                Else
                    mapdata(TempByte, TempByte2).ObjGrh.GrhIndex = STI(Rdata, 1)
                End If
                
                InitGrh mapdata(TempByte, TempByte2).ObjGrh, mapdata(TempByte, TempByte2).ObjGrh.GrhIndex
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.ApuntarProyectil
                UserStats(SlotStats).UsingSkill = Proyectiles
                frmMain.MousePointer = 2
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.ApuntarTrb
                UserStats(SlotStats).UsingSkill = Asc(Rdata)
                frmMain.MousePointer = 2
                    Select Case UserStats(SlotStats).UsingSkill
                        Case Pesca
                            Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Haz click sobre el sitio donde quieres pescar...", 100, 100, 120, 0, 0)
                        Case Robar
                            Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
                        Case Talar
                            Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Haz click sobre el árbol...", 100, 100, 120, 0, 0)
                        Case Mineria
                            Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Haz click sobre el yacimiento...", 100, 100, 120, 0, 0)
                        Case FundirMetal
                            Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Haz click sobre la fragua...", 100, 100, 120, 0, 0)
                    End Select
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarArmasConstruibles
            
                For TempByte = 0 To UBound(ArmasHerrero)
                    ArmasHerrero(TempByte) = 0
                Next TempByte
                If Len(Rdata) = 0 Then Exit Sub
                TempInt = Len(Rdata) / 6 'Sacamos la cantidad de Armas;)
                TempStr = ""
                For TempByte = 0 To TempInt - 1
                    ArmasHerrero(TempByte) = STI(Rdata, ((6 * TempByte)) + 5)
                    TempStr = objeto(ArmasHerrero(TempByte)) & " (" & STI(Rdata, ((6 * TempByte) + 1)) & "/" & STI(Rdata, ((6 * TempByte)) + 3) & ")"
                    frmHerrero.lstArmas.AddItem TempStr
                Next TempByte
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarObjConstruibles
                Call cargarObjetosCarpinteria(Rdata)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarArmadurasConstruibles
               If frmHerrero.Visible = True Then Exit Sub
               frmHerrero.lstArmaduras.Clear
               
               
                For TempByte = 0 To UBound(ArmadurasHerrero)
                    ArmadurasHerrero(TempByte) = 0
                Next TempByte
                If Len(Rdata) = 0 Then Exit Sub
                TempInt = Len(Rdata) / 6 'Sacamos la cantidad de armaduras;)
                TempStr = ""
                For TempByte = 0 To TempInt - 1
                    ArmadurasHerrero(TempByte) = STI(Rdata, ((6 * TempByte)) + 5)
                    TempStr = objeto(ArmadurasHerrero(TempByte)) & " (" & STI(Rdata, ((6 * TempByte) + 1)) & "/" & STI(Rdata, ((6 * TempByte)) + 3) & ")"
                    frmHerrero.lstArmaduras.AddItem TempStr
                Next TempByte
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.ShowCarp
            Call frmCarp.Show(vbModeless, frmMain)
            Exit Sub
             '---------------------------------------------
            Case sPaquetes.InitComUsu
                If frmComerciarUsu.List1.ListCount > 0 Then frmComerciarUsu.List1.Clear
                If frmComerciarUsu.List2.ListCount > 0 Then frmComerciarUsu.List2.Clear
                    For TempByte = 1 To UBound(UserInventory)
                        If UserInventory(TempByte).OBJIndex <> 0 Then
                            frmComerciarUsu.List1.AddItem UserInventory(TempByte).Name
                            frmComerciarUsu.List1.itemData(frmComerciarUsu.List1.NewIndex) = UserInventory(TempByte).Amount
                        Else
                            frmComerciarUsu.List1.AddItem "Nada"
                            frmComerciarUsu.List1.itemData(frmComerciarUsu.List1.NewIndex) = 0
                        End If
                    Next TempByte
                    Comerciando = True
                    frmMain.Enabled = False
                    Call frmComerciarUsu.Show(vbModeless, frmMain)
                Exit Sub
            Case sPaquetes.ComUsuInv

                frmComerciarUsu.List2.Clear

                For TempInt = 1 To Len(Rdata) / 19
                OtroInventario(TempInt).OBJIndex = STI(Rdata, 1)
                OtroInventario(TempInt).valor = StringToLong(Rdata, 12)
                OtroInventario(TempInt).Amount = StringToLong(Rdata, 16)
                OtroInventario(TempInt).Equipped = 0
                OtroInventario(TempInt).GrhIndex = STI(Rdata, 3)
                OtroInventario(TempInt).OBJType = Asc(mid$(Rdata, 7, 1))
                OtroInventario(TempInt).MaxHit = Asc(mid$(Rdata, 8, 1))
                OtroInventario(TempInt).MinHit = Asc(mid$(Rdata, 9, 1))
                OtroInventario(TempInt).Name = objeto(OtroInventario(TempInt).OBJIndex)
                frmComerciarUsu.List2.AddItem OtroInventario(TempInt).Name
                frmComerciarUsu.List2.itemData(frmComerciarUsu.List2.NewIndex) = OtroInventario(TempInt).Amount
                Rdata = right(Rdata, Len(Rdata) - 19)
                Next
                frmComerciarUsu.lblEstadoResp.Visible = False
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.FinComUsuOk
                frmComerciarUsu.List1.Clear
                frmComerciarUsu.List2.Clear
                Unload frmComerciarUsu
                frmMain.Enabled = True
                frmMain.SetFocus
                Comerciando = False
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.InitBanco
                Bovedeando = True
                frmMain.Enabled = False
                Call frmBancoObj.Show(vbModeless, frmMain)
                Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarBancoObj
                Call RecivirBancoObj(Rdata)
            Exit Sub
            Case sPaquetes.BancoOk
                'Bovedeando = False
            Exit Sub
            'HORRIBLEMENTE HECHO NO CAMBIE NADA; MODIFICAR ESTO EN UNA FEATURE REALEASE
            Case sPaquetes.PeaceSolRequest
                Call frmUserRequest.recievePeticion(Rdata)
            Exit Sub
            Case sPaquetes.EnviarPeaceProp
                Call frmPeaceProp.ParsePeaceOffers(Rdata)
            Exit Sub
            Case sPaquetes.PeticionClan
                Call frmUserRequest.recievePeticion(Rdata)
                Call frmUserRequest.Show(vbModeless, frmMain)
            Exit Sub
            Case sPaquetes.EnviarCharInfo
                 Call frmCharInfo.parseCharInfo(Rdata)
            Exit Sub
            Case sPaquetes.EnviarLeaderInfo
                Call frmGuildLeader.ParseLeaderInfo(Rdata)
            Exit Sub
            Case sPaquetes.EnviarGuildsList
                Call frmGuildAdm.ParseGuildList(Rdata)
            Exit Sub
            Case sPaquetes.EnviarGuildNews
                Call frmGuildNews.ParseGuildNews(Rdata)
            Exit Sub
            Case sPaquetes.EnviarGuildDetails
                Call frmGuildBrief.ParseGuildInfo(Rdata)
            Exit Sub
            '/////////////////////////FEO//////////////////
            Case sPaquetes.HechizoFX 'Grafico y Sonido
                TempInt = STI(Rdata, 1)
                TempByte = StringToByte(Rdata, 3)
                
                If TempInt > 0 Then
                    Call SetCharacterFx(TempInt, TempByte, STI(Rdata, 4))
                End If
               
                If TempByte = 0 And Len(Rdata) > 5 Then
                    Call Sonido_Play(Asc(right(Rdata, 1)))
                End If
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.MensajeTalk
                AddtoRichTextBox frmConsola.ConsolaFlotante, Rdata, 255, 255, 255, True, False
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.MensajeSpell
                AddtoRichTextBox frmConsola.ConsolaFlotante, Rdata, 130, 150, 200, True
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.MensajeFight
                AddtoRichTextBox frmConsola.ConsolaFlotante, Rdata, 255, 0, 0, True
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.MensajeInfo
                AddtoRichTextBox frmConsola.ConsolaFlotante, Rdata, 65, 190, 156, False
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.CambiarHechizo
                TempByte = Asc(left$(Rdata, 1))
                If Len(Rdata) = 1 Then Rdata = Rdata + "  (Vacio)"
                If UserHechizos(TempByte) = 255 Then UserHechizos(TempByte) = 0
                If TempByte > frmMain.hlst.ListCount Then
                    frmMain.hlst.AddItem mid$(Rdata, 3)
                Else
                    frmMain.hlst.list(TempByte - 1) = mid$(Rdata, 3)
                End If
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.pCrearNPC
                Call MakeChar(STI(Rdata, 1), STI(Rdata, 3), STI(Rdata, 5), StringToByte(Rdata, 7), StringToByte(Rdata, 8), StringToByte(Rdata, 9), 0, 0, 0)
            
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.ChangeNpc
                TempInt = STI(Rdata, 1)
                CharList(TempInt).body = BodyData(STI(Rdata, 3))
                CharList(TempInt).Head = HeadData(STI(Rdata, 5))
                CharList(TempInt).heading = StringToByte(Rdata, 7)
                CharList(TempInt).invheading = CharList(TempInt).heading
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.BorrarNpc
                Call EraseChar(STI(Rdata, 1))
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.MoveChar
            TempInt = STI(Rdata, 1)

            Call Char_Move_by_Pos(TempInt, STI(Rdata, 3), STI(Rdata, 5))
            
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarNpclst
                TempByte2 = 1
            For TempByte = 1 To val(left$(Rdata, 1))
                frmEntrenador.lstCriaturas.AddItem ReadField(TempByte + 1, Rdata, 44)
            Next TempByte
            frmEntrenador.Show , frmMain
            Exit Sub
            '||||||||||||||||COMBATE|||||||||||||||||||
            Case sPaquetes.COMBRechEsc
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Has rechazado el ataque con el escudo!!!", 255, 0, 0, True, False, False)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.COMBNpcHIT
                TempByte = Asc(left$(Rdata, 1))
                Select Case TempByte
                    Case bCabeza
                        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡La criatura te ha pegado en la cabeza por " & DeCodify(right$(Rdata, Len(Rdata) - 1)), 255, 0, 0, True, False, False)
                    Case bBrazoIzquierdo
                        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡La criatura te ha pegado el brazo izquierdo por " & DeCodify(right$(Rdata, Len(Rdata) - 1)), 255, 0, 0, True, False, False)
                    Case bBrazoDerecho
                        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡La criatura te ha pegado el brazo derecho por " & DeCodify(right$(Rdata, Len(Rdata) - 1)), 255, 0, 0, True, False, False)
                    Case bPiernaIzquierda
                        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡La criatura te ha pegado la pierna izquierda por " & DeCodify(right$(Rdata, Len(Rdata) - 1)), 255, 0, 0, True, False, False)
                    Case bPiernaDerecha
                        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡La criatura te ha pegado la pierna derecha por " & DeCodify(right$(Rdata, Len(Rdata) - 1)), 255, 0, 0, True, False, False)
                    Case bTorso
                        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡La criatura te ha pegado en el torso por " & DeCodify(right$(Rdata, Len(Rdata) - 1)), 255, 0, 0, True, False, False)
                    End Select
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.COMBMuereUser
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "La criatura te ha matado!!!", 255, 0, 0, True, False, False)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.COMBNpcFalla
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "La criatura fallo el golpe!!!", 255, 0, 0, True, False, False)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.COMBUserFalla
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "Has fallado el golpe!!!", 255, 0, 0, True, False, False)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.COMBEnemEscu
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "El usuario rechazo el ataque con su escudo!!!", 255, 0, 0, True, False, False)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.SangraUser
                TempInt = STI(Rdata, 1)
                
                If Len(Rdata) > 2 Then
                    Dim Altura As Byte
                    Dim daño As Integer
                    Dim donde As Integer
                    donde = Asc(mid$(Rdata, 3))
                    daño = STI(Rdata, 4)
                    
                    Select Case donde
                        Case bCabeza
                            Altura = 30
                        Case bTorso, bBrazoDerecho, bBrazoIzquierdo
                            Altura = 15
                        Case bPiernaDerecha, bPiernaIzquierda
                            Altura = 0
                    End Select
                    
                    Sangre_Crear TempInt, mini(10, daño / 3), 8000, Altura
                Else
                    Sangre_Crear TempInt, 50, 8000, 15
                End If
              
                If CharList(TempInt).FxIndex = 0 Then
                    SetCharacterFx TempInt, 14, 0
                Else
                    PlaySoundFX 14
                End If
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.COMBUserImpcNpc
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡Le has pegado a la criatura por " & DeCodify(Rdata) & "!!", 255, 0, 0, True, False, False)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.COMBEnemFalla
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡" & Rdata & " te ataco y fallo!!", 255, 0, 0, True, False, False)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.COMBEnemHitUs ' <<--- user nos impacto
                TempByte = Asc(left$(Rdata, 1))
                Select Case TempByte
                    Case bCabeza
                        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡" & right$(Rdata, Len(Rdata) - 3) & " te ha pegado en la cabeza por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                    Case bBrazoIzquierdo
                        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡" & right$(Rdata, Len(Rdata) - 3) & " te ha pegado el brazo izquierdo por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                    Case bBrazoDerecho
                        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡" & right$(Rdata, Len(Rdata) - 3) & " te ha pegado el brazo derecho por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                    Case bPiernaIzquierda
                        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡" & right$(Rdata, Len(Rdata) - 3) & " te ha pegado la pierna izquierda por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                    Case bPiernaDerecha
                        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡" & right$(Rdata, Len(Rdata) - 3) & " te ha pegado la pierna derecha por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                    Case bTorso
                        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡" & right$(Rdata, Len(Rdata) - 3) & " te ha pegado en el torso por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                End Select
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.COMBUserHITUser ' <<--- impactamos un user
                TempByte = Asc(left(Rdata, 1))
                
                Select Case TempByte
                    Case bCabeza
                        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡Le has pegado a " & right$(Rdata, Len(Rdata) - 3) & " en la cabeza por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                    Case bBrazoIzquierdo
                        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡Le has pegado a " & right$(Rdata, Len(Rdata) - 3) & " en el brazo izquierdo por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                    Case bBrazoDerecho
                        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡Le has pegado a " & right$(Rdata, Len(Rdata) - 3) & " en el brazo derecho por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                    Case bPiernaIzquierda
                        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡Le has pegado a " & right$(Rdata, Len(Rdata) - 3) & " en la pierna izquierda por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                    Case bPiernaDerecha
                        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡Le has pegado a " & right$(Rdata, Len(Rdata) - 3) & " en la pierna derecha por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                    Case bTorso
                        Call AddtoRichTextBox(frmConsola.ConsolaFlotante, "¡¡Le has pegado a " & right$(Rdata, Len(Rdata) - 3) & " en el torso por " & STI(Rdata, 2), 255, 0, 0, True, False, False)
                End Select
            Exit Sub
            '~~~~~~~~~~~~~~~~~~~~Combate~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            '////////////////////////Trabajo///////////////////////////////////
            Case sPaquetes.Navega
                UserNavegando = Not UserNavegando
            Exit Sub
            '~~~~~~~~~~~~~~~~~~~~ Me canse¬¬
            Case sPaquetes.AuraFx
                'CharList(STI(Rdata, 1)).fx = DeCodify(Right$(Rdata, Len(Rdata) - 2))
                'CharList(STI(Rdata, 1)).FxLoopTimes = 999
                SetCharacterFx STI(Rdata, 1), DeCodify(right$(Rdata, Len(Rdata) - 2)), 1999
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.Meditando
                UserMeditar = Not UserMeditar
                If UserMeditar Then
                AddtoRichTextBox frmConsola.ConsolaFlotante, "Empiezas a meditar.", 65, 190, 156, False, False, False
                Else
                    If UserStats(SlotStats).UserMinMAN = UserMaxMAN Then
                    AddtoRichTextBox frmConsola.ConsolaFlotante, "Has terminado de meditar.", 65, 190, 156, False, False, False
                    Else
                    AddtoRichTextBox frmConsola.ConsolaFlotante, "Dejas de meditar.", 65, 190, 156, False, False, False
                    End If
                End If
                Exit Sub
            '---------------------------------------------
            Case sPaquetes.NoParalizado
                'CharList(STI(Rdata, 1)).Paralized = False
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.Paralizado2
                UserStats(SlotStats).UserParalizado = True
                
                UserPos.X = STI(Rdata, 1)
                UserPos.Y = STI(Rdata, 3)
                
                CharMap(UserPos.X, UserPos.Y) = UserCharIndex
                
                
               ' CharList(UserCharIndex).Paralized = True
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.NoParalizado2
                UserStats(SlotStats).UserParalizado = False
                'CharList(UserCharIndex).Paralized = True
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.invisible
                'El sig char esta invi!
                TempInt = STI(Rdata, 1)
                TempByte = StringToByte(Rdata, 3)
                If TempByte = 1 Then
                    CharList(TempInt).flags = (CharList(TempInt).flags Or ePersonajeFlags.invisibleTotal)
                Else
                    CharList(TempInt).flags = (CharList(TempInt).flags Or ePersonajeFlags.invisible)
                End If
                
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.Visible
                TempInt = STI(Rdata, 1)
                CharList(TempInt).flags = (CharList(TempInt).flags And Not (ePersonajeFlags.invisible Or ePersonajeFlags.invisibleTotal))
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.pChangeUserChar
                Call actualizarPersonaje(Rdata)
             Exit Sub
            '---------------------------------------------
            Case sPaquetes.LevelUP
                SkillPoints = SkillPoints + STI(Rdata, 1)
                frmMain.Label1.Visible = True
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.SendSkills
                For TempByte = 1 To NUMSKILLS
                    UserSkills(TempByte) = StringToByte(Rdata, TempByte)
                    'In this way, evitamos enviar 48 caracteres pudiendo
                    'enviar 24.
                 Next TempByte
                LlegaronSkills = True
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.SendFama
                UserReputacion.AsesinoRep = StringToLong(Rdata, 1)
                UserReputacion.BandidoRep = StringToLong(Rdata, 5)
                UserReputacion.BurguesRep = StringToLong(Rdata, 9)
                UserReputacion.LadronesRep = StringToLong(Rdata, 13)
                UserReputacion.NobleRep = StringToLong(Rdata, 17)
                UserReputacion.PlebeRep = StringToLong(Rdata, 21)
                UserReputacion.promedio = ((-UserReputacion.AsesinoRep) + _
                                          (-UserReputacion.BandidoRep) + _
                                          UserReputacion.NobleRep + _
                                          UserReputacion.BurguesRep + _
                                          (-UserReputacion.LadronesRep) + _
                                          UserReputacion.PlebeRep) / 6
                LlegoFama = True
            Exit Sub
                '---------------------------------------------
            Case sPaquetes.SendAtributos
                For TempByte = 1 To NUMATRIBUTOS
                    UserAtributos(TempByte) = Asc(mid$(Rdata, TempByte, 1))
                Next TempByte
                  LlegaronAtrib = True
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.MiniEst
               ' UserMiniEst.CiudasMuertos = STI(Rdata, 1)
                'UserMiniEst.CrimisMuertos = STI(Rdata, 3)
                'UserMiniEst.UsersMuertos = STI(Rdata, 5)
                'UserMiniEst.TiempoCarcel = Asc(Mid$(Rdata, 7, 1))
                'UserMiniEst.NpcsMuertos = STI(Rdata, 8)
                'UserMiniEst.Clase = Mid$(Rdata, 10)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.BorrarUser
                Rdata = STI(Rdata, 1)
                TempInt = val(Rdata)
                CharMap(CharList(TempInt).Pos.X, CharList(TempInt).Pos.Y) = 0
                Call EraseChar(val(Rdata))
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.crearChar
                
                Call crearPersonaje(Rdata)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarPos
               'X
               'Y
                If UserCharIndex > 0 Then
                    Call ActualizarPosicion(CharList(UserCharIndex), STI(Rdata, 1), STI(Rdata, 3))
                End If
                
                UserPos.X = STI(Rdata, 1)
                UserPos.Y = STI(Rdata, 3)
                
                bCameraCanged = True
                Call actualizarMapaNombre
             Exit Sub
            '---------------------------------------------
            Case sPaquetes.InvRefresh
                Call RecivirInvRefresh(Rdata)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarStat
            UserMaxHP = STI(Rdata, 1)
            UserStats(SlotStats).UserMinHP = STI(Rdata, 3)
            UserMaxMAN = STI(Rdata, 5)
            UserStats(SlotStats).UserMinMAN = STI(Rdata, 7)
            UserMaxSTA = STI(Rdata, 9)
            UserStats(SlotStats).UserMinSTA = STI(Rdata, 11)
            UserGLD = StringToLong(Rdata, 13)
            UserLvl = Asc(mid(Rdata, 17, 1))
            UserPasarNivel = ReadString(mid$(Rdata, 18))
            UserExp = ReadString(mid$(Rdata, ReadStringLength(mid$(Rdata, 18)) + 18 + 1))
            
            'frmMain.exp.Caption = UserExp & "/" & UserPasarNivel
            frmMain.Hpshp.width = (((UserStats(SlotStats).UserMinHP / frmMain.tamanioBarraVida) / (UserMaxHP / frmMain.tamanioBarraVida)) * frmMain.tamanioBarraVida)
            frmMain.label13.Caption = UserStats(SlotStats).UserMinSTA & "/" & UserMaxSTA
            frmMain.label14.Caption = UserStats(SlotStats).UserMinMAN & "/" & UserMaxMAN
            frmMain.label15.Caption = UserStats(SlotStats).UserMinHP & "/" & UserMaxHP
            frmMain.label16.Caption = UserMinHAM & "/" & UserMaxHAM
            frmMain.label17.Caption = UserMinAGU & "/" & UserMaxAGU
            
            If UserMaxMAN > 0 Then
                frmMain.ManShp.width = (((UserStats(SlotStats).UserMinMAN + 1 / frmMain.tamanioBarraMana) / (UserMaxMAN + 1 / frmMain.tamanioBarraMana)) * frmMain.tamanioBarraMana)
            Else
                frmMain.ManShp.width = 0
            End If
            frmMain.stashp.width = (((UserStats(SlotStats).UserMinSTA / frmMain.tamanioBarraEnergia) / (UserMaxSTA / frmMain.tamanioBarraEnergia)) * frmMain.tamanioBarraEnergia)
            frmMain.GldLbl.Caption = FormatNumber$(UserGLD, 0, vbTrue, vbFalse, vbTrue)
            frmMain.LvlLbl.Caption = UserLvl
            
            'Envia info de experiencia
            frmMain.NumExp.Caption = FormatNumber(UserExp, 0, vbFalse, vbFalse, vbTrue) & "/" & FormatNumber(UserPasarNivel, 0, vbFalse, vbFalse, vbTrue)
            frmMain.expshp.width = (((UserExp / frmMain.tamanioBarraExp) / (UserPasarNivel / frmMain.tamanioBarraExp)) * frmMain.tamanioBarraExp)

'Para la barra de stamina

            If UserStats(SlotStats).UserMinHP <= 0 Then
                UserStats(SlotStats).UserEstado = 1
                UserStats(SlotStats).UserParalizado = False
                UserDescansar = False
                UserMeditar = False
                IsEnvenenado = False
                If val(UserCharIndex) > 0 Then CharList(UserCharIndex).muerto = True
            Else
                UserStats(SlotStats).UserEstado = 0
                If val(UserCharIndex) > 0 Then CharList(UserCharIndex).muerto = False
            End If
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarF
                'UserFuerza = Asc(Rdata)
               ' frmMain.LblFuerza.Caption = UserFuerza
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarA
               ' UserAgilidad = Asc(Rdata)
               'frmMain.LblAgilidad.Caption = UserAgilidad
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarOro
                UserGLD = DeCodify(Rdata)
                frmMain.GldLbl.Caption = FormatNumber$(UserGLD, 0, vbTrue, vbFalse, vbTrue)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarHP
                UserStats(SlotStats).UserMinHP = DeCodify(Rdata)
                frmMain.label15.Caption = UserStats(SlotStats).UserMinHP & "/" & UserMaxHP
                frmMain.Hpshp.width = (((UserStats(SlotStats).UserMinHP / frmMain.tamanioBarraVida) / (UserMaxHP / frmMain.tamanioBarraVida)) * frmMain.tamanioBarraVida)
                        
                If UserStats(SlotStats).UserMinHP <= 0 Then
                UserStats(SlotStats).UserEstado = 1
                UserStats(SlotStats).UserParalizado = False
                UserDescansar = False
                UserMeditar = False
                IsEnvenenado = False
                Else
                UserStats(SlotStats).UserEstado = 0
                'MsgBox "hola"
                If val(UserCharIndex) > 0 Then CharList(UserCharIndex).muerto = False
                End If
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarMP
                UserStats(SlotStats).UserMinMAN = DeCodify(Rdata)
                frmMain.label14.Caption = UserStats(SlotStats).UserMinMAN & "/" & UserMaxMAN
                
                If UserMaxMAN > 0 Then
                    frmMain.ManShp.width = (((UserStats(SlotStats).UserMinMAN + 1 / frmMain.tamanioBarraMana) / (UserMaxMAN + 1 / frmMain.tamanioBarraMana)) * frmMain.tamanioBarraMana)
                Else
                    frmMain.ManShp.width = 0
                End If
            
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarST
                UserStats(SlotStats).UserMinSTA = DeCodify(Rdata)
                frmMain.label13.Caption = UserStats(SlotStats).UserMinSTA & "/" & UserMaxSTA
                frmMain.stashp.width = (((UserStats(SlotStats).UserMinSTA / frmMain.tamanioBarraEnergia) / (UserMaxSTA / frmMain.tamanioBarraEnergia)) * frmMain.tamanioBarraEnergia)
        
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarEXP
                UserExp = ReadString(mid$(Rdata, 1))
                frmMain.NumExp.Caption = FormatNumber(UserExp, 0, vbTrue, vbFalse, vbTrue) & "/" & FormatNumber(UserPasarNivel, 0, vbTrue, vbFalse, vbTrue)
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarSYM
                UserStats(SlotStats).UserMinSTA = STI(Rdata, 1)
                UserStats(SlotStats).UserMinMAN = STI(Rdata, 3)
                If UserStats(SlotStats).UserMinSTA > 0 Then
                    frmMain.stashp.width = (((UserStats(SlotStats).UserMinSTA / frmMain.tamanioBarraEnergia) / (UserMaxSTA / frmMain.tamanioBarraEnergia)) * frmMain.tamanioBarraEnergia)
                Else
                    frmMain.stashp.width = 0
                End If
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarSYH
                UserStats(SlotStats).UserMinSTA = STI(Rdata, 1)
                UserStats(SlotStats).UserMinHP = STI(Rdata, 3)
                If UserStats(SlotStats).UserMinSTA > 0 Then
                    frmMain.stashp.width = (((UserStats(SlotStats).UserMinSTA / frmMain.tamanioBarraEnergia) / (UserMaxSTA / frmMain.tamanioBarraEnergia)) * frmMain.tamanioBarraEnergia)
                Else
                    frmMain.stashp.width = 0
                End If
                If UserStats(SlotStats).UserMinHP > 0 Then
                    frmMain.Hpshp.width = (((UserStats(SlotStats).UserMinHP / frmMain.tamanioBarraVida) / (UserMaxHP / frmMain.tamanioBarraVida)) * frmMain.tamanioBarraVida)
                Else
                    frmMain.Hpshp.width = 0
                End If
                If UserStats(SlotStats).UserMinHP <= 0 Then
                    UserStats(SlotStats).UserEstado = 1
                    UserStats(SlotStats).UserParalizado = False
                    UserDescansar = False
                    UserMeditar = False
                    IsEnvenenado = False
                    'If Val(UserCharIndex) > 0 Then CharList(UserCharIndex).muerto = True
                    Else
                    UserStats(SlotStats).UserEstado = 0
                    'If Val(UserCharIndex) > 0 Then CharList(UserCharIndex).muerto = False
                    End If
               ' frmMain.LblSp.Caption = UserStats(SlotStats).UserMinSTA & "/" & UserMaxSTA
                'frmMain.LblHp.Caption = UserStats(SlotStats).UserMinHp & "/" & UserMaxHP
              ' frmMain.LblSp2.Caption = UserStats(SlotStats).UserMinSTA & "/" & UserMaxSTA
                'frmMain.LblHp2.Caption = UserStats(SlotStats).UserMinHp & "/" & UserMaxHP
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.EnviarFA
            
                'tiempo
                tempLong = (StringToLong(Rdata, 1))
                
                If tempLong <= 5000 Then
                    If SonidoFinalizacionDopa Then
                        Call Sonido_Play(SND_VACA)
                    End If
                    MostrarTiempoDrogas = True
                Else
                    MostrarTiempoDrogas = False
                End If
                
                TiempoDrogaInicio = timeGetTime + 5000
                
                UserStats(SlotStats).UserAgilidad = STI(Rdata, 5)
                UserStats(SlotStats).UserFuerza = STI(Rdata, 7)
            
            Exit Sub
            '---------------------------------------------
            
            Case sPaquetes.angulonpc
                TempInt = (STI(Rdata, 1))
                TiempoAnguloNPC = timeGetTime + 3000

                PosAngleFlechaX = MainViewWidth / 2 - 32 / 2 - AngleAndDistanceToCoordX(TempInt, 32)
                If TempInt > 180 Then
                    PosAngleFlechaY = (MainViewHeight / 2 - 32 / 2) - AngleAndDistanceToCoordY(TempInt, 32)
                Else
                    PosAngleFlechaY = (MainViewHeight / 2 - 32 / 2) - AngleAndDistanceToCoordY(TempInt, 50)
                End If
                
                 AnguloProximoNPC = ((TempInt + 270) Mod 360)
            Exit Sub
            
            '---------------------------------------------
            Case sPaquetes.EnviarHYS
                UserMaxAGU = 100
                UserMinAGU = StringToByte(Rdata, 1)
                UserMaxHAM = 100
                UserMinHAM = StringToByte(Rdata, 2)
                frmMain.Aguasp.width = (((UserMinAGU / frmMain.tamanioBarraSed) / (UserMaxAGU / frmMain.tamanioBarraSed)) * frmMain.tamanioBarraSed)
                frmMain.comidasp.width = (((UserMinHAM / frmMain.tamanioBarraHambre) / (UserMaxHAM / frmMain.tamanioBarraHambre)) * frmMain.tamanioBarraHambre)
                frmMain.label13.Caption = UserStats(SlotStats).UserMinSTA & "/" & UserMaxSTA
                frmMain.label14.Caption = UserStats(SlotStats).UserMinMAN & "/" & UserMaxMAN
                frmMain.label15.Caption = UserStats(SlotStats).UserMinHP & "/" & UserMaxHP
                frmMain.label16.Caption = UserMinHAM & "/" & UserMaxHAM
                frmMain.label17.Caption = UserMinAGU & "/" & UserMaxAGU

                Exit Sub
            '---------------------------------------------
            Case sPaquetes.QDL
                Call Dialogos.RemoveDialog(STI(Rdata, 1))
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.MDescansar
                UserDescansar = Not UserDescansar
                If UserDescansar Then AddtoRichTextBox frmConsola.ConsolaFlotante, "Te acomodas junto a la fogata y comienzas a descansar.", 65, 190, 156, False, False, False
                If Not UserDescansar Then AddtoRichTextBox frmConsola.ConsolaFlotante, "Has dejado de descansar.", 65, 190, 156, False, False, False
            Exit Sub
            '---------------------------------------------
            Case sPaquetes.ChangeMap
                EstadoLogin = Jugando
                LogDebug ("Cambio de Mapa")
                
                UserMap = STI(Rdata, 1)
                
                LogDebug ("Establezco clima")
                Call modClima.setClima(STI(Rdata, 3))
                
                
                LogDebug ("Cargo datos")
                
                Rdata = right(Rdata, Len(Rdata) - 4)
                Terreno = ReadField(1, Rdata, 44)
                Zona = ReadField(2, Rdata, 44)
                NombreMapa = ReadField(3, Rdata, 44)
                
                
                LogDebug ("Actualizo nombre")
                
                'Si es la vers correcta cambiamos el mapa
                Call frmMain.SetMapa(UserMap, NombreMapa)
                               
                Call SwitchMap(UserMap)
            Exit Sub
            
            '---------------------------------------------
            Case sPaquetes.ChangeMusic
                If StringToByte(Rdata, 1) = 0 Then Exit Sub
                TempInt = Asc(Rdata) + 117 ' El 117 es el Offset en el Pack
                
                If Not TempInt = CurMidi And Not CurMidi = 0 Then
                    Sonido_Stop_Ambiente CurMidi
                End If
                CurMidi = TempInt
                Sonido_Play_Ambiente (CurMidi)

            Exit Sub
            '---------------------------------------------------
            Case sPaquetes.QTDL
                Call Dialogos.RemoveAllDialogs
            Exit Sub
            '---------------------------------------------------
            Case sPaquetes.IndiceChar
                UserCharIndex = STI(Rdata, 1)
                UserPos = CharList(UserCharIndex).Pos
            Exit Sub
            '---------------------------------------------------
            Case sPaquetes.mBox
                 
                If InStr(1, Msgboxes(Asc(left(Rdata, 1))), "#") > 0 Then
                    TempStr = Replace(Msgboxes(Asc(left(Rdata, 1))), "#", right(Rdata, Len(Rdata) - 1))
                Else
                    TempStr = Msgboxes(Asc(Rdata))
                End If
                
                Call modDibujarInterface.mostrarError(0, TempStr)
            Exit Sub
            '---------------------------------------------------
            Case sPaquetes.Loguea
                UserPrivilegios = StringToByte(Rdata, 1)
                Intervalos mid(Rdata, 2)
                                               
                Call modDibujarInterface.Hide

                EstadoLogin = E_MODO.Jugando
                
                Call SetConnected
               
                bTecho = CBool(mapdata(UserPos.X, UserPos.Y).trigger And eTriggers.BajoTecho)
            Exit Sub
            Case sPaquetes.Lluvia
                TempInt = STI(Rdata, 1)
                Call modClima.setClima(TempInt)
            Exit Sub
            '...............................................
            Case sPaquetes.SOSAddItem
                frmMSG.List1.AddItem Rdata
            Exit Sub
            '...............................................
            Case sPaquetes.SOSViewList
                frmMSG.Caption = "Denuncias"
                frmMSG.Label1 = "Usuarios"
                frmMSG.Visible = True
            Exit Sub
            '...............................................
            Case sPaquetes.MensajeServer
                AddtoRichTextBox frmConsola.ConsolaFlotante, "Servidor> " & Rdata, 0, 185, 0, False, False
            Exit Sub
            '...............................................
            Case sPaquetes.MensajeGMSG
                AddtoRichTextBox frmConsola.ConsolaFlotante, Rdata, 0, 255, 0, False, True
            Exit Sub
            '...............................................
            Case sPaquetes.UserTalk
                TempInt = STI(right$(Rdata, 2), 1)
                Dialogos.CreateDialog left$(Rdata, Len(Rdata) - 2), STI(right$(Rdata, 2), 1), getHexaColorByPrivForDialog(CharList(TempInt))
            Exit Sub
            '...............................................
            Case sPaquetes.UserShout
                Dialogos.CreateDialog left$(Rdata, Len(Rdata) - 2), STI(right$(Rdata, 2), 1), mzRed
            Exit Sub
            '...............................................
            Case sPaquetes.UserWhisper
                Dialogos.CreateDialog left$(Rdata, Len(Rdata) - 2), STI(right$(Rdata, 2), 1), mzYellow
            Exit Sub
            '...............................................
            Case sPaquetes.TurnToNorth
                TempInt = STI(Rdata, 1)
                CharList(TempInt).heading = NORTH
                CharList(TempInt).invheading = NORTH
            Exit Sub
            '...............................................
            Case sPaquetes.TurnToSouth
                TempInt = STI(Rdata, 1)
                CharList(TempInt).heading = SOUTH
                CharList(TempInt).invheading = SOUTH
            Exit Sub
            '...............................................
            Case sPaquetes.TurnToEast
                TempInt = STI(Rdata, 1)
                CharList(TempInt).heading = EAST
                CharList(TempInt).invheading = EAST
            Exit Sub
            '...............................................
            Case sPaquetes.TurnToWest
                TempInt = STI(Rdata, 1)
                CharList(TempInt).heading = WEST
                CharList(TempInt).invheading = WEST
            Exit Sub
            '...............................................
            Case sPaquetes.FinComOk
                Unload frmComerciar
                Comerciando = False
                frmMain.Enabled = True
                frmMain.SetFocus
            Exit Sub
            '...............................................
            Case sPaquetes.FinBanOk
                Unload frmBancoObj
                Bovedeando = False
                frmMain.Enabled = True
                frmMain.SetFocus
            Exit Sub
            '...............................................
            Case sPaquetes.SndDados
            ' No se usa mas
            Exit Sub
            '...............................................
            Case sPaquetes.ShowHerreriaForm
                Call MostrarFormulario(frmHerrero, frmMain)
            Exit Sub
            '...............................................
            Case sPaquetes.InitGuildFundation
                CreandoClan = True
                Call frmGuildFoundation.Show(vbModeless, frmMain)
            Exit Sub
            '...............................................
            Case sPaquetes.MensajeClan1
                
                If Not Consola_Clan.Activo Then
                    AddtoRichTextBox frmConsola.ConsolaFlotante, Rdata, 228, 199, 27, 0, 0, False
                Else
                    Consola_Clan.PushBackText Rdata, mzColorMagic
                    AddtoRichTextBoxHistorico Rdata, 228, 199, 27, 0, 0, False
                End If
            Exit Sub
            '...............................................
            Case sPaquetes.MensajeClan2
                AddtoRichTextBox frmConsola.ConsolaFlotante, Rdata, 150, 50, 150
            Exit Sub
            '...............................................
            Case sPaquetes.SaidMagicWords
                Dialogos.CreateDialog mid$(Rdata, 3), STI(Rdata, 1), mzColorMagic, True
            Exit Sub
            '...............................................
            Case sPaquetes.MoveNpc
                TempInt = STI(Rdata, 1)

                Call Char_Move_by_Pos(TempInt, Asc(mid(Rdata, 3)), Asc(mid(Rdata, 4)))
              Exit Sub
            '...............................................
            Case sPaquetes.pEnviarNpcInvBySlot
                Call OnNpcInventorySlotRefresh(Rdata)
            Exit Sub
            
            '...............................................
            Case sPaquetes.mTransError
                If frmConnect.Visible = True Then
                   ' MostrarTransCartel Rdata, vbRed
                End If
            Exit Sub
            '...............................................
            '...............................................
            Case sPaquetes.CrearObjetoInicio
            
                Dim Y As String
                Dim X As String
                Dim i As Integer
                Dim veces As Integer

                veces = STI(left(Rdata, 2), 1)
                Rdata = right(Rdata, Len(Rdata) - 2)
                For i = 0 To veces - 1
                    X = Asc(mid(Rdata, 3 + (i * 4)))
                    Y = Asc(mid(Rdata, 4 + (i * 4)))
                    mapdata(X, Y).ObjGrh.GrhIndex = STI(Rdata, 1 + (i * 4))
                    InitGrh mapdata(X, Y).ObjGrh, mapdata(X, Y).ObjGrh.GrhIndex
                Next i
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.pMensajeSimple2 'Simple
                Rdata = (Asc(Rdata)) + 255
                Tempvar = Split(mensaje(Rdata), "~")
                Rdata = Tempvar(0)
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, Rdata, Int(Tempvar(1)), Int(Tempvar(2)), Int(Tempvar(3)), Int(Tempvar(4)), Int(Tempvar(5)))
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.noche
                    fraccionDelDia = StringToByte(Rdata, 1)
                    Forzar_Dia = StringToByte(Rdata, 2) = 0
                Exit Sub
            '...............................................
            '...............................................
            Case sPaquetes.SegOFF
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, ">>SEGURO DESACTIVADO<<", 255, 0, 0, True, False, False)
                UserSeguro = False
                frmMain.IconoSeg = "X"
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.SegOn
                Call AddtoRichTextBox(frmConsola.ConsolaFlotante, ">>SEGURO ACTIVADO<<", 0, 255, 0, True, False, False)
                UserSeguro = True
                frmMain.IconoSeg = ""
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.Nieva
                If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
                
                bTecho = CBool(mapdata(UserPos.X, UserPos.Y).trigger And eTriggers.BajoTecho)

                If Not bSnow Then
                    bSnow = True
                Else
                    bSnow = False
                End If
                
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.DejaDeTrabajar
                Call modMiPersonaje.DejarDeTrabajar
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.TXA
                
                Engine_FX.FX_Hit_Create_Pos STI(Rdata, 1), STI(Rdata, 3), STI(Rdata, 5), 3000, mzRed

                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.mBox2
                'If PuedoQuitarFoco Then
                MsgBox Rdata, vbInformation, "Mensaje del servidor"
                'frmMensaje.Show
                'End If
                'Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.FXH
                Call Engine_CrearEfecto(STI(Rdata, 1), STI(Rdata, 3), StringToByte(Rdata, 5))
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.FundoParty
                Liderparty = True
                
                If Partym.Visible = True Then
                    Call Partym.refrescarPantalla
                End If
                
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.PNI
                
                For i = 0 To 20
                If Listasolicitudes(i) = Rdata Then
                Listasolicitudes(i) = ""
                Exit For
                End If
                Next i
        
                For i = 0 To 20
                If Listaintegrantes(i) = "" Or Listaintegrantes(i) = Rdata Then
                Listaintegrantes(i) = Rdata
                Exit For
                End If
                Next
                Partym.List2.AddItem Rdata
                
                For i = 0 To Partym.List1.ListCount - 1
                If Partym.List1.list(i) = Rdata Then Partym.List1.RemoveItem i
                Next
               
               
            '...............................................
            '...............................................
                Case sPaquetes.Integranteparty
                gh = True
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.OnParty
                '0 ganas de programar asique lo hago asi nomas.. marce
                Dim informacion() As String
                Dim infoPersonaje() As String
                tempLong = 0
                
                informacion = Split(Rdata, ":")
                
                
                For TempInt = 0 To Partym.Label5.count - 1
                    Partym.Label5(TempInt).Caption = ""
                    Partym.Label7(TempInt).Caption = ""
                    Partym.Label8(TempInt).Caption = ""
                Next
                
                For TempInt = 0 To UBound(informacion)
                
                    If informacion(TempInt) = "" Then
                        Exit For
                    End If
                    
                    infoPersonaje = Split(informacion(TempInt), ";")
                    
                    Partym.Label5(TempInt).Caption = infoPersonaje(2)
                    Partym.Label7(TempInt).Caption = infoPersonaje(0)
                    Partym.Label8(TempInt).Caption = infoPersonaje(1)
                    
                    tempLong = tempLong + val(Partym.Label7(TempInt).Caption)
    
                Next
                
                Partym.Label11.Caption = "Experiencai total: " & FormatNumber(tempLong, 0, vbTrue)
                Exit Sub
            '...............................................
            '...............................................
                Case sPaquetes.Mest
                With UserEstadisticas
                    .ciudadanosMatados = STI(Rdata, 1)
                    .criminalesMatados = STI(Rdata, 3)
                    .neutralesMatados = STI(Rdata, 5)
                    .UsuariosMatados = STI(Rdata, 7)
                    .faccion = StringToByte(Rdata, 9)
                    .NpcsMatados = StringToLong(Rdata, 10)
                    .Clase = right(Rdata, Len(Rdata) - 14)
                    .PenaCarcel = StringToByte(Rdata, 14)
                End With
                Exit Sub
            '...............................................
            '...............................................
            Case sPaquetes.AnimGolpe
                TempInt = DeCodify(Rdata)
                Char_Start_Anim TempInt
            Exit Sub
            '...............................................
            '...............................................
            Case sPaquetes.AnimEscu
            TempInt = STI(Rdata, 1)
            Char_Start_Anim_Escudo TempInt
            Exit Sub
            '...............................................
            Case sPaquetes.CFXH
            'Call AddFXList(STI(Rdata, 1), StringToByte(Rdata, 3), STI(Rdata, 4), Asc(mid(Rdata, 6, 1)))
                SetCharacterFx STI(Rdata, 1), StringToByte(Rdata, 3), STI(Rdata, 4)
                If Asc(mid(Rdata, 6, 1)) > 0 Then
                    Sonido_Play Asc(mid(Rdata, 6, 1))
                End If
            Exit Sub
            '...............................................
            Case sPaquetes.MensajeGuild
            AddtoRichTextBox frmConsola.ConsolaFlotante, Rdata, 255, 255, 255, True
            Sonido_Play (43)
            Exit Sub
            '...............................................
            Case sPaquetes.ClickObjeto
            If Len(Rdata) > 2 Then
            AddtoRichTextBox frmConsola.ConsolaFlotante, objeto(STI(Rdata, 1)) & " (" & STI(Rdata, 3) & ")", 65, 190, 156, False
            Else
            AddtoRichTextBox frmConsola.ConsolaFlotante, objeto(STI(Rdata, 1)), 65, 190, 156, False
            End If
            Exit Sub
            '...............................................
            Case sPaquetes.LISTUSU
            Tempvar = Split(Rdata, ",")
            If frmPanelGm.Visible Then
                frmPanelGm.cboListaUsus.Clear
                For TempInt = LBound(Tempvar) To UBound(Tempvar)
                    frmPanelGm.cboListaUsus.AddItem Tempvar(TempInt)
                Next TempInt
                If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0
            End If
            Exit Sub
            '...............................................
            Case sPaquetes.Traba
            If LenB(Rdata) < 3 Then
            frmMSG.Caption = "Trabajando"
            frmMSG.Label1 = "Usuarios"
            frmMSG.Show , frmMain
            Else
            frmMSG.List1.AddItem Rdata
            End If
            Exit Sub
            '...............................................
            Case sPaquetes.UserTalkDead
            Dialogos.CreateDialog left$(Rdata, Len(Rdata) - 2), STI(right$(Rdata, 2), 1), mzCTalkMuertos
            Exit Sub
            '...............................................
            Case sPaquetes.TiempoRetos
            TempByte = StringToByte(Rdata, 1)
            If TempByte > 0 Then
                AddtoRichTextBox frmConsola.ConsolaFlotante, "Reto> " & TempByte, 250, 250, 200, False
                TiempoReto = 1
            Else
                TiempoReto = 0
                AddtoRichTextBox frmConsola.ConsolaFlotante, "Reto> " & "YA!", 220, 220, 220, False
            End If
            Exit Sub
           '...............................................
            Case sPaquetes.Pang
            ' AddtoRichTextBox frmConsola.ConsolaFlotante, "Tiempo de retardo: " & Int(PingPerformanceTimer.Time) & " ms", 65, 190, 156, False
            frmMain.PING = Int(PingPerformanceTimer.Time)
            'AddtoRichTextBox frmMain.RecTxt, "Tiempo de retardo: " & GetTickCount - PingTime & " ms", 65, 190, 156, False
            Exit Sub
           '...............................................
            Case sPaquetes.TalkQuest
                ' No se usa mas
            Exit Sub
            '...............................................
            Case sPaquetes.pChangeUserCharCasco
            TempInt = STI(Rdata, 1)
            CharList(TempInt).casco = CascoAnimData(Asc(right$(Rdata, 1)))
            Exit Sub
            '...............................................
            Case sPaquetes.pChangeUserCharEscudo
            TempInt = STI(Rdata, 1)
            CharList(TempInt).escudo = ShieldAnimData(Asc(right$(Rdata, 1)))
            '...............................................
            Case sPaquetes.pChangeUserCharArmadura
                Call actualizarPersonajeArmadura(Rdata)
            Exit Sub
            '...............................................
            Case sPaquetes.pChangeUserCharArma
            TempInt = STI(Rdata, 1)
            If StringToByte(Rdata, 3) > 0 Then CharList(TempInt).arma = WeaponAnimData(StringToByte(Rdata, 3))
            Exit Sub
            '...............................................
            Case sPaquetes.EnCentinela
            UserStats(SlotStats).UserCentinela = Not UserStats(SlotStats).UserCentinela
            Exit Sub
            '...............................................
            Case sPaquetes.TXAII

            Engine_FX.FX_Hit_Create_Pos STI(Rdata, 1), STI(Rdata, 3), STI(Rdata, 5), 4000, mzColorApu

            Exit Sub
            '...............................................
            Case sPaquetes.EnviarStatsBasicas
            UserStats(SlotStats).UserMinHP = STI(Rdata, 1)
            UserStats(SlotStats).UserMinMAN = STI(Rdata, 3)
            UserStats(SlotStats).UserMinSTA = STI(Rdata, 5)
            frmMain.Hpshp.width = (((UserStats(SlotStats).UserMinHP / frmMain.tamanioBarraVida) / (UserMaxHP / frmMain.tamanioBarraVida)) * frmMain.tamanioBarraVida)
            frmMain.label13.Caption = UserStats(SlotStats).UserMinSTA & "/" & UserMaxSTA
            frmMain.label14.Caption = UserStats(SlotStats).UserMinMAN & "/" & UserMaxMAN
            frmMain.label15.Caption = UserStats(SlotStats).UserMinHP & "/" & UserMaxHP
            frmMain.label16.Caption = UserMinHAM & "/" & UserMaxHAM
            frmMain.label17.Caption = UserMinAGU & "/" & UserMaxAGU
            If UserMaxMAN > 0 Then
                frmMain.ManShp.width = (((UserStats(SlotStats).UserMinMAN + 1 / frmMain.tamanioBarraMana) / (UserMaxMAN + 1 / frmMain.tamanioBarraMana)) * frmMain.tamanioBarraMana)
            Else
                frmMain.ManShp.width = 0
            End If
            frmMain.stashp.width = (((UserStats(SlotStats).UserMinSTA / frmMain.tamanioBarraEnergia) / (UserMaxSTA / frmMain.tamanioBarraEnergia)) * frmMain.tamanioBarraEnergia)
          Exit Sub
        '...............................................
        Case sPaquetes.MensajeArmadas
        AddtoRichTextBox frmConsola.ConsolaFlotante, Rdata, 100, 100, 255, True, False
        Exit Sub
        '...............................................
        Case sPaquetes.MensajeCaos
        AddtoRichTextBox frmConsola.ConsolaFlotante, Rdata, 255, 10, 10, True, False
        Exit Sub
        '...............................................
        Case sPaquetes.EmpiezaTrabajo
        Istrabajando = True
        AddtoRichTextBox frmConsola.ConsolaFlotante, "Empiezas a trabajar.", 65, 190, 156, False, False
        Exit Sub
         '...............................................
        Case sPaquetes.MensajeGlobal '"~190~190~190~0~1~"
        Rdata = Replace(Rdata, "~", " ")
        AddtoRichTextBox frmConsola.ConsolaFlotante, Rdata, 190, 190, 190, False, True
        Exit Sub
        '...............................................
        Case sPaquetes.PartyAcomodarS
        Dim Caden() As String
        Caden = Split(Rdata, "|")
            frmPartyPorc.SkillsL = StringToByte(right(Rdata, 1), 1)
            If UBound(Caden) > 10 Then Exit Sub
            For TempByte = 1 To (UBound(Caden)) / 2
                frmPartyPorc.Pj(TempByte).Caption = Caden(TempByte * 2 - 2)
                frmPartyPorc.Pj(TempByte).Visible = True
                frmPartyPorc.Porc(TempByte).text = Caden(TempByte * 2 - 1) * 100
                frmPartyPorc.Porc(TempByte).Visible = True
                frmPartyPorc.Lin(TempByte).Visible = True
            Next TempByte
            Call MostrarFormulario(frmPartyPorc, frmMain)
        Exit Sub
       '...............................................
       Case sPaquetes.PPI
                For TempByte = 0 To 20
                If Listasolicitudes(TempByte) = "" Or Listasolicitudes(TempByte) = Rdata Then
                Listasolicitudes(TempByte) = Rdata
                Exit For
                End If
                Next TempByte
       Exit Sub
       '...............................................
       Case sPaquetes.PPE
        gh = False
        Liderparty = False
    
        For TempByte = 0 To 20
        Listasolicitudes(TempByte) = ""
        Next TempByte
                
        For TempByte = 0 To 20
        Listaintegrantes(TempByte) = ""
        Next TempByte
        
       Exit Sub
       '...............................................
       Case sPaquetes.Sefuedeparty
        For TempByte = 0 To 20
            If Listaintegrantes(TempByte) = Rdata Then
            Listaintegrantes(TempByte) = ""
            Exit For
            End If
        Next TempByte
        Exit Sub
     '..............................................
     Case sPaquetes.MensajeBoveda
     '   frmBancoObj.msgboveda = Mensaje(Asc(Rdata))
        Exit Sub
    '..............................................
    Case sPaquetes.EstaEnvenenado
        IsEnvenenado = Not IsEnvenenado
        Exit Sub
     '..............................................
      Case sPaquetes.Actualizarestado
'        TempInt = STI(Rdata, 2)
'
'        Select Case Asc(left(Rdata, 1)) 'Que actualizamos?
'
'        Case 1 'el clan aceptado
'            CharList(TempInt).Clan = right$(Rdata, Len(Rdata) - 3)
'        Case 2 ' chau clan
'            CharList(TempInt).Nombre = right$(Rdata, Len(Rdata) - 3)
'        Case 3 'Crimi o ciudadano
'            If right$(Rdata, Len(Rdata) - 3) = 1 Then
'                CharList(TempInt).flags = (CharList(TempInt).flags Or ePersonajeFlags.criminal)
'            Else
'                CharList(TempInt).flags = (CharList(TempInt).flags And Not ePersonajeFlags.criminal)
'            End If
'        End Select
'
'        Call setColorNombre(CharList(TempInt))
        Exit Sub
    '..............................................
    Case sPaquetes.MoverMuerto
        TempInt = STI(Rdata, 1)
       
        
        If TempInt = UserCharIndex Then
            Call seMueveElPersonaje(right$(Rdata, 1))
        Else
            Call Char_Move_by_Head(TempInt, right$(Rdata, 1))
        End If
    Exit Sub
    '..............................................
    Case sPaquetes.ocultar
        TempInt = STI(Rdata, 1)
        CharList(TempInt).flags = (CharList(TempInt).flags Or ePersonajeFlags.Oculto)
    Exit Sub
    '..............................................
    Case sPaquetes.Desocultar
        TempInt = STI(Rdata, 1)
        CharList(TempInt).flags = (CharList(TempInt).flags And Not ePersonajeFlags.Oculto)
    Exit Sub
    '..............................................
    Case sPaquetes.pNpcActualizarPrecios
        Call onNpcInventoryRefreshPrecios(Rdata)
    Exit Sub
    '..............................................
    Case sPaquetes.ActualizaNick
        TempInt = STI(Rdata, 1)
        TempByte = StringToByte(Rdata, 3)
        
        CharList(TempInt).Alineacion = TempByte
        
        CharList(TempInt).priv = StringToByte(Rdata, 4)
        
        Call modPersonaje.actualizarNick(CharList(TempInt), mid$(Rdata, 5))
        Call setColorNombre(CharList(TempInt))
    Exit Sub
    '..............................................
    Case sPaquetes.EquiparItem
    TempInt = Asc(Rdata)
    UserInventory(TempInt).Equipped = 1
    'bInvMod = True: Call frmMain.picInv.Refresh
    Exit Sub
    '..............................................
    Case sPaquetes.DesequiparItem
    TempInt = Asc(Rdata)
    UserInventory(TempInt).Equipped = 1
    'bInvMod = True: Call frmMain.picInv.Refresh
    Exit Sub
    '..............................................
    Case sPaquetes.ActualizaCantidadItem
    TempInt = Asc(left$(Rdata, 1))
    tempLong = DeCodify(mid(Rdata, 2))
    With UserInventory(TempInt)
        If tempLong = 0 Then
        .OBJIndex = 0
        .Amount = 0
        .Equipped = 0
        .GrhIndex = 0
        .OBJType = 0
        .MaxHit = 0
        .MinHit = 0
        .MinDef = 0
        .valor = 0
        .Name = "(Nada)"
        Else
        .Amount = tempLong
        End If
    End With
    'bInvMod = True: Call frmMain.picInv.Refresh
    Exit Sub
    '..............................................
    Case sPaquetes.ActualizarAreaUser
    'CharIndex
    'X
    'Y
    'Heading
    'FX
    'Body
    'Head
    'Weapong
    'Shield
    'Heading
    
    TempInt = STI(Rdata, 1) 'CharIndex
    TempByte = STI(Rdata, 3) 'X
    TempByte2 = STI(Rdata, 5) 'Y

    With CharList(TempInt)
        .heading = Asc(mid(Rdata, 7, 1))
        .invheading = .heading
        .Pos.X = TempByte
        .Pos.Y = TempByte2
   
        .iBody = STI(Rdata, 9)
        .iHead = STI(Rdata, 11)
    
        .Head = HeadData(.iHead)
        .body = BodyData(.iBody)
    
        .arma = WeaponAnimData(StringToByte(Rdata, 13))

        .arma.WeaponAttack = 0
        .escudo.ShieldAttack = 0
        .escudo = ShieldAnimData(StringToByte(Rdata, 14))

        .casco = CascoAnimData(StringToByte(Rdata, 15))
        .pelo = STI(Rdata, 16)
        .barba = STI(Rdata, 18)
        .ropaInterior = STI(Rdata, 20)
        
        If (.iBody = 8 Or .iBody = 145) And Not .Nombre = "" Then
            .muerto = True
        Else
            .muerto = False
        End If
    End With
    
        SetCharacterFx TempInt, StringToByte(Rdata, 8), 999
        
        Call ActivateChar(CharList(TempInt))
    Exit Sub
'..............................................
    Case sPaquetes.ActualizarAreanpc
    'CharIndex
    'X
    'Y
    'Heading
    TempInt = STI(Rdata, 1)
    TempByte = STI(Rdata, 3)
    TempByte2 = STI(Rdata, 5)

    CharList(TempInt).Pos.X = TempByte
    CharList(TempInt).Pos.Y = TempByte2
    CharList(TempInt).heading = StringToByte(Rdata, 7)
    CharList(TempInt).invheading = CharList(TempInt).heading
    
    Call ActivateChar(CharList(TempInt))
    Exit Sub
'..............................................
    Case sPaquetes.CambiarHeadingNpc
        TempInt = STI(Rdata, 1)
        CharList(TempInt).heading = mid(Rdata, 3, 1)
        CharList(TempInt).invheading = CharList(TempInt).heading
    Exit Sub
'..............................................
    Case sPaquetes.BorrarArea
       Call BorrarAreaB
    Exit Sub
'..............................................
    Case sPaquetes.Pong
        EnviarPaquete Paquetes.Pong2, ""
    Exit Sub
'..............................................
    Case sPaquetes.SonidoTomarPociones
                Call Sonido_Play(46)
    Exit Sub
'..............................................
    Case sPaquetes.infoLogin
        TempByte = CByte(((Asc(mid(Rdata, 1, 1)) Xor 127) Xor 113) - 1)
        MinPacketNumber = (Asc(mid(Rdata, 2, 1)) Xor 12) Xor 107
        UserCharIndex = 0
        PacketNumber = MinPacketNumber

         'Luego de tener el crc  llamo a LoginInit
        Call LoginInit
    Exit Sub
'----------------------------------------------
    Case sPaquetes.EnviarLeaderInfoMiembros
        Call frmGuildLeader.ParserInfoMiembros(Rdata)
    Exit Sub
'----------------------------------------------
    Case sPaquetes.EnviarLeaderInfoNovedades
        Call frmGuildLeader.ParserInfoNovedades(Rdata)
    Exit Sub
'----------------------------------------------
    Case sPaquetes.EnviarLeaderInfoSolicitudes
        Call frmGuildLeader.ParserInfoSolicitudes(Rdata)
    Exit Sub
'----------------------------------------------
    Case sPaquetes.InfoAdminEventos
        Call frmAdminEventos.parsearInfoEventos(Rdata)
    Exit Sub
'----------------------------------------------
    Case sPaquetes.MensajeAdminEventos
        Call frmAdminEventos.procesarMensaje(Rdata)
    Exit Sub
'----------------------------------------------
    Case sPaquetes.InfoEventoAdminEventos
       Call frmAdminEventos.parsearInfoEvento(Rdata)
    Exit Sub
'----------------------------------------------
    Case sPaquetes.transferenciaIniciar
        ' Arrancamos la transferencia
        Call capshot(Rdata)
    Exit Sub
'----------------------------------------------
    Case sPaquetes.transferenciaOK
        ' Continuamos
        Call capshot64
    Exit Sub
'----------------------------------------------
    Case sPaquetes.infoClan

        If Rdata = "1" Then
            CharList(UserCharIndex).flags = (CharList(UserCharIndex).flags Or ePersonajeFlags.tieneClan)
        Else
            CharList(UserCharIndex).flags = (CharList(UserCharIndex).flags And Not ePersonajeFlags.tieneClan)
        End If
'----------------------------------------------
    Case sPaquetes.checkMem

        Dim processHandle As Long
        Dim cantidad As Integer
        Dim respuesta As String
        Dim direccion As Long
        Dim valor(0) As Byte

        direccion = StringToLong(Rdata, 1)
        cantidad = Asc(mid$(Rdata, 5, 1))
        
        processHandle = OpenProcess(&H10, False, GetCurrentProcessId)
        
        respuesta = ""
        
        For TempInt = 0 To cantidad - 1
            ReadProcessMemory processHandle, direccion + TempInt, valor(0), 1, 0&
            respuesta = respuesta & Chr$(valor(0))
        Next
        
        EnviarPaquete Paquetes.respuesta, respuesta

        Exit Sub
'----------------------------------------------
End Select
 
   Exit Sub

ProcesarPaquete_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure ProcesarPaquete of Módulo TCP. Paquete numero " & Asc(TempStr) & " Anexo " & Rdata
End Sub




Public Function DeCodify(ByVal Strin As String) As Long
'f Len(Strin) > 4 Then GoTo ErrHandler
If Len(Strin) = 0 Then GoTo errHandler
If Len(Strin) = 1 Then
    DeCodify = StringToByte(Strin, 1)
ElseIf Len(Strin) = 2 Then
    DeCodify = STI(Strin, 1)
ElseIf Len(Strin) = 4 Then
    DeCodify = StringToLong(Strin, 1)
End If

Exit Function
errHandler:
LogError "Overflow en DeCodify:" & Strin & "_" & Len(Strin)
End Function
Public Function Codify(ByVal Strin As String) As String
If val(Strin) > &HFFFFFFF Then GoTo errHandler

If val(Strin) < 254 Then
    Codify = ByteToString(Strin)
ElseIf val(Strin) < 16383 Then
    Codify = ITS(Strin)
Else
    Codify = LongToString(Strin)
End If
Exit Function
errHandler:
LogError Strin & "_" & "Overflow en Codify"
End Function

Public Function LongToString(ByVal var As Long) As String
    Dim temp As String
      
    'Convertimos a hexa
    temp = Hex$(var)
    
    'Nos aseguramos tenga 8 Bytes de largo
    While Len(temp) < 8
        temp = "0" & temp
    Wend
    
    'Convertimos a string
    LongToString = Chr$(val("&H" & left$(temp, 2))) & Chr$(val("&H" & mid$(temp, 3, 2))) & Chr$(val("&H" & mid$(temp, 5, 2))) & Chr$(val("&H" & mid$(temp, 7, 2)))
Exit Function
errHandler:
LogError "LongToString:" & var
End Function

Public Function StringToSingle(ByVal Str As String, Start As Byte) As Single
   StringToSingle = StringToLong(Str, Start) + (Asc(mid$(Str, Start + 4, 1)) / 100)
End Function

Public Function WriteString(valor As String) As String
    WriteString = Chr$(Len(valor)) & valor
End Function

Public Function ReadString(valor As String) As String
    ReadString = mid$(valor, 2, Asc(left$(valor, 1)))
End Function

Public Function ReadStringLength(valor As String) As Integer
    ReadStringLength = Asc(left$(valor, 1))
End Function

Public Function ArrayByteToString(vector() As Byte) As String
    Dim i As Integer
    Dim cadena As String
    
    cadena = ""
    For i = LBound(vector) To UBound(vector)
        cadena = cadena & Chr$(vector(i))
    Next i
    
    ArrayByteToString = cadena
End Function


Public Function StringToLong(ByVal Str As String, ByVal Start As Byte) As Long
    If Len(Str) < Start - 3 Then Exit Function
    
    Dim TempStr As String
    Dim tempstr2 As String
    Dim tempstr3 As String
    'Tomamos los últimos 3 Bytes y convertimos sus valroes ASCII a hexa
    TempStr = Hex$(Asc(mid$(Str, Start + 1, 1)))
    tempstr2 = Hex$(Asc(mid$(Str, Start + 2, 1)))
    tempstr3 = Hex$(Asc(mid$(Str, Start + 3, 1)))
    
    'Nos aseguramos todos midan 2 Bytes (los ceros a la izquierda cuentan por ser Bytes 2, 3 y 4)
    While Len(TempStr) < 2
        TempStr = "0" & TempStr
    Wend
    
    While Len(tempstr2) < 2
        tempstr2 = "0" & tempstr2
    Wend
    
    While Len(tempstr3) < 2
        tempstr3 = "0" & tempstr3
    Wend
    
    'Convertimos a una única cadena hexa
    StringToLong = CLng("&H" & Hex$(Asc(mid$(Str, Start, 1))) & TempStr & tempstr2 & tempstr3)
End Function

Public Function ByteToString(ByVal var As Byte) As String
    ByteToString = Chr$(var)
Exit Function

errHandler:
End Function
Public Function StringToByte(ByVal Str As String, ByVal Start As Integer) As Byte
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    If Len(Str) < Start Then Exit Function
    
    StringToByte = Asc(mid$(Str, Start, 1))
End Function
Public Function ITS(ByVal var As Integer) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    'No aceptamos valores que utilicen los últimos bits, pues los usamos como flag para evitar chr$(0)s
    Dim temp As String
       
    'Convertimos a hexa
    temp = Hex$(var)
    
    'Nos aseguramos tenga 4 Bytes de largo
    While Len(temp) < 4
        temp = "0" & temp
    Wend
    
    'Convertimos a string
    ITS = Chr$(val("&H" & left$(temp, 2))) & Chr$(val("&H" & right$(temp, 2)))
Exit Function

errHandler:

End Function
Public Function STI(ByVal Str As String, ByVal Start As Integer) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    Dim TempStr As String
    
    'Asergurarse sea válido
    If Len(Str) < Start - 1 Then Exit Function
    'Convertimos a hexa el valor ascii del segundo Byte
    TempStr = Hex$(Asc(mid$(Str, Start + 1, 1)))
    
    'Nos aseguramos tenga 2 Bytes (los ceros a la izquierda cuentan por ser el segundo Byte)
    While Len(TempStr) < 2
        TempStr = "0" & TempStr
    Wend
    
    'Convertimos a integer
    STI = val("&H" & Hex$(Asc(mid$(Str, Start, 1))) & TempStr)
End Function

Sub LoginInit()

If EstadoLogin = PantallaCreacion Then
    EstadoLogin = CrearPersonajeSeteado
    Call modCrearPersonaje.crearPersonaje
'ElseIf EstadoLogin = CreandoPersonaje Then
'    EstadoLogin = CrearPersonajeSeteado
End If

End Sub

Sub RecivirInvRefresh(ByVal Rdata As String)
Dim Slot As Byte
If LenB(Rdata) = 0 Then Exit Sub
    Slot = Asc(left$(Rdata, 1))

    If Len(Rdata) > 1 Then
    UserInventory(Slot).OBJIndex = STI(Rdata, 2)
    UserInventory(Slot).Amount = STI(Rdata, 4)
    UserInventory(Slot).Equipped = mid(Rdata, 6, 1)
    UserInventory(Slot).GrhIndex = STI(Rdata, 7)
    UserInventory(Slot).OBJType = Asc(mid(Rdata, 9, 1))
    UserInventory(Slot).MaxHit = STI(Rdata, 10)
    UserInventory(Slot).MinHit = STI(Rdata, 12)
    UserInventory(Slot).MinDef = STI(Rdata, 14)
    UserInventory(Slot).valor = DeCodify(mid(Rdata, 16))
    UserInventory(Slot).Name = objeto(UserInventory(Slot).OBJIndex)
    Else
    UserInventory(Slot).OBJIndex = 0
    UserInventory(Slot).Amount = 0
    UserInventory(Slot).Equipped = 0
    UserInventory(Slot).GrhIndex = 0
    UserInventory(Slot).OBJType = 0
    UserInventory(Slot).MaxHit = 0
    UserInventory(Slot).MinHit = 0
    UserInventory(Slot).MinDef = 0
    UserInventory(Slot).valor = 0
    UserInventory(Slot).Name = "(Nada)"
    End If
    TempStr = ""
    
    If UserInventory(Slot).Equipped = 1 Then
    TempStr = TempStr & "(Eqp)"
    End If
    If UserInventory(Slot).Amount > 0 Then
    TempStr = TempStr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).Name
    Else
    TempStr = TempStr & UserInventory(Slot).Name
    End If
                  
    Exit Sub

End Sub
Private Sub RecivirBancoObj(ByVal Data As String)

    Dim Slot As Byte
    Dim inventario As Inventory
    
    Slot = Asc(left$(Data, 1))
    
    If Len(Data) > 2 Then
        inventario.OBJIndex = STI(Data, 2)
        inventario.Amount = STI(Data, 4)
        inventario.GrhIndex = STI(Data, 6)
        inventario.OBJType = Asc(mid$(Data, 8, 1))
        inventario.MaxHit = STI(Data, 9)
        inventario.MinHit = STI(Data, 11)
        inventario.MinDef = StringToByte(Data, 13)
        inventario.MaxDef = StringToByte(Data, 14)
        inventario.Name = objeto(inventario.OBJIndex)
    Else
        inventario.OBJIndex = 0
        inventario.MinDef = 0
        inventario.MaxDef = 0
        inventario.Amount = 0
        inventario.GrhIndex = 0
        inventario.OBJType = 0
        inventario.MaxHit = 0
        inventario.MinHit = 0
        inventario.Name = "(None)"
    End If
    
    Call frmBancoObj.setSlot(Slot, inventario)
End Sub


Function EsUnAura(ByVal fX As Byte) As Boolean 'Funcion que devuelve True si el fx que recive es'una meditacion.[Parte del sistema de AB de auras]
    Select Case fX
        Case 4, 5, 6
        EsUnAura = True
        Case Else
        EsUnAura = False
        End Select
        Exit Function
End Function
Public Function STI2(ByVal Str As String, ByVal Start As Single) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    Dim TempStr As String
    
    'Asergurarse sea válido
    If Len(Str) < Start - 1 Then Exit Function
    'Convertimos a hexa el valor ascii del segundo Byte
    TempStr = Hex$(Asc(mid$(Str, Start + 1, 1)))
    
    'Nos aseguramos tenga 2 Bytes (los ceros a la izquierda cuentan por ser el segundo Byte)
    While Len(TempStr) < 2
        TempStr = "0" & TempStr
    Wend
    
    'Convertimos a integer
    STI2 = val("&H" & Hex$(Asc(mid$(Str, Start, 1))) & TempStr)
    
    'Vemos si el primer Byte era cero
    If STI2 And &H8000 Then _
        STI2 = STI2 Xor &H8001
    
    'Si el segundo Byte era cero
    If STI2 And &H4000 Then _
        STI2 = STI2 Xor &H4000
End Function


Public Sub AgregarMensaje(numero As Integer)
Dim Tempvar() As String
Tempvar = Split(mensaje(numero), "~")
Call AddtoRichTextBox(frmConsola.ConsolaFlotante, Tempvar(0), Int(Tempvar(1)), Int(Tempvar(2)), Int(Tempvar(3)), Int(Tempvar(4)), Int(Tempvar(5)))
End Sub


Public Sub Intervalos(cadena As String)
Dim Interval As Variant
Interval = Split(cadena, "-")

UserStats(SlotStats).IntervaloPegar = val(Replace(Interval(0), ",", ".")) + 0.000101
UserStats(SlotStats).IntervaloLanzarMagias = val(Replace(Interval(1), ",", ".")) + 0.000111
UserStats(SlotStats).IntervaloLanzarFlechas = val(Replace(Interval(2), ",", ".")) + 0.000011
UserStats(SlotStats).intervaloNoChupU = val(Replace(Interval(3), ",", ".")) + 0.0011
UserStats(SlotStats).IntervaloNoChupClick = val(Replace(Interval(4), ",", ".")) + 0.000011

IntervaloPegarB = CryptStr(val(Replace(Interval(0), ",", ".")) + 0.000101, 0)
IntervaloLanzarMagiasB = CryptStr(val(Replace(Interval(1), ",", ".")) + 0.000111, 0)
IntervaloLanzarFlechasB = CryptStr(val(Replace(Interval(2), ",", ".")) + 0.000011, 0)
intervaloNoChupUB = CryptStr(val(Replace(Interval(3), ",", ".")) + 0.0011, 0)
IntervaloNoChupClickB = CryptStr(val(Replace(Interval(4), ",", ".")) + 0.000011, 0)

'Anticheat
UserStats(SlotStats).IntervaloSolapaLanzar = val(Interval(5))
UserStats(SlotStats).IntervaloSolapaLanzarSuper = val(Interval(6))
UserStats(SlotStats).IntervaloHechizoLanzar = val(Interval(7))
UserStats(SlotStats).IntervaloHechizoLanzarSuper = val(Interval(8))
UserStats(SlotStats).UmbralAlerta = val(Interval(9))

IntervaloSolapaLanzarB = CryptStr(Interval(5), 0)
IntervaloSolapaLanzarSuperB = CryptStr(Interval(6), 0)
IntervaloHechizoLanzarB = CryptStr(Interval(7), 0)
IntervaloHechizoLanzarSuperB = CryptStr(Interval(8), 0)
UmbralAlertaB = CryptStr(Interval(9), 0)
End Sub

Private Sub crearPersonaje(ByRef Rdata As String)
    Dim CharIndex As Integer
    Dim fX As Byte
    Dim body As Integer
    Dim Head As Integer
    Dim heading As Byte
    Dim X As Integer
    Dim Y As Integer
    Dim weapon As Byte
    Dim escudo As Byte
    Dim casco As Byte
    Dim tieneClan As Byte
    Dim privilegio As Byte
    Dim Nombre As String
    Dim Alineacion As eAlineaciones
    Dim pelo As Integer
    Dim barba As Integer
    Dim ropaInterior As Integer
    
    CharIndex = STI(Rdata, 1)
    fX = StringToByte(Rdata, 3)
    body = STI(Rdata, 4)
    Head = STI(Rdata, 6)
    heading = StringToByte(Rdata, 8)
    X = STI(Rdata, 9)
    Y = STI(Rdata, 11)
    weapon = StringToByte(Rdata, 13)
    escudo = StringToByte(Rdata, 14)
    casco = StringToByte(Rdata, 15)
    privilegio = StringToByte(Rdata, 17)
    Alineacion = StringToByte(Rdata, 16)
    tieneClan = StringToByte(Rdata, 18)
    pelo = STI(Rdata, 19)
    barba = STI(Rdata, 21)
    ropaInterior = STI(Rdata, 23)
    Nombre = mid$(Rdata, 25)
    
    CharList(CharIndex).Alineacion = Alineacion
    
    Call modPersonaje.actualizarNick(CharList(CharIndex), Nombre)
        
    CharList(CharIndex).priv = privilegio
    CharList(CharIndex).pelo = pelo
    CharList(CharIndex).barba = barba
    CharList(CharIndex).ropaInterior = ropaInterior
    
    Call MakeChar(CharIndex, body, Head, heading, X, Y, weapon, escudo, casco)
    
    Call setColorNombre(CharList(CharIndex))
    
    SetCharacterFx CharIndex, fX, 999
    
    If CharIndex = UserCharIndex Then
        Call actualizarMapaNombre
        Call rm2a
        Cachear_Tiles = True
    End If
End Sub
