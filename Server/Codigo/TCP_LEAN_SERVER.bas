Attribute VB_Name = "TCP_LEAN"
Option Explicit
'Modulo desarrollado por Leandro (Wizard_II)
'Contacto: Lean_ar@hotmail.com
Rem Enums de Comandos y Paquetes.
Public Enum Paquetes
    MensajeSimple = 1 'Mensajes simples q solo mandan "Y259" x ejemplo...
    MensajeCompuesto = 2 'Mensajes q manden alguna data para reemplazar en cliente y mostrar en consola
    PrenderFogata = 3 'Prende la fogata no ay mucho q explicar
    MostrarCartel = 4 'Muestra carteles cuando se clikean
    MensajeForo = 5 'Ves contenido de foro
    MensajeForo2 = 6 'Ves mensaje en foro
    WavSnd = 7 'reproduccion de WAVS
    pNpcInventory = 8 'Actualiza el inventario de Npc comerciante
    TransOK = 9 'TRansacion ok
    Pausa = 10 'BKW:> Pausa en el WS
    pIniciarComercioNPC = 11 'Se inicia comercio con npc, abrimos ventana
    loguea = 12 'El user loguea, le decimos q muestre frmmain
    VeUser = 13 'ClickeHeading.EASTe un usuario
    VeObjeto = 14 'clikeHeading.EASTe un objeto
    VeNpc = 15 'ClikeHeading.EASTe Npc
    DescNpc = 16 'el npc dice la desc guardada en cliente
    DescNpc2 = 17 'Npc dice Desc q no ta guardada
    ' ModCeguera = 18 'Modifica Ciego-NoCiego
    ' ModEstupidez = 19 'Modifica Estupido-Noestupido
    BloquearTile = 20 'Bloquea un Tile
    pEnviarSpawnList = 21 'muestra la lista de spawn con el /CC
    EquiparItem = 22 'Equipa el Slot
    DesequiparItem = 23 'Desequipa el item del slot
    BorrarObj = 24 'Borra un objeto
    CrearObjeto = 25 'Crea un objeto
    ApuntarProyectil = 26 'Muestra la mira de arco
    ApuntarTrb = 27 'Muestra la mira en trabajos como minar talar
    EnviarArmasConstruibles = 28 'Muestra lista de armas de herreria
    EnviarObjConstruibles = 29 'Muestra lista de objetos en carpinteria
    EnviarArmadurasConstruibles = 30 'Muestra lista de armaduras
    ShowCarp = 31 'Muetra formulario carpintero
    InitComUsu = 32 'inicia comercio con otro usuario
    ComUsuInv = 33 'mandamos inventario del user
    FinComUsuOk = 34 'Termina comercio con user
    InitBanco = 35 'Inicia bovedeo
    EnviarBancoObj = 36 'Nos manda el inventario del banco
    BancoOk = 37 'cierra banco
    PeaceSolRequest = 38 'Pide Sol paz
    EnviarPeaceProp = 39 'Manda sol paz
    PeticionClan = 40 'Nos devuelve peticion de clan
    EnviarCharInfo = 41 'Nos manda info de un char en clan
    EnviarLeaderInfo = 42 'nos manda la info para el lider, numero de miembros etc
    EnviarGuildsList = 43 'lista de clanes
    EnviarGuildDetails = 44 'Detalles de 1 clan
    HechizoFX = 45 'Manda FX,Loop y Wav
    MensajeTalk = 46 'Mensaje consola blanco
    MensajeSpell = 47 'Mensaje Consola para hechizos
    MensajeFight = 48 'Mensaje Consola en rojo
    mensajeinfo = 49 'Mensaje consola en Fonttype_info
    CambiarHechizo = 50 'Cambiamos hechizo
    pCrearNPC = 51 'Crea npc
    ChangeNpc = 52 'Modifica el npc
    BorrarNpc = 53 'Borra npc
    MoveChar = 54 'Mueve 1 char
    EnviarNpclst = 55 'Envia la lista de npcs al entenador
    COMBRechEsc = 56 'Rechazamos con escudo
    COMBNpcHIT = 57 'El Npc nos pega
    COMBMuereUser = 58 'Morimos
    COMBNpcFalla = 59 'Falla el npc
    COMBUserFalla = 60 'Fallamos el golpe
    COMBEnemEscu = 61 'Un user nos rechaza con escu
    SangraUser = 62 'sangra tal user
    COMBUserImpcNpc = 63 'Impactamos Npc
    COMBEnemFalla = 64 'Nos falla 1 enemigo
    COMBEnemHitUs = 65 'Nos pega 1 enemigo
    COMBUserHITUser = 66 'Le pegamos a un gil COMELA PETE
    Navega = 67 'Avisamos q navega
    AuraFx = 68 'Mostramos fx de meditar
    Meditando = 69 'Esta paralizado
    NoParalizado = 70 'No paralizado
    Paralizado2 = 71 'Paralizado con clan
    NoParalizado2 = 72 'Paralizado sin clan
    Invisible = 73 'invisible
    Visible = 74 'Visible
    pChangeUserChar = 75 'modificamos el char
    leveLUp = 76 'Subiste de lvl
    SendSkills = 77 'mandamos skills
    SendFama = 78 'Mandamos fama para ESTADISTICAS
    SendAtributos = 79 'Mandamos atributos
    MiniEst = 80                    ' Mandamos minest tiempo en carcel etc
    BorrarUser = 81                 ' Borramos un user
    CrearChar = 82                  ' Creamos un Persnaje
    EnviarPos = 83                  ' Enviamos pos
    EnviarStat = 85                 ' Enviamos todas las stats
    EnviarF = 86                    ' Enviamos fuerza
    EnviarA = 87                    ' Enviamos Agilidad
    EnviarOro = 88                  ' Enviamos oro
    EnviarHP = 89                   ' Enviamos Vida
    EnviarMP = 90                   ' Enviamos Mana
    EnviarST = 91                   ' Enviamos stamina
    EnviarEXP = 92                  ' Enviamos EXP
    EnviarSYM = 93                  ' Enviamos Stamina y mana
    EnviarSYH = 94                  ' Enviamos Estamina y vida
    EnviarFA = 95                   ' Enviamos Fuerza y agi
    EnviarHYS = 96                  ' Envimoas Hambre y sed
    QDL = 97                        ' Borramos el mensaje de tal char
    MDescansar = 98                 ' Descansa no descansa
    ChangeMap = 99                  ' Cambiamos mapa
    ChangeMusic = 100               ' Cambiamos musica
    QTDL = 101                      ' Borramos Toods los mensajes
    IndiceChar = 102                ' Nos da el Indice del char
    mbox = 103                      ' MSGBOX
    lluvia = 104                    ' Llueve no llueve
    SOSAddItem = 105                ' Agregamos un item al SOS
    SOSViewList = 106               ' Miramos la lista sos
    MensajeServer = 107             ' Mensaje q empieza con "SERVIDOR>" en verde
    MensajeGMSG = 108               ' Mensaje a gms
    UserTalk = 109                  ' Tal usuario dice tal boludes
    UserShout = 110                 ' Tal usuario grita tal boludes
    UserWhisper = 111               ' Tal usuario me susurra tal idiotes
    TurnToNORTH = 112               ' Tal mira al norte
    TurnToSOUTH = 113               ' Tal mira al sur
    TurnToWEST = 114                ' Tal mira al oeste
    TurnToEAST = 115                ' Tal mira al este
    FinComOk = 116                  ' Termina el comercio
    FinBanOk = 117                  ' Termina el banco
    SndDados = 118                  ' Enviamos resultados de dados
    ShowHerreriaForm = 119          ' Mostramos la herreria
    ' EnviarUI = 120                ' Enviamos el Userindex del usuario
    EnviarGuildNews = 121           ' Enviamos GuildNews
    InvRefresh = 122                ' Refrescamos el inventario
    initGuildFundation = 123        ' Iniciamos la fundacion
    MensajeClan1 = 124              ' Mensaje a clan de 1 color
    MensajeClan2 = 125              ' Mensaje a clan de otr para diferenciar entre vivo y muerto
    SaidMagicWords = 126            ' Palabras magicas de un hechizo
    MoveNpc = 127                   ' Movemos el npc
    pEnviarNpcInvBySlot = 128       ' Mandamos 1 slot del inventario del npc
    mTransError = 129               ' Error transparente una boludes de silver
    CrearObjetoInicio = 130         ' Para resolver el error de colgar en los mapas donde hay muchos items. Tendria que solucionar con este nuevo tcp pero por las dudas... Marche
    MensajeSimple2 = 131            ' de 255 para arriba
    Noche = 132
    SegOFF = 133
    SegOn = 134
    ' nieva = 135
    DejaDeTrabaja = 136
    TXA = 137
    mBox2 = 138
    FXh = 139
    FundoParty = 140
    Pni = 141
    Integranteparty = 142
    OnParty = 143
    mest = 144
    AnimGolpe = 145                 'El Yind - Animaciones de golpe y esucdo
    AnimEscu = 146
    CFXH = 147
    MensajeGuild = 148
    ClickObjeto = 149
    LISTUSU = 150
    traba = 151
    UserTalkDead = 152
    TiempoReto = 153
    Pang = 154
    TalkQuest = 155
    pChangeUserCharCasco = 156      ' Cambia solo el casco
    pChangeUserCharEscudo = 157     ' Cambia solo el escudo
    pChangeUserCharArmadura = 158   ' Cambia solo la armadura
    pChangeUserCharArma = 159       ' Cambia solo el arma
    EnCentinelaPa = 160
    TXAII = 161
    EnviarStatsBasicas = 162        ' Echo para mandar los stats que se modifican en el gametimer (antes mandaba la experiencia, el nivel, y asi)
    MensajeArmadas = 163
    MensajeCaos = 164
    EmpiezaTrabajo = 165
    mensajeGlobal = 166             ' Mensaje del chat global
    PartyAcomodarS = 167
    PPI = 168                       ' Pedir ingreso a la party
    ppe = 169                       ' Expulso a alguien del party
    Sefuedeparty = 170
    MensajeBoveda = 171
    IniciarAutoupdater = 172
    EstaEnvenenado = 173
    ' ActualizarEstado = 174
    MoverMuerto = 175               ' Mueve el muerto y al vivo con animacion para que no paresca que se laguea
    ocultar = 176                   ' Es distinto la invisibilidad que el ocultar
    Desocultar = 177                ' por ende utiliza paquetes distintos
    pNpcActualizarPrecios = 178
    ActualizaNick = 179             ' Actualiza el nick del personaje (nombre, clan, alineacion, privilegios)
    ActualizaCantidadItem = 180
    ActualizarAreaUser = 181
    ActualizarAreaNpc = 182
    CambiarHeadingNpc = 183
    BorrarArea = 184
    '185, 186, 187,188,189
    Pong = 190
    SonidoTomarPociones = 191       ' Nos ahorramos un byte en decirle que sonido es ya que esta en el cliente
    infoLogin = 192
    
    EnviarLeaderInfoSolicitudes = 193
    EnviarLeaderInfoMiembros = 194
    EnviarLeaderInfoNovedades = 195
    
    InfoAdminEventos = 196          ' Lista de eventos automaticos
    MensajeAdminEventos = 197       ' Mensaje al administrador remoto
    InfoEventoAdminEventos = 198    ' Toda la info de un evento
    
    transferenciaIniciar = 199      ' Se solicita que comience la transferencia
    transferenciaOK = 200           ' Paso de la transferencia se hizo ok
    infoClan = 201                  ' Informacion del clan
    checkMem = 202                  ' Chequea que no se este editando la memoria
    DuracionAributos = 203          ' Informa el tiempo que le queda de dopa.
    AnguloNPC = 204
End Enum

Public Enum cPaquetes
    comandos = 1
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
    'preConnect = 34
    MemberInfo = 35
    DeclararWar = 36
    MEast = 37
    ChangeItemsSlot = 38
    OfrecerComUsu = 39
    MOesteM = 40
    ComandosConse = 41
    GuildDetail = 42
    Expulsarparty = 43
   ' revivirAutomaticamente = 44
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
    'CrearRetoD = 56
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
    'LanzarHechizo = 90
    CrearParty = 91
    InfoHechizo = 92
    ' ArrojarDados = 93
    RechazarGuild = 94
    iParty = 95
    ccParty = 96
    MCombate = 97
    'ConnectPj = 98
    MirarOeste = 99
    GuildSol = 100
    MirarEste = 101
    'Spawn = 102
    GuildDSend = 103
    CreatePj = 104
    URLChange = 105
    SkillSetOcultar = 106
    ExitOk = 107
    Pegar = 108
        
    obtClanMiembros = 109
    obtClanSolicitudes = 110
    obtClanNews = 111

    infoTransferencia = 112         ' Informacion de transferencia de usuario
    respuestaMemCheck = 113         ' Respuesta al anticheat de chequeo de memoria
    NODO_INFO_HASH = 250            ' FALSO MENSAJE DEL CLIENTE. Esta data en realidad me la envia el nodo.
End Enum

Public Enum CmdUsers
    online = 1 'ok
    CCOMERCIAR = 2 'ok
    BOVEDA = 3 'ok
    CMEDITAR = 4 'ok
    ENLISTAR = 5 'ok
    informacion = 6 'ok
    DESINVOCAR = 7 'Tengo q programarlo(No lo hice para no agregar comandos al pedo en tds)
    Resucitar = 8 'ok
    CURAR = 9 'ok
    CDESCANSAR = 10 'ok
    ENTRENAR = 11 'ok
    ACOMPAÑAR = 12 'ok
    QUIETO = 13 'ok
    CBALANCE = 14 'ok
    SALIRCLAN = 15 'ok
    FUNDARCLAN = 16 'ok
    ONLINECLAN = 17 'ok
    DONDECLAN = 18 'Tengo q programarlo(No lo hice para no agregar comandos al pedo en tds)
    CSALIR = 19 'ok
    CAYUDA = 20 'ok
    EST = 21 'ok
    RECOMPENSA = 22 'ok
    CMOTD = 23 'ok
    GM = 24 'ok
    MOVER = 25 'ok
    'ONLINEP = 26 'ok
    Retirar = 27 'ok
    Depositar = 28 'ok
    PASARORO = 29 'ok
    APOSTAR = 30 'No lo puse porq no lo uso...
    VOTO = 31 'ok
    PASSWD = 32 'ok
    CDESC = 33 'ok
    BUG = 34 'Al pedo!!!!!!!!!!!!!!
    CMSG = 35 'ok
    DENUNCIAR = 95
    centinela = 26
    cheque = 97
    Activar = 98
    Ping = 99
    RetarS = 100
    AcomodarPorcentajesDeParty = 101
    Retirartodo = 102
    DepositarTodo = 103
    Abandonar = 104
    Fianza = 105
    PartyPorcecntaje = 106
    Penas = 107
    pmsg = 108
    disolverclan = 109
    ReanudarClan = 110
    Participar = 111 'Comando para participar en los eventos
    aceptar = 112 'Un usuario acepta participar en un evento con un usuario
    eventosInfo = 113 ' Obtener informacion de los eventos en curso o de un evento.
    decirEnTorneo = 114 ' Cuando finaliza un torneo el ganador puede dar un mensaje al pueblo
    minutoEnTorneo = 115 'Pide minuto
    rechazar = 116 ' Rechaza una propuesta de un usuario para participar de un evento
    tiempo = 117    ' Tiempo restante para jugar TDSF. SOLO VALIDO EN TDSF
End Enum

Public Enum CmdConse
    Hora = 36 'ok
    TELEPLOC = 37 'ok
    SHOW_SOS = 38 'ok
    CINVISIBLE = 39 'taTa
    PANELGM = 40 'FALTA!
    CTRABAJANDO = 41
    CREM = 42 'ok
    TELEP = 43 'ok
    NENE = 44 'Pa q?!?!
    donde = 45 'ok
    IRA = 46 'ok
    LISTUSU = 47
    
    ONLINEMAP = 53 'ok
    
    carcel = 54 'ok
        
    GMSG = 95 'ok
    
    Penas = 106
    
    RMSG = 107 'ok
End Enum

Public Enum CmdSemi
    CINFO = 47 'ok
    INV = 48 'ok
    BOV = 49 'ok
    CSKILLS = 50 'ok
    CREVIVIR = 51 'ok
    IP2NICK = 52 'ok
    
    
    PERDON = 55 'ok
    CECHAR = 56 'ok
    CBAN = 57 'ok
    CUNBAN = 58 'ok
    CSUM = 59 'ok
    RESETINV = 60 'ok
    RMSG = 61 'ok
    NICK2IP = 62 'ok
    ONLINEGM = 63 'ok
    cc = 64 'ok
    LIMPIAR = 65 'ok
    SEGUIR = 66 'ok
    
    ejecutar = 96
    Qtalk = 97
    AUCO = 98
    
    LargarCentinelas = 107
    Aotromap = 109
    Secaeest = 108
    AntiHpts = 111
    
    ROSG = 113
    MaxLevelMap = 114
    MinLevelMap = 115
    MapaFrio = 116
    
    HabilitarRobo = 117
    
    Spawn = 118
    
    InfoMap = 119
    OnlyCiuda = 120
    OnlyCrimi = 121
    LimiteUserMap = 122
    Cname = 123
    OnlyCaos = 124
    OnlyArmada = 125
    
    'Eventos automaticos
    CrearEvento = 126
    ObtenerEventos = 127
    obtenerInfoEvento = 128
    cancelarEvento = 129
    
    publicarEvento = 133
    inscribirEvento = 134
    
    DT = 130
    BLOQ = 131
    CT = 132
End Enum

Public Enum CmdDios
    MASSDEST = 67 'ok
    ' BANIPLIST = 68 'ok
    ' BANIPRELOAD = 69 'ok
    PASSDAY = 70 'ok
    
    CDEST = 72 'ok
    '73
    MATA = 74 'ok
    MASSKILL = 75 'ok
    'MOTDCAMBIA = 76 'NO SIRVE
    ACC = 77 'ok
    'NAVE = 78 ' NO SIRVE
    APAGAR = 79 'ok
    'CDOBACKUP = 80 'ok
    BORRAR_SOS = 82 'ok
    SHOW_INT = 83 'NO SIRVE
    CLLUVIA = 84 'ok
    '85
    CTRIGGER = 86 'ok
    ' CBANIP = 87 'ok
    ' UNBANIP = 88 'ok
    LastIP = 89 'ok
    RACC = 90 'ok
    'CONDEN = 91 'ok
    'RAJAR = 92 'ok
    RAJARCLAN = 93 'ok
    'CMod = 94 'ok
    'CI = 95
    ZONEST = 96 'Cambia el mapa de zona segura a zona insegura para quest.
    RETEST = 97
    CHATEST = 98
    'TCPESTATS = 99
    ECHARTODOSPJS = 100
    NickMac = 101
    banMac = 102
    unBanMac = 103
    'ReloadServer = 104
    'Cname = 105
    Retorings = 106
    Habilitar = 107
    AceptarConsejo = 108
    ExpulsarConsejo = 109
    ModoRol = 110
    CTE = 112 'Para crear telepors sin el grafico
    CheqCli = 113
    VerPongs = 114
    CNameClan = 115
    EcharClan = 116
    capturarPantalla = 117 ' Captura la pantalla del usuario
    consultarMem = 118 ' Se fuerza el chequeo de la memoria del usuario
End Enum

Public Enum cmdAdmin
    GRABAR = 1
    APAGAR = 2
    BackUp = 3
    CONDEN = 4
    RAJAR = 5
    CMod = 6
    CI = 7
    TCPEST = 8
    ReloadServer = 9
    BANIPRELOAD = 10
End Enum

Public Function GenCrc(NumeroPaquete As Long, MinimoPaquete As Byte, paquete As String) As Byte
    GenCrc = (((NumeroPaquete ^ 2.33) + Len(paquete) + MinimoPaquete) Mod 249) + 1
End Function


Public Sub sHandleData(ByVal rdata As String, ByVal UserIndex As Integer)
'Recibimos el Paquete principal.
'Segun la cabecera lo mandamos inteligentemente a la funcion
'a la q debe mandarse, y revisamos el CRC
Dim tempbyte As Byte

' Chequeamos
If Not (Asc(Left(rdata, 1)) = GenCrc(UserList(UserIndex).PacketNumber, UserList(UserIndex).MinPacketNumber, rdata)) Then
    
    If UserList(UserIndex).flags.UserLogged = True Then
        LogHack ("Paquete falso " & UserList(UserIndex).Name & " " & UserList(UserIndex).MacAddress)
    Else
        LogHack ("Paquete falso. UserIndex " & UserIndex)
    End If
        
    EnviarPaquete Paquetes.IniciarAutoupdater, "", UserIndex, ToIndex
    
    ' Cerramos ya!!!!
    If Not CloseSocket(UserIndex) Then Call LogError("sHandleData")
        
    Exit Sub
End If

tempbyte = Asc(mid(rdata, 2, 1))

UserList(UserIndex).PacketNumber = (UserList(UserIndex).PacketNumber + tempbyte * Asc(Left$(rdata, 1))) Mod 5003

UserList(UserIndex).Counters.IdleCount = 0
 
If UserList(UserIndex).flags.UserLogged = True Then
    Select Case tempbyte
        Case cPaquetes.comandos  'Comandos de usuario
            ProcesarComando0 Asc(mid$(rdata, 3, 1)), UserIndex, Right$(rdata, Len(rdata) - 3)
        Case cPaquetes.ComandosConse  'Comandos de Consejero
            ProcesarComando1 Asc(mid$(rdata, 3, 1)), UserIndex, Right$(rdata, Len(rdata) - 3)
        Case cPaquetes.ComandosSemi  'Comandos de semidios
            ProcesarComando2 Asc(mid$(rdata, 3, 1)), UserIndex, Right$(rdata, Len(rdata) - 3)
        Case cPaquetes.ComandosDios  'Comandos de Dios
            ProcesarComando3 Asc(mid$(rdata, 3, 1)), UserIndex, Right$(rdata, Len(rdata) - 3)
        Case cPaquetes.ComandosAdmin  'Comandos de administradores
            ProcesarComandoAdmin Asc(mid$(rdata, 3, 1)), UserIndex, Right$(rdata, Len(rdata) - 3)
        Case Else 'Paquetes comunes
        '    Debug.Print "Procesamos el Paquete:" & "Nro:" & TempByte & " Argumentos:" & Right$(rdata, Len(rdata) - 1)
            ProcesarPaqueteON tempbyte, UserIndex, Right$(rdata, Len(rdata) - 2)
    End Select
Else
    ProcesarPaqueteOFF tempbyte, UserIndex, Right$(rdata, Len(rdata) - 2)
End If
End Sub
Private Sub ProcesarPaqueteON(ByVal numero As Byte, ByVal UserIndex As Integer, Optional anexo As String)

Dim tempbyte As Byte
Dim tempbyte2 As Byte
Dim tempLong As Long
Dim TempInt As Integer
Dim tempInt2 As Integer
Dim tempSingle As Single


If ProfilePaquetes Then
    Logs.logProfilePaquete ("O" & numero)
End If
            
Select Case numero
    Case cPaquetes.Hablar


        If Len(anexo) > 500 Then
            EnviarPaquete Paquetes.mensajeinfo, UserList(UserIndex).Name & "-> Spamea", UserIndex, ToAdmins
            Exit Sub
        End If
        
        'Me fijo si tiene el chat activado
        If mid(anexo, 1, 1) = "." And UserList(UserIndex).Stats.GlobAl = 2 Then
         If UserList(UserIndex).flags.Muerto = 1 Then
            EnviarPaquete Paquetes.mensajeinfo, "No puedes hablar por global si estas muerto.", UserIndex, ToIndex
         Else
            If charlageneral = False Then
                EnviarPaquete Paquetes.MensajeSimple2, Chr$(129), UserIndex, ToIndex
            Else
                If UserList(UserIndex).Stats.ELV < 10 Then
                    EnviarPaquete Paquetes.mensajeinfo, "Debes ser nivel 10 o superior.", UserIndex, ToIndex
                Else
                    Call modChatGlobal.enviarMensaje(UserList(UserIndex), anexo)
                End If
            End If
         End If
        Else
            'Guarda en Log
            If UserList(UserIndex).flags.Privilegios > 0 Then
                Call LogGM(UserList(UserIndex).id, "Dijo: " & anexo)
            End If

            If UserList(UserIndex).flags.Oculto > 0 Then
                UserList(UserIndex).flags.Oculto = 0
                UserList(UserIndex).flags.Invisible = 0
                EnviarPaquete Paquetes.Desocultar, ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToMap
                EnviarPaquete Paquetes.MensajeSimple2, Chr$(288 - 255), UserIndex, ToIndex
            End If

            If UserList(UserIndex).flags.Muerto = 1 Then
                EnviarPaquete Paquetes.UserTalkDead, anexo & ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToDeadArea, UserList(UserIndex).pos.map
                EnviarPaquete Paquetes.UserTalkDead, anexo & ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToAdminsArea, UserList(UserIndex).pos.map
            Else
                'Quitar el dialogo
                If Len(anexo) = 1 And anexo = " " Then
                    EnviarPaquete Paquetes.QDL, ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToMap, UserList(UserIndex).pos.map
                Else
                    EnviarPaquete Paquetes.UserTalk, anexo & ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToPCArea, UserList(UserIndex).pos.map
                End If
            End If
        End If

           Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.Gritar
    
        If UserList(UserIndex).flags.Privilegios = 1 Then
            Call LogGM(UserList(UserIndex).id, "Grito" & anexo)
        End If
        
        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        
        If UserList(UserIndex).flags.Oculto > 0 Then
            UserList(UserIndex).flags.Oculto = 0
            UserList(UserIndex).flags.Invisible = 0
            EnviarPaquete Paquetes.Desocultar, ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToMap
            EnviarPaquete Paquetes.MensajeSimple2, Chr$(288 - 255), UserIndex, ToIndex
        End If
            
        If LenB(anexo) = 0 Then
        EnviarPaquete Paquetes.QDL, ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToMap, UserList(UserIndex).pos.map
        Else
        EnviarPaquete Paquetes.UserShout, anexo & ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToPCArea, UserList(UserIndex).pos.map
        End If
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.Susurrar
        Dim tempstr As String
    
        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        
        If UserList(UserIndex).flags.Privilegios >= 1 Then
            Call LogGM(UserList(UserIndex).id, "SUSORRO: " & anexo)
        End If
        
        tempstr = ReadField(1, anexo, 44)
        TempInt = NameIndex(tempstr)
        
        If TempInt > 0 Then
            'Le digo que esta offline siempre si es un GM
            If UserList(TempInt).flags.Privilegios > 0 Then
                EnviarPaquete Paquetes.MensajeSimple2, Chr(290 - 255), UserIndex, ToIndex
                Exit Sub
            End If
            
            If Not EstaPCarea(UserIndex, TempInt) Then
                'Le digo que esta muy lejos
                EnviarPaquete Paquetes.MensajeSimple2, Chr(290 - 255), UserIndex, ToIndex
                Exit Sub
            End If
            
            'Esta en el area de vision?
            EnviarPaquete Paquetes.UserWhisper, Right(anexo, Len(anexo) - Len(tempstr) - 1) & ITS(UserList(UserIndex).Char.charIndex), TempInt, ToIndex, 0
            EnviarPaquete Paquetes.UserWhisper, Right(anexo, Len(anexo) - Len(tempstr) - 1) & ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToIndex, 0
            'Le aviso a los GMS
            EnviarPaquete Paquetes.UserWhisper, "a " & tempstr & "> " & Right(anexo, Len(anexo) - Len(tempstr) - 1) & ITS(UserList(UserIndex).Char.charIndex), TempInt, ToAdminsArea, 0
         Else
            'Le digo que esta offline
            EnviarPaquete Paquetes.MensajeSimple2, Chr(290 - 255), UserIndex, ToIndex
         End If
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.MEast
        Call Moverse(UserList(UserIndex))
        Call moverPersonajeHacia(UserList(UserIndex), UserList(UserIndex).pos.x + 1, UserList(UserIndex).pos.y, eHeading.EAST)
        Call posMoverse(UserList(UserIndex))
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.MWest
        Call Moverse(UserList(UserIndex))
        Call moverPersonajeHacia(UserList(UserIndex), UserList(UserIndex).pos.x - 1, UserList(UserIndex).pos.y, eHeading.WEST)
        Call posMoverse(UserList(UserIndex))
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.MNorth
        Call Moverse(UserList(UserIndex))
        Call moverPersonajeHacia(UserList(UserIndex), UserList(UserIndex).pos.x, UserList(UserIndex).pos.y - 1, eHeading.NORTH)
        Call posMoverse(UserList(UserIndex))
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.MSouth
        Call Moverse(UserList(UserIndex))
        Call moverPersonajeHacia(UserList(UserIndex), UserList(UserIndex).pos.x, UserList(UserIndex).pos.y + 1, eHeading.SOUTH)
        Call posMoverse(UserList(UserIndex))
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.PostearForo
        LogHack (UserList(UserIndex).Name & "tiene cheat. BANEAR T0. IP: " & HelperIP.longToIP(UserList(UserIndex).ip))
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.Usar
        tempbyte = Asc(mid$(anexo, 1, 1))
        tempSingle = StringToSingle(anexo, 3)
        tempbyte2 = mid$(anexo, 2, 1)
                
        If UserList(UserIndex).Invent.Object(tempbyte).ObjIndex = 0 Or UserList(UserIndex).flags.Meditando Then Exit Sub
        
        UseInvItem UserIndex, tempbyte, tempbyte2, tempSingle
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.Agarrar
        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        If UserList(UserIndex).flags.Privilegios = 1 Then EnviarPaquete Paquetes.MensajeSimple, Chr(220), UserIndex, ToIndex: Exit Sub
        GetObj UserIndex
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.Tirar
        ' Chequeamos si no está en un estado invalido para tirar un objeto
        If UserList(UserIndex).flags.Comerciando Then Exit Sub
        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        If UserList(UserIndex).flags.Privilegios = 1 Then Exit Sub
        If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
        
        ' Intervalo para que no llenen el mapa de oro
        If Not IntervaloPermiteAtacar(UserIndex, True) Then
            Exit Sub
        End If
            
        ' Obtenemos Parametros
        tempbyte = Asc(anexo)
        tempLong = Int(DeCodify(Right$(anexo, Len(anexo) - 1)))
        
        If tempLong = 0 Then Exit Sub

        If tempbyte = FLAGORO Or tempbyte = 254 Then
            Call Acciones.TirarOroAlSuelo(UserList(UserIndex), tempLong)
            
            EnviarPaquete Paquetes.EnviarOro, Codify(UserList(UserIndex).Stats.GLD), UserIndex, ToIndex, 0
        Else
            If UserList(UserIndex).Invent.Object(tempbyte).ObjIndex = 0 Then Exit Sub
            DropObj UserList(UserIndex), tempbyte, tempLong, UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y
        End If
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
   ' Case cPaquetes.pLanzarHechizo
   '     If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
   '     UserList(UserIndex).flags.Hechizo = Asc(Anexo)
   ' Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.Pegar
        Call modCombate.usuarioPegar(UserList(UserIndex), StringToSingle(anexo, 1))
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.MCombate
        UserList(UserIndex).flags.modoCombate = Not UserList(UserIndex).flags.modoCombate
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.Seguro
        UserList(UserIndex).flags.Seguro = Not UserList(UserIndex).flags.Seguro
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.Drag
        UserList(UserIndex).flags.PermitirDragAndDrop = Not UserList(UserIndex).flags.PermitirDragAndDrop
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.MirarNorte
        UserList(UserIndex).Char.heading = eHeading.NORTH
        EnviarPaquete Paquetes.TurnToNORTH, ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToArea
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.MirarSur
        UserList(UserIndex).Char.heading = eHeading.SOUTH
        EnviarPaquete Paquetes.TurnToSOUTH, ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToArea
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.MirarEste
        UserList(UserIndex).Char.heading = eHeading.EAST
        EnviarPaquete Paquetes.TurnToEAST, ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToArea
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.MirarOeste
        UserList(UserIndex).Char.heading = eHeading.WEST
        EnviarPaquete Paquetes.TurnToWEST, ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToArea
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.comprar
        ' Comprar un Objeto en un Comerciante
        If UserList(UserIndex).flags.Muerto Then Exit Sub

        ' Â¿Tiene a una criatura seleccionada?
        If UserList(UserIndex).flags.TargetNPC = 0 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(4), UserIndex
            Exit Sub
        End If
        
        ' Â¿Es un comerciante?
        If NpcList(UserList(UserIndex).flags.TargetNPC).Comercia = 0 Then
            EnviarPaquete Paquetes.DescNpc2, Chr$(3) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
            Exit Sub
        End If
        
        'Â¿Esta demasiado lejos?
        If distancia(UserList(UserIndex).pos, NpcList(UserList(UserIndex).flags.TargetNPC).pos) > 5 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(7), UserIndex
            Exit Sub
        End If
        
        ' Compramos un objeto (La Criatura lo vende)
        NPCVentaItem UserIndex, Asc(Left$(anexo, 1)), DeCodify(Right$(anexo, Len(anexo) - 1)), UserList(UserIndex).flags.TargetNPC
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.Vender
    
        Call Comercio.personajeVendeObjetoACriatura(UserList(UserIndex), anexo)
        
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.ComUsuOk
        AceptarComercioUsu UserIndex
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.ComOk
        UserList(UserIndex).flags.Comerciando = False
        EnviarPaquete Paquetes.FinComOk, "", UserIndex
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.Retirar
        If UserList(UserIndex).flags.Muerto Then Exit Sub
        If UserList(UserIndex).flags.TargetNPC = 0 Then EnviarPaquete Paquetes.MensajeSimple, Chr$(131), UserIndex: Exit Sub
        
        'Â¿Es un banquero?
        If UserList(UserIndex).flags.TargetNpcTipo <> NPCTYPE_BANQUERO Then
            EnviarPaquete Paquetes.DescNpc, Chr$(90) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
            Exit Sub
        End If
        
        'Â¿Esta demasiado lejos?
        If distancia(UserList(UserIndex).pos, NpcList(UserList(UserIndex).flags.TargetNPC).pos) > 5 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(7), UserIndex
            Exit Sub
        End If
        
        UserRetiraItem UserIndex, Asc(Left$(anexo, 1)), DeCodify(Right$(anexo, Len(anexo) - 1))
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.Depositar
        If UserList(UserIndex).flags.Muerto Then Exit Sub
        
        If UserList(UserIndex).flags.TargetNPC = 0 Then EnviarPaquete Paquetes.MensajeSimple, Chr$(131), UserIndex: Exit Sub
        
        'Â¿Es un banquero?
        If UserList(UserIndex).flags.TargetNpcTipo <> NPCTYPE_BANQUERO Then
            EnviarPaquete Paquetes.DescNpc, Chr$(90) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
            Exit Sub
        End If
        
        'Â¿Esta demasiado lejos?
        If distancia(UserList(UserIndex).pos, NpcList(UserList(UserIndex).flags.TargetNPC).pos) > 5 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(7), UserIndex
            Exit Sub
        End If
        
        UserDepositaItem UserIndex, Asc(Left$(anexo, 1)), DeCodify(Right$(anexo, Len(anexo) - 1))
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.BancoOk
        UserList(UserIndex).flags.Comerciando = False
        EnviarPaquete Paquetes.FinBanOk, "", UserIndex
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.CCarpintero
        Call Trabajo.personajeCarpinteria(UserList(UserIndex), anexo)
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.FinComUsu
            If UserList(UserIndex).ComUsu.DestUsu > 0 And _
                UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                EnviarPaquete Paquetes.MensajeCompuesto, Chr$(40) & UserList(UserIndex).Name, UserList(UserIndex).ComUsu.DestUsu
                Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
            End If
            Call FinComerciarUsu(UserIndex)
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.RechazarComUsu
        If UserList(UserIndex).ComUsu.DestUsu > 0 Then
            EnviarPaquete Paquetes.MensajeCompuesto, Chr$(40) & UserList(UserIndex).Name, UserList(UserIndex).ComUsu.DestUsu
            Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
        End If
        EnviarPaquete Paquetes.MensajeSimple, Chr$(226), UserIndex
        Call FinComerciarUsu(UserIndex)
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.OfrecerComUsu
        Call OfrecerItemsComercio(UserIndex, anexo)
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.GuildInfo
        If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
           EnviarInformacionALider UserIndex
        Else
           EnviarListaClanes UserIndex
       End If
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.AceptarGuild
        mdClanes.AceptarMiembro UserIndex, anexo
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.RechazarGuild
        mdClanes.DenegarSolicitud UserIndex, anexo
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.EcharGuild
        mdClanes.EcharMiembro UserIndex, anexo
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.EnviarGuildComen
        mdClanes.EnviarComentarioPeticion UserIndex, Trim(anexo)
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.EnvPeaceOffer
       ' RecievePeaceOffer UserIndex, anexo
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.DeclararAlly
       ' DeclareAllie UserIndex, anexo
    Exit Sub
'.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.DeclararWar
       '  DeclareWar UserIndex, anexo
     Exit Sub
 '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.GuildDetail
        mdClanes.EnviarDetallesDeClan UserIndex, anexo
     Exit Sub
 '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.GuildDSend
        TempInt = clanes.getCantidad()
            If mdClanes.CrearClan(UserIndex, anexo) Then
                If TempInt = 0 Then
                   EnviarPaquete Paquetes.MensajeSimple, Chr$(249), UserIndex
                  Else
                   EnviarPaquete Paquetes.MensajeCompuesto, Chr$(21) & TempInt + 1, UserIndex
                  End If
                
                EnviarPaquete Paquetes.MensajeGuild, "Â¡Â¡Â¡" & UserList(UserIndex).Name & " fundo el clan '" & UserList(UserIndex).GuildInfo.GuildName & "'!!!", 0, ToAll
              End If
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.GuildCode
        mdClanes.ActualizarCodecsYDesc anexo, UserIndex
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.ActualizarGNews
        If Len(anexo) <= 1000 Then
            mdClanes.ActualizarNovedades anexo, UserIndex
          Else
            EnviarPaquete Paquetes.mensajeinfo, "El texto de las novedades es demasiado largo.", UserIndex, ToIndex
          End If
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.MemberInfo
        If Len(Trim(anexo)) > 0 Then
            mdClanes.EnviarInformacionPersonaje Trim(anexo), UserIndex
          End If
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.GuildSol
        SolicitudIngresoClan UserIndex, anexo
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.URLChange
        mdClanes.SetNewURL UserIndex, anexo
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.CHerrero
        Call Trabajo.personajeHerreria(UserList(UserIndex), anexo)
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
     ' Case cPaquetes.FinReto
  '        CerrarReto1 UserIndex
      'Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.ExitOk
        If Not CloseSocket(UserIndex) Then Call LogError("Exit ok")
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.ClickIzquierdo
        LookatTile UserIndex, UserList(UserIndex).pos.map, STI(anexo, 1), STI(anexo, 3)
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.ClickSkill
        AccionConSkill UserList(UserIndex), anexo
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.ClickAccion
        Call accion(UserList(UserIndex), Asc(Left$(anexo, 1)), Asc(mid(anexo, 2, 1)))
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.InfoHechizo
        If Len(anexo) = 0 Then Exit Sub
        tempbyte = UserList(UserIndex).Stats.UserHechizos(Asc(anexo))
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(22) & hechizos(tempbyte).nombre & "," & hechizos(tempbyte).desc & "," & hechizos(tempbyte).MinSkill & "," & getManaRequeridoHechizoParaPersonaje(UserList(UserIndex), hechizos(tempbyte)) & "," & hechizos(tempbyte).StaRequerido, UserIndex
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.Equipar
        tempbyte = Asc(anexo)
        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        If UserList(UserIndex).Invent.Object(tempbyte).ObjIndex = 0 Then Exit Sub
        EquiparInvItem UserIndex, tempbyte
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.SkillSetDomar
        EnviarPaquete Paquetes.ApuntarTrb, Chr$(eSkills.Domar), UserIndex
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.SkillSetRobar
        EnviarPaquete Paquetes.ApuntarTrb, Chr$(eSkills.Robar), UserIndex
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.SkillSetOcultar
        Call DoOcultarse(UserList(UserIndex))
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.CallForSkill
        EnviarSkills UserIndex
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.CallForFama
        EnviarFama UserIndex
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.CallForAtributos
        EnviarAtrib UserIndex
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.SosDone
        TempInt = NameIndex(anexo)
        Ayuda.eliminar TempInt
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.DIClick
           
          ' Intervalo para que no llenen el mapa de items
        If Not IntervaloPermiteAtacar(UserIndex, True) Then
              Exit Sub
          End If
        
        DraguedClick UserList(UserIndex), Asc(Left$(anexo, 1)), Asc(mid$(anexo, 2, 1)), Asc(mid$(anexo, 3, 1)), STI(anexo, 4)
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
      'Case cPaquetes.RetoAccpt
          's 'elect Case UserList(UserIndex).Retos.RetoActivo / 2
              'Case 1
            '      Call IniciarReto1vs1(UserIndex, UserList(UserIndex).Retos.Contrincante(1))
         ' End Select
      'Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    '  Case cPaquetes.RetoCncl
         ' Select Case UserList(UserIndex).Retos.RetoActivo / 2
            '  Case 1
                ' Call CancelarReto1(UserIndex, UserList(UserIndex).Retos.Contrincante(1))
          'End Select
     ' Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.PeaceProp
         ' SendPeacePropositions UserIndex
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.PeaceAccpt
         ' AcceptPeaceOffer UserIndex, anexo
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.SkillMod
              'Codigo para prevenir el hackeo de los skills
              '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            For tempbyte = 1 To NUMSKILLS
               tempInt2 = StringToByte(anexo, tempbyte)
                
                If tempInt2 < 0 Then
                    Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & HelperIP.longToIP(UserList(UserIndex).ip) & " trato de hackear los skills.")
                    UserList(UserIndex).Stats.SkillPts = 0
                    If Not CloseSocket(UserIndex) Then Call LogError("Skill MOD")
                      Exit Sub
                  End If
                
                TempInt = TempInt + tempInt2
            Next tempbyte
            
            If TempInt > UserList(UserIndex).Stats.SkillPts Then
                Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & HelperIP.longToIP(UserList(UserIndex).ip) & " trato de hackear los skills.")
                    If Not CloseSocket(UserIndex) Then Call LogError("Skillpts")
                  Exit Sub
              End If
              '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            
            For tempbyte = 1 To NUMSKILLS
                tempInt2 = StringToByte(anexo, tempbyte)
                UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts - tempInt2
                UserList(UserIndex).Stats.UserSkills(tempbyte) = UserList(UserIndex).Stats.UserSkills(tempbyte) + tempInt2
                If UserList(UserIndex).Stats.UserSkills(tempbyte) > 100 Then UserList(UserIndex).Stats.UserSkills(tempbyte) = 100
            
                If tempInt2 > 0 Then Call LogAsignaSkill(UserList(UserIndex).id, tempbyte, tempInt2, UserList(UserIndex).ip, UserList(UserIndex).Stats.UserSkills(tempbyte))
            Next tempbyte
          Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.entrenador
        Call modEntrenador.solicitarCriatura(UserList(UserIndex), anexo)
    Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
   Case cPaquetes.MNorteM
          Dim tindex As Integer
          Dim x As Byte
          Dim y As Byte
        If UserList(UserIndex).flags.Navegando Then Exit Sub
        If MapInfo(UserList(UserIndex).pos.map).Pk = False Then Exit Sub
        tindex = MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y - 1).UserIndex
        If tindex = 0 Then Exit Sub
        If UserList(tindex).flags.Muerto = 0 Then Exit Sub
        
        x = UserList(UserIndex).pos.x
        y = UserList(UserIndex).pos.y
        

        Call WarpUserCharEspecial(UserIndex, UserList(tindex).pos.map, UserList(tindex).pos.x, UserList(tindex).pos.y, eHeading.NORTH)
        Call WarpUserCharEspecial(tindex, UserList(UserIndex).pos.map, x, y, eHeading.SOUTH)
        MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).UserIndex = UserIndex
          Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
   Case cPaquetes.MOesteM
        If UserList(UserIndex).flags.Navegando Then Exit Sub
        If MapInfo(UserList(UserIndex).pos.map).Pk = False Then Exit Sub
        tindex = MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x - 1, UserList(UserIndex).pos.y).UserIndex
        If tindex = 0 Then Exit Sub
        If UserList(tindex).flags.Muerto = 0 Then Exit Sub
        x = UserList(UserIndex).pos.x
        y = UserList(UserIndex).pos.y
        Call WarpUserCharEspecial(UserIndex, UserList(tindex).pos.map, UserList(tindex).pos.x, UserList(tindex).pos.y, eHeading.WEST)
        Call WarpUserCharEspecial(tindex, UserList(UserIndex).pos.map, x, y, eHeading.EAST)
        MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).UserIndex = UserIndex
          Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.MSurM
        If UserList(UserIndex).flags.Navegando Then Exit Sub
        If MapInfo(UserList(UserIndex).pos.map).Pk = False Then Exit Sub
        tindex = MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y + 1).UserIndex
        If tindex = 0 Then Exit Sub
        If UserList(tindex).flags.Muerto = 0 Then Exit Sub
        x = UserList(UserIndex).pos.x
        y = UserList(UserIndex).pos.y
        Call WarpUserCharEspecial(UserIndex, UserList(tindex).pos.map, UserList(tindex).pos.x, UserList(tindex).pos.y, eHeading.SOUTH)
        Call WarpUserCharEspecial(tindex, UserList(UserIndex).pos.map, x, y, eHeading.NORTH)
        MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).UserIndex = UserIndex
          Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.MEsteM
        If UserList(UserIndex).flags.Navegando Then Exit Sub
        If MapInfo(UserList(UserIndex).pos.map).Pk = False Then Exit Sub
        tindex = MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x + 1, UserList(UserIndex).pos.y).UserIndex
        If tindex = 0 Then Exit Sub
        If UserList(tindex).flags.Muerto = 0 Then Exit Sub
        x = UserList(UserIndex).pos.x
        y = UserList(UserIndex).pos.y
        Call WarpUserCharEspecial(UserIndex, UserList(tindex).pos.map, UserList(tindex).pos.x, UserList(tindex).pos.y, eHeading.EAST)
        Call WarpUserCharEspecial(tindex, UserList(UserIndex).pos.map, x, y, eHeading.WEST)
        MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).UserIndex = UserIndex
          Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.FEST
        Call EnviarMiniEstadisticas(UserIndex)
          Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.iParty
        Call mdParty.SolicitarIngresoAParty(UserIndex)
          Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.ccParty
        Call CParty(UserIndex)
          Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.CrearParty
        If Not mdParty.PuedeCrearParty(UserIndex) Then Exit Sub
        Call mdParty.CrearParty(UserIndex)
          Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.Moverhechi
        Call DesplazarHechizo(UserIndex, CInt(Asc(mid(anexo, 1, 1))), CInt(Asc(mid(anexo, 2, 1))))
          Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.Encarcelame
        Call Encarcelar(UserList(UserIndex), TIEMPO_CARCEL_PIQUETE)
          Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.UNLAG
        Call enviarPosicion(UserList(UserIndex))
         Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.MTrabajar
        Call Trabajo.personajeTrabajar(UserList(UserIndex))
          Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.DejadeLaburar
        UserList(UserIndex).Trabajo.tipo = eTrabajos.Ninguno
        UserList(UserIndex).Trabajo.modo = 0
        UserList(UserIndex).Trabajo.cantidad = 0
        UserList(UserIndex).Trabajo.modificador = 0
        UserList(UserIndex).flags.Trabajando = False
        RemoverTrabajador UserIndex
          Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.Salirparty
        Call mdParty.SalirDeParty(UserIndex)
          Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.Aprobaringresoparty
        TempInt = NameIndex(anexo)
            If TempInt > 0 Then
                Call mdParty.AprobarIngresoAParty(UserIndex, TempInt)
              Else
                EnviarPaquete Paquetes.MensajeSimple2, Chr$(47), UserIndex, ToIndex
              End If
          Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.Expulsarparty
            TempInt = NameIndex(anexo)
            If TempInt > 0 Then
                Call mdParty.ExpulsarDeParty(UserIndex, TempInt)
              Else
                EnviarPaquete Paquetes.MensajeSimple2, Chr$(47), UserIndex, ToIndex
              End If
          Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.CrearReto
    
        Call modRetos.crear(UserList(UserIndex), anexo)

          Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.FaccionMsg
     If UserList(UserIndex).flags.ModoRol = True Then
           If anexo <> "" Then
              If UserList(UserIndex).faccion.ArmadaReal = 1 Then
                 For TempInt = 1 To LastUser
                 If UserList(TempInt).faccion.ArmadaReal = 1 Then EnviarPaquete Paquetes.MensajeArmadas, UserList(UserIndex).Name & "> " & Right(anexo, Len(anexo) - 1), TempInt
                   Next
                   Exit Sub
              ElseIf UserList(UserIndex).faccion.FuerzasCaos = 1 Then
                 For TempInt = 1 To LastUser
                 If UserList(TempInt).faccion.FuerzasCaos = 1 Then EnviarPaquete Paquetes.MensajeCaos, UserList(UserIndex).Name & "> " & Right(anexo, Len(anexo) - 1), TempInt
                   Next
                   Exit Sub
                End If
             End If
      Else
              'Sino tiene activado el modo rol hablar normalmente
            If UserList(UserIndex).flags.Privilegios > 1 Then
                Call LogGM(UserList(UserIndex).id, "Dijo: " & anexo)
              End If
        
            If UserList(UserIndex).flags.Oculto > 0 Then
                UserList(UserIndex).flags.Oculto = 0
                UserList(UserIndex).flags.Invisible = 0
                EnviarPaquete Paquetes.Desocultar, ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToMap
                EnviarPaquete Paquetes.MensajeSimple2, Chr$(288 - 255), UserIndex, ToIndex
              End If
            
            If UserList(UserIndex).flags.Muerto = 1 Then
                EnviarPaquete Paquetes.UserTalkDead, anexo & ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToDeadArea, UserList(UserIndex).pos.map
                EnviarPaquete Paquetes.UserTalkDead, anexo & ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToAdminsArea, UserList(UserIndex).pos.map
              Else
                If Len(anexo) = 1 And anexo = " " Then
                EnviarPaquete Paquetes.QDL, ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToMap, UserList(UserIndex).pos.map
                  Else
                EnviarPaquete Paquetes.UserTalk, anexo & ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToPCArea, UserList(UserIndex).pos.map
                  End If
              End If
      End If
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.ChangeItemsSlot
        If UserList(UserIndex).flags.Comerciando = True Then Exit Sub
        tempbyte = Asc(Left$(anexo, 1))
        tempbyte2 = Asc(Right$(anexo, 1))
        Call ChangeItemSlot(tempbyte2, tempbyte, UserIndex)
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.ChangeItemsSlotboveda
        tempbyte = Asc(Left$(anexo, 1))
        tempbyte2 = Asc(Right$(anexo, 1))
        Call ChangeItemSlotBoveda(tempbyte2, tempbyte, UserIndex)
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.PEACEDET
         ' Call SendPeaceRequest(UserIndex, anexo)
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.Lachiteo
        Call anticheat.anticheatCliente(UserList(UserIndex), anexo)
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.EFotoDenuncia
        If UserList(UserIndex).Counters.FotoDenuncia = 0 Then
            Call modFotodenuncias.reportarFotodenuncia(UserList(UserIndex), anexo)
          End If
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.LaChiteo2
        EnviarPaquete Paquetes.mensajeinfo, UserList(UserIndex).Name & " posible macro para chupar. Info: " & anexo, 0, ToAdmins
        
        Call LogAnticheat(UserList(UserIndex), macro, "Posible macro para chupar. Info: " & anexo)
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.obtClanMiembros
        If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
            Call mdClanes.EnviarInformacionALiderMiembros(UserIndex)
          End If
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.obtClanNews
        If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
            Call mdClanes.EnviarInformacionALiderNovedades(UserIndex)
          End If
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.obtClanSolicitudes
        If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
            Call mdClanes.EnviarInformacionALiderSolicitudes(UserIndex)
          End If
      Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.infoTransferencia
        Call modCapturarPantalla.agregarDatos(UserList(UserIndex), anexo)
          Exit Sub
  '.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.:':.
    Case cPaquetes.respuestaMemCheck
        Call Anticheat_MemCheck.respuestaPersonaje(UserList(UserIndex), anexo)
          Exit Sub
  End Select

End Sub

Public Function usuarioYaConectado(MacAddress) As Integer

Dim N As Integer

'Â¿Esta pc ya esta conectada??
If Len(MacAddress) <> 0 And MacAddress <> "000000000000" Then
    For N = 1 To LastUser
      If UserList(N).flags.UserLogged = True Then
        If UserList(N).MacAddress = MacAddress Then
                usuarioYaConectado = N
                Exit Function
            End If
        End If
    Next N
End If

usuarioYaConectado = 0

End Function




Private Sub ProcesarPaqueteOFF(ByVal numero As Byte, ByVal UserIndex As Integer, Optional anexo As String)
Dim tempstr As String
Dim TempStr2 As String
Dim tempbyte As Byte
Dim tempbyte2 As Byte
Dim MacAddress As String
Dim StrCheckSum As String

Dim donde As Byte


If ProfilePaquetes Then
    Logs.logProfilePaquete ("C" & numero)
End If
    
UserList(UserIndex).ConfirmacionConexion = 1


donde = 0

Select Case numero
    Case cPaquetes.NODO_INFO_HASH
        
        Call ProcesarPaqueteNodo(UserList(UserIndex), anexo)
        
        Exit Sub
    '-----------------------------------------------------------------------------------------------
    Case cPaquetes.CreatePj
        Dim Genero As Byte
        Dim clase As Byte
        Dim Raza As Byte
        Dim nombre As String
        Dim Password As String
        Dim Email As String
        Dim alineacion As Byte
        
        Genero = StringToByte(anexo, 1)
        clase = StringToByte(anexo, 2)
        Raza = StringToByte(anexo, 3)
        alineacion = StringToByte(anexo, 4)
        
        anexo = mid$(anexo, 5)
        
        nombre = ReadField(1, anexo, 44)
        Password = ReadField(2, anexo, 44)
        Email = ReadField(3, anexo, 44)
                
        crearPersonaje UserIndex, nombre, Password, Email, Genero, clase, Raza, alineacion

    Exit Sub
    '----------------------------------------------------------------------------------------------------
    Case 76
        EnviarPaquete Paquetes.IniciarAutoupdater, "", UserIndex, ToIndex
        If Not CloseSocket(UserIndex) Then Call LogError("Paquete 76")
    Case Else
        UserList(UserIndex).ConfirmacionConexion = 0
    Exit Sub
    '----------------------------------------------------------------------------------------------------
End Select

End Sub

Private Sub ProcesarComandoAdmin(ByVal numero As Byte, ByVal UserIndex As Integer, Optional anexo As String)
'Solo Dioses
Dim TempInt As Integer
Dim tempInt2 As Integer
Dim tempLong As Long
Dim tempstr As String
Dim tempbyte As Byte

If UserList(UserIndex).flags.Privilegios < PRIV_ADMINISTRADOR Then Exit Sub

Select Case numero
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case cmdAdmin.APAGAR
        Call LogMain("Server apagado por " & UserList(UserIndex).Name & ". ")

        Call frmServidor.cerrarServidorGracefull
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case cmdAdmin.CONDEN
        'TempInt = NameIndex(anexo)
       ' If TempInt > 0 Then Call modPersonaje.VolverCriminal(UserList(TempInt))
       EnviarPaquete Paquetes.mensajeinfo, "Este comando no posee validez.", UserIndex, ToIndex
    Exit Sub
     '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case cmdAdmin.BackUp
        DoBackUp
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case cmdAdmin.GRABAR
        Call modEventos.guardarEstadoEventos
        EnviarPaquete Paquetes.MensajeTalk, "Eventos guardados", UserIndex, ToIndex
        'GuardarUsuarios (1)
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case cmdAdmin.CMod
        TempInt = NameIndex(ReadField(1, anexo, Asc(" ")))
        tempstr = ReadField(2, anexo, Asc(" "))
        tempLong = val(ReadField(3, anexo, Asc(" ")))
        If TempInt <= 0 Then EnviarPaquete Paquetes.MensajeSimple, Chr$(5), UserIndex: Exit Sub
        Select Case UCase$(tempstr)
            Case "EXP"
                Call modUsuarios.agregarExperiencia(TempInt, tempLong)
                SendUserStatsBox TempInt

                If TempInt = UserIndex Then
                    EnviarPaquete Paquetes.mensajeinfo, "Sumaste " & tempLong & " puntos de experiencia", TempInt, ToIndex
                Else
                    EnviarPaquete Paquetes.mensajeinfo, "Otorgaste a " & UserList(UserIndex).Name & " " & tempLong & " puntos de experiencia", UserIndex, ToIndex
                End If
            Case "ORO"
                If tempLong > 10000000 Then EnviarPaquete Paquetes.mensajeinfo, "No se puede valores mayores a 10.000.000", UserIndex, ToIndex: Exit Sub
                UserList(TempInt).Stats.GLD = UserList(TempInt).Stats.GLD + tempLong
                SendUserStatsBox TempInt
            Case "MANA"
                If tempLong > 30000 Then Exit Sub
                UserList(TempInt).Stats.MaxMAN = tempLong
                SendUserStatsBox TempInt
            Case "BODY"
                ChangeUserChar ToMap, TempInt, UserList(TempInt).pos.map, TempInt, tempLong, UserList(TempInt).Char.Head, UserList(TempInt).Char.heading, UserList(TempInt).Char.WeaponAnim, UserList(TempInt).Char.ShieldAnim, UserList(TempInt).Char.CascoAnim
            Case "HEAD"
                ChangeUserChar ToMap, TempInt, UserList(TempInt).pos.map, TempInt, UserList(TempInt).Char.Body, tempLong, UserList(TempInt).Char.heading, UserList(TempInt).Char.WeaponAnim, UserList(TempInt).Char.ShieldAnim, UserList(TempInt).Char.CascoAnim
            Case "CRI"
                UserList(TempInt).faccion.CriminalesMatados = tempLong
            Case "CIU"
                UserList(TempInt).faccion.CiudadanosMatados = tempLong
            Case "VIDA"
                If tempLong > 30000 Then Exit Sub
                UserList(TempInt).Stats.MaxHP = tempLong
            Case "SKILL"
                If tempLong > 30000 Then Exit Sub
                UserList(TempInt).Stats.SkillPts = tempLong
                Call EnviarSubirNivel(TempInt, UserList(TempInt).Stats.SkillPts)
            Case "LEVEL"
                If tempLong > STAT_MAXELV Then Exit Sub
                UserList(TempInt).Stats.ELV = tempLong
                
                EnviarPaquete Paquetes.mensajeinfo, UserList(TempInt).Name & " es nivel " & tempLong & ".", UserIndex, ToIndex
            Case "GENERO"
                tempstr = UCase$(ReadField(3, anexo, Asc(" ")))
                tempLong = generoToByte(tempstr)
                
                If tempLong = 0 Then
                    EnviarPaquete Paquetes.mensajeinfo, "El generó ingresado es inválida. Debe ser HOMBRE O MUJER. Revisa los tildes.", UserIndex
                    Exit Sub
                End If
                
                UserList(TempInt).Genero = tempLong
                
                If TempInt = UserIndex Then
                    EnviarPaquete Paquetes.mensajeinfo, "Ahora tu género es " & tempstr & ".", UserIndex
                Else
                    EnviarPaquete Paquetes.mensajeinfo, "Ahora " & UserList(TempInt).Name & " tiene el género " & tempstr & ".", TempInt
                End If
                
            Case "RAZA"
                tempstr = UCase$(ReadField(3, anexo, Asc(" ")))
                tempLong = razaToByte(tempstr)
                
                If tempLong = 0 Then
                    EnviarPaquete Paquetes.mensajeinfo, "La raza ingresada es inválida. Revisa los tildes.", UserIndex
                    Exit Sub
                End If
                
                UserList(TempInt).Raza = tempLong
                
                If TempInt = UserIndex Then
                    EnviarPaquete Paquetes.mensajeinfo, "Ahora tu raza es " & tempstr & ".", UserIndex
                Else
                    EnviarPaquete Paquetes.mensajeinfo, "Ahora " & UserList(TempInt).Name & " es de raza " & tempstr & ".", TempInt
                End If
                
            Case "CLASE"
                tempstr = UCase$(ReadField(3, anexo, Asc(" ")))
                tempLong = claseToByte(tempstr)
                
                If tempLong = 0 Then
                    EnviarPaquete Paquetes.mensajeinfo, "La clase ingresada es inválida. Revisa los tildes.", UserIndex
                    Exit Sub
                End If
                
                UserList(TempInt).clase = tempLong
                UserList(TempInt).ClaseNumero = claseToConfigID(tempstr)
                
                If TempInt = UserIndex Then
                    EnviarPaquete Paquetes.mensajeinfo, "Ahora eres es un " & tempstr & ".", UserIndex
                Else
                    EnviarPaquete Paquetes.mensajeinfo, "Ahora " & UserList(TempInt).Name & " es un " & tempstr & ".", TempInt
                End If
        End Select
        Call SendUserStatsBox(TempInt)
         
        LogGM UserList(UserIndex).id, UserList(TempInt).Name & " su " & tempstr & " hasta " & tempLong, "MOD"
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case cmdAdmin.CI
        TempInt = STI(anexo, 1) 'Numero de Objeto
        tempLong = STI(anexo, 3) ' Cantidad para crear
        
        If ObjData(TempInt).Name = "Nada" Then Exit Sub
        
        If MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y - 1).OBJInfo.ObjIndex > 0 Then
            Exit Sub
        End If
        
        If Not MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y - 1).accion Is Nothing Then
            Exit Sub
        End If

        If TempInt < 1 Or TempInt > NumObjDatas Then Exit Sub
        If tempLong < 1 Or tempLong > 10000 Then Exit Sub
        Dim objeto As obj
        objeto.ObjIndex = TempInt
        objeto.Amount = tempLong
            
        Call LogGM(UserList(UserIndex).id, tempLong & " " & ObjData(objeto.ObjIndex).Name, "CI")
           
        Call MakeObj(ToMap, 0, UserList(UserIndex).pos.map, objeto, UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y - 1)
        Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case cmdAdmin.TCPEST
        
        Call modEstadisticasTCP.enviarEstadisticas(UserIndex)
        
        Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case cmdAdmin.ReloadServer
        frmCargando.Show
      
        Unload frmCargando
        
        ' Call LoadSini
        Call CargaNpcsDat
        ' Call CargarHechizos
        Call LoadOBJData
        EnviarPaquete Paquetes.mensajeinfo, "El server ha recargado la información de los dats.", UserIndex, ToIndex
        'Call LogGM(UserList(UserIndex).Name, "Recargo la informacion del servidor")
        Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
End Select

End Sub

Private Sub ProcesarComando3(ByVal numero As Byte, ByVal UserIndex As Integer, Optional anexo As String)
Dim tempstr As String
Dim TempStr2 As String
Dim tempbyte As Byte
Dim tempbyte2 As Byte
Dim tempByte3 As Byte
Dim MiObj As obj
Dim TempInt As Integer
'Solo GMS


If ProfilePaquetes Then
    Logs.logProfilePaquete ("3" & numero)
End If

If UserList(UserIndex).flags.Privilegios < 3 Then Exit Sub

'Solo GMS
Select Case numero
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.MASSDEST
        For tempbyte = UserList(UserIndex).pos.y - RangoY + 1 To UserList(UserIndex).pos.y + RangoY - 1
            For tempbyte2 = UserList(UserIndex).pos.x - RangoX + 1 To UserList(UserIndex).pos.x + RangoX - 1
                If tempbyte2 >= SV_Constantes.X_MINIMO_JUGABLE And tempbyte >= SV_Constantes.Y_MINIMO_JUGABLE And tempbyte2 <= SV_Constantes.X_MAXIMO_JUGABLE And tempbyte2 <= SV_Constantes.Y_MAXIMO_JUGABLE Then _
                    If MapData(UserList(UserIndex).pos.map, tempbyte2, tempbyte).OBJInfo.ObjIndex > 0 Then _
                    If ItemNoEsDeMapa(MapData(UserList(UserIndex).pos.map, tempbyte2, tempbyte).OBJInfo.ObjIndex) Then Call EraseObj(ToMap, UserIndex, UserList(UserIndex).pos.map, 10000, UserList(UserIndex).pos.map, tempbyte2, tempbyte)
            Next tempbyte2
    Next tempbyte
   
        LogGM UserList(UserIndex).id, UserList(UserIndex).pos.map & " " & UserList(UserIndex).pos.x & "-" & UserList(UserIndex).pos.y, "MASSDEST"
    Exit Sub
    
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.CTRIGGER
        MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).Trigger = val(anexo)
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.CDEST
        EraseObj ToMap, UserIndex, UserList(UserIndex).pos.map, _
        MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).OBJInfo.Amount, _
        UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y
        
        LogGM UserList(UserIndex).id, UserList(UserIndex).pos.map & " " & UserList(UserIndex).pos.x & "-" & UserList(UserIndex).pos.y, "DEST"
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.MATA
        If UserList(UserIndex).flags.TargetNPC <> 0 Then
            LogGM UserList(UserIndex).id, NpcList(UserList(UserIndex).flags.TargetNPC).Name, "MATA"
            QuitarNPC UserList(UserIndex).flags.TargetNPC
        End If
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.MASSKILL
        For tempbyte = UserList(UserIndex).pos.y - RangoY + 1 To UserList(UserIndex).pos.y + RangoY - 1
            For tempbyte2 = UserList(UserIndex).pos.x - RangoX + 1 To UserList(UserIndex).pos.x + RangoX - 1
                If tempbyte2 > 0 And tempbyte > 0 And tempbyte2 < 101 And tempbyte < 101 Then _
                If MapData(UserList(UserIndex).pos.map, tempbyte2, tempbyte).npcIndex > 0 Then Call QuitarNPC(MapData(UserList(UserIndex).pos.map, tempbyte2, tempbyte).npcIndex)
            Next tempbyte2
        Next tempbyte
        
        Call LogGM(UserList(UserIndex).id, UserList(UserIndex).pos.map & " " & UserList(UserIndex).pos.x & "-" & UserList(UserIndex).pos.y, "MASSKILL")
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.LastIP
         
         TempInt = NameIndex(anexo)
         EnviarPaquete Paquetes.mensajeinfo, HelperIP.longToIP(UserList(TempInt).ip), UserIndex
   '      LogGM UserList(UserIndex).Name, "LASTIP: " & Anexo
    
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.ACC
        TempInt = val(anexo)
        
        If TempInt > 0 Then 'TODO falta agregar que se fije el numero maximo de npcs
            Call invocarCriatura(TempInt, False, UserList(UserIndex))
        End If
        
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.RACC
        TempInt = val(anexo)
        
        If TempInt > 0 Then 'TODO falta agregar que se fije el maximo tambien
            'Â¿Cuantas criaturas hay?
            Call invocarCriatura(TempInt, True, UserList(UserIndex))
        End If
        
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.RAJARCLAN
        If isNombreValido(anexo) Then
            Call mdClanes.EcharIntegranteDeClan(anexo, UserIndex)
        Else
            EnviarPaquete Paquetes.mensajeinfo, "Nombre inválido", UserIndex, ToIndex
        End If
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.BORRAR_SOS
        Ayuda.vaciar
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.CLLUVIA
        Call modClima.cambiarClima
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.PASSDAY
        'TO-DO
        'DayElapsed
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
     Case CmdDios.ZONEST
            If MapInfo(UserList(UserIndex).pos.map).Pk = True Then
            MapInfo(UserList(UserIndex).pos.map).Pk = False
            EnviarPaquete Paquetes.MensajeSimple2, Chr(326 - 255), UserIndex, ToIndex
            Else
            MapInfo(UserList(UserIndex).pos.map).Pk = True
            EnviarPaquete Paquetes.MensajeSimple2, Chr(325 - 255), UserIndex, ToIndex
            End If
            Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.RETEST
    
            If Not modRetos.ACT_RETO Then
            
                modRetos.ACT_RETO = True
                EnviarPaquete Paquetes.MensajeServer, "Los retos han sido activados.", 0, ToAll
                
            Else
            
                modRetos.ACT_RETO = False
                EnviarPaquete Paquetes.MensajeServer, "Los retos han sido desactivados.", 0, ToAll
            End If
            Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.CHATEST
            If Not charlageneral Then
                Call modChatGlobal.activarChatGlobal
            Else
                Call modChatGlobal.desactivarChatGlobal
            End If
            Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.ECHARTODOSPJS
            'Call LogGM(UserList(UserIndex).Name, "Echo a todos los personajes")
            Call EcharPjsNoPrivilegiados
            Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.NickMac
            TempInt = NameIndex(anexo)
            If TempInt > 0 Then
                EnviarPaquete Paquetes.mensajeinfo, UserList(TempInt).MacAddress, UserIndex
                LogGM UserList(UserIndex).id, anexo, "NICKMAC"
            Else
                EnviarPaquete Paquetes.mensajeinfo, "Usuario inexistente", UserIndex
            End If
            Exit Sub
   '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   Case CmdDios.banMac
   
            If Len(anexo) > 0 Then
            
                tempstr = Trim(ReadField(1, anexo, Asc("@"))) 'Nick o MacAddress
                TempStr2 = Trim(ReadField(2, anexo, Asc("@"))) ' Razon
            
                TempInt = NameIndex(tempstr)
                'Banea la mac de un personaje o directo la mac Address?
                If TempInt > 0 Then
                    If Not AdminMacAddress.isMacBaneada(UserList(TempInt).MacAddress) Then
                        'Le baneo la mac y lo echo del juego
                        Call AdminMacAddress.banMac(UserList(TempInt).MacAddress, UserList(UserIndex).id, TempStr2)
                        CloseSocket TempInt
                        'Le aviso
                        EnviarPaquete Paquetes.mensajeinfo, "Has baneado la Mac Address de " & tempstr & " por " & TempStr2, UserIndex
                        'Guardo en el log
                        LogGM UserList(UserIndex).id, UserList(TempInt).MacAddress & "Razon: " & TempStr2, "BANMAC"
                    Else
                        EnviarPaquete Paquetes.mensajeinfo, "La MacAddress ya esta baneada.", UserIndex, ToIndex
                    End If
                Else
                    If Not AdminMacAddress.isMacBaneada(tempstr) Then
                        'baneo la MAC
                        Call AdminMacAddress.banMac(tempstr, UserList(UserIndex).id, TempStr2)
                        'Le aviso
                        EnviarPaquete Paquetes.mensajeinfo, "Has baneado la Mac Address " & tempstr & " por " & TempStr2, UserIndex
                        'Guardo el log
                        LogGM UserList(UserIndex).id, tempstr & "Razon: " & TempStr2, "BANMAC"
                    Else
                        EnviarPaquete Paquetes.mensajeinfo, "La MacAddress ya esta baneada.", UserIndex, ToIndex
                    End If
                End If
            Else
                EnviarPaquete Paquetes.mensajeinfo, "La Mac addres no es valida.", UserIndex
            End If
            Exit Sub
   '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   Case CmdDios.unBanMac
            If Len(anexo) >= 1 Then
                If AdminMacAddress.isMacBaneada(anexo) = True Then
                
                    Call AdminMacAddress.unBanMac(anexo)
        
                    EnviarPaquete Paquetes.mensajeinfo, "La MacAddres fue desbaneada.", UserIndex
                Else
                    EnviarPaquete Paquetes.mensajeinfo, "La MacAddres no esta baneada.", UserIndex
                End If
                
                LogGM UserList(UserIndex).id, anexo, "UNBANMAC"
            Else
            EnviarPaquete Paquetes.mensajeinfo, "La Mac addres no es valida.", UserIndex
            End If
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.Retorings
    
        EnviarPaquete Paquetes.mensajeinfo, "Rings disponibles " & modRings.getCantidadRingsDisponibles & ".", UserIndex, ToIndex
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.Habilitar
    If ServerSoloGMs = 1 Then
        ServerSoloGMs = 0
        EnviarPaquete Paquetes.mensajeinfo, "Servidor habilitado.", 0, ToAll
    Else
        ServerSoloGMs = 1
        EnviarPaquete Paquetes.mensajeinfo, "Servidor deshabilitado.", 0, ToAll
    End If
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.AceptarConsejo
        TempInt = NameIndex(anexo)
        
        If TempInt > 0 Then
        
            If UserList(TempInt).faccion.alineacion = eAlineaciones.caos Then
                UserList(TempInt).flags.PertAlConsCaos = 1
            ElseIf UserList(TempInt).faccion.alineacion = eAlineaciones.Real Then
                UserList(TempInt).flags.PertAlCons = 1
            End If
            
            Call WarpUserChar(TempInt, UserList(TempInt).pos.map, UserList(TempInt).pos.x, UserList(TempInt).pos.y, False)
        End If
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.ExpulsarConsejo
    
    TempInt = NameIndex(anexo)
    
    If TempInt > 0 Then
        UserList(TempInt).flags.PertAlConsCaos = 0
        UserList(TempInt).flags.PertAlCons = 0
        Call WarpUserChar(TempInt, UserList(TempInt).pos.map, UserList(TempInt).pos.x, UserList(TempInt).pos.y, False)
    End If
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.ModoRol
        TempInt = NameIndex(anexo)
        If TempInt <= 0 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(56), UserIndex
        Else
            UserList(TempInt).flags.ModoRol = Not UserList(TempInt).flags.ModoRol
        End If
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.CTE
        TempInt = val(ReadField(1, anexo, Asc(" ")))
        tempbyte2 = val(ReadField(2, anexo, Asc(" ")))
        tempByte3 = val(ReadField(3, anexo, Asc(" ")))
        
        Call modComandos.CrearPortal(UserList(UserIndex), TempInt, tempbyte2, tempByte3)
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   Case CmdDios.CheqCli
        EnviarPaquete Paquetes.Pong, "", UserIndex, ToAll
        Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   Case CmdDios.VerPongs
   
        ProfilePaquetes = Not ProfilePaquetes
        
        If ProfilePaquetes Then
            EnviarPaquete Paquetes.MensajeFight, "PROFILE: está activado.", UserIndex, ToIndex
        Else
            EnviarPaquete Paquetes.MensajeFight, "PROFILE: desactivado.", UserIndex, ToIndex
        End If
        'tempstr = ""
        
        'For TempInt = 1 To LastUser
        '    If UserList(TempInt).flags.UserLogged = True And UserList(TempInt).ConnIDValida And UserList(TempInt).PongLlego = 0 Then
        '        tempstr = tempstr & UserList(TempInt).Name & ", "
        '    End If
        'Next
        
        'If tempstr = "" Then tempstr = "Todo en orden."
        'EnviarPaquete Paquetes.mensajeinfo, tempstr, UserIndex, ToIndex
        Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   Case CmdDios.CNameClan
        tempstr = Trim(ReadField(1, anexo, Asc("@"))) 'Nombre actual
        TempStr2 = Trim(ReadField(2, anexo, Asc("@"))) ' Nuevo nombre
        Call mdClanes.cambiarNombreClan(tempstr, TempStr2, UserIndex)
        Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdDios.capturarPantalla
        Call modCapturarPantalla.capturarPantalla(UserList(UserIndex), anexo)
        Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   Case CmdDios.consultarMem
        TempInt = NameIndex(anexo)
        
        If TempInt = 0 Then
            EnviarPaquete Paquetes.mensajeinfo, "El personaje '" & anexo & "' se encuentra offline.", UserIndex, ToIndex
            Exit Sub
        End If
        
        If Anticheat_MemCheck.existeChequeoPara(UserList(TempInt)) Then
            EnviarPaquete Paquetes.mensajeinfo, "Ya existe un chequeo en curso para este personaje.", UserIndex, ToIndex
            Exit Sub
        End If
                
        Call Anticheat_MemCheck.chequearPersonaje(UserList(TempInt))
        
        EnviarPaquete Paquetes.mensajeinfo, "Se hace un chequeo sobre '" & UCase$(anexo) & "'.", UserIndex, ToIndex
        
        Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
End Select

End Sub

Private Sub ProcesarComando2(ByVal numero As Byte, ByVal UserIndex As Integer, Optional anexo As String)
Dim tempstr As String
Dim TempStr2 As String
Dim tempbyte As Byte
Dim tempbyte2 As Byte
Dim tempByte3 As Byte
Dim TempInt As Integer
Dim tempInt2 As Integer
Dim MiObj As obj


If ProfilePaquetes Then
    Logs.logProfilePaquete ("2" & numero)
End If
    
If UserList(UserIndex).flags.Privilegios < 2 Then Exit Sub

'Solo GMS
Select Case numero
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.CINFO
        'Â¿El usuario esta online?
        TempInt = NameIndex(anexo)
        If TempInt <= 0 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr(62), UserIndex
            Exit Sub
        End If
        SendUserStatsTxt UserIndex, TempInt
        

        LogGM UserList(UserIndex).id, anexo, "INFO"
        
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.INV
        'Â¿El usuario esta online?
        TempInt = NameIndex(anexo)
        If TempInt <= 0 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(62), UserIndex
            Exit Sub
        End If
        SendUserInvTxt UserIndex, TempInt
        
        LogGM UserList(UserIndex).id, anexo, "INV"
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.BOV
        'Â¿El usuario esta online?
        TempInt = NameIndex(anexo)
        If TempInt <= 0 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(56), UserIndex
            Exit Sub
        End If
        SendUserBovedaTxt UserIndex, TempInt
        
        LogGM UserList(UserIndex).id, anexo, "BOV"
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.CSKILLS
        'Â¿El usuario esta online?
        TempInt = NameIndex(anexo)
        If TempInt <= 0 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(56), UserIndex
            Exit Sub
        End If
        SendUserSkillsTxt UserIndex, TempInt
        
        LogGM UserList(UserIndex).id, anexo, "SKILLS"
        
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.CREVIVIR
        'Â¿El usuario esta online?
        TempInt = NameIndex(anexo)
        If TempInt <= 0 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(56), UserIndex
            Exit Sub
        End If
        UserList(TempInt).flags.Muerto = 0
        UserList(TempInt).Stats.minHP = UserList(TempInt).Stats.MaxHP
        Call DarCuerpoDesnudo(UserList(TempInt))
        Call ChangeUserChar(ToMap, 0, UserList(TempInt).pos.map, val(TempInt), UserList(TempInt).Char.Body, UserList(TempInt).OrigChar.Head, UserList(TempInt).Char.heading, UserList(TempInt).Char.WeaponAnim, UserList(TempInt).Char.ShieldAnim, UserList(TempInt).Char.CascoAnim)
        Call SendUserStatsBox(TempInt)
       ' EnviarPaquete Paquetes.MensajeCompuesto, Chr$(28) & UserList(UserIndex).Name, TempInt
        LogGM UserList(UserIndex).id, anexo, "RESU"
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.ONLINEGM
        For TempInt = 1 To LastUser
            If UserList(TempInt).flags.Privilegios > 0 And UserList(TempInt).Name <> "" Then
                tempstr = tempstr & UserList(TempInt).Name & ", "
            End If
        Next TempInt
        tempstr = Left$(tempstr, Len(tempstr) - 2)
        
        EnviarPaquete Paquetes.mensajeinfo, tempstr, UserIndex
    
        LogGM UserList(UserIndex).id, tempstr, "ONLINEGM"
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   ' Case CmdSemi.PERDON
   '     TempInt = NameIndex(anexo)
   '     If TempInt > 0 Then
  '         VolverCiudadano TempInt
  '      End If
        
 '       LogGM UserList(UserIndex).id, anexo, "PERDON"
 '   Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.CECHAR
        TempInt = NameIndex(anexo)
        If TempInt <= 0 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(56), UserIndex
            Exit Sub
        Else
            CloseSocket TempInt
            LogGM UserList(UserIndex).id, anexo, "ECHAR"
        End If
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.CBAN
        tempstr = ReadField(2, anexo, Asc("@")) 'Nick
        TempInt = NameIndex(tempstr) 'Index del nick
        tempbyte = Int(ReadField(3, anexo, Asc("@"))) ' Cantidad de dias de baeno
    
        If UserIndex = TempInt Then Exit Sub ' no se puede banear asi mismo, sino se buguea
    
        Call BanearUsuario(UserList(UserIndex).Name, tempstr, ReadField(1, anexo, Asc("@")), tempbyte, True)
        
        Call LogGM(UserList(UserIndex).id, tempstr, "BAN")
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.CUNBAN
        ' TODO. hacer bien
        
        If UserList(UserIndex).flags.Privilegios < PRIV_DIOS Then
            Exit Sub
        End If
        
        anexo = Trim$(anexo)
        
        Call usuarios.UnbanearUsuario(UserList(UserIndex).Name, anexo, True)
        
        LogGM UserList(UserIndex).id, anexo, "UNBAN"

    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.SEGUIR
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.CSUM
        TempInt = NameIndex(anexo)
        If TempInt <= 0 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(56), UserIndex
            Exit Sub
        End If
    '    EnviarPaquete Paquetes.MensajeCompuesto, Chr$(29) & UserList(UserIndex).Name, TempInt
        LogGM UserList(UserIndex).id, UserList(TempInt).Name & " Map:" & UserList(UserIndex).pos.map & " X:" & UserList(UserIndex).pos.x & " Y:" & UserList(UserIndex).pos.y & " (Desde " & UserList(TempInt).pos.map & ")", "SUM"
    
        WarpUserChar TempInt, UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y + 1, True
            
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.cc
        EnviarSpawnList UserIndex
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.RESETINV
        ' Obtenemos el nombre
        tempstr = Trim$(anexo)
        
        ' Obtengo el Index
        TempInt = NameIndex(anexo)
        
        ' Â¿Esta Online?
        If TempInt <= 0 Then
            EnviarPaquete Paquetes.mensajeinfo, "El personaje " & anexo & " no está Online.", UserIndex, ToIndex
            Exit Sub
        End If
        
        '
        If UserList(TempInt).Stats.minHP > 1 Then
            UserList(TempInt).Stats.minHP = UserList(TempInt).Stats.minHP - 1
            
            Call SendUserStatsBox(TempInt)
            EnviarPaquete Paquetes.mensajeinfo, "Se le quito un punto de vida a " & anexo, UserIndex, ToIndex
        Else
            EnviarPaquete Paquetes.mensajeinfo, "No se le puede sacar vida a " & anexo & ". Tiene solo punto de vida.", UserIndex, ToIndex
        End If
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.LIMPIAR
        LimpiarMundo
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.ROSG
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(14) & anexo, UserIndex, ToAll
        LogGM UserList(UserIndex).id, anexo, "ROSG"
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.NICK2IP
        TempInt = NameIndex(anexo)
        If TempInt <= 0 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(56), UserIndex
            Exit Sub
        End If
        If UserList(UserIndex).flags.Privilegios > UserList(TempInt).flags.Privilegios Then
            EnviarPaquete Paquetes.mensajeinfo, "El ip de " & anexo & " es " & HelperIP.longToIP(UserList(TempInt).ip) & ".", UserIndex
        End If
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.IP2NICK
        TempInt = NameIndex(anexo)
        If TempInt <= 0 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(56), UserIndex
            Exit Sub
        End If
        
        For tempInt2 = 1 To LastUser
            If UserList(tempInt2).ip = UserList(TempInt).ip Then
                tempstr = tempstr & UserList(tempInt2).Name & ", "
            End If
        Next tempInt2
        If Len(tempstr) > 1 Then tempstr = Left$(tempstr, Len(tempstr) - 2)
        EnviarPaquete Paquetes.mensajeinfo, "Los personajes con esa ip son: " & tempstr, UserIndex
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.ejecutar
        TempInt = NameIndex(anexo)
        If TempInt > 0 Then
            'If UserList(UserIndex).flags.Privilegios > UserList(TempInt).flags.Privilegios Then
                Call UserDie(TempInt, False)
                
                LogGM UserList(UserIndex).id, anexo, "EJECUTAR"
            'End If
        Else
            EnviarPaquete Paquetes.MensajeSimple, Chr(69), UserIndex, ToIndex
        End If
        Exit Sub
   '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.Qtalk
        'Ya que no se usa este comando.. ponemos para confirmar el fin correcto de un evento que paga apuestas
        'Nombre evento@Desc del equipo
        If modEventos.establecerGanadorEvento(Trim(ReadField(1, anexo, Asc("@"))), Trim(ReadField(2, anexo, Asc("@")))) Then
            EnviarPaquete Paquetes.MensajeTalk, "Ok.. espera unos segundos", UserIndex, ToIndex
        Else
            EnviarPaquete Paquetes.MensajeTalk, "El evento no existe o el equipo esta mal puesto.", UserIndex, ToIndex
        End If
       Exit Sub
   '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.AUCO
       Conteo = anexo
       frmMain.Timer2.Enabled = True
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.LargarCentinelas
    
        If frmMain.AntiMacrosCen.Enabled Then
        
            EnviarPaquete Paquetes.mensajeinfo, "Los centinelas ya estan trabajando.", UserIndex, ToIndex
            LogGM UserList(UserIndex).id, "Intento llamar", "CENTINELA"
            
        Else
        
            LogGM UserList(UserIndex).id, "Los llamo", "CENTINELA"
            
            modCentinelas.TiempoMin = 999 'Fuerzo la ejecucion
            Call modCentinelas.AntiMacrosL
            
            EnviarPaquete Paquetes.mensajeinfo, "Se ha llamado a los centinelas.", UserIndex, ToIndex
        End If
        
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.Secaeest
        If MapInfo(UserList(UserIndex).pos.map).SeCaeiItems = 1 Then
            MapInfo(UserList(UserIndex).pos.map).SeCaeiItems = 0
            EnviarPaquete Paquetes.mensajeinfo, "Se caen los items en este mapa.", UserIndex, ToIndex
        Else
            MapInfo(UserList(UserIndex).pos.map).SeCaeiItems = 1
            EnviarPaquete Paquetes.mensajeinfo, "NO se caen los items en este mapa.", UserIndex, ToIndex
        End If
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.Aotromap
        TempInt = UserList(UserIndex).pos.map
    
        If val(ReadField(1, anexo, Asc(" "))) = 0 Or (SV_PosicionesValidas.existeMapa(val(ReadField(1, anexo, Asc(" ")))) And SV_PosicionesValidas.existePosicionMundo(val(ReadField(1, anexo, Asc(" "))), val(ReadField(2, anexo, Asc(" "))), val(ReadField(3, anexo, Asc(" "))))) Then
            MapInfo(TempInt).Aotromapa.map = val(ReadField(1, anexo, Asc(" ")))
            MapInfo(TempInt).Aotromapa.x = val(ReadField(2, anexo, Asc(" ")))
            MapInfo(TempInt).Aotromapa.y = val(ReadField(3, anexo, Asc(" ")))
            
            EnviarPaquete Paquetes.mensajeinfo, "Ok.", UserIndex, ToIndex
            
            Call LogGM(UserList(UserIndex).id, "Mapa " & MapInfo(TempInt).Aotromapa.map & " X: " & MapInfo(TempInt).Aotromapa.x & " Y: " & MapInfo(TempInt).Aotromapa.y, "SEVA")
        Else
            EnviarPaquete Paquetes.mensajeinfo, "El mapa o la posición no es válida.", UserIndex, ToIndex
        End If
     Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.AntiHpts
        If MapInfo(UserList(UserIndex).pos.map).AntiHechizosPts = 1 Then
            MapInfo(UserList(UserIndex).pos.map).AntiHechizosPts = 0
            EnviarPaquete Paquetes.mensajeinfo, "El mapa volvio a la normalidad.", UserIndex, ToIndex
        Else
            MapInfo(UserList(UserIndex).pos.map).AntiHechizosPts = 1
            EnviarPaquete Paquetes.mensajeinfo, "NO se puede tirar invisibilidad, elementales o robar en este mapa.", UserIndex, ToIndex
        End If
    Exit Sub
   '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.MapaFrio
         If MapInfo(UserList(UserIndex).pos.map).Frio = 1 Then
            MapInfo(UserList(UserIndex).pos.map).Frio = 0
            EnviarPaquete Paquetes.mensajeinfo, "Volvio el calor!!.", UserIndex, ToIndex
         Else
            MapInfo(UserList(UserIndex).pos.map).Frio = 1
            EnviarPaquete Paquetes.mensajeinfo, "Que frio!.", UserIndex, ToIndex
         End If
    Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.HabilitarRobo
         If MapInfo(UserList(UserIndex).pos.map).PermiteRoboNPC = 1 Then
            MapInfo(UserList(UserIndex).pos.map).PermiteRoboNPC = 0
            EnviarPaquete Paquetes.mensajeinfo, "En este mapa NO se permite el robo de npcs.", UserIndex, ToIndex
         Else
            MapInfo(UserList(UserIndex).pos.map).PermiteRoboNPC = 1
            EnviarPaquete Paquetes.mensajeinfo, "En este mapa SI se permite el robo de npcs.", UserIndex, ToIndex
         End If
   Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.MinLevelMap
        anexo = Int(val(anexo))
        MapInfo(UserList(UserIndex).pos.map).Nivel = anexo
        EnviarPaquete Paquetes.mensajeinfo, "El nivel minimo para ingresar a este mapa es: " & anexo, UserIndex, ToIndex
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.MaxLevelMap
        anexo = Int(val(anexo))
        MapInfo(UserList(UserIndex).pos.map).MaxLevel = anexo
        EnviarPaquete Paquetes.mensajeinfo, "El nivel maximo para ingresar a este mapa es: " & anexo, UserIndex, ToIndex
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.Cname
        tempstr = UCase$(Trim(ReadField(1, anexo, Asc("@")))) 'Nombre actual
        TempStr2 = UCase$(Trim(ReadField(2, anexo, Asc("@")))) ' Nuevo nombre

        Call modPersonaje.CambiarNombre(UserIndex, tempstr, TempStr2)
    Exit Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   Case CmdSemi.Spawn
        tempInt2 = Asc(anexo)
        
        If tempInt2 >= LBound(SpawnList) And tempInt2 <= UBound(SpawnList) Then
            Call modComandos.invocarCriatura(SpawnList(tempInt2).npcIndex, False, UserList(UserIndex))
        End If
        
    Exit Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   Case CmdSemi.InfoMap
        'Informacion del estado del mapa
        TempInt = UserList(UserIndex).pos.map
        
        ' Cantidad de Criaturas y cantidad de personajes en el mapa
        tempstr = "Usuarios en el mapa: " & MapInfo(TempInt).usuarios.getCantidadElementos & vbCrLf
        tempstr = tempstr & "Criaturas en el mapa: " & MapInfo(TempInt).NPCs.getCantidadElementos & vbCrLf
        
        ' Limite en las magias utilizadas
        tempstr = tempstr & "Anti nw: "
        If MapInfo(TempInt).AntiHechizosPts = 0 Then
            tempstr = tempstr & "No" & vbCrLf
          Else
            tempstr = tempstr & "Si" & vbCrLf
          End If
        
          ' Caida o no de objetos al morir
        tempstr = tempstr & "Se caen los items: "
        If MapInfo(TempInt).SeCaeiItems = 0 Then
            tempstr = tempstr & "Si" & vbCrLf
          Else
            tempstr = tempstr & "No" & vbCrLf
          End If
        
          ' Seva
        If MapInfo(TempInt).Aotromapa.map = 0 Then
            tempstr = tempstr & "Destino al morir: Ninguno" & vbCrLf
          Else
            tempstr = tempstr & "Destino al morir: " & MapInfo(TempInt).Aotromapa.map & " (" & MapInfo(TempInt).Aotromapa.x & "," & MapInfo(TempInt).Aotromapa.y & ")" & vbCrLf
          End If
        
          ' Permite robo de criaturas
        If MapInfo(TempInt).PermiteRoboNPC = 1 Then
            tempstr = tempstr & "AntiRobo: No" & vbCrLf
          Else
            tempstr = tempstr & "AntiRobo: Si" & vbCrLf
          End If
        
          ' Zona segura o insegura
        If MapInfo(TempInt).Pk = True Then
            tempstr = tempstr & "Zona Insegura" & vbCrLf
          Else
            tempstr = tempstr & "Zona Segura" & vbCrLf
          End If
        
          ' Nivel minimo/máximo y cantidad de personajes en el mapa
        tempstr = tempstr & "Nivel Minima/Maximo: " & MapInfo(TempInt).Nivel & "/" & MapInfo(TempInt).MaxLevel & vbCrLf
        tempstr = tempstr & "Max usuarios en el mapa: " & MapInfo(TempInt).UsuariosMaximo & vbCrLf
        
          ' Ingreso permitido de alineaciones
        If MapInfo(TempInt).SoloCiudas = 1 Then
            tempstr = tempstr & "El mapa solo puede ser accedido por integrantes del Ejército Índigo." & vbCrLf
        ElseIf MapInfo(TempInt).SoloCrimis = 1 Then
            tempstr = tempstr & "El mapa solo puede ser accedido por integrantes del Ejército Escarlata." & vbCrLf
          Else
            tempstr = tempstr & "El mapa puede ser accedido por todos los status." & vbCrLf
          End If
        
          ' Ingreso permitido de facciones
        If MapInfo(TempInt).SoloArmada = 1 And MapInfo(TempInt).SoloCaos = 1 Then
            tempstr = tempstr & "El mapa solo puede ser accedido por legionarios." & vbCrLf
        ElseIf MapInfo(TempInt).SoloArmada = 1 Then
            tempstr = tempstr & "El mapa solo puede ser accedido por armadas reales." & vbCrLf
        ElseIf MapInfo(TempInt).SoloCaos = 1 Then
            tempstr = tempstr & "El mapa solo puede ser accedido por integrates del ejercito del caos" & vbCrLf
          Else
            tempstr = tempstr & "El mapa no exige ser de alguna faccion para ingresar." & vbCrLf
          End If
    
          ' Envio informacion
        EnviarPaquete Paquetes.mensajeinfo, tempstr, UserIndex, ToIndex
      Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.LimiteUserMap
        TempInt = val(anexo)
        MapInfo(UserList(UserIndex).pos.map).UsuariosMaximo = TempInt
        EnviarPaquete Paquetes.mensajeinfo, "El mapa soporta " & TempInt & " usuarios.", UserIndex, ToIndex
      Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.OnlyCiuda
        If MapInfo(UserList(UserIndex).pos.map).SoloCiudas = 0 Then
            MapInfo(UserList(UserIndex).pos.map).SoloCiudas = 1
            EnviarPaquete Paquetes.mensajeinfo, "El mapa solo puede ser accedido por ciudadanos.", UserIndex, ToIndex
          Else
            MapInfo(UserList(UserIndex).pos.map).SoloCiudas = 0
            EnviarPaquete Paquetes.mensajeinfo, "El mapa puede ser accedido por todos.", UserIndex, ToIndex
          End If
        
        MapInfo(UserList(UserIndex).pos.map).SoloCrimis = 0
      Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.OnlyCrimi
        If MapInfo(UserList(UserIndex).pos.map).SoloCrimis = 0 Then
            MapInfo(UserList(UserIndex).pos.map).SoloCrimis = 1
            EnviarPaquete Paquetes.mensajeinfo, "El mapa solo puede ser accedido por integrantes del Ejército Escarlata.", UserIndex, ToIndex
          Else
            MapInfo(UserList(UserIndex).pos.map).SoloCrimis = 0
            EnviarPaquete Paquetes.mensajeinfo, "El mapa puede ser accedido por todos.", UserIndex, ToIndex
          End If
        
        MapInfo(UserList(UserIndex).pos.map).SoloCiudas = 0
      Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   Case CmdSemi.OnlyCaos
        If MapInfo(UserList(UserIndex).pos.map).SoloCaos = 0 Then
            MapInfo(UserList(UserIndex).pos.map).SoloCaos = 1
            EnviarPaquete Paquetes.mensajeinfo, "El mapa solo puede ser accedido por integrantes del ejercito del Caos.", UserIndex, ToIndex
          Else
            MapInfo(UserList(UserIndex).pos.map).SoloCaos = 0
            EnviarPaquete Paquetes.mensajeinfo, "El mapa no impone restricciones de ser del caos para ingresar.", UserIndex, ToIndex
          End If
      Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   Case CmdSemi.OnlyArmada
        If MapInfo(UserList(UserIndex).pos.map).SoloArmada = 0 Then
            MapInfo(UserList(UserIndex).pos.map).SoloArmada = 1
            EnviarPaquete Paquetes.mensajeinfo, "El mapa solo puede ser accedido por integrantes del ejercito real.", UserIndex, ToIndex
          Else
            MapInfo(UserList(UserIndex).pos.map).SoloArmada = 0
            EnviarPaquete Paquetes.mensajeinfo, "El mapa no impone restricciones de ser de la armada real para ingresar.", UserIndex, ToIndex
          End If
      Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.CrearEvento
        Call parsearInfo(anexo, UserList(UserIndex))
      Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.ObtenerEventos
        EnviarPaquete Paquetes.InfoAdminEventos, modEventos.obtenerEstadoTorneos, UserIndex, ToIndex
      Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.obtenerInfoEvento
        Call modAdminEventosGM.obtenerInfoEvento(anexo, UserList(UserIndex))
      Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.cancelarEvento
        If modEventos.cancelarEvento(anexo) = True Then
            EnviarPaquete Paquetes.MensajeAdminEventos, "El evento ha sido cancelado.", UserIndex, ToIndex
          Else
            EnviarPaquete Paquetes.MensajeAdminEventos, "ERROR. El evento no existe o no pudo ser cancelado.", UserIndex, ToIndex
          End If
      Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.publicarEvento
        If modEventos.publicarEvento(anexo) = True Then
            EnviarPaquete Paquetes.MensajeAdminEventos, "El evento ha sido publicado.", UserIndex, ToIndex
          Else
            EnviarPaquete Paquetes.MensajeAdminEventos, "ERROR. El evento no existe o no pudo ser publicado.", UserIndex, ToIndex
          End If
      Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.inscribirEvento
        
        Call modAdminEventosGM.inscribirPersonajes(anexo, UserList(UserIndex))
    
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.BLOQ
    
 '         'Puede hacer esto?
 '        If UserList(UserIndex).flags.Privilegios = PRIV_GAMEMASTER Then
 '            If Not (modGameMasterEventos.esMapaDeEvento(UserList(UserIndex).Pos.map)) Then
 '                EnviarPaquete Paquetes.mensajeinfo, "Los GameMasters solo pueden bloquear mapas de eventos. Segui ayudando a Tierras del Sur y algún día serás Dios.", UserIndex, ToIndex
 '                 Exit Sub
 '             End If
 '         End If
 '
 '        If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y).Blocked = 1 Then
 '            MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y).Blocked = 0
 '            Bloquear ToMap, UserIndex, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y, 0
 '
 '            LogGM UserList(UserIndex).id, UserList(UserIndex).Pos.map & " " & UserList(UserIndex).Pos.x & "-" & UserList(UserIndex).Pos.y, "UNBLOCK"
 '
 '         Else
 '            MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y).Blocked = 1
 '            Bloquear ToMap, UserIndex, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y, 1
 '
 '            LogGM UserList(UserIndex).id, UserList(UserIndex).Pos.map & " " & UserList(UserIndex).Pos.x & "-" & UserList(UserIndex).Pos.y, "BLOCK"
 '
 '         End If
      Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.CT
          'Posicion del Game Master
        TempInt = val(ReadField(1, anexo, Asc(" ")))
        tempbyte2 = val(ReadField(2, anexo, Asc(" ")))
        tempByte3 = val(ReadField(3, anexo, Asc(" ")))
        
        Call modComandos.CrearPortal(UserList(UserIndex), TempInt, tempbyte2, tempByte3)
   
      Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdSemi.DT
        TempInt = UserList(UserIndex).flags.TargetMap
        tempbyte2 = UserList(UserIndex).flags.TargetX
        tempByte3 = UserList(UserIndex).flags.TargetY
        
        Call EliminarPortal(UserList(UserIndex), TempInt, tempbyte2, tempByte3)
        
      Exit Sub
    
    Case CmdSemi.RMSG
        EnviarPaquete Paquetes.MensajeTalk, UserList(UserIndex).Name & "> " & anexo, UserIndex, ToAll
        
        LogGM UserList(UserIndex).id, anexo, "RMSG"
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
  End Select

End Sub

Private Sub ProcesarComando1(ByVal numero As Byte, ByVal UserIndex As Integer, Optional anexo As String)
'Solo GMS
If UserList(UserIndex).flags.Privilegios = 0 Then Exit Sub
'Solo GMS
Dim tempstr As String
Dim tempbyte As Byte
Dim tempbyte2 As Byte
Dim tempByte3 As Byte
Dim TempInt As Integer
Dim tempInt2 As Integer

    If ProfilePaquetes Then
        Logs.logProfilePaquete ("1" & numero)
    End If
    
Select Case numero

        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdConse.CREM
            LogGM UserList(UserIndex).id, anexo, "REM"
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdConse.Hora
            EnviarPaquete Paquetes.mensajeinfo, Time & " " & Date, UserIndex, ToAll
            LogGM UserList(UserIndex).id, "", "HORA"
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdConse.donde
            TempInt = NameIndex(anexo)
            If TempInt <= 0 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(56), UserIndex
                Exit Sub
            End If
            EnviarPaquete Paquetes.mensajeinfo, "Ubicacion: " & UserList(TempInt).pos.map & "[" & UserList(TempInt).pos.x & "." & UserList(TempInt).pos.y & "]", UserIndex
            
            LogGM UserList(UserIndex).id, UserList(TempInt).Name, "DONDE"
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdConse.NENE
            TempInt = CIntSeguro(anexo)
            
            If Not SV_PosicionesValidas.existeMapa(TempInt) Then
                EnviarPaquete Paquetes.mensajeinfo, "Mapa inválido. Escribí correctamente el número de mapa que queres consultar.", UserIndex, ToIndex
                Exit Sub
            End If
            
            EnviarPaquete Paquetes.mensajeinfo, "Cantidad de personajes en el mapa " & TempInt & ": " & MapInfo(anexo).usuarios.getCantidadElementos, UserIndex
            LogGM UserList(UserIndex).id, str$(MapInfo(anexo).usuarios.getCantidadElementos), "NENE"
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdConse.TELEPLOC
            If Not SV_PosicionesValidas.existeMapa(val(anexo)) Then Exit Sub
            
            WarpUserChar UserIndex, UserList(UserIndex).flags.TargetMap, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, IIf(UserList(UserIndex).flags.AdminInvisible = 1, False, True)
            
            LogGM UserList(UserIndex).id, UserList(UserIndex).flags.TargetMap & " " & UserList(UserIndex).flags.TargetX & " " & UserList(UserIndex).flags.TargetY, "TELEPLOC"
        
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdConse.TELEP
        
            tempstr = UCase$(ReadField(1, anexo, Asc(" "))) ' Usuario
            tempInt2 = CIntSeguro(ReadField(2, anexo, Asc(" "))) ' Mapa
            tempbyte2 = CByteSeguro(ReadField(3, anexo, Asc(" "))) ' X
            tempByte3 = CByteSeguro(ReadField(4, anexo, Asc(" "))) ' Y
                        
            ' Â¿Puso nombre del usuario?
            If Len(tempstr) = 0 Then Exit Sub
            
            ' Â¿Mapa valido?
            If Not SV_PosicionesValidas.existeMapa(tempInt2) Then
                EnviarPaquete Paquetes.mensajeinfo, "El mapa al cual deseas trasportar no existe.", UserIndex, ToIndex
                Exit Sub
            End If
                    
            ' Â¿Posicion valida?
            If Not SV_PosicionesValidas.existePosicionMundo(tempInt2, tempbyte2, tempByte3) Then
                EnviarPaquete Paquetes.mensajeinfo, "La posición a la cual deseas transportar no es válida.", UserIndex, ToIndex
                Exit Sub
            End If
            
            If tempstr <> "YO" And tempstr <> UCase$(UserList(UserIndex).Name) Then
            
                'Transporta a otra persona.. Puede hacerlo?
                'Los consejeros no pueden transportar
                If UserList(UserIndex).flags.Privilegios > 1 Then
                
                    TempInt = NameIndex(tempstr)
                    
                    If TempInt > 0 Then
                        Call LogGM(UserList(UserIndex).id, UserList(TempInt).Name & " hacia " & "Mapa" & tempInt2 & " X:" & tempbyte2 & " Y:" & tempByte3 & " (desde Mapa " & UserList(TempInt).pos.map & " X: " & UserList(TempInt).pos.x & " Y:" & UserList(TempInt).pos.y & ")", "TELEP")
                        
                        Call WarpUserChar(TempInt, tempInt2, tempbyte2, tempByte3, True)
                    Else
                        EnviarPaquete Paquetes.mensajeinfo, "El personaje " & tempstr & " no está online.", UserIndex, ToIndex
                    End If
                    
                End If
                
            Else
                WarpUserChar UserIndex, tempInt2, tempbyte2, tempByte3, True
                Call LogGM(UserList(UserIndex).id, "YO " & "Mapa" & tempInt2 & " X:" & tempbyte2 & " Y:" & tempByte3, "TELEP")
            End If
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdConse.SHOW_SOS
            Ayuda.itIniciar
            
            Do While Ayuda.ithasNext
                TempInt = Ayuda.itnext
                EnviarPaquete Paquetes.SOSAddItem, UserList(TempInt).Name, UserIndex
            Loop
            
            EnviarPaquete Paquetes.SOSViewList, "", UserIndex
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdConse.IRA
            Dim nPos As WorldPos
            
            TempInt = NameIndex(anexo)

            If TempInt <= 0 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(56), UserIndex
                LogGM UserList(UserIndex).id, anexo & " (offline)", "IRA"
                Exit Sub
            End If
              
            'Obtenemos una posicion cercana al personaje
            Call ClosestLegalPos(UserList(TempInt).pos, nPos, UserList(UserIndex))
            
            'Esta ok la pos?
            If nPos.map = 0 Then
                EnviarPaquete Paquetes.mensajeinfo, "No hay una posición cercana al personaje legal.", UserIndex, ToIndex
                Exit Sub
            End If
            
            'Lo transportamos
            If UserList(UserIndex).flags.AdminInvisible = 0 Then
                EnviarPaquete Paquetes.mensajeinfo, UserList(UserIndex).Name & " se trasporto hacia ti.", TempInt, ToIndex
                WarpUserChar UserIndex, nPos.map, nPos.x, nPos.y, True
            Else
                WarpUserChar UserIndex, nPos.map, nPos.x, nPos.y, False
            End If
            
            'Log
            LogGM UserList(UserIndex).id, anexo & " ( " & nPos.map & ":" & nPos.x & "," & nPos.y & ")", "IRA"
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdConse.LISTUSU
            For TempInt = 1 To LastUser
                If (UserList(TempInt).Name <> "") And UserList(TempInt).flags.Privilegios = 0 Then
                    tempstr = tempstr & UserList(TempInt).Name & ","
                End If
            Next TempInt
            If Len(tempstr) > 7 Then
                tempstr = Left$(tempstr, Len(tempstr) - 2)
            End If
            EnviarPaquete Paquetes.LISTUSU, tempstr, UserIndex, ToIndex
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdConse.CINVISIBLE
            DoAdminInvisible UserIndex
            
            LogGM UserList(UserIndex).id, anexo, "INVI"
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
       ' Case CmdConse.PANELGM
           ' EnviarPaquete Paquetes.ShowGMPANEL, "", UserIndex
        'Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdConse.CTRABAJANDO
            For TempInt = 1 To LastUser
                If (UserList(TempInt).Name <> "") And UserList(TempInt).flags.Trabajando = True Then
                    EnviarPaquete Paquetes.traba, UserList(TempInt).Name, UserIndex
                End If
            Next TempInt
            EnviarPaquete Paquetes.traba, "0", UserIndex, ToIndex
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
       Case CmdConse.Penas
            If anexo = "" Then Exit Sub
            
            Call modComandos.enviarPenas(anexo, UserList(UserIndex))
            Exit Sub
   '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdConse.carcel
        TempInt = NameIndex(ReadField(1, anexo, Asc("@")))
        tempbyte = ReadField(3, anexo, Asc("@"))
        tempstr = ReadField(2, anexo, Asc("@"))
        
        If TempInt <= 0 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(56), UserIndex
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.Privilegios < UserList(TempInt).flags.Privilegios Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(151), UserIndex
            Exit Sub
        End If
        
        If tempbyte > 60 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(152), UserIndex
            Exit Sub
        End If
        
        Encarcelar UserList(TempInt), tempbyte, UserList(UserIndex).Name
        
        If LenB(UserList(TempInt).flags.Penasas) = 0 Then
            UserList(TempInt).flags.Penasas = "Encarcelado " & tempbyte & "m. Razón: " & tempstr & ". Gm: " & UserList(UserIndex).Name & " " & Now
        Else
            UserList(TempInt).flags.Penasas = UserList(TempInt).flags.Penasas & vbCrLf & "Encarcelado " & tempbyte & "m. Razón: " & tempstr & ".Gm: " & UserList(UserIndex).Name & " " & Now
        End If
        
        LogGM UserList(UserIndex).id, UserList(TempInt).Name & " por " & tempstr & " " & tempbyte & " minutos.", "CARCEL"
    
    Exit Sub
     '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdConse.ONLINEMAP
        TempInt = CIntSeguro(anexo)
        
        If TempInt = 0 Then
            TempInt = UserList(UserIndex).pos.map
        End If
        
        If Not SV_PosicionesValidas.existeMapa(TempInt) Then
            EnviarPaquete Paquetes.mensajeinfo, "Número de mapa no válido.", UserIndex, ToIndex
            Exit Sub
        End If
           
        tempstr = listarPersonajesOnline(MapInfo(CInt(anexo)))
        
        EnviarPaquete Paquetes.mensajeinfo, tempstr, UserIndex
    Exit Sub
     '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdConse.GMSG
        EnviarPaquete Paquetes.MensajeGMSG, UserList(UserIndex).Name & "> " & anexo, UserIndex, ToAdmins
        
        LogGM UserList(UserIndex).id, anexo, "GMSG"
    Exit Sub
     '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdConse.RMSG

        EnviarPaquete Paquetes.MensajeTalk, UserList(UserIndex).Name & "> " & anexo, UserIndex, ToAll
        
        LogGM UserList(UserIndex).id, anexo, "RMSG"
    Exit Sub
    
End Select

End Sub


Private Sub ProcesarComando0(ByVal numero As Byte, ByVal UserIndex As Integer, Optional anexo As String)
Rem //////////////////////////////////////////////////////////////////////
Rem/Esta Sub Procesa los comandos, primero los de users
Rem/luego los de conse,semi y dioses...
Rem/Q simpatico el "Rem"
Rem/=)
Rem//////////////////////////////////////////////////////////////////////
Dim tempstr As String
Dim TempStr2 As String
Dim tempLong As Long
Dim bucle As Integer
Dim TempInt As Integer
Dim nPos As WorldPos

'If CheckCrCode(Asc(Mid$(rdata, 2, 1)), Asc(Right$(rdata, 1))) = 1 Then
 '   rdata = Left$(rdata, Len(rdata) - 1)
'Else 'Si el crc existia y estaba Okey, lo quitamos
'    MsgBox "alguien la ta bugueado"
'End If

    If ProfilePaquetes Then
        Logs.logProfilePaquete ("0" & numero)
    End If
    
    Select Case numero
        Case CmdUsers.online
'            If NumUsers < 700 Then
'
'                tempstr = vbNullString
'                TempInt = 0
'
'                For bucle = 1 To LastUser
'                    If UserList(bucle).Name <> "" And UserList(bucle).flags.Privilegios = 0 Then
'                        tempstr = tempstr & UserList(bucle).Name & ", "
'                        TempInt = TempInt + 1
'                    End If
'                Next bucle
'
'                If Len(tempstr) > 2 Then 'Le sacamos la ultima coma
'                tempstr = Left(tempstr, Len(tempstr) - 2)
'                tempstr = tempstr & "."
'                End If
'
'                EnviarPaquete Paquetes.mensajeinfo, tempstr, UserINdex
'                EnviarPaquete Paquetes.MensajeCompuesto, Chr$(26) & TempInt, UserINdex
'
'            Else
            EnviarPaquete Paquetes.MensajeCompuesto, Chr$(26) & NumUsers, UserIndex
'            End If
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.CSALIR
            Call Cerrar_Usuario(UserList(UserIndex))
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.FUNDARCLAN
            Dim clan As cClan
            If UserList(UserIndex).GuildInfo.FundoClan = 1 Then
               Set clan = clanes.getClan(UserList(UserIndex).GuildInfo.ClanFundadoID)
               If clan.getEstado = Activo Then
                    EnviarPaquete Paquetes.MensajeSimple2, Chr$(2), UserIndex
               Else
                    If clan.getIDLider = UserList(UserIndex).id Then
                        EnviarPaquete Paquetes.mensajeinfo, "Ya has fundado el clan " & clan.getNombre & " que fue disuelto por ti. Escribe /reanudarclan para volverlo a activar.", UserIndex, ToIndex
                    Else
                        EnviarPaquete Paquetes.mensajeinfo, "Ya has fundado el clan " & clan.getNombre & " que fue disuelto por su lider.", UserIndex, ToIndex
                    End If
               End If
            ElseIf PuedeCrearClan(UserIndex) Then
                EnviarPaquete Paquetes.initGuildFundation, "", UserIndex
            End If
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.SALIRCLAN
            mdClanes.SalirDeClan (UserIndex)
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.CBALANCE
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  EnviarPaquete Paquetes.MensajeSimple, Chr$(4), UserIndex, ToIndex
                  Exit Sub
            End If
            If distancia(NpcList(UserList(UserIndex).flags.TargetNPC).pos, UserList(UserIndex).pos) > 3 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(7), UserIndex, ToIndex
                Exit Sub
            End If
            Select Case NpcList(UserList(UserIndex).flags.TargetNPC).NPCtype
            Case NPCTYPE_BANQUERO
                EnviarPaquete Paquetes.DescNpc2, ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex) & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta.", UserIndex, ToIndex
            Case NPCTYPE_TIMBERO
                If UserList(UserIndex).flags.Privilegios > 0 Then
                Dim tlong As Long
                Dim N As Long
                    tlong = apuestas.Ganancias - apuestas.Perdidas
                    N = 0
                    If tlong >= 0 And apuestas.Ganancias <> 0 Then
                        N = Int(tlong * 100 / apuestas.Ganancias)
                    End If
                    If tlong < 0 And apuestas.Perdidas <> 0 Then
                        N = Int(tlong * 100 / apuestas.Perdidas)
                    End If
                    EnviarPaquete Paquetes.mensajeinfo, "Entradas: " & apuestas.Ganancias & " Salida: " & apuestas.Perdidas & " Ganancia Neta: " & tlong & " (" & N & "%) Jugadas: " & apuestas.Jugadas, UserIndex, ToIndex
                End If
            End Select
            Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.QUIETO
            'Â¿Clikeo un npc?
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(131), UserIndex
                Exit Sub
            End If
            
            If NpcList(UserList(UserIndex).flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
            
            Call NPCs.ponerEstatico(UserList(UserIndex).flags.TargetNPC)
        
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.ACOMPAÑAR
            'Â¿Clikeo un npc?
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(131), UserIndex
                Exit Sub
            End If
            If NpcList(UserList(UserIndex).flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
            Call FollowAmo(UserList(UserIndex).flags.TargetNPC)
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.ENTRENAR
            'Â¿Clikeo un npc?
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(131), UserIndex
                Exit Sub
            End If
            'Â¿Esta demasiado lejos?
            If distancia(UserList(UserIndex).pos, NpcList(UserList(UserIndex).flags.TargetNPC).pos) > 10 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(4), UserIndex
                Exit Sub
            End If
            If NpcList(UserList(UserIndex).flags.TargetNPC).NPCtype <> NPCTYPE_ENTRENADOR Then Exit Sub
            Call EnviarListaCriaturas(UserIndex, UserList(UserIndex).flags.TargetNPC)
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.CDESCANSAR
            If HayOBJarea(UserList(UserIndex).pos, FOGATA) Then
                EnviarPaquete Paquetes.MDescansar, "", UserIndex
                UserList(UserIndex).flags.Descansar = Not UserList(UserIndex).flags.Descansar
            Else
                If UserList(UserIndex).flags.Descansar Then
                    UserList(UserIndex).flags.Descansar = False
                    EnviarPaquete Paquetes.MDescansar, "", UserIndex
                    Exit Sub
                End If
                EnviarPaquete Paquetes.MensajeSimple2, Chr$(6), UserIndex, ToIndex
            End If
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.CMEDITAR
            Call GamePLay.Meditar(UserList(UserIndex))
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.Resucitar
            'Â¿Clikeo un npc?
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(4), UserIndex
                Exit Sub
            End If
            'Â¿Esta demasiado lejos?
            If distancia(UserList(UserIndex).pos, NpcList(UserList(UserIndex).flags.TargetNPC).pos) > 10 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(4), UserIndex
                Exit Sub
            End If
            'Â¿Es un sacerdote?
            If UserList(UserIndex).flags.TargetNpcTipo <> 1 Then
                EnviarPaquete Paquetes.DescNpc, Chr$(91) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.Muerto = 0 Then Exit Sub
            
            Call RevivirUsuario(UserList(UserIndex), 1, 50, 50)
            
            EnviarPaquete Paquetes.MensajeSimple2, Chr$(41), UserIndex
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.CURAR
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
            '   Call Senddata(ToIndex, UserIndex, 0, "Y4")
               Exit Sub
           End If
           If NpcList(UserList(UserIndex).flags.TargetNPC).NPCtype <> 1 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           If distancia(UserList(UserIndex).pos, NpcList(UserList(UserIndex).flags.TargetNPC).pos) > 10 Then
               EnviarPaquete Paquetes.MensajeSimple, Chr$(8), UserIndex, ToIndex
               Exit Sub
           End If
           UserList(UserIndex).Stats.minHP = UserList(UserIndex).Stats.MaxHP
           Call SendUserStatsBox(val(UserIndex))
           EnviarPaquete Paquetes.MensajeSimple, Chr$(17), UserIndex, ToIndex
           Exit Sub
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.CAYUDA
            SendHelp UserIndex
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.EST
            SendUserStatsTxt UserIndex, UserIndex
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.CCOMERCIAR
             'Â¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(3), UserIndex, ToIndex
                Exit Sub
            End If
            If UserList(UserIndex).flags.Comerciando Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(27), UserIndex, ToIndex
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.Privilegios = 1 Then
                EnviarPaquete Paquetes.MensajeSimple2, Chr$(118), UserIndex, ToIndex
                Exit Sub
            End If
            'Â¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                  'Â¿El NPC puede comerciar?
                  If NpcList(UserList(UserIndex).flags.TargetNPC).Comercia = 0 Then
                     If Len(NpcList(UserList(UserIndex).flags.TargetNPC).desc) > 0 Then EnviarPaquete Paquetes.DescNpc, Chr$(3) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex, ToPCArea
                     Exit Sub
                  End If
                  If distancia(NpcList(UserList(UserIndex).flags.TargetNPC).pos, UserList(UserIndex).pos) > 3 Then
                      EnviarPaquete Paquetes.MensajeSimple, Chr$(7), UserIndex, ToIndex
                      Exit Sub
                  End If
                  'Iniciamos la rutina pa' comerciar.
                  Call IniciarCOmercioNPC(UserIndex)
             '[Alejo]
            ElseIf UserList(UserIndex).flags.TargetUser > 0 Then
                'Call SendData(ToIndex, UserIndex, 0, "||COMERCIO SEGURO ENTRE USUARIOS TEMPORALMENTE DESHABILITADO" & FONTTYPE_INFO)
                'Exit Sub
                'Comercio con otro usuario
                'Puede comerciar ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 1 Then
                    EnviarPaquete Paquetes.MensajeSimple2, Chr$(9), UserIndex, ToIndex
                    Exit Sub
                End If
                 If UserList(UserIndex).flags.Navegando = 1 Then
                    EnviarPaquete Paquetes.MensajeSimple2, Chr$(42), UserIndex, ToIndex
                    Exit Sub
                 End If
                'soy yo ?
                If UserList(UserIndex).flags.TargetUser = UserIndex Then
                    EnviarPaquete Paquetes.MensajeSimple2, Chr$(10), UserIndex, ToIndex
                    Exit Sub
                End If
                'ta muy lejos ?
                If distancia(UserList(UserList(UserIndex).flags.TargetUser).pos, UserList(UserIndex).pos) > 3 Then
                    EnviarPaquete Paquetes.MensajeSimple2, Chr$(43), UserIndex, ToIndex
                    Exit Sub
                End If
                '[Wizard 03/09/05]Es Consejero????o yo Soy consejero?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Privilegios = 1 Then
                    EnviarPaquete Paquetes.MensajeSimple2, Chr$(118), UserIndex, ToIndex
                    Exit Sub
                End If
                '[/Wizard]
                'Ya ta comerciando ? es con migo o con otro ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Comerciando = True And _
                    UserList(UserList(UserIndex).flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                    EnviarPaquete Paquetes.MensajeSimple2, Chr$(11), UserIndex, ToIndex
                    Exit Sub
                End If
                'inicializa unas variables...
                UserList(UserIndex).ComUsu.DestUsu = UserList(UserIndex).flags.TargetUser
                UserList(UserIndex).ComUsu.DestNick = UserList(UserList(UserIndex).flags.TargetUser).Name
                'UserList(UserIndex).ComUsu.cant(TempInt2) = 0
                'UserList(UserIndex).ComUsu.Objeto(1) = 0
                UserList(UserIndex).ComUsu.Acepto = False
                'Rutina para comerciar con otro usuario
                Call IniciarComercioConUsuario(UserIndex, UserList(UserIndex).flags.TargetUser)
            Else
                EnviarPaquete Paquetes.MensajeSimple2, Chr$(44), UserIndex, ToIndex
            End If
            Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.BOVEDA
            'Â¿Clikeo un npc?
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                EnviarPaquete Paquetes.MensajeSimple2, Chr$(44), UserIndex
                Exit Sub
            End If
            'Â¿Esta demasiado lejos?
            If distancia(UserList(UserIndex).pos, NpcList(UserList(UserIndex).flags.TargetNPC).pos) > 5 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(7), UserIndex
                Exit Sub
            End If
            'Â¿Es un banquero?
            If UserList(UserIndex).flags.TargetNpcTipo <> NPCTYPE_BANQUERO Then
                EnviarPaquete Paquetes.DescNpc, Chr$(90) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
                Exit Sub
            End If
            IniciarDeposito UserIndex
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.ENLISTAR
            #If TDSFacil = 0 Then
                Call ModFacciones.EnlistarPersonaje(UserList(UserIndex))
            #End If
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.informacion
            'Â¿Clikeo un npc?
            
            tempstr = Trim(anexo)
        
            If Len(anexo) = 0 Then
                If UserList(UserIndex).flags.TargetNPC = 0 Then
                    EnviarPaquete Paquetes.MensajeSimple, Chr$(131), UserIndex
                    Exit Sub
                End If
                'Â¿Esta demasiado lejos?
                If distancia(UserList(UserIndex).pos, NpcList(UserList(UserIndex).flags.TargetNPC).pos) > 10 Then
                    EnviarPaquete Paquetes.MensajeSimple, Chr$(4), UserIndex
                    Exit Sub
                End If
                'Â¿Es un lider?
                If UserList(UserIndex).flags.TargetNpcTipo <> NPCTYPE_NOBLE Then
                    EnviarPaquete Paquetes.DescNpc, Chr$(92) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
                    Exit Sub
                End If
                If NpcList(UserList(UserIndex).flags.TargetNPC).faccion = eAlineaciones.Real Then
                'Es el rey:|
                    If UserList(UserIndex).faccion.ArmadaReal = 0 Then
                        EnviarPaquete Paquetes.DescNpc, Chr$(16) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
                    Else
                        EnviarPaquete Paquetes.DescNpc, Chr$(17) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
                    End If
                ElseIf NpcList(UserList(UserIndex).flags.TargetNPC).faccion = eAlineaciones.caos Then
                    If UserList(UserIndex).faccion.FuerzasCaos = 0 Then
                        EnviarPaquete Paquetes.DescNpc, Chr$(18) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
                    Else
                        EnviarPaquete Paquetes.DescNpc, Chr$(19) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
                    End If
                End If
            End If
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.RECOMPENSA
            'Â¿Clikeo un npc?
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(131), UserIndex
                Exit Sub
            End If
            'Â¿Esta demasiado lejos?
            If distancia(UserList(UserIndex).pos, NpcList(UserList(UserIndex).flags.TargetNPC).pos) > 10 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(4), UserIndex
                Exit Sub
            End If
            'Â¿Es un lider?
            If UserList(UserIndex).flags.TargetNpcTipo <> NPCTYPE_NOBLE Then
                EnviarPaquete Paquetes.DescNpc, Chr$(92) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
                Exit Sub
            End If
            If NpcList(UserList(UserIndex).flags.TargetNPC).faccion = eAlineaciones.Real Then
            'Es el rey:|
                If UserList(UserIndex).faccion.ArmadaReal = 0 Then
                    EnviarPaquete Paquetes.DescNpc, Chr$(16) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
                Else
                    RecompensaArmadaReal UserIndex
                End If
            ElseIf NpcList(UserList(UserIndex).flags.TargetNPC).faccion = eAlineaciones.caos Then
                If UserList(UserIndex).faccion.FuerzasCaos = 0 Then
                    EnviarPaquete Paquetes.DescNpc, Chr$(18) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
                Else
                    RecompensaCaos UserIndex
                End If
            End If
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.CMOTD
            SendMOTD UserIndex
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
       Case CmdUsers.CMSG
            'Â¿Tiene clan?
            If UserList(UserIndex).GuildInfo.id > 0 Then
                    EnviarPaquete Paquetes.MensajeClan1, UserList(UserIndex).Name & "> " & anexo, UserIndex, ToGuildMembers
            End If
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.pmsg
        Call mdParty.BroadCastParty(UserIndex, UserList(UserIndex).Name & "> " & Replace(anexo, "~", " "))
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.ONLINECLAN
            'Â¿Tiene clan?
            If UserList(UserIndex).GuildInfo.id = 0 Then Exit Sub
            
            Dim IntegrantesOnline As EstructurasLib.ColaConBloques
            Set IntegrantesOnline = UserList(UserIndex).ClanRef.getIntegrantesOnline
            
            IntegrantesOnline.itIniciar
            
            Do While IntegrantesOnline.ithasNext
                        tempstr = tempstr & UserList(IntegrantesOnline.itnext).Name & ", "
            Loop
            
            'Le sacamos la ultima coma..
            tempstr = Left$(tempstr, Len(tempstr) - 2)
            tempstr = tempstr & ". (" & UserList(UserIndex).ClanRef.getCantidadOnline() & ")"
            EnviarPaquete Paquetes.MensajeClan1, tempstr, UserIndex
        Exit Sub
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.GM
        If UserList(UserIndex).Stats.ELV > 4 Then
            anexo = mysql_real_escape_string(anexo)
            sql = "INSERT INTO " & DB_NAME_PRINCIPAL & ".sos(Usuario,Mail,Fecha,Mensaje) values('" & UserList(UserIndex).Name & "','" & UserList(UserIndex).Email & "','" & Now & "','" & anexo & "')"


                  'If Ayuda.Existe(UserList(UserIndex).Name) Then
                  '    Ayuda.Quitar UserList(UserIndex).Name
                  '    Ayuda.Push anexo, UserList(UserIndex).Name
                  'Else
                  '    Ayuda.Push anexo, UserList(UserIndex).Name
                  'End If
            conn.Execute sql, , adExecuteNoRecords
          End If
          Exit Sub
          '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.CDESC
            If Not AsciiValidos(anexo) Then
               EnviarPaquete Paquetes.MensajeSimple2, Chr$(15), UserIndex, ToIndex
                 Exit Sub
              End If
            If Len(anexo) > 50 Then
            EnviarPaquete Paquetes.mensajeinfo, "La descripción solicitada es muy larga.", UserIndex, ToIndex
              Exit Sub
              Else
            EnviarPaquete Paquetes.MensajeSimple2, Chr$(16), UserIndex, ToIndex
            UserList(UserIndex).desc = anexo
              End If
          Exit Sub
          '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.VOTO
            If UserList(UserIndex).GuildInfo.id > 0 Then
                Call mdClanes.Votar(UserIndex, Trim(anexo))
              Else
                  'TO-DO No perteneces a ningun clan..
              End If
          Exit Sub
          '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.PASSWD
            tempstr = ReadField(1, anexo, Asc("@"))
            TempStr2 = ReadField(2, anexo, Asc("@"))
            If MD5String(tempstr) = UserList(UserIndex).Password Then
                If Len(TempStr2) < 6 Then
                    EnviarPaquete Paquetes.MensajeSimple2, Chr$(17), UserIndex
                  Else
                    EnviarPaquete Paquetes.MensajeSimple2, Chr$(18), UserIndex
                    UserList(UserIndex).Password = MD5String(TempStr2)
                  End If
              Else
            EnviarPaquete Paquetes.mensajeinfo, "El password ingresado no pertenece al personaje.", UserIndex, ToIndex
              End If
              Exit Sub
          '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.Retirar
              'Â¿Clikeo un npc?
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(4), UserIndex
                  Exit Sub
              End If
              'Â¿Esta demasiado lejos?
            If distancia(UserList(UserIndex).pos, NpcList(UserList(UserIndex).flags.TargetNPC).pos) > 10 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(5), UserIndex
                  Exit Sub
              End If
              'Quiere retirar del banco?
            If UserList(UserIndex).flags.TargetNpcTipo = NPCTYPE_BANQUERO Then
                If val(anexo) > 0 And val(anexo) <= UserList(UserIndex).Stats.Banco Then
                    UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(anexo)
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + val(anexo)
                    EnviarPaquete Paquetes.DescNpc, Chr$(26) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex) & UserList(UserIndex).Stats.Banco & ",", UserIndex
                    SendUserStatsBox (UserIndex)
                  Else
                    EnviarPaquete Paquetes.DescNpc, Chr$(27) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
                  End If
              'Quiere salir de la faccion?
            ElseIf UserList(UserIndex).flags.TargetNpcTipo = NPCTYPE_NOBLE Then
                If UserList(UserIndex).faccion.ArmadaReal = 1 Then
                    If NpcList(UserList(UserIndex).flags.TargetNPC).faccion = eAlineaciones.Real Then
                        Call ExpulsarFaccionReal(UserIndex)
                        EnviarPaquete Paquetes.DescNpc, Chr$(20) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
                      Else
                        EnviarPaquete Paquetes.DescNpc, Chr$(21) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
                      End If
                ElseIf UserList(UserIndex).faccion.FuerzasCaos = 1 Then
                    If NpcList(UserList(UserIndex).flags.TargetNPC).faccion = eAlineaciones.caos Then
                        Call ExpulsarFaccionCaos(UserIndex)
                        EnviarPaquete Paquetes.DescNpc, Chr$(22) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
                      Else
                        EnviarPaquete Paquetes.DescNpc, Chr$(23) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
                      End If
                  Else
                    EnviarPaquete Paquetes.DescNpc, Chr$(25) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
                  End If
              End If
          Exit Sub
          '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.Depositar
              'Â¿Clikeo un npc?
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(4), UserIndex
                  Exit Sub
              End If
              'Â¿Esta demasiado lejos?
            If distancia(UserList(UserIndex).pos, NpcList(UserList(UserIndex).flags.TargetNPC).pos) > 10 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(5), UserIndex
                  Exit Sub
              End If
            If UserList(UserIndex).flags.TargetNpcTipo = NPCTYPE_BANQUERO Then
                If val(anexo) > 0 And val(anexo) <= UserList(UserIndex).Stats.GLD Then
                    UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco + val(anexo)
                    EnviarPaquete Paquetes.DescNpc, Chr$(26) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex) & UserList(UserIndex).Stats.Banco & ",", UserIndex
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(anexo)
                    EnviarPaquete Paquetes.EnviarOro, Codify(UserList(UserIndex).Stats.GLD), UserIndex, ToIndex
                  Else
                    EnviarPaquete Paquetes.DescNpc, Chr$(27) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
                  End If
              End If
              Exit Sub
          '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
          'Case CmdUsers.ONLINEP
          '    TempInt = NameIndex(Anexo)
          '    If TempInt > 0 Then
          '        EnviarPaquete Paquetes.MensajeSimple, Chr$(197), UserIndex
          '    Else
          '        EnviarPaquete Paquetes.MensajeSimple, Chr$(56), UserIndex
          '    End If
          'Exit Sub
          '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
          'Case CmdUsers.PASARORO
              'Â¿Clickeo un usuario?
           '   If UserList(UserIndex).flags.TargetUser = 0 Then
           '       EnviarPaquete Paquetes.MensajeSimple, Chr$(131), UserIndex
           '       Exit Sub
           '   End If
           '   TempInt = NameIndex(UserList(UserIndex).flags.TargetUser)
           '   Anexo = DeCodify(Anexo)
           '  If val(Anexo) <= 0 Then Exit Sub
           '   If TempInt <= 0 Then
           '       EnviarPaquete Paquetes.MensajeSimple, Chr$(56), UserIndex
           '       Exit Sub
           '  End If
           '   If UserList(UserIndex).Stats.GLD < val(Anexo) Then
           '       EnviarPaquete Paquetes.MensajeSimple, Chr$(40), UserIndex
           '       Exit Sub
           '   End If
            
           '   If UserList(UserIndex).flags.Privilegios > 0 Then
           '       LogGM UserList(UserIndex).Name, "Paso oro a " & UserList(TempInt).Name
           '   End If
            
           '   UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(Anexo)
           '   UserList(TempInt).Stats.GLD = UserList(TempInt).Stats.GLD + val(Anexo)
         ' Exit Sub
          '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.MOVER
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(131), UserIndex
                  Exit Sub
              End If
              'Â¿Esta demasiado lejos?
            If distancia(UserList(UserIndex).pos, UserList(UserList(UserIndex).flags.TargetUser).pos) > 4 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(4), UserIndex
                  Exit Sub
              End If
            TempInt = NameIndex(UserList(UserIndex).flags.TargetUser)
            If TempInt <= 0 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(56), UserIndex
                  Exit Sub
              End If
            If MapData(UserList(TempInt).pos.map, UserList(TempInt).pos.x, UserList(TempInt).pos.y).OBJInfo.ObjIndex = 0 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(205), UserIndex
                  Exit Sub
              End If
            Call ClosestLegalPos(UserList(TempInt).pos, nPos, UserList(UserIndex))
            If nPos.x <> 0 And nPos.y <> 0 Then
                Call WarpUserChar(TempInt, nPos.map, nPos.x, nPos.y, False)
              End If
          Exit Sub
          '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.DENUNCIAR
            If denunciarActivado Then
                EnviarPaquete Paquetes.MensajeClan1, UserList(UserIndex).Name & " denuncia: " & anexo, UserIndex, ToAdmins
                  'Agrego la denuncia en los logs de los gms que estan logeados, para ver si estan
                  'de vagos o no.
                  'For TempInt = 1 To GmsGroup.Count
                  '    If Not GmsGroup.Item(TempInt) = 0 Then
                  '        If UserList(GmsGroup.Item(TempInt)).flags.Privilegios > 0 Then Call LogGM(UserList(GmsGroup.Item(TempInt)).id, UserList(UserIndex).Name & " denuncia: " & anexo)
                  '    End If
                  'Next TempInt
                  'Lo guardo como log general, para todos.
                Call LogGM(0, UserList(UserIndex).Name & " denuncia: " & anexo)
                
                If Ayuda.existeElemento(UserIndex) = 0 Then
                    Ayuda.agregar UserIndex
                  End If
              Else
                EnviarPaquete Paquetes.MensajeFight, "Sistema de denunciar desactivado temporalmente.", UserIndex, ToIndex
              End If
          Exit Sub
          '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.centinela
            Call modCentinelas.ponerCodigo(UserList(UserIndex), anexo)
          Exit Sub
          '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.cheque
             Call cobrarCheque(UserList(UserIndex), anexo)
          Exit Sub
          '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.APOSTAR
            N = CLng(val(anexo))
            If N > 32000 Then N = 32000
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  'Se asegura que el target es un npc
                EnviarPaquete Paquetes.MensajeSimple, Chr$(4), UserIndex, ToIndex
            ElseIf distancia(NpcList(UserList(UserIndex).flags.TargetNPC).pos, UserList(UserIndex).pos) > 10 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(5), UserIndex, ToIndex
            ElseIf NpcList(UserList(UserIndex).flags.TargetNPC).NPCtype <> NPCTYPE_TIMBERO Then
                EnviarPaquete Paquetes.DescNpc2, ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex) & "No tengo ningun interes en apostar.", UserIndex, ToIndex
            ElseIf N < 1 Then
                EnviarPaquete Paquetes.DescNpc2, ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex) & "El minimo de apuesta es 1 moneda.", UserIndex, ToIndex
            ElseIf N > 5000 Then
                EnviarPaquete Paquetes.DescNpc2, ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex) & "El maximo de apuesta es 5000 monedas.", UserIndex, ToIndex
            ElseIf UserList(UserIndex).Stats.GLD < N Then
                EnviarPaquete Paquetes.DescNpc2, ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex) & "No tienes esa cantidad.", UserIndex, ToIndex
              Else
                If RandomNumber(1, 100) <= 47 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + N
                    EnviarPaquete Paquetes.DescNpc2, ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex) & "Felicidades! Has ganado " & CStr(N) & " monedas de oro!", UserIndex, ToIndex
                    apuestas.Perdidas = apuestas.Perdidas + N
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(apuestas.Perdidas))
                  Else
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - N
                    EnviarPaquete Paquetes.DescNpc2, ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex) & "Lo siento, has perdido " & CStr(N) & " monedas de oro.", UserIndex, ToIndex
                    apuestas.Ganancias = apuestas.Ganancias + N
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(apuestas.Ganancias))
                  End If
                apuestas.Jugadas = apuestas.Jugadas + 1
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(apuestas.Jugadas))
                EnviarPaquete Paquetes.EnviarOro, Codify(UserList(UserIndex).Stats.GLD), UserIndex, ToIndex
              End If
          Exit Sub
          '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.Activar
            If charlageneral Then
                If UserList(UserIndex).Stats.GlobAl = 2 Then
                    UserList(UserIndex).Stats.GlobAl = 0
                    EnviarPaquete Paquetes.mensajeinfo, "No podras enviar ni recibir mensajes globales.", UserIndex, ToIndex
                  Else
                    UserList(UserIndex).Stats.GlobAl = 2
                    EnviarPaquete Paquetes.mensajeinfo, "Podras escribir y recibir mensajes globales.", UserIndex, ToIndex
                  End If
              Else
                EnviarPaquete Paquetes.mensajeinfo, "El chat global no esta disponible.", UserIndex, ToIndex
              End If
              Exit Sub
          '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.Ping
            EnviarPaquete Paquetes.Pang, "", UserIndex, ToIndex
              Exit Sub
          '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.RetarS
            Call modRetos.aceptarSolicitud(UserIndex, anexo)
              Exit Sub
          '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.rechazar
            Call modRetos.rechazarSolicitud(UserList(UserIndex), anexo)
              Exit Sub
          '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.AcomodarPorcentajesDeParty
            Call mdParty.AcomodarP(UserIndex, anexo)
              Exit Sub
          '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.Retirartodo
              'Â¿Clikeo un npc?
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(4), UserIndex
                  Exit Sub
              End If
              'Â¿Esta demasiado lejos?
            If distancia(UserList(UserIndex).pos, NpcList(UserList(UserIndex).flags.TargetNPC).pos) > 10 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(5), UserIndex
                  Exit Sub
              End If
            If UserList(UserIndex).flags.TargetNpcTipo = NPCTYPE_BANQUERO Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.Banco + UserList(UserIndex).Stats.GLD
                    UserList(UserIndex).Stats.Banco = 0
                    EnviarPaquete Paquetes.DescNpc, Chr$(26) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex) & UserList(UserIndex).Stats.Banco & ",", UserIndex
                    EnviarPaquete Paquetes.EnviarOro, Codify(UserList(UserIndex).Stats.GLD), UserIndex
              End If
          Exit Sub
          '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        Case CmdUsers.DepositarTodo
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(4), UserIndex
                  Exit Sub
              End If
              'Â¿Esta demasiado lejos?
            If distancia(UserList(UserIndex).pos, NpcList(UserList(UserIndex).flags.TargetNPC).pos) > 10 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(5), UserIndex
                  Exit Sub
              End If
            If UserList(UserIndex).flags.TargetNpcTipo = NPCTYPE_BANQUERO Then
                    UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco + UserList(UserIndex).Stats.GLD
                    EnviarPaquete Paquetes.DescNpc, Chr$(26) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex) & UserList(UserIndex).Stats.Banco & ",", UserIndex
                    UserList(UserIndex).Stats.GLD = 0
                    EnviarPaquete Paquetes.EnviarOro, ByteToString(0), UserIndex, ToIndex
              End If
          Exit Sub
      '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdUsers.Abandonar
          'El abandonar sirve para abandonar mascotas o abandonar un evento
        
          'Se asegura que el target es un npc
          If UserList(UserIndex).flags.TargetNPC > 0 Then
              'Se asegura que el npc sea mascota del usuario
            If NpcList(UserList(UserIndex).flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
              'Elimina al npc
            Call QuitarNPC(UserList(UserIndex).flags.TargetNPC)
              'Mensaje de que ha abandoado a la mascota
            EnviarPaquete Paquetes.MensajeSimple2, Chr$(100), UserIndex, ToIndex
            Else
            If UserList(UserIndex).evento Is Nothing Then
                  'No tiene marcado a un npc ni esta en un evento.
                EnviarPaquete Paquetes.MensajeSimple2, Chr$(96), UserIndex, ToIndex: Exit Sub
            ElseIf UserList(UserIndex).evento.getEstadoEvento = eEstadoEvento.Desarrollandose Then
                  'Esta en un evento, entonces lo abandono
                Call UserList(UserIndex).evento.usuarioAbandono(UserIndex)
              Else 'No esta en un evento el cual esta desarrollandose
                EnviarPaquete Paquetes.MensajeSimple2, Chr$(96), UserIndex, ToIndex: Exit Sub
              End If
            End If
          
          Exit Sub
      '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdUsers.Fianza
        Call modFianza.pagarFianza(UserList(UserIndex), UCase$(anexo))
      Exit Sub
      '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdUsers.PartyPorcecntaje
        Call mdParty.Acomodar(UserIndex)
      Exit Sub
      '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Case CmdUsers.Penas
        If LenB(UserList(UserIndex).flags.Penasas) > 0 Then
            EnviarPaquete Paquetes.mensajeinfo, UserList(UserIndex).flags.Penasas, UserIndex, ToIndex
         Else
            EnviarPaquete Paquetes.mensajeinfo, "No tienes penas.", UserIndex, ToIndex
         End If
      Exit Sub
      '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Case CmdUsers.disolverclan
    mdClanes.disolverclan UserIndex
  Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Case CmdUsers.ReanudarClan
    If Trim$(anexo) = "" Then
        mdClanes.ReanudarClan UserIndex
      Else
        mdClanes.ReanudarClan UserIndex, Trim$(anexo)
      End If
  Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Case CmdUsers.Participar
    Call modSolicitudesEventos.crear(UserIndex, Trim(anexo))
  Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Case CmdUsers.aceptar
    Call modSolicitudesEventos.aceptarSolicitud(UserIndex, Trim(anexo))
  Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Case CmdUsers.eventosInfo
    If Len(anexo) = 0 Then
        Call modEventos.enviarListaTorneos(UserIndex)
      Else
        Call modEventos.enviarInformacionEvento(UserIndex, anexo)
      End If
  Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Case CmdUsers.decirEnTorneo
    Debug.Print "Minuto en torneo"
  Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Case CmdUsers.minutoEnTorneo
    Debug.Print "Minuto en torneo"
  Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Case CmdUsers.tiempo
      #If TDSFacil = 1 Then
    Call enviarTiempoGratisTDSF(UserList(UserIndex))
      #End If
  Exit Sub
  '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
  End Select

End Sub

'---------------------------------------------------------------------------------------
' Procedure : EnviarPaquete
' DateTime  : 18/02/2007 20:02
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub EnviarPaquete(ByVal paquete As Byte, ByVal Argumentos As String, ByVal sndIndex As Integer, Optional sndRoute As Byte = 0, Optional ByVal map As Integer, Optional ByVal nick As String)
    Senddata sndRoute, sndIndex, map, Chr$(paquete) & Argumentos
End Sub


'---------------------------------------------------------------------------------------
' Procedure : StringToByte
' DateTime  : 18/02/2007 20:03
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
'CSEH: Nada
Public Function StringToByte(ByVal str As String, ByVal Start As Byte) As Byte
    If Len(str) < Start Then Exit Function
    
    StringToByte = Asc(mid$(str, Start, 1))
End Function
'CSEH: Nada
Public Function ITS(ByVal Var As Integer, Optional llamada As Byte = 0) As String
    Dim temp As String
       
    'Convertimos a hexa
    temp = Hex$(Var)
    
    'Nos aseguramos tenga 4 Bytes de largo
    While Len(temp) < 4
        temp = "0" & temp
    Wend
    
    'Convertimos a string
    ITS = Chr$(val("&H" & Left$(temp, 2))) & Chr$(val("&H" & Right$(temp, 2)))
End Function

'---------------------------------------------------------------------------------------
' Procedure : STI
' DateTime  : 18/02/2007 20:03
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
'CSEH: Nada
Public Function STILong(ByVal str As String, ByVal Start As Long) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    Dim tempstr As String
    
    If Len(str) < Start - 1 Then Exit Function
    
    'Convertimos a hexa el valor ascii del segundo Byte
    tempstr = Hex$(Asc(mid$(str, Start + 1, 1)))
    
    'Nos aseguramos tenga 2 Bytes (los ceros a la izquierda cuentan por ser el segundo Byte)
    While Len(tempstr) < 2
        tempstr = "0" & tempstr
    Wend
    
    'Convertimos a integer
    STILong = val("&H" & Hex$(Asc(mid$(str, Start, 1))) & tempstr)
End Function

'---------------------------------------------------------------------------------------
' Procedure : STI
' DateTime  : 18/02/2007 20:03
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
'CSEH: Nada
Public Function STI(ByVal str As String, ByVal Start As Byte) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    Dim tempstr As String
    
    'Asergurarse sea válido
    If Len(str) < Start - 1 Then Exit Function
    
    'Convertimos a hexa el valor ascii del segundo Byte
    tempstr = Hex$(Asc(mid$(str, Start + 1, 1)))
    
    'Nos aseguramos tenga 2 Bytes (los ceros a la izquierda cuentan por ser el segundo Byte)
    While Len(tempstr) < 2
        tempstr = "0" & tempstr
    Wend
    
    'Convertimos a integer
    STI = val("&H" & Hex$(Asc(mid$(str, Start, 1))) & tempstr)

End Function


'---------------------------------------------------------------------------------------
' Procedure : CrearPersonaje
' DateTime  : 18/02/2007 20:00
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub crearPersonaje(ByVal UserIndex As Integer, ByVal Name As String, ByVal Password As String, ByVal Email As String, ByVal Genero As Byte, ByVal clase As Byte, ByVal Raza As Byte, ByVal alineacion As Byte)
Dim tempbyte As Byte
Dim tempbyte2 As Byte
Dim tempByte3 As Byte
Dim IDCuenta As Long
Dim idPersonaje As Long
Dim loopAtributos As Byte
Dim MiInt As Single
 
If Not isNombreValido(Name) Then
    If Not CloseSocket(UserIndex) Then Call LogError("create nombre invalido")
    Exit Sub
End If

'//Verifica que el personaje ya exista
If modUsuarios.existePersonaje_Nombre(Name) Then
    EnviarPaquete Paquetes.mbox, Chr$(9), UserIndex, ToIndex
    Exit Sub
End If

If modPersonaje_Repository.isNickInapropiado(Name) Then
    EnviarPaquete mbox, Chr$(14) & "El nombre elegido no cumple con los requisitos de nombres. ¡Usa tu creatividad!", UserIndex, ToIndex
    Exit Sub
End If

#If TDSFacil Then
    '// Tierras Del Sur Facil SOLO PARA PREMIUMS
    ' Me fijo que este en la lista de nicks habuilitados
    Dim infoCuenta As modCuentas.tInfoCuenta
    
    infoCuenta = obtenerInfoCuentaByMail(Email)

    'El nombre esta reservado?
    If infoCuenta.id = -1 Then
        'EnviarPaquete Paquetes.mbox, Chr$(15), UserIndex, ToIndex
         EnviarPaquete mbox, Chr$(14) & "No hay ninguna cuenta registrada al mail ingresado.", UserIndex, ToIndex
        Exit Sub
    End If
    
    'Si esta en la lista me fijo si esta es la persona habilitada
    ' para crear el personaje, osea si el PIN y el MAIL ingresads a la
    'hora del crear el personaje son los de la cuenta
    If infoCuenta.id = 0 Then
        EnviarPaquete mbox, Chr$(14) & "Hay un error con la reseva del personaje. Envía soporte.", UserIndex, ToIndex
        LogError ("se quiere crear uun personaje que tiene id cuenta 0. nick:" & Name)
        Exit Sub
    End If
    
    'La cuenta existe?
    If infoCuenta.id = -2 Then
        EnviarPaquete mbox, Chr$(14) & "Hay un problema con la Cuenta. Por favor, enviá soporte.", UserIndex, ToIndex
        Exit Sub
    End If
    
    IDCuenta = infoCuenta.id
    
    'Â¿Tiene permiso para crear este nick?
    'El mail puede variar en las mayusculas y minusculas
    'If UCase$(infoCuenta.mail) <> UCase$(Email) Then
    '    EnviarPaquete Paquetes.mbox, Chr$(16), UserIndex, ToIndex
    '    Exit Sub
    'End If
    
    'Â¿Es premium? El server es solo para premium o gente que tiene horas free
    'If Not infoCuenta.Premium Then
    '    If infoCuenta.segundosTDSF <= 0 Then
            'Liberamos la memoria
   '         EnviarPaquete mbox, Chr$(18), UserIndex
   '         Exit Sub
   '     End If
   ' End If
    
    'Esta bloqueada?
    If infoCuenta.bloqueada Then
        EnviarPaquete mbox, Chr$(14) & "Tu cuenta se encuentra BLOQUEADA. Para más información ingresá a tu Cuenta.", UserIndex, ToIndex
        Exit Sub
    End If

    'Esta activada?
    'If Not infoCuenta.Estado = "ACTIVADA" Then
    '    EnviarPaquete mbox, Chr$(14) & "Antes de ingresar al juego debes activar tu Cuenta.", UserIndex, ToIndex
    '    Exit Sub
    'End If
#End If

idPersonaje = modUsuarios.crearPersonaje(Name)

With UserList(UserIndex)
    
    .id = idPersonaje
    .Name = Name
        
    .FechaIngreso = Now
    
    .Raza = razaToByte(listaRazas(Raza))
    .clase = claseToByte(modClases.clasesConfig(clase).nombre)

    .ClaseNumero = clase
        
    .Password = Password
    .Email = Email
    .Stats.MaxItems = 20
    .Stats.SkillPts = 10
    '.pin = pin
        
    #If TDSFacil Then
        .IDCuenta = IDCuenta
        .Premium = True
    #Else
        .IDCuenta = 0
        .Premium = False
    #End If
    
    'Seteamos Genero y Hogar de una manera burda porque marce no me dio pelota
    If Genero = 1 Then
        .Genero = eGeneros.Hombre
    Else
        .Genero = eGeneros.Mujer
    End If
    

    .Hogar = "Nix"

    .flags.Muerto = 0

    .Reputacion.AsesinoRep = 0
    .Reputacion.BandidoRep = 0
    .Reputacion.BurguesRep = 0
    .Reputacion.LadronesRep = 0
    .Reputacion.NobleRep = 1000
    .Reputacion.PlebeRep = 30
    .Reputacion.promedio = 30 / 6

    For loopAtributos = 1 To NUMATRIBUTOS
        .Stats.UserAtributos(loopAtributos) = Constantes_Generales.razasConfig(razaToConfigID(.Raza)).atributos(loopAtributos)
        .Stats.UserAtributosBackUP(loopAtributos) = .Stats.UserAtributos(loopAtributos)
    Next
        
    Call modPersonaje_Creacion.GenerarCuerpoYCabeza(UserList(UserIndex))

    .Char.heading = eHeading.SOUTH
    .OrigChar = .Char
    .Char.WeaponAnim = 7
    .Char.ShieldAnim = NingunEscudo
    .Char.CascoAnim = NingunCasco

    Call getStatsIniciales(UserList(UserIndex), .Stats.MaxHP, .Stats.MaxSta, .Stats.MaxMAN, .Stats.MinHIT, .Stats.MaxHIT)
    
    ' Arranca con todo en el maximo
    .Stats.minHP = .Stats.MaxHP
    .Stats.MinSta = .Stats.MaxSta
    .Stats.MinMAN = .Stats.MaxMAN
    
    .Stats.MaxAGU = 100
    .Stats.minAgu = 100
    .Stats.MaxHam = 100
    .Stats.minham = 100

    'Hechizos que ya vienen incluidos
    If .clase = eClases.Mago Or _
        .clase = eClases.Clerigo Or _
        .clase = eClases.Druida Or _
        .clase = eClases.Bardo Or _
        .clase = eClases.asesino Then
            .Stats.UserHechizos(1) = 2
    End If


    .Stats.GLD = 0
    .Stats.Exp = 0
    .Stats.ELV = 1
        
    .Stats.ELU = Constantes_Generales.obtenerExperienciaNecesaria(.Stats.ELV)

'???????????????? INVENTARIO Â¿Â¿Â¿Â¿Â¿Â¿Â¿Â¿Â¿Â¿Â¿Â¿Â¿Â¿Â¿Â¿Â¿Â¿Â¿Â¿
    .Invent.NroItems = 4

    .Invent.Object(1).ObjIndex = 467
    .Invent.Object(1).Amount = 150

    .Invent.Object(2).ObjIndex = 468
    .Invent.Object(2).Amount = 150

    .Invent.Object(3).ObjIndex = 460
    .Invent.Object(3).Amount = 1
    .Invent.Object(3).Equipped = 1

    ' Vestimenta
    If .Genero = eGeneros.Mujer Then
        If .Raza = eRazas.Gnomo Or .Raza = eRazas.Enano Then
            .Invent.Object(4).ObjIndex = 683
        Else
            .Invent.Object(4).ObjIndex = 664
        End If
    ElseIf .Genero = eGeneros.Hombre Then
        Select Case .Raza
            Case eRazas.Humano
                .Invent.Object(4).ObjIndex = 463
            Case eRazas.Elfo
                .Invent.Object(4).ObjIndex = 464
            Case eRazas.ElfoOscuro
                .Invent.Object(4).ObjIndex = 465
            Case eRazas.Enano
                .Invent.Object(4).ObjIndex = 466
            Case eRazas.Gnomo
                .Invent.Object(4).ObjIndex = 466
        End Select
    End If

    .Invent.Object(4).Amount = 1
    .Invent.Object(4).Equipped = 1
    
    
    .Invent.Object(5).ObjIndex = POCION_ROJA_NEWBIE
    .Invent.Object(5).Amount = 1000
    
    .Invent.Object(6).ObjIndex = POCION_VIOLETA
    .Invent.Object(6).Amount = 100

    .Invent.Object(7).ObjIndex = POCION_AMARILLA_NEWBIE
    .Invent.Object(7).Amount = 500
    
    .Invent.Object(8).ObjIndex = POCION_VERDE_NEWBIE
    .Invent.Object(8).Amount = 500
    
    ' Clases Magicas arrancan con pociones
    If .clase = eClases.Mago Or .clase = eClases.Druida _
        Or .clase = eClases.Clerigo Or .clase = eClases.Bardo _
        Or .clase = eClases.asesino Or .clase = eClases.Paladin Then
        
            .Invent.Object(9).ObjIndex = POCION_AZUL_NEWBIE
            .Invent.Object(9).Amount = 1000
            
    End If

    .Invent.ArmourEqpSlot = 4
    .Invent.ArmourEqpObjIndex = .Invent.Object(4).ObjIndex
    .Invent.WeaponEqpObjIndex = .Invent.Object(3).ObjIndex
    .Invent.WeaponEqpSlot = 3
        
    .Char.Body = ObjData(.Invent.Object(.Invent.ArmourEqpSlot).ObjIndex).Ropaje

End With

'Guardamos
Call SaveUser(UserIndex)

'Open User
Call ConnectUser(UserIndex, UserList(UserIndex).id, Password)
    
End Sub


Public Sub RemoverTrabajador(UserIndex As Integer)
    Call TrabajadoresGroup.eliminar(UserIndex)
End Sub

Function DobleEspacios(UserName As String) As Boolean
Dim Antes As Boolean
Dim i As Integer
For i = 1 To Len(UserName)
    If mid(UserName, i, 1) = " " Then
        If Antes = True Then
            DobleEspacios = True
            Exit Function
        Else
        Antes = True
        End If
    Else
    Antes = False
    End If
Next
DobleEspacios = False
End Function

Public Function CheckSum(cadena As String, key As Byte) As String
Dim Salto As String

Salto = "01luoq"

CheckSum = mid(MD5String(cadena & key & Salto), 3, 11)
Debug.Print CheckSum
End Function



Public Sub QuitarDelMercadoAo(nombre As String)

Dim sql As String

sql = "DELETE FROM " & DB_NAME_PRINCIPAL & ".ventas WHERE (Personaje='" & nombre & "' or Personaje like '" & nombre & "-%' or Personaje like '%-" & nombre & "' or Personaje like '%-" & nombre & "-%')"

conn.Execute sql, , adCmdText + adExecuteNoRecords

sql = "DELETE FROM " & DB_NAME_PRINCIPAL & ".confirmaciones WHERE (Vendedor='" & nombre & "' or Vendedor like '" & nombre & "-%' or Vendedor like '%-" & nombre & "' or Vendedor like '%-" & nombre & "-%')"

conn.Execute sql, , adCmdText + adExecuteNoRecords


sql = "DELETE FROM " & DB_NAME_PRINCIPAL & ".confirmaciones WHERE (Comprador='" & nombre & "' or Comprador like '" & nombre & "-%' or Comprador like '%-" & nombre & "' or Comprador like '%-" & nombre & "-%')"

conn.Execute sql, , adCmdText + adExecuteNoRecords
        
End Sub
'MD5Mod: Aca va el MD5String del MD5File del .exe esperado + el salto
Private Function isVersionCorrecta(md5mod As String) As Boolean
    Dim md5modReal As String
    
    #If TDSFacil Then
        md5modReal = "458e7fbc4c76e4f9812021c95368453f"
    #Else
        md5modReal = "9da301ee59b3fa378ddaf6523752227d"
    #End If
   
    If md5modReal <> md5mod Then
        isVersionCorrecta = False
    Else
        isVersionCorrecta = True
    End If
    
End Function

Private Function isCheckSumCorrecto(Dato1 As String, Dato2 As String, key As Byte, CheckSum_ As String) As Boolean
    
    #If TDSFacil Then
        If Not CheckSum(Dato1 & Dato2, key) = CheckSum_ Then
    #Else
        If Not CheckSum(Dato1 & Dato2, key) = CheckSum_ Then
    #End If
        isCheckSumCorrecto = False
    Else
        isCheckSumCorrecto = True
    End If
End Function
