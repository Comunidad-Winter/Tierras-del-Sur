Attribute VB_Name = "modPersonaje"
Option Explicit

Enum eTipoSalida
    NoSaliendo = 0              ' El personaje no está saliendo. Esta jugando normalmente
    SaliendoNaturalmente = 1    ' El personaje escribio /SALIR
    SaliendoForsozamente = 2    ' Se produjo un corte abrupto en la conexión entre el Cliente y el Servidor
End Enum

 'Fama del usuario
Type tReputacion
    NobleRep As Double
    BurguesRep As Double
    PlebeRep As Double
    LadronesRep As Double
    BandidoRep As Double
    AsesinoRep As Double
    promedio As Double
End Type

' Estado del Personaje
Type UserStats
    GLD As Long 'Dinero
    GldBackup As Long
    Banco As Long
    MaxHP As Integer
    minHP As Integer
    MaxSta As Integer
    MinSta As Integer
    MaxMAN As Integer
    MinMAN As Integer
    MaxHIT As Integer
    MinHIT As Integer
    MaxHam As Integer
    minham As Integer
    MaxAGU As Integer
    minAgu As Integer
    Def As Integer
    ELV As Integer
    Exp As Currency
    ELU As Currency
    UserSkills(1 To NUMSKILLS) As Integer
    UserAtributos(1 To NUMATRIBUTOS) As Integer
    UserAtributosBackUP(1 To NUMATRIBUTOS) As Integer
    UserHechizos(1 To MAXUSERHECHIZOS) As Integer
    UsuariosMatados As Integer
    NPCsMuertos As Long 'Variable cmabiada por marche
    SkillPts As Integer
    
    OroGanado As Long
    OroPerdido As Long
    RetosGanadoS As Integer
    RetosPerdidosB As Integer
    
    Veceshechado As Integer
    GlobAl As Integer
    MaxItems As Integer
End Type

'Flags
Type UserFlags
    Saliendo As eTipoSalida     ' El personaje esta saliendo del juego?
    Banrazon As String
    Penasas As String
    Muerto As Byte              ' ¿Esta muerto?
    Comerciando As Boolean      ' ¿Esta comerciando?
    UserLogged As Boolean       ' Este UserIndex tiene un personaje cargado?
    Meditando As Boolean
    modoCombate As Boolean
    Hambre As Byte
    Sed As Byte
    ModoRol As Boolean
    Envenenado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    paralizadoPor As Integer    ' El personaje que lo paralizo/inmovilizo
    Invisible As Byte
    Oculto As Byte
    Desnudo As Byte
    Descansar As Boolean
    hechizo As Integer
    Navegando As Byte
    Seguro As Boolean
    PermitirDragAndDrop As Boolean
    DuracionEfecto As Long
    TargetNPC As Integer ' Npc señalado por el usuario
    TargetNpcTipo As Integer ' Tipo del npc señalado
    Ban As Byte
    TargetUser As Integer ' Usuario señalado
    TargetObj As Integer ' Obj señalado
    TargetObjMap As Integer
    TargetObjX As Integer
    TargetObjY As Integer
    TargetMap As Integer
    TargetX As Integer
    TargetY As Integer
    Privilegios As Byte
    LastCrimMatado As String
    LastCiudMatado As String
    LastNeutralMatado As String
    OldBody As Integer
    OldHead As Integer
    AdminInvisible As Byte
    Trabajando As Boolean
    UltimoMensaje As Byte
    PertAlCons As Byte
    PertAlConsCaos As Byte
    Silenciado As Byte
    Mimetizado As Byte
    Unban As String
    ShowDopa As Boolean
End Type


'Intervalos del usuario.
Type UserIntervalos
    Golpe As Long
    Magia As Long
    Flecha As Long
    UsarU As Long
    UsarClick As Long
End Type

' Contadores
Type UserCounters
    Mimetismo As Long
    IdleCount As Long           ' Inactividad
    HPCounter As Long           ' Sana automaticamente
    STACounter As Long          ' Recupera energia
    Frio As Long                ' Frio
    Calor As Long               ' Calor
    COMCounter As Long          ' Hambre
    AGUACounter As Long         ' Sed
    Veneno As Long              ' Veneno
    Meditacion As Long          ' Meditacion
    Paralisis As Long           ' Paralisis
    Invisibilidad As Long       ' Invisibilidad
    Pena As Long                ' Pena
    Salir As Long               ' Tiempo restante para salir del juego
    
    combateRegresiva As Byte    ' Sistema de Eventos. Cuenta regresiva.
    FotoDenuncia As Byte        ' Cantidad de fotodenuncias enviadas por Minuto
    
    ' Marcas de tiempo de Intervalos del SERVIDOR
    TimerLanzarSpell As Long    ' Ultima vez lanzo magia
    TimerPuedeAtacar As Long    ' Ultima vez que ataco
    TimerUsarU As Long          ' Ultima vez que hizo clic en la U
    TimerUsarClic As Long       ' Ultimo Clic de Usar
    TimerOculto As Long         ' Cantidad de tiempo desde que esta oculto
    ultimoIntentoOcultar As Long   ' Ultima momento en el que se intento ocultar.

    ' Marcas de tiempo de intervalos del CLIENTE
    ultimoTickClicUsar As Single
    ultimoTickU As Single
    ultimoTickProyectiles As Single
    ultimoTickMagia As Single
    ultimoTickPegar As Single
End Type

' Informacion de Faccion
Type tFacciones
    alineacion As eAlineaciones
    ArmadaReal As Byte                  '   Pertenece a la armada?
    FuerzasCaos As Byte                 '   Pertenece al caos?
    CriminalesMatados As Long           '   Cuantos criminales mato
    CiudadanosMatados As Long           '   Cuantos ciudadanos mato
    NeutralesMatados As Long            '   Cuantos ciudadanos mato
    RecompensasReal As Long             '   Nivel en la faccion Real
    RecompensasCaos As Long             '   Nivel en la faccion Caos
    RecibioExpInicialReal As Byte
    RecibioExpInicialCaos As Byte
    RecibioArmaduraReal As Byte
    RecibioArmaduraCaos As Byte
End Type

' Informacion de Clan
Type tGuild
    GuildName As String
    Solicitudes As Long
    SolicitudesRechazadas As Long
    echadas As Long
    VecesFueGuildLeader As Long
    EsGuildLeader As Byte
    FundoClan As Byte
    ClanFundadoID As Long
    id As Long
    ClanesParticipo As Long
    GuildPoints As Double
End Type

' Informacion de la actividad de Trabajo
Type tTrabajo
    tipo  As eTrabajos  'Tipo de trabajo
    modo As Integer 'Red de pesca o caña, hacha dorada o comun
    cantidad As Integer 'Cantidad para hacer. Necesesario en
    modificador As Integer 'Probabilidad de que el trabajo de sus frutos, o cantidad (maxima) que puede hacer cada vez que trabaja
    rangoGeneracion As tRango 'Minimo y maximo de elementos que puede generar
End Type

' Informacion sobre el Mimetismo de un personaje
Type tMimetizado
    nombre As String            ' Nombre del personaje con el cual se mimetizo
    alineacion As eAlineaciones          ' El personaje con el cual nos mimetizamos es criminal?
    Privilegios As Byte         ' Privilegios del personaje con el cual se mimetizo
    Apareciencia As Char        ' Informacion Estetica
End Type

' Datos para el control de cheats
Type tControlCheat
    VecesAtack As Integer       ' Cantidad de hechizos / golpe que pego por minuto
    rompeIntervalo As Byte      ' Cantidad de veces que rompio el intervalo en el CLIENTE
    vecesCheatEngine As Byte    ' Cantidad de vecs que le salto el cheat engine para caminar
End Type


Type tEventoOcultar
    fecha As Long
    Posicion As Position
End Type

' Personaje
Type User
    id As Long          ' Identificador del Personaje
    IDCuenta As Long    ' Identificador de la Cuenta
    Name As String      ' Nombre del Personaje
    Premium As Boolean  ' ¿El usuario es premium?
    
    #If TDSFacil = 1 Then
        segundosPremium As Long     ' Segundos que puedo jugar TDSF Gratis
    #End If
    
    TokSolicitudDePersonaje As Long 'Esto esta para que no se le entregue la info de login de un usuario,
                                    'a otro que ingreso en ese slot mientras el otro esperaba y cerro el juego
    
    ' Datos de ingreso
    MacAddress As String
    NombrePC As String
    ip As Currency                  'Se guarda en formato entero
    
    ' Datos de registro
    Email As String
    pin As String
    Password As String
    
    ' Mimestimo. Información que toma el personaje al ser mimetizado
    Mimetizado As tMimetizado
    
    Char As Char                    ' Define la apariencia
    OrigChar As Char
    desc As String                  ' Descripcion
    
    clase As eClases
    ClaseNumero As Byte             ' Mantengo cacheado cual es el número de la clase en el config... Problema de diseño.
    Raza As eRazas
    Genero As eGeneros
    
       
    Hogar As String
    CentinelaID As Integer          ' El ID del centinela que tiene asociado
    Invent As inventario
    pos As WorldPos
    
    RDBuffer As String 'Buffer roto
    
    BancoInvent As BancoInventario
    Counters As UserCounters
    intervalos As UserIntervalos
    
    Stats As UserStats
    flags As UserFlags
    
    Reputacion As tReputacion
    faccion As tFacciones
    
    'Clan
    GuildInfo As tGuild
    ClanRef As cClan
    
    'Comercio
    ComUsu As tCOmercioUsuario
    
    'Mascotas
    NroMacotas As Integer
    NroMascotasGuardadas As Byte
    MascotasIndex(1 To MAXMASCOTAS) As Integer
    MascotasType(1 To MAXMASCOTAS) As Integer
    MascotasGuardadas(1 To MAXMASCOTAS) As Integer
    
    'Conexion
    ConnID As Long 'Identificador del Socket. -1 Si es valido
    InicioConexion As Long 'Momento que se acepto la conexión
    ConfirmacionConexion As Byte 'El usuario mando un paquete valido
    UserIndex As Integer
    
    'Party
    PartyIndex As Integer   'index a la party q es miembro
    PartySolicitud As Integer   'index a la party q solicito
        
    'Anti robo de npc
    LuchandoNPC As Integer
    
    'Trabajo
    Trabajo As tTrabajo
    
    'Seguridad
    'De los paquetes
    PacketNumber As Long
    MinPacketNumber As Byte
    CryptOffset As Integer 'Gorlok
    
    ' Eventos.
    evento As iEvento
    solicitudEvento As cSolicitudEvento
    
    ' Fecha ingreso
    FechaIngreso As Date
    
    ' Anticheat
    controlCheat As tControlCheat
    
    ' Resucitador
    resucitacionPendiente As resucitacionPendiente
    
    ' Eventos que hizo
    eventoOcultar As tEventoOcultar
End Type

Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean
EsNewbie = UserList(UserIndex).Stats.ELV <= LimiteNewbie
End Function

Sub DarCuerpoDesnudo(ByRef Usuario As User)
'[Wizard] Si navega damos el Body de la barca:)
If Not Usuario.flags.Navegando Then
    Select Case Usuario.Raza
        Case eRazas.Humano
                Select Case Usuario.Genero
                    Case eGeneros.Hombre
                            Usuario.Char.Body = 21
                    Case eGeneros.Mujer
                            Usuario.Char.Body = 39
                End Select
        Case eRazas.ElfoOscuro
                Select Case Usuario.Genero
                    Case eGeneros.Hombre
                            Usuario.Char.Body = 32
                    Case eGeneros.Mujer
                            Usuario.Char.Body = 40
                End Select
        Case eRazas.Enano
          Select Case Usuario.Genero
                    Case eGeneros.Hombre
                            Usuario.Char.Body = 53
                    Case eGeneros.Mujer
                            Usuario.Char.Body = 60
          End Select
        Case eRazas.Gnomo
                Select Case Usuario.Genero
                    Case eGeneros.Hombre
                            Usuario.Char.Body = 53
                    Case eGeneros.Mujer
                            Usuario.Char.Body = 60
                End Select
        Case Else
                Select Case Usuario.Genero
                    Case eGeneros.Hombre
                            Usuario.Char.Body = 21
                    Case eGeneros.Mujer
                            Usuario.Char.Body = 39
                End Select
    End Select
Else
    Usuario.Char.Body = ObjData(Usuario.Invent.BarcoObjIndex).Ropaje
End If

Usuario.flags.Desnudo = 1
End Sub

' Determina si el usuario esta en condicion de subir resistencia magica de manera natural
Private Function puedeSubirResistenciaMagica(ByRef personaje As User) As Boolean

    If personaje.clase = eClases.Paladin Or personaje.clase = eClases.Guerrero Or personaje.clase = eClases.Cazador Then
        puedeSubirResistenciaMagica = True
        Exit Function
    End If
    
    If personaje.Invent.AnilloEqpObjIndex > 0 Then
        If ObjData(personaje.Invent.AnilloEqpObjIndex).DefensaMagicaMin > 0 Then
            puedeSubirResistenciaMagica = True
            Exit Function
        End If
    End If
    
    If personaje.Invent.BrasaleteEqpObjIndex > 0 Then
        If ObjData(personaje.Invent.BrasaleteEqpObjIndex).DefensaMagicaMin > 0 Then
            puedeSubirResistenciaMagica = True
            Exit Function
        End If
    End If
    
    puedeSubirResistenciaMagica = False
End Function

Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'Marce 12-6-6 . Cambiado a pedido de balance
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

If UserList(UserIndex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub

'Restricciones
If Skill = Apuñalar And UserList(UserIndex).Stats.UserSkills(eSkills.Apuñalar) < 10 And Not (UserList(UserIndex).clase = eClases.asesino) Then
    Exit Sub
End If

If UserList(UserIndex).flags.Hambre = 1 Or UserList(UserIndex).flags.Sed = 1 Or UserList(UserIndex).Stats.ELV > UBound(LevelSkill) Then
    Exit Sub
End If

Dim Aumenta As Integer
Dim Prob As Integer
Dim lvl As Integer

If Skill = Ocultarse And Skill = Apuñalar Then
    If UserList(UserIndex).Stats.ELV <= 3 Then
        Prob = 7  '15%
    ElseIf UserList(UserIndex).Stats.ELV <= 6 Then
        Prob = 10 '10%
    ElseIf UserList(UserIndex).Stats.ELV <= 10 Then
        Prob = 20 '5%
    ElseIf UserList(UserIndex).Stats.ELV <= 20 Then
        Prob = 25 '4%
    Else
        Prob = 29 '3.5%
    End If
Else
    If UserList(UserIndex).Stats.ELV <= 3 Then
        Prob = 7
    ElseIf UserList(UserIndex).Stats.ELV > 3 And UserList(UserIndex).Stats.ELV < 6 Then
        Prob = 10
    ElseIf UserList(UserIndex).Stats.ELV >= 6 And UserList(UserIndex).Stats.ELV < 10 Then
        Prob = 20
    ElseIf UserList(UserIndex).Stats.ELV >= 10 And UserList(UserIndex).Stats.ELV < 20 Then
        Prob = 25
    Else
        Prob = 28
    End If
End If

Aumenta = Int(RandomNumber(1, Prob))
lvl = UserList(UserIndex).Stats.ELV
  
#If TDSFacil Then
    If Aumenta < 10 And UserList(UserIndex).Stats.UserSkills(Skill) < LevelSkill(lvl) Then
#Else
    If Aumenta = 2 And UserList(UserIndex).Stats.UserSkills(Skill) < LevelSkill(lvl) Then
#End If
        Call AddtoVar(UserList(UserIndex).Stats.UserSkills(Skill), 1, MAXSKILLPOINTS)
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(32) & SkillsNames(Skill) & "," & UserList(UserIndex).Stats.UserSkills(Skill), UserIndex
        
        Call modUsuarios.agregarExperiencia(UserIndex, 50)
        
        EnviarPaquete Paquetes.MensajeSimple, Chr$(48), UserIndex
                    
        'Si esta trabajando le actualizo las probabilidades
        If UserList(UserIndex).Trabajo.tipo > 0 Then
            Call Trabajo.CalcularModificador(UserList(UserIndex))
        'Si esta comerciando le actualizo el inventario con los nuevos precios
        ElseIf UserList(UserIndex).flags.Comerciando = True Then
            Call ActualizarPrecios(UserIndex, UserList(UserIndex).flags.TargetNPC)
        End If
End If

End Sub

Public Sub CambiarAlineacion(UserIndex As Integer, NuevaAlineacion As eAlineaciones)
'Si tiene un clan y sale de la faccion es expulsado del clan
If UserList(UserIndex).GuildInfo.id > 0 Then
    Dim clan As cClan
    Set clan = UserList(UserIndex).ClanRef
    'Si la nueva alineacion no corresponde con la del clan entonces lo saco del clan.
    'Este if podría estar de más ya que se supone que si pertenece al clan es por que tiene la alineacion
    'del clan y si cambia la perderia. Pero podría suceder el caso de que la alineacion sea cambiada de la base de
    'datos y no desde el juego. Al  no ser desde el juego no se expulsarian a los usuarios y no se harian las
    'activadades correspondientes al cambio de alineacion del clan. Por eso esta este IF.
    If Not clan.isAlineacionCompatible(NuevaAlineacion) Then
        'Si es el lider se hecha a todos los integrantes y se hace neutro al lcan
        If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then
            Call mdClanes.SacarMiembroOnline(UserIndex, clan.id, False)
            EnviarPaquete Paquetes.MensajeGuild, "Has sido expulsado del clan ya que no tienes la alineación correspondiente.", UserIndex, ToIndex
        Else
            Call mdClanes.cambiarAlineacionClan(clan, NuevaAlineacion)
            EnviarPaquete Paquetes.MensajeGuild, "La alineación del clan a cambiado. Todos los integrantes que no cumplían con la misma fueron expulsados.", UserIndex, ToIndex
        End If
    End If
End If
End Sub

Public Function obtenerAlineacion(UserIndex As Integer) As eAlineaciones
    If UserList(UserIndex).faccion.ArmadaReal = 1 Then
      obtenerAlineacion = eAlineaciones.Real
    ElseIf UserList(UserIndex).faccion.FuerzasCaos = 1 Then
      obtenerAlineacion = eAlineaciones.caos
    Else
      obtenerAlineacion = eAlineaciones.Neutro
    End If
End Function


Public Sub GuardarMascotas(ByRef personaje As User)

Dim i As Byte

If personaje.NroMacotas = 0 Then Exit Sub

For i = 1 To MAXMASCOTAS
    If personaje.MascotasIndex(i) > 0 Then
        Call guardarMascota(personaje, personaje.MascotasType(i))
        Call QuitarNPC(personaje.MascotasIndex(i))
        personaje.MascotasType(i) = 0
        personaje.MascotasIndex(i) = 0
    End If
Next i

personaje.NroMacotas = 0

End Sub
Public Sub BorrarMascotas(UserIndex As Integer)
Dim i As Byte
    With UserList(UserIndex)
        For i = 1 To MAXMASCOTAS
            If .MascotasIndex(i) > 0 Then
                Call QuitarNPC(.MascotasIndex(i))
            End If
        Next i
    End With
End Sub
Public Sub cambiarNombreInapropaido(UserIndex As Integer, NombreOriginal As String, NombreNuevo As String)
    Call CambiarNombre(UserIndex, NombreOriginal, NombreNuevo)

    Call modPersonaje_Repository.saveNickInapropiado(NombreOriginal)
End Sub
Public Sub CambiarNombre(UserIndex As Integer, NombreOriginal As String, NombreNuevo As String)
Dim TempInt As Integer
Dim infoPersonaje As ADODB.Recordset
Dim error As Byte
Dim sql As String

'Anti bobo
If NombreOriginal = "" Or NombreNuevo = "" Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(85), UserIndex, ToIndex
    Exit Sub
End If

'Chequeo que el nick nuevo sea correcto
If AsciiValidos(NombreNuevo) = False Or DobleEspacios(NombreNuevo) = True Then
    EnviarPaquete Paquetes.mensajeinfo, "Nick nuevo invalido.", UserIndex, ToIndex
    Exit Sub
End If


'El personaje a cambiarle el nick esta online? Lo echo
TempInt = NameIndex(NombreOriginal)
If TempInt > 0 Then
    If Not CloseSocket(TempInt) Then Call LogError("Cambiar nombre")
End If

'Ya existe un personaje con ese nombre?
Call General.cargarAtributosPersonajeOffline(NombreNuevo, infoPersonaje, "ID", False)

If Not infoPersonaje.EOF Then error = 1 Else error = 0

'Liberamos memoria
infoPersonaje.Close
Set infoPersonaje = Nothing

If error > 0 Then
    EnviarPaquete Paquetes.MensajeSimple2, Chr$(80), UserIndex, ToIndex
    Exit Sub
End If
      
'Si es Tierras del Sur Facil tengo los nombres reservados
#If TDSFacil = 1 Then
    Dim infoReserva As ADODB.Recordset
    Dim IDCuentaReservado As Long
    
    sql = "SELECT IDCuenta FROM " & DB_NAME_PRINCIPAL & ".nicks_reservados WHERE Nombre='" & NombreNuevo & "'"
    Set infoReserva = conn.Execute(sql, , adCmdText)
        
    'Me fijo si el nombre ya esta habilitado. Si esta habilitado me fijo de quien
    If infoReserva.EOF = False Then
        IDCuentaReservado = infoReserva!IDCuenta
    Else
        IDCuentaReservado = 0
    End If
    
    'Libero memoria
    infoReserva.Close
    Set infoReserva = Nothing
#End If

If error = 1 Then Exit Sub

'TODO ok. 'Cargamos el personaje a modificar
Call General.cargarAtributosPersonajeOffline(NombreOriginal, infoPersonaje, "NICKB, PENASASB, IDCUENTA", True)

If Not infoPersonaje.EOF Then

    #If TDSFacil = 1 Then
        If IDCuentaReservado > 0 And (val(infoPersonaje!IDCuenta) <> IDCuentaReservado) Then
            EnviarPaquete Paquetes.mensajeinfo, "El nombre que se desea ya lo tiene reservado otra cuenta.", UserIndex, ToIndex
            'Libero memoria
            infoPersonaje.Close
            Set infoPersonaje = Nothing
            Exit Sub
        End If
    #End If

    infoPersonaje!nickb = NombreNuevo
    infoPersonaje!penasasb = infoPersonaje!penasasb & vbCrLf & "Cambio de nick. Antes era " & LCase$(NombreOriginal) & " " & Date & " " & Time
    infoPersonaje.Update

    'Aviso al usuario
    EnviarPaquete Paquetes.mensajeinfo, NombreOriginal & " paso a llamarse " & NombreNuevo, UserIndex, ToIndex
    
    'Liberamos antes de ejecutar otra consulta
    infoPersonaje.Close
    Set infoPersonaje = Nothing
    
    #If TDSFacil = 1 Then
        If IDCuentaReservado = 0 Then
            'Sino esta habilitado por nadie le cambio el nombre viejo por el nuevo en nicks reservados
            sql = "UPDATE " & DB_NAME_PRINCIPAL & ".nicks_reservados SET Nombre='" & NombreNuevo & "' WHERE Nombre='" & NombreOriginal & "'"
            Call modMySql.ejecutarSQL(sql)
        End If
    #End If
Else
    infoPersonaje.Close
    Set infoPersonaje = Nothing
    'El personaje no existe
    EnviarPaquete Paquetes.mensajeinfo, "El usuario al cual se le quiere cambiar el nick no existe.", UserIndex, ToIndex
End If
          
'Guardo en el log del gm
Call LogGM(UserList(UserIndex).id, NombreOriginal & " ahora es " & NombreNuevo, "CNAME")
End Sub

Sub RevivirUsuarioEnREeto(ByVal UserIndex As Integer)

UserList(UserIndex).flags.Muerto = 0
UserList(UserIndex).Stats.minHP = UserList(UserIndex).Stats.MaxHP
UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta

' Si esta navegando, deja de navegar
If UserList(UserIndex).flags.Navegando Then
    UserList(UserIndex).flags.Navegando = 0
    EnviarPaquete Paquetes.Navega, "", UserIndex
End If

Call DarCuerpoDesnudo(UserList(UserIndex))

Call ChangeUserChar(ToMap, 0, UserList(UserIndex).pos.map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

Call SendUserStatsBox(UserIndex)
Call EnviarHambreYsed(UserIndex)

End Sub

Sub UserDie(ByVal UserIndex As Integer, PocionNegra As Boolean, Optional ByVal asesino As Integer = 0)

Dim i As Integer

With UserList(UserIndex)
    
    'Sonido de muerte
    EnviarPaquete Paquetes.WavSnd, Chr$(SND_USERMUERTE), UserIndex, ToPCArea
    'Quitar el dialogo del user muerto
    EnviarPaquete Paquetes.QDL, ITS(.Char.charIndex), UserIndex, ToPCArea

    .Stats.minHP = 0
    .Stats.MinSta = 0
    .flags.Envenenado = 0

    '[Wizard 03/09/05]
    .Counters.Veneno = 0

    '[Wizard]
    .flags.Muerto = 1

    'Cambiamos un poco esto para ahorrarnos muchos paquetes en algunos casos y en otros no tantos
    'pero algo es algo no?. Marche
    '<<<< Paralisis >>>>
    If .flags.Paralizado = 1 Then
        .flags.Paralizado = 0
    End If
'<<<< Descansando >>>>
    If .flags.Descansar Then
        .flags.Descansar = False
    End If
'<<<< Meditando >>>>
    If .flags.Meditando Then
        .flags.Meditando = False
        .Char.FX = 0
        .Char.loops = 0
        ' EnviarPaquete Paquetes.HechizoFX, ITS(.Char.charindex) & ByteToString(0) & ITS(0), UserIndex, ToPCArea, .Pos.Map
    End If
    
'<<<< Duracion Dopa >>>>

    .flags.DuracionEfecto = 0
    Dim loopX As Integer
    For loopX = 1 To NUMATRIBUTOS
            UserList(UserIndex).Stats.UserAtributos(loopX) = UserList(UserIndex).Stats.UserAtributosBackUP(loopX)
    Next
    UserList(UserIndex).flags.ShowDopa = False

'<<<< Invisible >>>>
    If .flags.Invisible = 1 Then
        .flags.Oculto = 0
        .flags.Invisible = 0
        EnviarPaquete Paquetes.Desocultar, ITS(.Char.charIndex), UserIndex, ToMap
        EnviarPaquete Paquetes.Visible, ITS(.Char.charIndex), UserIndex, ToMap
    End If

    If EsPosicionParaAtacarSinPenalidad(UserList(UserIndex).pos) = False Then
        If MapInfo(UserList(UserIndex).pos.map).SeCaeiItems = 0 Then
            If EsNewbie(UserIndex) Then
                'Si esta en un mapa no newbie pierde los items no newbies
                'A excepcion de que se haya matado a si mismo con la pocion negra
                'Si es newbie y se mato no pierde nada
                If MapInfo(.pos.map).restringir <> 1 And Not PocionNegra Then
                    Call TirarTodosLosItemsNoNewbies(UserList(UserIndex), asesino)
                End If
            ElseIf Not EsNewbie(UserIndex) Then
                Call TirarTodo(UserList(UserIndex), asesino)
            End If
        End If
    End If

    ' DESEQUIPA TODOS LOS OBJETOS
    'desequipar armadura
    If .Invent.ArmourEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
    End If

    If .Invent.WeaponEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
    End If

    If .Invent.EscudoEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
    End If

    If .Invent.CascoEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, .Invent.CascoEqpSlot)
    End If

    If .Invent.HerramientaEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, .Invent.HerramientaEqpSlot)
    End If

    If .Invent.MunicionEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
    End If
    
    If .Invent.AnilloEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)
    End If
        
    If .Invent.BrasaleteEqpObjIndex > 0 Then
        Call desequiparByItem(UserList(UserIndex), .Invent.BrasaleteEqpObjIndex)
    End If

    Call quitarObjetos(Objetos_Constantes.COLLAR, 1, .UserIndex)

    
' << Reseteamos los posibles FX sobre el personaje >>
    If .Char.loops = LoopAdEternum Then
        .Char.FX = 0
        .Char.loops = 0
    End If

    ' Debido a todos los cambios
    Call modPersonaje.DarAparienciaCorrespondiente(UserList(.UserIndex))
    
    ' Lo actualizo
    Call modPersonaje_TCP.ActualizarEstetica(UserList(.UserIndex))

    'Si murio el usuario, tambien desaparecen sus mascotas
    For i = 1 To MAXMASCOTAS
        If .MascotasIndex(i) > 0 Then
            Call guardarMascota(UserList(UserIndex), .MascotasType(i))
            Call QuitarNPC(.MascotasIndex(i))
            .MascotasIndex(i) = 0
            .MascotasType(i) = 0
        End If
    Next i
    
    .NroMacotas = 0
    
    'Restauro sus datos por si el personaje estba drogado
    .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributosBackUP(eAtributos.Fuerza)
    .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributosBackUP(eAtributos.Agilidad)

    ' Sistema para quests
    If MapInfo(.pos.map).Aotromapa.map > 0 Then
        Call WarpUserChar(UserIndex, MapInfo(.pos.map).Aotromapa.map, MapInfo(.pos.map).Aotromapa.x + Rnd(10), MapInfo(.pos.map).Aotromapa.y + Rnd(10), False)
    End If

    ' Anti robo de npcs. Libera al npc
    If UserList(UserIndex).LuchandoNPC > 0 Then
        Call AntiRoboNpc.resetearLuchador(NpcList(UserList(UserIndex).LuchandoNPC))
    End If
    
    Call SendUserStatsBox(UserIndex)
    Call UpdateUserInv(True, UserIndex, 0)
        
    'El usuario esta participando en un evento?
    If Not .evento Is Nothing Then
        'El evento esta en pleno desarrollo?
        If .evento.getEstadoEvento = eEstadoEvento.Desarrollandose Then
            'Le digo al evento que tenga en cuenta que el usuario murio
            Call .evento.usuarioMuere(UserIndex)
        End If
    End If
    
End With

End Sub


Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean
 
 If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
 PuedeApuñalar = _
 ((UserList(UserIndex).Stats.UserSkills(Apuñalar) >= MIN_APUÑALAR) _
 And (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1)) _
 Or _
  ((UserList(UserIndex).clase = eClases.asesino) And _
  (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1))
Else
 PuedeApuñalar = False
End If

End Function

Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean
    PuedeAtravesarAgua = UserList(UserIndex).flags.Navegando = 1
End Function

Public Sub RevivirUsuario(personajeResucitado As User, Optional ByVal minHP As Integer = 1, Optional ByVal minAgu As Integer = 0, Optional ByVal minham As Integer = 0)

personajeResucitado.flags.Muerto = 0
personajeResucitado.Stats.minHP = minHP
personajeResucitado.Stats.MinSta = 0
personajeResucitado.Stats.MinMAN = 0
personajeResucitado.Stats.minAgu = minAgu
personajeResucitado.Stats.minham = minham

If minAgu = 0 Then
    personajeResucitado.flags.Sed = 1
Else
    personajeResucitado.flags.Sed = 0
End If

If minham = 0 Then
    personajeResucitado.flags.Hambre = 1
Else
    personajeResucitado.flags.Hambre = 0
End If



' Si esta navegando, deja de navegar
If personajeResucitado.flags.Navegando Then
    personajeResucitado.flags.Navegando = 0
    EnviarPaquete Paquetes.Navega, "", personajeResucitado.UserIndex
End If

' Cabeza
personajeResucitado.Char.Head = personajeResucitado.OrigChar.Head

Call DarCuerpoDesnudo(personajeResucitado)

' Actualizamos la estetica
Call modPersonaje_TCP.ActualizarEstetica(personajeResucitado)

Call SendUserStatsBox(personajeResucitado.UserIndex)
Call EnviarHambreYsed(personajeResucitado.UserIndex)
    
End Sub

Public Function puedeAyudar(ByRef atacante As User, ByRef victima As User) As Boolean
   If atacante.faccion.alineacion = eAlineaciones.Neutro Then
        puedeAyudar = True
        Exit Function
    End If
    
    If atacante.faccion.alineacion = victima.faccion.alineacion Then
        puedeAyudar = True
        Exit Function
    End If
    
    puedeAyudar = False
End Function
Public Function puedeAtacar(ByRef atacante As User, ByRef victima As User) As Boolean

If atacante.flags.modoCombate = False Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(218), atacante.UserIndex
    puedeAtacar = False
    Exit Function
End If

If victima.flags.Muerto = 1 Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(175), atacante.UserIndex
    puedeAtacar = False
    Exit Function
End If

'No se puede atacar en mapas de descanso
If MapInfo(victima.pos.map).zona = "DESCANSO" Then
    puedeAtacar = False
    Exit Function
End If

If EsPosicionParaAtacarSinPenalidad(atacante.pos) And EsPosicionParaAtacarSinPenalidad(victima.pos) Then
    puedeAtacar = True
    Exit Function
End If

If MapInfo(victima.pos.map).Pk = False Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(176), atacante.UserIndex
    puedeAtacar = False
    Exit Function
End If

If (MapData(victima.pos.map, victima.pos.x, victima.pos.y).Trigger And eTriggers.PosicionSegura) Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(177), atacante.UserIndex
    puedeAtacar = False
    Exit Function
End If

If Not puedeAtacarFaccion(atacante, victima) Then
    EnviarPaquete Paquetes.mensajeinfo, "Tu alineación no permite atacar a tu objetivo.", atacante.UserIndex, ToIndex
    puedeAtacar = False
    Exit Function
End If


'If (Victima.flags.Paralizado = 1 Or Victima.flags.Inmovilizado = 1) And atacante.flags.Invisible = 1 Then
'    EnviarPaquete Paquetes.mensajeinfo, "No podés atacar a personajes paralizados si estas invisible.", atacante.userIndex, ToIndex
'    puedeAtacar = False
'    Exit Function
'End If

'Se asegura que la victima no es un GM
If victima.flags.Privilegios >= 1 Then
    puedeAtacar = False
    Exit Function
End If

If atacante.flags.Muerto = 1 Then
    EnviarPaquete Paquetes.mensajeinfo, "No podes atacar porque estas muerto.", atacante.UserIndex
    puedeAtacar = False
    Exit Function
End If


puedeAtacar = True
End Function

Public Function PoderEvasionEscudo(ByRef personaje As User) As Long
    PoderEvasionEscudo = (personaje.Stats.UserSkills(Defensa) * ModEvasionDeEscudoClase(personaje.ClaseNumero)) / 2
End Function

Public Function PoderEvasion(ByRef personaje As User) As Long
    Dim PoderEvasionTemp As Long
    Dim cantidadSkills As Integer
    
    cantidadSkills = personaje.Stats.UserSkills(tacticas)
    
    If cantidadSkills <= 31 Then
        PoderEvasionTemp = (cantidadSkills * ModificadorEvasion(personaje.ClaseNumero))
    ElseIf cantidadSkills <= 61 Then
        PoderEvasionTemp = ((cantidadSkills + personaje.Stats.UserAtributos(eAtributos.Agilidad)) * ModificadorEvasion(personaje.ClaseNumero))
    ElseIf cantidadSkills <= 99 Then
        PoderEvasionTemp = ((cantidadSkills + (2 * personaje.Stats.UserAtributos(eAtributos.Agilidad))) * ModificadorEvasion(personaje.ClaseNumero))
    Else
        PoderEvasionTemp = ((cantidadSkills + (2.5 * personaje.Stats.UserAtributos(eAtributos.Agilidad))) * ModificadorEvasion(personaje.ClaseNumero))
    End If
    
    PoderEvasion = (PoderEvasionTemp + (2.5 * HelperMatematicas.maxs(personaje.Stats.ELV - 12, 0)))
End Function

Public Function PoderAtaqueArma(ByRef personaje As User) As Long
    Dim PoderAtaqueTemp As Long
    Dim cantidadSkills As Integer
   
    cantidadSkills = personaje.Stats.UserSkills(Armas)
    
    If cantidadSkills < 31 Then
        PoderAtaqueTemp = (cantidadSkills + (0.5 * personaje.Stats.UserAtributos(eAtributos.Agilidad)) * ModificadorPoderAtaqueArmas(personaje.ClaseNumero))
    ElseIf cantidadSkills < 61 Then
        PoderAtaqueTemp = ((cantidadSkills + personaje.Stats.UserAtributos(eAtributos.Agilidad)) * ModificadorPoderAtaqueArmas(personaje.ClaseNumero))
    ElseIf cantidadSkills < 91 Then
        PoderAtaqueTemp = ((cantidadSkills + (1.5 * personaje.Stats.UserAtributos(eAtributos.Agilidad))) * ModificadorPoderAtaqueArmas(personaje.ClaseNumero))
    ElseIf cantidadSkills < 100 Then
        PoderAtaqueTemp = ((cantidadSkills + (2 * personaje.Stats.UserAtributos(eAtributos.Agilidad))) * ModificadorPoderAtaqueArmas(personaje.ClaseNumero))
    Else
       PoderAtaqueTemp = ((cantidadSkills + (3 * personaje.Stats.UserAtributos(eAtributos.Agilidad))) * ModificadorPoderAtaqueArmas(personaje.ClaseNumero))
    End If
    PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * HelperMatematicas.maxs(personaje.Stats.ELV - 12, 0)))
End Function

Public Function PoderAtaqueProyectil(ByRef personaje As User) As Long
    Dim PoderAtaqueTemp As Long
    Dim cantidadSkills As Integer
    
    cantidadSkills = personaje.Stats.UserSkills(eSkills.proyectiles)
    
    If cantidadSkills < 31 Then
        PoderAtaqueTemp = (cantidadSkills + 0.5 * personaje.Stats.UserAtributos(eAtributos.Agilidad)) * ModificadorPoderAtaqueProyectiles(personaje.ClaseNumero)
    ElseIf cantidadSkills < 61 Then
        PoderAtaqueTemp = (cantidadSkills + personaje.Stats.UserAtributos(eAtributos.Agilidad)) * ModificadorPoderAtaqueProyectiles(personaje.ClaseNumero)
    ElseIf cantidadSkills < 91 Then
        PoderAtaqueTemp = (cantidadSkills + (1.5 * personaje.Stats.UserAtributos(eAtributos.Agilidad))) * ModificadorPoderAtaqueProyectiles(personaje.ClaseNumero)
    ElseIf cantidadSkills < 99 Then
        PoderAtaqueTemp = (cantidadSkills + (2 * personaje.Stats.UserAtributos(eAtributos.Agilidad))) * ModificadorPoderAtaqueProyectiles(personaje.ClaseNumero)
    Else
        PoderAtaqueTemp = (cantidadSkills + (3 * personaje.Stats.UserAtributos(eAtributos.Agilidad))) * ModificadorPoderAtaqueProyectiles(personaje.ClaseNumero)
    End If
    
    PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * HelperMatematicas.maxs(personaje.Stats.ELV - 12, 0)))
End Function

Public Function PoderAtaqueWresterling(ByVal UserIndex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    If UserList(UserIndex).Stats.UserSkills(Wresterling) < 31 Then
        PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(Wresterling) * _
        ModificadorPoderAtaqueArmas(UserList(UserIndex).ClaseNumero))
    ElseIf UserList(UserIndex).Stats.UserSkills(Wresterling) < 61 Then
            PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Wresterling) + _
            UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad)) * _
            ModificadorPoderAtaqueArmas(UserList(UserIndex).ClaseNumero))
    ElseIf UserList(UserIndex).Stats.UserSkills(Wresterling) < 91 Then
            PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Wresterling) + _
            (2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
            ModificadorPoderAtaqueArmas(UserList(UserIndex).ClaseNumero))
    Else
           PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Wresterling) + _
           (3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
           ModificadorPoderAtaqueArmas(UserList(UserIndex).ClaseNumero))
    End If
    PoderAtaqueWresterling = (PoderAtaqueTemp + (2.5 * HelperMatematicas.maxs(UserList(UserIndex).Stats.ELV - 12, 0)))
End Function

' Devuelve verdadero si se le puede quitar, falso de lo contrario
Public Function QuitarEnergia(personaje As User, ByVal cantidad As Integer) As Boolean

    ' ¿Tiene la energia que le quier sacar?
    If personaje.Stats.MinSta - cantidad >= 0 Then
        personaje.Stats.MinSta = personaje.Stats.MinSta - cantidad
        QuitarEnergia = True
    Else
        QuitarEnergia = False
    End If
   
End Function

Public Sub quitarParalisis(ByRef personaje As User)
    personaje.flags.Paralizado = 0
    EnviarPaquete Paquetes.NoParalizado2, "", personaje.UserIndex
End Sub

' El personaje deja de trabajar
Public Sub DejarDeTrabajar(personaje As User)

    personaje.Trabajo.tipo = eTrabajos.Ninguno
    personaje.Trabajo.modo = 0
    personaje.Trabajo.cantidad = 0
    personaje.Trabajo.modificador = 0
    
    personaje.flags.Trabajando = False
               
    EnviarPaquete Paquetes.DejaDeTrabaja, "", personaje.UserIndex
    
    RemoverTrabajador personaje.UserIndex
End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal cantidad As Integer)
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - cantidad
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
End Sub


Public Function tieneArmaduraCazador(ByRef personaje As User) As Boolean

    If personaje.Invent.ArmourEqpObjIndex > 0 Then
        If (personaje.Invent.ArmourEqpObjIndex = ARMADURA_DE_CAZADOR Or _
            personaje.Invent.ArmourEqpObjIndex = ARMADURA_DE_CAZADOR_G Or _
            personaje.Invent.ArmourEqpObjIndex = ARMADURA_DE_CAZADOR_2 Or _
            personaje.Invent.ArmourEqpObjIndex = EQUIPO_INVERNAL_HH Or _
            personaje.Invent.ArmourEqpObjIndex = EQUIPO_INVERNAL_HM Or _
            personaje.Invent.ArmourEqpObjIndex = EQUIPO_INVERNAL_EG) Then
            
            tieneArmaduraCazador = True
            Exit Function
        End If
    End If

    tieneArmaduraCazador = False
End Function
Private Sub dejarDeOcultarseAlMoverse(ByRef personaje As User)

    ' Regla
    If personaje.clase = eClases.Cazador Then
        If personaje.Invent.ArmourEqpObjIndex > 0 Then
            If tieneArmaduraCazador(personaje) Then
                If personaje.eventoOcultar.Posicion.x > 0 And personaje.eventoOcultar.Posicion.y > 0 Then
                    If Distance(personaje.pos.x, personaje.pos.y, personaje.eventoOcultar.Posicion.x, personaje.eventoOcultar.Posicion.y) < 2 Then
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    
    
    Call quitarOcultamiento(personaje)

    EnviarPaquete Paquetes.MensajeSimple, Chr$(23), personaje.UserIndex

End Sub
Public Sub Moverse(ByRef personaje As User)
    ' Si esta Saliendo naturalmente, lo cancelamos
    If personaje.flags.Saliendo = eTipoSalida.SaliendoNaturalmente Then
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(36), personaje.UserIndex, ToIndex
        personaje.flags.Saliendo = NoSaliendo
        personaje.Counters.Salir = 0
    End If
    
    If Not personaje.resucitacionPendiente Is Nothing Then
        Call modResucitar.cancelarResucitacion(personaje.resucitacionPendiente)
    End If
        
    ' Si el personaje esta descansando, deja de descansar
    If personaje.flags.Descansar Then
        personaje.flags.Descansar = False
        EnviarPaquete Paquetes.MDescansar, "", personaje.UserIndex, ToIndex
    ElseIf personaje.flags.Meditando Then
        ' Si el personaje esta meditando, deja de hacerlo
        personaje.flags.Meditando = False
        personaje.Char.FX = 0
        personaje.Char.loops = 0
        
        EnviarPaquete Paquetes.Meditando, "", personaje.UserIndex, ToIndex
        EnviarPaquete Paquetes.HechizoFX, ITS(personaje.Char.charIndex) & ByteToString(0) & ITS(0), personaje.UserIndex, ToMap, personaje.pos.map
    End If

    'Si bien esto esta en el cliente, lo pongo aca por las dudas
    If personaje.flags.Trabajando = True Then Call DejarDeTrabajar(personaje)

End Sub

Public Sub posMoverse(ByRef personaje As User)
    ' Si se mueve y no es ladron, deja de ocultarse
    If personaje.flags.Oculto = 1 Then
        Call dejarDeOcultarseAlMoverse(personaje)
    End If
    
     If personaje.flags.Muerto = 1 Then
        If (MapData(personaje.pos.map, personaje.pos.x, personaje.pos.y).Trigger And eTriggers.RevivirAutomatico) Then
            Call RevivirUsuario(personaje, 1, 50, 50)
            EnviarPaquete Paquetes.MensajeSimple2, Chr$(41), personaje.UserIndex
        End If
    End If

End Sub
' Establece en el personaje la Apariencia que deberia tener
' en base a su estado
Public Sub DarAparienciaCorrespondiente(ByRef personaje As User)

' ¿Esta Navegando?
If personaje.flags.Navegando = 1 Then
    personaje.Char.Head = 0
    
    Dim numerobody As Integer
    
    If personaje.flags.Muerto = 0 Then
        
        Select Case personaje.faccion.alineacion
        
            Case eAlineaciones.caos
                Select Case personaje.Invent.BarcoObjIndex
                    Case 474 ' Barca
                        numerobody = 307
                    Case 475 ' Galera
                        numerobody = 309
                    Case 476 ' Galeon
                        numerobody = 305
                End Select
            
            Case eAlineaciones.Neutro
                
                Select Case personaje.Invent.BarcoObjIndex
                    Case 474 ' Barca
                        numerobody = 308
                    Case 475 ' Galera
                        numerobody = 311
                    Case 476 ' Galeon
                        numerobody = 225
                End Select
            
            Case eAlineaciones.Real
                Select Case personaje.Invent.BarcoObjIndex
                    Case 474 ' Barca
                        numerobody = 84
                    Case 475 ' Galera
                        numerobody = 310
                    Case 476 ' Galeon
                        numerobody = 306
                End Select
        
        End Select

        personaje.Char.Body = numerobody
    Else
        Select Case personaje.Invent.BarcoObjIndex
            Case 474 ' Barca
                numerobody = 314
            Case 475 ' Galera
                numerobody = 312
            Case 476 ' Galeon
                numerobody = 313
        End Select
    
        personaje.Char.Body = numerobody
    End If
        
    personaje.Char.ShieldAnim = NingunEscudo
    personaje.Char.WeaponAnim = NingunArma
    personaje.Char.CascoAnim = NingunCasco
    Exit Sub
End If

' Apareciencia de Muerto?
If personaje.flags.Muerto = 1 Then
    ' ¿Es Caos?
    If personaje.faccion.FuerzasCaos <> 0 Then
        personaje.Char.Body = iCuerpoMuertoCrimi
        personaje.Char.Head = iCabezaMuertoCrimi
    Else ' No es Caos.
        personaje.Char.Body = iCuerpoMuerto
        personaje.Char.Head = iCabezaMuerto
    End If
        
    personaje.Char.ShieldAnim = NingunEscudo
    personaje.Char.WeaponAnim = NingunArma
    personaje.Char.CascoAnim = NingunCasco
    Exit Sub
End If

' Apareciencia de Mimetizado?
If personaje.flags.Mimetizado = 1 Then
    personaje.Char.Body = personaje.Mimetizado.Apareciencia.Body
    personaje.Char.Head = personaje.Mimetizado.Apareciencia.Head
    personaje.Char.CascoAnim = personaje.Mimetizado.Apareciencia.CascoAnim
    personaje.Char.ShieldAnim = personaje.Mimetizado.Apareciencia.ShieldAnim
    personaje.Char.WeaponAnim = personaje.Mimetizado.Apareciencia.WeaponAnim
    Exit Sub
End If

' Si llegamos hasta acá es porque no hay nada (barca, fantasma, mimetismo)
' que este sobre el personaje

' Apareciencia Original
personaje.Char.Head = personaje.OrigChar.Head
        
' Armadura
If personaje.Invent.ArmourEqpObjIndex > 0 Then
    personaje.Char.Body = ObjData(personaje.Invent.ArmourEqpObjIndex).Ropaje
Else
    Call DarCuerpoDesnudo(personaje)
End If
             
' Escudo
If personaje.Invent.EscudoEqpObjIndex > 0 Then
    personaje.Char.ShieldAnim = ObjData(personaje.Invent.EscudoEqpObjIndex).ShieldAnim
Else
    personaje.Char.ShieldAnim = NingunEscudo
End If

' Arma
If personaje.Invent.WeaponEqpObjIndex > 0 Then
    personaje.Char.WeaponAnim = ObjData(personaje.Invent.WeaponEqpObjIndex).WeaponAnim
Else
    personaje.Char.WeaponAnim = NingunArma
End If
 
' Casco
If personaje.Invent.CascoEqpObjIndex > 0 Then
    personaje.Char.CascoAnim = ObjData(personaje.Invent.CascoEqpObjIndex).CascoAnim
Else
    personaje.Char.CascoAnim = NingunCasco
End If


End Sub


Public Sub VolverCriminal(ByRef personaje As User)
    If EsPosicionParaAtacarSinPenalidad(personaje.pos) Then Exit Sub
    
    If personaje.flags.Privilegios > 2 Then Exit Sub
    
    ' Pierde toda la reputacion positiva
    personaje.Reputacion.BurguesRep = 0
    personaje.Reputacion.NobleRep = 0
    personaje.Reputacion.PlebeRep = 0
        
    ' Estadisticas
    Call AddtoVar(personaje.Reputacion.BandidoRep, vlASALTO, MAXREP)
            
    Call modPersonaje_TCP.actualizarNick(personaje)
         
    ' Si es de la facción, es expulsado
    If personaje.faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(personaje.UserIndex)
End Sub

Public Sub VolverCiudadano(ByRef personaje As User)
    If EsPosicionParaAtacarSinPenalidad(personaje.pos) Then Exit Sub
    
    personaje.Reputacion.LadronesRep = 0
    personaje.Reputacion.BandidoRep = 0
    personaje.Reputacion.AsesinoRep = 0
    
    ' Estadisticas
    Call AddtoVar(personaje.Reputacion.PlebeRep, vlASALTO, MAXREP)
    
    ' Actualizamos la estetica
    Call modPersonaje_TCP.actualizarNick(personaje)
End Sub

Public Function estaALaIntemperie(ByRef personaje As User) As Boolean
    Dim Trigger As Integer
    
    Trigger = MapData(personaje.pos.map, personaje.pos.x, personaje.pos.y).Trigger
    
    estaALaIntemperie = False
    
    If MapInfo(personaje.pos.map).zona <> "DUNGEON" Then
        If (Trigger And eTriggers.BajoTecho) = False And (Trigger And eTriggers.PosicionSegura) Then
           estaALaIntemperie = True
        End If
    End If
End Function

Public Function incrementarFuerza(ByRef personaje As User, cantidad As Integer, tiempo As Long)
    Dim maximoIncremento As Integer
    
    maximoIncremento = getMaximoIncrementoFuerza(personaje)
                
    Call AddtoVar(personaje.Stats.UserAtributos(eAtributos.Fuerza), cantidad, maximoIncremento)
    
    personaje.flags.DuracionEfecto = tiempo
    
    If personaje.Stats.UserAtributos(eAtributos.Fuerza) = maximoIncremento Or personaje.Stats.UserAtributos(eAtributos.Agilidad) = getMaximoIncrementoAgilidad(personaje) Then
        'Aca llamo a que se muestre el contador
        EnviarPaquete Paquetes.EnviarFA, LongToString(personaje.flags.DuracionEfecto) & ITS(personaje.Stats.UserAtributos(eAtributos.Agilidad)) & ITS(personaje.Stats.UserAtributos(eAtributos.Fuerza)), personaje.UserIndex, ToIndex
    End If
                   
    
    personaje.flags.ShowDopa = False
End Function

Public Function getMaximoIncrementoAgilidad(ByRef personaje As User) As Integer
    getMaximoIncrementoAgilidad = mini(MAXATRIBUTOS, 2 * personaje.Stats.UserAtributosBackUP(eAtributos.Agilidad))
End Function

Public Function getMaximoIncrementoFuerza(ByRef personaje As User) As Integer
    getMaximoIncrementoFuerza = mini(MAXATRIBUTOS, 2 * personaje.Stats.UserAtributosBackUP(eAtributos.Fuerza))
End Function

Public Function incrementarAgilidad(ByRef personaje As User, cantidad As Integer, tiempo As Long)
    Dim maximoIncremento As Integer
    
    maximoIncremento = getMaximoIncrementoAgilidad(personaje)
                
    Call AddtoVar(personaje.Stats.UserAtributos(eAtributos.Agilidad), cantidad, maximoIncremento)
    
    personaje.flags.DuracionEfecto = tiempo
    
    If personaje.Stats.UserAtributos(eAtributos.Agilidad) = maximoIncremento Or getMaximoIncrementoFuerza(personaje) = personaje.Stats.UserAtributos(eAtributos.Fuerza) Then
        'Aca llamo a que se muestre el contador
        EnviarPaquete Paquetes.EnviarFA, LongToString(personaje.flags.DuracionEfecto) & ITS(personaje.Stats.UserAtributos(eAtributos.Agilidad)) & ITS(personaje.Stats.UserAtributos(eAtributos.Fuerza)), personaje.UserIndex, ToIndex
    End If

    personaje.flags.ShowDopa = False
End Function

Public Function reducirFuerza(ByRef personaje As User, cantidad As Integer, tiempo As Long)
    personaje.Stats.UserAtributos(eAtributos.Fuerza) = personaje.Stats.UserAtributos(eAtributos.Fuerza) - cantidad
        
    If personaje.Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then
        personaje.Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
    End If
    
    personaje.flags.DuracionEfecto = tiempo
End Function

Public Function reducirAgilidad(ByRef personaje As User, cantidad As Integer, tiempo As Long)
    personaje.Stats.UserAtributos(eAtributos.Agilidad) = personaje.Stats.UserAtributos(eAtributos.Agilidad) - cantidad
        
    If personaje.Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then
        personaje.Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
    End If
    
    personaje.flags.DuracionEfecto = tiempo
End Function

Public Function getIntervaloParalizado(ByRef personaje As User) As Long
    If personaje.clase = eClases.Guerrero Then
        getIntervaloParalizado = IntervaloParalizadoGuerrero
    ElseIf personaje.clase = eClases.Cazador Then
        getIntervaloParalizado = IntervaloParalizadoCazador
    Else
        getIntervaloParalizado = IntervaloParalizado
    End If
End Function


Public Sub getStatsIniciales(ByRef personaje As User, ByRef vidaInicial As Integer, ByRef staminaInicial As Integer, ByRef manaInicial As Integer, ByRef hitMinimoInicial As Integer, ByRef hitMaximoInicial As Integer)

    ' Vida
    If personaje.clase = eClases.asesino Then
        vidaInicial = 30
    Else
        vidaInicial = 15 + Int(getPromedioAumentoVida(personaje) + 0.5)
    End If

    staminaInicial = 40
    
'<-----------------MANA----------------------->
    If personaje.clase = eClases.Mago Then
        manaInicial = 100
    ElseIf personaje.clase = eClases.Clerigo Or personaje.clase = eClases.Druida Or personaje.clase = eClases.Bardo Or personaje.clase = eClases.asesino Then
        manaInicial = 50
    Else
        manaInicial = 0
    End If
    
            
    hitMaximoInicial = 2
    hitMinimoInicial = 1

End Sub

Public Function GetCiudad(ByRef personaje As User) As WorldPos
    GetCiudad = Nix
End Function

Public Function getCarcel(ByRef personaje As User) As carcel
    getCarcel = NixCarcel
End Function
