Attribute VB_Name = "mdClanes"
Option Explicit
'************
' Objetivo de la clase mdClanes:
'   Contiene la logica de los clanes la cual aplica al los objetos clanes y clan
'   Envia información a los usuarios


#If TDSFacil Then
    ' TDSF
    Private Const NIVEL_MINIMO = 45
    Private Const ORO_MINIMO = 20000000
#Else
    ' TDS
    Private Const NIVEL_MINIMO = 25
    Private Const ORO_MINIMO = 0
#End If

Private Const LIDERAZGO_MINIMO = 90
Private Const GUILD_POINTS_CREAR = 5000
Private Const GUILD_POINTS_INGRESO = 25
Private Const MAX_CANTIDAD_MIEMBROS = 20
'El tiempo que como minimo debe estar disuelto un clan
'cuando se lo disuelve
Private Const DIAS_MINIMO_DISOLUCION = 7

'Constantes de condicion de disolucion
    'Cantidad de integrantes minimo que tiene que tener el clan
Private Const DISOLUCION_MINIMOS_INTEGRANTES = 4
    'Días durante el cual se puede mantener la condicion antes de ser disuelt
Private Const DISOLUCION_TIEMPO_PERMANENCIA_CONDICION = 2 ' Luego cambiar a 60
'\

Public Enum eAlineaciones
    indefinido = 0
    Neutro = 1
    Real = 2
    caos = 3
End Enum

Public Enum eEstadoClan
    Activo = 1
    Disuelto = 2
End Enum

Private Const CANTIDAD_CODECS = 8

' Ultima fecha en la cual se realizo el chequeo del estado de los clanes
Public UltimaFechaProcesada As Date

Public Sub iniciar()
    Set clanes = New cClanes
    clanes.cargar
End Sub

Public Function PuedeCrearClan(UserIndex As Integer)

    If UserList(UserIndex).Stats.UserSkills(Liderazgo) < LIDERAZGO_MINIMO Then
        PuedeCrearClan = False
        EnviarPaquete Paquetes.MensajeSimple, Chr$(206), UserIndex
        Exit Function
    End If
    
    If UserList(UserIndex).Stats.ELV < NIVEL_MINIMO Then
        PuedeCrearClan = False
        EnviarPaquete Paquetes.mensajeinfo, "Para fundar un clan necesitas ser nivel " & NIVEL_MINIMO & " o superior.", UserIndex
        Exit Function
    End If
    
    If UserList(UserIndex).Stats.GLD < ORO_MINIMO Then
        PuedeCrearClan = False
        EnviarPaquete Paquetes.mensajeinfo, "Necesitas " & ORO_MINIMO & "  de monedas de oro por fundar un clan.", UserIndex
        Exit Function
    End If
    
    PuedeCrearClan = True

End Function



Public Function CrearClan(UserIndex As Integer, info As String) As Boolean
' Formato de info
' Nombre Clan¬Descripcion¬Sitio¬Cantidad de Codecs¬Alineacion¬Codecs
Dim NombreClan As String
Dim descripcion As String
Dim Mandamientos As Integer
Dim codecs(1 To 8) As String
Dim infoClan() As String
Dim CAlineacion As Byte
Dim URL As String
Dim resultado As Integer  ' El ID del nuevo clan creado

Dim i As Integer

If Not PuedeCrearClan(UserIndex) Then
    CrearClan = False
    Exit Function
End If

' Parseo la información que me llego
infoClan = Split(info, "¬")

NombreClan = Trim(infoClan(0))
descripcion = Trim(infoClan(1))
URL = Trim(infoClan(2))

CAlineacion = CByte(val(infoClan(4)))

Mandamientos = CInt(val(infoClan(3)))

For i = 1 To Mandamientos
    codecs(i) = Trim(infoClan(i + 4))
Next i

' Validaciones
If Not NombreClanValido(NombreClan) Then
    EnviarPaquete Paquetes.mensajeinfo, "El nombre del clan no es válido.", UserIndex, ToIndex
    CrearClan = False
    Exit Function
End If

If clanes.ExisteClan(NombreClan) Then
    EnviarPaquete Paquetes.mensajeinfo, "El clan que queres crear ya existe.", UserIndex, ToIndex
    CrearClan = False
    Exit Function
End If

If PuedeUsarNombre(NombreClan, UserList(UserIndex).id) = False Then
    EnviarPaquete Paquetes.mensajeinfo, "No tienes permisos para utilizar este nombre de clan.", UserIndex, ToIndex
    CrearClan = False
    Exit Function
End If


Select Case CAlineacion
    Case eAlineaciones.Neutro
        If UserList(UserIndex).faccion.FuerzasCaos = 1 Or UserList(UserIndex).faccion.ArmadaReal = 1 Then
            CrearClan = False
            EnviarPaquete Paquetes.mensajeinfo, "Para fundar un clan neutro no puedes ser de las Fuerzas del Caos ni de la Arnada Real", UserIndex, ToIndex
            Exit Function
        End If
    Case eAlineaciones.Real
        If UserList(UserIndex).faccion.ArmadaReal = 0 Then
            CrearClan = False
            EnviarPaquete Paquetes.mensajeinfo, "Para fundar un clan de la Armada Real debes ser de la misma.", UserIndex, ToIndex
            Exit Function
        End If
    Case eAlineaciones.caos
        If UserList(UserIndex).faccion.FuerzasCaos = 0 Then
            CrearClan = False
            EnviarPaquete Paquetes.mensajeinfo, "Para fundar un clan del caos debes ser del mismo.", UserIndex, ToIndex
            Exit Function
        End If
    Case Else 'ErrOR
        Call LogError("Error en la ALINEACION!!!!")
        CrearClan = False: Exit Function
End Select

Dim NuevoClan As cClan
Set NuevoClan = New cClan

resultado = NuevoClan.crear(UserList(UserIndex).id, UserList(UserIndex).Name, NombreClan, descripcion, CAlineacion, URL, codecs)

'Se pudo crear el clan?
If resultado > 0 Then

    ' Creo el clan
    Call clanes.NuevoClan(NuevoClan)
    
    'Seteo como online al personaje
    Call NuevoClan.agregarMiembro(UserList(UserIndex).Name, UserList(UserIndex).id)
    Call NuevoClan.setOnline(UserIndex)
    
    'Modifico los flags del usuairos asignadole el clan que creo y lidera
    Call modUsuarios.establecerClanFundado(UserList(UserIndex), NuevoClan)
    
    ' Quito el oro
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - ORO_MINIMO
    
    ' Estadisticas
    Call AddtoVar(UserList(UserIndex).GuildInfo.VecesFueGuildLeader, 1, 10000)
    Call AddtoVar(UserList(UserIndex).GuildInfo.ClanesParticipo, 1, 10000)
    Call GiveGuildPoints(GUILD_POINTS_CREAR, UserIndex)
    
    ' Actualizamos el oro
    EnviarPaquete Paquetes.EnviarOro, Codify(UserList(UserIndex).Stats.GLD), UserIndex, ToIndex
    
    ' Le aviso al usuario que ahora pertenece a unc lan
    EnviarPaquete Paquetes.infoClan, "1", UserIndex, ToIndex
    
    'Actualizo el clan del usuario en la pantalla
    Call modPersonaje_TCP.actualizarNick(UserList(UserIndex))
    
    'Por las dudas de que se cierra abruptamente el server guardo el personaje
    Call SaveUser(UserIndex, 1)
End If

CrearClan = True
 
Exit Function
errhandler:
End Function

Private Function NombreClanValido(ByVal NombreClan As String) As Boolean

Dim car As Byte
Dim i As Integer

NombreClan = LCase$(NombreClan)

If Len(NombreClan) = 0 Or Len(NombreClan) > 15 Then
    NombreClanValido = False
    Exit Function
End If

If DobleEspacios(NombreClan) Then
    NombreClanValido = False
    Exit Function
End If

For i = 1 To Len(NombreClan)
    car = Asc(mid$(NombreClan, i, 1))
    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        NombreClanValido = False
        Exit Function
    End If
Next i

NombreClanValido = True

End Function
Public Function PuedeUsarNombre(NombreClan As String, idUsuario As Long) As Boolean
'El clan esta reservado?
Dim clanReservado As ADODB.Recordset

PuedeUsarNombre = True

Set clanReservado = conn.Execute("SELECT IDUsuario FROM clanes_reservas WHERE Nombre='" & NombreClan & "'")
'Esta reservado?
If clanReservado.EOF = False Then
    If clanReservado!idUsuario <> idUsuario Then
        PuedeUsarNombre = False
    End If
End If

clanReservado.Close
End Function

Public Sub EnviarListaClanes(ByVal UserIndex As Integer)
    If clanes.getCantidad > 0 Then
        EnviarPaquete Paquetes.EnviarGuildsList, clanes.getClanesString(), UserIndex, ToIndex
    Else
        EnviarPaquete Paquetes.EnviarGuildsList, "", UserIndex, ToIndex
    End If
End Sub
Public Sub EnviarDetallesDeClan(ByVal UserIndex As Integer, ByVal GuildName As String)

Dim clan As cClan
Dim cadena As String

If clanes.getCantidad() = 0 Then Exit Sub

Set clan = clanes.getClanPorNombre(GuildName)

If clan Is Nothing Then Exit Sub

cadena = clan.getNombre() & "¬" & clan.getFundador() & "¬" & clan.getDiaFundacion() & "¬" & clan.getLider() & "¬" & clan.getWeb() & _
          "¬" & clan.getCantidadMiembros() & "¬" & clan.getDiasProximaEleccion() & "¬" & clan.getAlineacion() & "¬" & clan.getDescripcion()

Dim i As Byte

For i = 0 To CANTIDAD_CODECS - 1
    cadena = cadena & "¬" & clan.getCodec(i)
Next i

EnviarPaquete Paquetes.EnviarGuildDetails, cadena, UserIndex, ToIndex
End Sub

Public Sub SolicitudIngresoClan(ByVal UserIndex As Integer, ByVal data As String)

Dim MiSol As cSolicitud
Dim clan As cClan
Dim IndexLider As Integer

If EsNewbie(UserIndex) Then
   EnviarPaquete Paquetes.MensajeSimple, Chr$(195), UserIndex
  Exit Sub
End If

Set clan = clanes.getClanPorNombre(Trim(ReadField(1, data, Asc("¬"))))

If clan Is Nothing Then Exit Sub
If clan.isMiembro(UserList(UserIndex).id) Then Exit Sub

If Not clan.isAlineacionCompatible(modPersonaje.obtenerAlineacion(UserIndex)) Then
    EnviarPaquete Paquetes.MensajeGuild, "No tienes la alineacion requerida para entrar a este clan.", UserIndex
    Exit Sub
End If

'Ya existe una solicitud?
If Not clan.existeSolicitud(UserList(UserIndex).id) Then
        'Creo la solicitud
        Set MiSol = New cSolicitud
        Call MiSol.iniciar(UserList(UserIndex).Name, UserList(UserIndex).id, ReadField(2, data, Asc("¬")), Now)
        'Agrego la solicitud al clan
        Call clan.agregarSolicitud(MiSol)
        'Aumento en uno la estadistica que se lleva sobre la cantidad de solicitudes que el usuario envio a todos los clanes
        Call AddtoVar(UserList(UserIndex).GuildInfo.Solicitudes, 1, 1000)
        'Le aviso que la solicitud fue recibida
        EnviarPaquete Paquetes.MensajeSimple, Chr$(196), UserIndex
        'Le aviso al lider que recibio una nueva solicitud
        IndexLider = IDIndex(clan.getIDLider())
        
        If IndexLider > 0 Then 'El lider esta online
            EnviarPaquete Paquetes.MensajeGuild, "Has recibidio una solicitud de ingreso al clan de " & UserList(UserIndex).Name & ".", IndexLider, ToIndex
        End If
Else
        'Si ya habia mandado le digo que espere
        EnviarPaquete Paquetes.MensajeSimple, Chr$(197), UserIndex
End If
End Sub

Public Sub EnviarInformacionALider(ByVal UserIndex As Integer)
Dim clan As cClan

'Tomo el clan
Set clan = UserList(UserIndex).ClanRef
' Esta todo ok?
If clan Is Nothing Then Exit Sub

EnviarPaquete Paquetes.EnviarLeaderInfo, clanes.getClanesString(), UserIndex, ToIndex, 0
End Sub

Public Sub EnviarInformacionALiderSolicitudes(ByVal UserIndex As Integer)
Dim clan As cClan

'Tomo el clan
Set clan = UserList(UserIndex).ClanRef
' Esta todo ok?
If clan Is Nothing Then Exit Sub

'<------- Preparo la información a enviar ---------->
EnviarPaquete Paquetes.EnviarLeaderInfoSolicitudes, clan.getSolicitudesString, UserIndex, ToIndex, 0
End Sub

Public Sub EnviarInformacionALiderMiembros(ByVal UserIndex As Integer)
Dim clan As cClan

'Tomo el clan
Set clan = UserList(UserIndex).ClanRef
' Esta todo ok?
If clan Is Nothing Then Exit Sub

EnviarPaquete Paquetes.EnviarLeaderInfoMiembros, clan.getIntegrantesString, UserIndex, ToIndex, 0
End Sub

Public Sub EnviarInformacionALiderNovedades(ByVal UserIndex As Integer)
Dim info As String
Dim clan As cClan

'Tomo el clan
Set clan = UserList(UserIndex).ClanRef
' Esta todo ok?
If clan Is Nothing Then Exit Sub

EnviarPaquete Paquetes.EnviarLeaderInfoNovedades, clan.getNovedades, UserIndex, ToIndex, 0
End Sub
'CSEH: TDS_LINEA
Public Sub EnviarInformacionPersonaje(ByVal UserName As String, ByVal UserIndex As Integer)

If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim TempInt As Integer
Dim IDPJ As Long
Dim nombre As String
Dim Raza As Byte
Dim clase As String
Dim Genero As String
Dim Nivel As Byte
Dim oro As Long
Dim Banco As Long
Dim promedio As Long
Dim FundoClan As Byte
Dim NombreClanFundado As String
Dim echadas As Byte
Dim Solicitudes As Integer
Dim SolicitudesRechazadas As Integer
Dim VecesLider As Byte
Dim armada As Byte
Dim caos As Byte
Dim CiudadanosMatados As Long
Dim CriminalesMatados As Long
Dim criminal As Byte
Dim cadena As String
Dim solicitud As cSolicitud
Dim clanFundado As cClan
    
criminal = 0

Dim clan As cClan

Set clan = UserList(UserIndex).ClanRef

IDPJ = clan.getIDMiembro(UserName)

If IDPJ = 0 Then 'Es miemrbo?
    Set solicitud = clan.getSolicitudPorNombre(UserName)
    If solicitud Is Nothing Then
        'Solo puede ver la información de usuarios que sean miembro del clan o que quieran entrar
        Exit Sub
    Else
      IDPJ = solicitud.getIDPJ()
    End If
End If

TempInt = IDIndex(IDPJ)

If TempInt > 0 Then 'esta online
With UserList(TempInt)
    nombre = .Name
    Raza = razaToConfigID(.Raza)
    clase = byteToClase(.clase)
    Genero = byteToGenero(.Genero)
    Nivel = .Stats.ELV
    oro = .Stats.GLD
    Banco = .Stats.Banco
    promedio = .Reputacion.promedio
        
    FundoClan = .GuildInfo.FundoClan
        
    If .GuildInfo.FundoClan = 1 Then
        Set clanFundado = clanes.getClan(.GuildInfo.ClanFundadoID)
        If Not clanFundado Is Nothing Then
            NombreClanFundado = clanFundado.getNombre()
        Else
            NombreClanFundado = "nombre desconocido"
        End If
    Else
        NombreClanFundado = ""
    End If
        
    echadas = .GuildInfo.echadas
    Solicitudes = .GuildInfo.Solicitudes
    SolicitudesRechazadas = .GuildInfo.SolicitudesRechazadas
    VecesLider = .GuildInfo.VecesFueGuildLeader
        
    armada = .faccion.ArmadaReal
    caos = .faccion.FuerzasCaos
    CiudadanosMatados = .faccion.CiudadanosMatados
    CriminalesMatados = .faccion.CriminalesMatados
End With
Else ' esta offline
sql = "SELECT * FROM " & DB_NAME_PRINCIPAL & ".usuarios WHERE ID = " & IDPJ
    
Dim infoPersonaje  As ADODB.Recordset
    
General.cargarAtributosPersonajeOffline UserName, infoPersonaje, "NICKB, razaB, claseb, generoB, elvb, gldb, bancob, promedioB, ClanFundadoID, Echadasb, SolicitudesB, SolicitudesRechazadasB, VecesFueGuildLeaderB, EjercitoRealB, EjercitoCaosB, CiudMatadosB, CrimMatadosB", False

If Not infoPersonaje.EOF Then
    nombre = infoPersonaje!nickb
    Raza = razaToConfigID(razaToByte(infoPersonaje!razaB))
    clase = infoPersonaje!claseb
    Genero = infoPersonaje!generoB
    Nivel = infoPersonaje!elvb
    oro = infoPersonaje!gldb
    Banco = infoPersonaje!bancob
    promedio = infoPersonaje!promedioB
             
    If infoPersonaje!ClanFundadoID > 0 Then
        FundoClan = 1
            
        Set clanFundado = clanes.getClan(infoPersonaje!ClanFundadoID)
        If Not clanFundado Is Nothing Then
            NombreClanFundado = clanFundado.getNombre()
        Else
            NombreClanFundado = "nombre desconocido"
        End If
    Else
        FundoClan = 0
        NombreClanFundado = ""
    End If
        
    echadas = infoPersonaje!Echadasb
    Solicitudes = infoPersonaje!SolicitudesB
    SolicitudesRechazadas = infoPersonaje!SolicitudesRechazadasB
    VecesLider = infoPersonaje!VecesFueGuildLeaderB
        
    armada = infoPersonaje!EjercitoRealB
    caos = infoPersonaje!EjercitoCaosB
    CiudadanosMatados = infoPersonaje!CiudMatadosB
    CriminalesMatados = infoPersonaje!CrimMatadosB
Else
    Exit Sub 'TO-DO: Hay un personaje que pertenece a un clan y no existe.
End If
    
'Libero
infoPersonaje.Close
Set infoPersonaje = Nothing
End If

If promedio < 0 Then criminal = 1
    
cadena = Chr$(Raza) & Chr$(claseToConfigID(clase)) & Chr$(generoToByte(Genero)) & _
     Chr$(Nivel) & LongToString(oro) & LongToString(Banco) & LongToString(Abs(promedio)) & criminal & FundoClan & ITS(echadas, 7) & _
     ITS(Solicitudes, 8) & ITS(SolicitudesRechazadas, 9) & ByteToString(VecesLider) & armada & caos & ITS(CiudadanosMatados, 10) & _
     ITS(CriminalesMatados, 11) & NombreClanFundado & "¬" & nombre
        
EnviarPaquete Paquetes.EnviarCharInfo, cadena, UserIndex, ToIndex

End Sub

Public Sub AceptarMiembro(ByVal UserIndex As Integer, ByVal Solicitante As String)
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim clan As cClan
Dim solicitud As cSolicitud
Dim MemberIndex As Integer
Dim sql As String
'Variables donde voy a guardar los datos necesarios del usuario solicitante.
Dim IDPJ As Long
Dim IDClan As Long
Dim armada As Byte
Dim caos As Byte
Dim nombre As String
Dim alineacion As eAlineaciones

'Obtengo el clan
Set clan = UserList(UserIndex).ClanRef
If clan Is Nothing Then Exit Sub

'Antes que nada me fijo si tiene espacio para guardar mas integrantes
If clan.getCantidadMiembros() >= MAX_CANTIDAD_MIEMBROS Then
    EnviarPaquete Paquetes.MensajeGuild, "El clan alcanzo el máximo de " & MAX_CANTIDAD_MIEMBROS & " integrantes.", UserIndex
    Exit Sub
End If

'Si puede aceptar nuevos integrantes, obtengo la solicitud
Set solicitud = clan.getSolicitudPorNombre(Trim(Solicitante))
If solicitud Is Nothing Then Exit Sub

'Obtenfo info de la solicitud
IDPJ = solicitud.getIDPJ()
MemberIndex = IDIndex(IDPJ)

If MemberIndex > 0 Then 'El usuario esta online
    nombre = UserList(MemberIndex).Name
    IDClan = UserList(MemberIndex).GuildInfo.id
    armada = UserList(MemberIndex).faccion.ArmadaReal
    caos = UserList(MemberIndex).faccion.FuerzasCaos
Else
    Dim infoPersonaje As ADODB.Recordset
    
    Call General.cargarAtributosPersonajeOffline(Solicitante, infoPersonaje, "NickB, IDClan, EjercitoRealB, EjercitoCaosB", False)
    
    If Not infoPersonaje.EOF Then
        nombre = infoPersonaje!nickb
        IDClan = infoPersonaje!IDClan
        armada = infoPersonaje!EjercitoRealB
        caos = infoPersonaje!EjercitoCaosB
    End If
    
    infoPersonaje.Close
    Set infoPersonaje = Nothing
    
    If nombre = "" Then
        EnviarPaquete Paquetes.MensajeGuild, "El personaje del cual queres aceptar la solicitud ya no existe.", UserIndex
        Exit Sub 'TODO Manehar este error
    End If
    
End If

If caos = 1 Then
    alineacion = eAlineaciones.caos
ElseIf armada = 1 Then
    alineacion = eAlineaciones.Real
Else
    alineacion = eAlineaciones.Neutro
End If

'El usuario ya pertenece a un clan?
If IDClan = 0 Then
     ' La alineacion del personaje es correcta?
     If Not clan.isAlineacionCompatible(alineacion) Then
        EnviarPaquete Paquetes.MensajeGuild, "El personaje no tiene la alineación requerida.", UserIndex
        Exit Sub
     End If
    
    'Si esta todo ok, agrego el integrante al clan
    Call clan.agregarMiembro(nombre, IDPJ)
    Call clan.removerSolicitud(IDPJ)
    
    'Le aviso a todos los de clan
    EnviarPaquete Paquetes.WavSnd, Chr$(SND_ACEPTADOCLAN), UserIndex, ToGuildMembers
    EnviarPaquete Paquetes.MensajeGuild, nombre & " ha sido aceptado en el clan.", UserIndex, ToGuildMembers 'Actualizo los datos del personaje que ingreso
    
    'Actualizo los stats del personaje
    If MemberIndex > 0 Then
        Call modUsuarios.establecerClanAUsuarioOnline(MemberIndex, clan)
        
        Call AddtoVar(UserList(MemberIndex).GuildInfo.ClanesParticipo, 1, 1000)
        'Gana guilds points por ingresar al clan
        Call GiveGuildPoints(GUILD_POINTS_INGRESO, MemberIndex)
         
        EnviarPaquete Paquetes.MensajeSimple, Chr$(193), MemberIndex, ToIndex
        EnviarPaquete Paquetes.MensajeGuild, "Ahora sos un miembro activo del clan " & UserList(MemberIndex).GuildInfo.GuildName, MemberIndex, ToIndex
    Else
       Call modUsuarios.establecerClanUsuarioOffline(IDPJ, clan.id)
    End If
    
    'El nuevo ingreso afecta en la condicion de riesgo del clan de disolucion?
    Call cumpleCondicionDisolucionAutomatica(clan)
Else ' El usuario pertenece a un clan
    EnviarPaquete Paquetes.MensajeGuild, "No podes aceptar esta solicitud, el personaje pertenece a otro clan", UserIndex
    Exit Sub
End If
End Sub

'Envia el comentario que dejo el usuario (Nombre) al momento de pedir la solicitud de ingreso
Public Sub EnviarComentarioPeticion(ByVal UserIndex As Integer, ByVal nombre As String)

If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub
    
Dim clan As cClan

Set clan = UserList(UserIndex).ClanRef

If clan Is Nothing Then Exit Sub

Dim solicitud As cSolicitud

Set solicitud = clan.getSolicitudPorNombre(nombre)

If Not solicitud Is Nothing Then
    EnviarPaquete Paquetes.PeaceSolRequest, solicitud.getDescripcion, UserIndex, ToIndex
End If
End Sub
Public Sub SalirDeClan(UserIndex As Integer)

    If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(3), UserIndex
        Exit Sub
    ElseIf UserList(UserIndex).GuildInfo.id = 0 Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(223), UserIndex
        Exit Sub
    Else
        'Le aviso al clan
        EnviarPaquete Paquetes.MensajeGuild, UserList(UserIndex).Name & " decidió dejar al clan.", UserIndex, ToGuildMembers
        'Quito efecitvamente al usuario
        Call SacarMiembroOnline(UserIndex, UserList(UserIndex).GuildInfo.id, False)
    End If
End Sub
'Un integrante puede ser sacado de un clan tanto por que fue echado,
'ya no tiene la alineacion que corresponde para el clan, el mismo quizo salir
'ATENCION: Esta funcion no hace chequeo, solo saca al usuario del clan
Public Function SacarMiembroOnline(UserIndex As Integer, IDClanDesdeDondeSeLoQuiereQuitar As Long, echado As Boolean) As Boolean
'Lo quito del clan
Dim clan As cClan

SacarMiembroOnline = False
'El clan del cual es el usuario es el mismo que el que lo quiere sacar?
If UserList(UserIndex).GuildInfo.id = IDClanDesdeDondeSeLoQuiereQuitar Then
    'Quito la relacion que tiene el clan con el usuario
    Set clan = UserList(UserIndex).ClanRef
    Call clan.quitarMiembro(UserList(UserIndex).id, UserIndex)
    'El egreso afecta en la condicion de riesgo del clan de disolucion?
    Call cumpleCondicionDisolucionAutomatica(clan)
    'Todo ok
    SacarMiembroOnline = True
End If
End Function

'Un integrante puede ser sacado de un clan tanto por que fue echado,
'ya no tiene la alineacion que corresponde para el clan, el mismo quizo salir
'ATENCION: Esta funcion no hace chequeo, solo saca al usuario del clan
Public Function SacarMiembroOffline(IDMiembro As Long, IDClanDesdeDondeSeLoQuiereQuitar As Long, echado As Boolean) As Boolean

Dim IDClan As Long

SacarMiembroOffline = False

IDClan = getIDClanUsuarioOffline(IDMiembro)

If IDClan > 0 Then
    'Lo quito efecivamente del clan
    If IDClan = IDClanDesdeDondeSeLoQuiereQuitar Then
    
        Call clanes.getClan(IDClan).quitarMiembro(IDMiembro, 0)
        'El egreso afecta en la condicion de riesgo del clan de disolucion?
        Call cumpleCondicionDisolucionAutomatica(clanes.getClan(IDClan))
        
        SacarMiembroOffline = True
    End If 'TODO Log. Se intenta echar a un personaje que no pertenece al clan
End If 'TODO Log Se intenta echar a un personaje que no existe

End Function

Public Sub EcharMiembro(ByVal UserIndex As Integer, ByVal nombre As String)

Dim MemberIndex As Integer
Dim NombreMiembro As String
Dim IDMiembro As Long

Dim clan As cClan

If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Set clan = UserList(UserIndex).ClanRef

If clan Is Nothing Then Exit Sub

'No se puede echar miembros en el día de las elecciones.
If clan.isElecciones Then
    EnviarPaquete Paquetes.MensajeGuild, "No puedes expulsar integrantes durante el día de elecciones.", UserIndex
    Exit Sub
End If

NombreMiembro = Trim(nombre)
IDMiembro = clan.getIDMiembro(NombreMiembro)

If IDMiembro > 0 Then

    MemberIndex = IDIndex(IDMiembro)
    
    If MemberIndex <> 0 Then 'esta online
    
        If UserList(MemberIndex).GuildInfo.EsGuildLeader = 1 Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(188), UserIndex, ToIndex
            Exit Sub
        End If
        
        'Saco al mimebro que esta online
        Call SacarMiembroOnline(MemberIndex, clan.id, True)
        'Le envio el mensae de que fue echado
        EnviarPaquete Paquetes.MensajeSimple, Chr$(189), MemberIndex, ToIndex
    Else ' El personaje esta offline
    
        Call SacarMiembroOffline(IDMiembro, clan.id, True)
    End If
    
    EnviarPaquete Paquetes.MensajeGuild, NombreMiembro & " fue expulsado del clan.", UserIndex, ToGuildMembers
Else
    'TO-DO Quiere echar a un personaje que no es miebro
End If

End Sub

Public Sub DenegarSolicitud(UserIndex As Integer, nombre As String)

Dim clan As cClan
Dim solicitud As cSolicitud
Dim MemberIndex As Integer

If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Set clan = UserList(UserIndex).ClanRef

If clan Is Nothing Then Exit Sub

Set solicitud = clan.getSolicitudPorNombre(nombre)

If Not solicitud Is Nothing Then    'Existe la solicitud?

    MemberIndex = IDIndex(solicitud.getIDPJ())

    If MemberIndex > 0 Then 'esta online
        EnviarPaquete Paquetes.MensajeSimple, Chr$(191), MemberIndex
        Call AddtoVar(UserList(MemberIndex).GuildInfo.SolicitudesRechazadas, 1, 10000)
    Else
        sql = "UPDATE " & DB_NAME_PRINCIPAL & ".usuarios SET SolicitudesRechazadasB=SolicitudesRechazadasB+1 WHERE ID = " & solicitud.getIDPJ()
        conn.Execute (sql), , adExecuteNoRecords
    End If
    
    'Quito la solicitud
    Call clan.removerSolicitud(solicitud.getIDPJ())
End If
End Sub

Public Sub ActualizarCodecsYDesc(ByVal rdata As String, ByVal UserIndex As Integer)
Dim descripcion As String
Dim info() As String
Dim codecs(0 To CANTIDAD_CODECS - 1) As String
Dim contador As Byte

info = Split(rdata, "¬")

descripcion = info(0)

For contador = 1 To CANTIDAD_CODECS
    codecs(contador - 1) = info(contador + 1)
Next contador

' Es lider y tiene clan?
If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 And Not UserList(UserIndex).ClanRef Is Nothing Then
    Call UserList(UserIndex).ClanRef.setDescripcion(descripcion)
    Call UserList(UserIndex).ClanRef.setCodecs(codecs)
Else
    Exit Sub
End If
End Sub
Public Sub ActualizarNovedades(ByVal rdata As String, ByVal UserIndex As Integer)
' Es lider y tiene clan?
If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 And Not UserList(UserIndex).ClanRef Is Nothing Then
    Call UserList(UserIndex).ClanRef.setNovedades(rdata)
    EnviarPaquete Paquetes.MensajeGuild, "Las novedades del clan fueron actualizadas.", UserIndex
Else
    Exit Sub
End If
End Sub

Public Sub SetNewURL(ByVal UserIndex As Integer, ByVal rdata As String)
' Es lider y tiene clan?
If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 And Not UserList(UserIndex).ClanRef Is Nothing Then
    Call UserList(UserIndex).ClanRef.setURL(rdata)
    EnviarPaquete Paquetes.MensajeSimple, Chr$(198), UserIndex
Else
    Exit Sub
End If
End Sub
Public Sub GiveGuildPoints(ByVal Pts As Integer, ByVal UserIndex As Integer, Optional ByVal SendNotice As Boolean = True)
If SendNotice Then _
    EnviarPaquete Paquetes.MensajeGuild, "¡¡Has recibido " & Pts & " guildpoints!!", UserIndex, ToIndex
Call AddtoVar(UserList(UserIndex).GuildInfo.GuildPoints, Pts, 9000000)
End Sub

Public Sub EnviarNovedadesClan(ByVal UserIndex As Integer)
    'Tiene clan
    Dim clan As cClan
    
    If UserList(UserIndex).GuildInfo.id > 0 Then
    
        Set clan = UserList(UserIndex).ClanRef

        ' Sacamos este cartel molesto
        ' EnviarPaquete Paquetes.EnviarGuildNews, clan.getNovedades(), UserIndex
        
        'El clan esta en elecciones?. Le avisamos que tiene que votar
        If clan.isElecciones Then
            'Le aviso que hoy son las elecciones
            EnviarPaquete Paquetes.MensajeSimple, Chr$(201), UserIndex
            If Not clan.YaVoto(UserList(UserIndex).id) Then
                'Le digo como votar
                EnviarPaquete Paquetes.MensajeSimple, Chr$(202), UserIndex
                EnviarPaquete Paquetes.MensajeSimple, Chr$(203), UserIndex
                EnviarPaquete Paquetes.MensajeSimple, Chr$(204), UserIndex
            Else
                EnviarPaquete Paquetes.MensajeGuild, "Al final del día se conocerán los resultados. Gracias por votar.", UserIndex, ToIndex
            End If
        End If
        
        'El clan esta en riesgo? Le avisams al lider
        If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
            Dim fechaInfraccion As Date
            
            fechaInfraccion = clan.getFechaInfraccion()
            
            If Not fechaInfraccion = 0 Then
                 EnviarPaquete Paquetes.MensajeGuild, "ATENCION: tu clan no cumple con las condiciones mínimas de existencia. Si esto se mantiene así durante los próximos " & DateDiff("y", Date, DateAdd("y", DISOLUCION_TIEMPO_PERMANENCIA_CONDICION, fechaInfraccion)) & " días el clan será disuelto.", UserIndex
            End If
        End If
    End If
End Sub

Public Sub Votar(UserIndex As Integer, NombreVotado As String)

Dim IDVotado As Long
Dim IDVotante As Long
Dim clan As cClan

Set clan = UserList(UserIndex).ClanRef

IDVotante = UserList(UserIndex).id
'Es miembro del clan?
If clan.isMiembro(IDVotante) Then
    'El clan esta en elecciones?
    If clan.isElecciones() Then
        If Not clan.YaVoto(IDVotante) Then
            'Obtengo el ID del usuario a quein voto
            IDVotado = ObtenerIDUsuario(NombreVotado)
            If IDVotado > 0 Then ' Existe el personaje?
                If clan.isMiembro(IDVotado) Then '¿El personaje es miembro del clan?
                        'Esta todo en orden, computo el voto
                        Call clan.ComputarVoto(UserList(UserIndex).id, IDVotado)
                        'Le aviso que el voto se computo
                        EnviarPaquete Paquetes.MensajeSimple, Chr$(16), UserIndex
                Else
                 'El personaje no es miembro del clan
                  EnviarPaquete Paquetes.MensajeSimple, Chr$(15), UserIndex
                End If
            Else
            'El personaje a quien quiere votar no existe
            EnviarPaquete Paquetes.MensajeGuild, "El personaje al cual quieres votar no existe.", UserIndex, ToIndex
            End If
        Else
            'El usuario ya voto
            EnviarPaquete Paquetes.MensajeSimple, Chr$(14), UserIndex
        End If
    Else
        'El clan no esta en elecciones
        EnviarPaquete Paquetes.MensajeSimple, Chr$(13), UserIndex
    End If
End If

End Sub

Public Sub echarDelJuegoIntegrantes(NombreClan As String)

Dim clan As cClan
Dim IntegrantesOnline As EstructurasLib.ColaConBloques
Set clan = clanes.getClanPorNombre(NombreClan)

If Not clan Is Nothing Then
    Set IntegrantesOnline = clan.getIntegrantesOnline
    IntegrantesOnline.itIniciar
    Do While IntegrantesOnline.ithasNext
        If Not CloseSocket(IntegrantesOnline.itnext) Then LogError ("Echar integrantes")
    Loop
End If

End Sub

Public Sub cambiarNombreClan(NombreActual As String, NombreNuevo As String, UserIndex As Integer)
Dim clan As cClan
' Validaciones
If Not NombreClanValido(NombreNuevo) Then
    EnviarPaquete Paquetes.mensajeinfo, "El nombre del clan no es válido.", UserIndex, ToIndex
    Exit Sub
End If

If clanes.ExisteClan(NombreNuevo) Then
    EnviarPaquete Paquetes.mensajeinfo, "El clan que queres crear ya existe.", UserIndex, ToIndex
    Exit Sub
End If

Set clan = clanes.getClanPorNombre(NombreActual)
If Not clan Is Nothing Then
    'El nombre lo cambia un gm...
    If PuedeUsarNombre(NombreNuevo, clan.getIDFundador()) = False Then
        EnviarPaquete Paquetes.mensajeinfo, "No tienes permisos para utilizar este nombre de clan.", UserIndex, ToIndex
        Exit Sub
    End If

    Call mdClanes.echarDelJuegoIntegrantes(NombreActual)
    Call clanes.cambiarNombreClan(clan, NombreNuevo)

    EnviarPaquete Paquetes.mensajeinfo, "El clan fue cambiado de nombre. De " & NombreActual & " a " & NombreNuevo, UserIndex, ToIndex
Else
    'El clan no existe
    EnviarPaquete Paquetes.mensajeinfo, "El clan al cual quieres cambiarle el nombre no existe.", UserIndex, ToIndex
End If

End Sub
Public Sub cambiarAlineacionClan(clan As cClan, NuevaAlineacion As eAlineaciones)
    Call clan.ExpulsarTodosIntegrantesMenosLider
    Call clan.setAlineacion(NuevaAlineacion)
End Sub
Public Sub ReanudarClan(UserIndex As Integer, Optional NombreClan As String = vbNullString)


Dim clan As cClan
Dim fechaMinima As Date
Dim diferenciaHoras As Integer

If UserList(UserIndex).GuildInfo.id = 0 Then
    'Obtengo cual es el clan que disolvio....
    If NombreClan = vbNullString Then
        Set clan = clanes.obtenerUltimoClanDisueltoPorUsuario(UserList(UserIndex).id)
    Else
        Set clan = clanes.getClanPorNombre(NombreClan)
        
        If Not clan.getIDLider = UserList(UserIndex).id Then
             EnviarPaquete Paquetes.mensajeinfo, "No puedes reanudar este clan. El único que lo puede hacer es el último Líder, el personaje que lo disolvió.", UserIndex, ToIndex
             Exit Sub
        End If
    End If
    'Si disolvio alguno, me fijo la fecha de disolucion
    If Not clan Is Nothing Then
        'Paso el minimo de tiempo sin reanudar?
       
        fechaMinima = DateAdd("y", DIAS_MINIMO_DISOLUCION, clan.getFechaDisolucion)
        
        If Date > fechaMinima Then
            'lo reanudo
            Call clanes.ReanudarClan(clan)
            'El usuario cambio de alineacion en este tiempo?
            If Not clan.getAlineacion = obtenerAlineacion(UserIndex) Then
                Call mdClanes.cambiarAlineacionClan(clan, obtenerAlineacion(UserIndex))
            End If
            'Agrego al usuario nuevamente al clan
            Call clan.agregarMiembro(UserList(UserIndex).Name, UserList(UserIndex).id)
            'Le pongo al users los datos del clan
            Call modUsuarios.establecerClanAUsuarioOnline(UserIndex, clan)
            'Lo pongo como lider
            UserList(UserIndex).GuildInfo.EsGuildLeader = 1
            'Aviso a todo el mundo
            EnviarPaquete Paquetes.MensajeGuild, "¡" & UserList(UserIndex).Name & " reanudo el clan " & clan.getNombre & "!", UserIndex, ToIndex
        ElseIf Date = fechaMinima Then
            EnviarPaquete Paquetes.mensajeinfo, "No puedes reanudar el clan todavía. Debes esperar a que termine el día para poder reanudarlo.", UserIndex, ToIndex
        Else
            EnviarPaquete Paquetes.mensajeinfo, "No puedes reanudar el clan. Debes esperar al menos " & DateDiff("y", Date, fechaMinima) & " días para poder reanudarlo.", UserIndex, ToIndex
        End If
    Else
        EnviarPaquete Paquetes.mensajeinfo, "No has disuelto ningún clan anteriormente.", UserIndex, ToIndex
    End If
Else
    EnviarPaquete Paquetes.mensajeinfo, "No puedes reanudar un clan si perteneces a otro.", UserIndex, ToIndex
End If

End Sub

'Realiza las validaciones y se comunica con el usuario
'cuando este quiere disolver su clan
Public Sub disolverclan(UserIndex As Integer)
Dim clan As cClan

    If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
    
        Set clan = UserList(UserIndex).ClanRef

        If Not clan Is Nothing Then
            If clan.getEstado = eEstadoClan.Activo Then
                If Not clan.isElecciones Then
                    'Le aviso a todos los integrantes
                    mdClanes.EnviarPaqueteAClan MensajeGuild, UserList(UserIndex).Name & " ha  disuelto el clan.", clan
                    '
                    LogDesarrollo ("Se disuelve el clan " & clan.getNombre)
                    'disuelvo al clan en si
                    Call clanes.disolverclan(clan)
                    'Le aviso que el clan fue disuelto al que lo disolvio
                    EnviarPaquete Paquetes.MensajeGuild, "Has disuelto tú clan. Para reanudarlo deberás esperar al menos " & DIAS_MINIMO_DISOLUCION & " días y escribir el comando /REANUDARCLAN.", UserIndex, ToIndex
                Else
                    EnviarPaquete Paquetes.mensajeinfo, "No puedes disolver el clan justo el día de las elecciones.", UserIndex, ToIndex
                End If
            Else
                'Error de sistemas. Tiene como referencia un clan inactivo
            End If
        Else
            'Error de SISTEMA. No tiene ref de clan
        End If
    Else
        'Error de USUARIO no es el lider, no puede disolverlo
    End If
End Sub

' ESTA FUNCION ES DE UN COMANDO PARA GMS
Public Sub EcharIntegranteDeClan(nombrePersonaje As String, UserIndex As Integer)

Dim index As Integer
Dim IDClan As Long
Dim IDPJ As Long
Dim EsLider As Boolean
Dim infoPersonaje As ADODB.Recordset

index = NameIndex(nombrePersonaje)
If index > 0 Then '¿Esta online?
    IDClan = UserList(index).GuildInfo.id
    If IDClan > 0 Then
        If UserList(index).GuildInfo.EsGuildLeader = 0 Then
            Call SacarMiembroOnline(index, IDClan, False)
        Else
            EnviarPaquete Paquetes.mensajeinfo, "El personaje es líder del clan.", UserIndex, ToIndex
        End If
    Else
        EnviarPaquete Paquetes.mensajeinfo, "El personaje no tiene clan.", UserIndex, ToIndex
    End If
Else
    
    Call General.cargarAtributosPersonajeOffline(nombrePersonaje, infoPersonaje, "ID, IDCLAN, EsGuildLeaderB", False)
    
    'Obtenemos la informacion del personaje
    If Not infoPersonaje.EOF Then
        IDPJ = infoPersonaje!id
        IDClan = infoPersonaje!IDClan
        EsLider = (infoPersonaje!EsGuildLeaderB = 1)
    Else
        IDPJ = 0
        EnviarPaquete Paquetes.mensajeinfo, "El personaje no existe.", UserIndex, ToIndex
    End If
    
    'Liberamos
    infoPersonaje.Close
    Set infoPersonaje = Nothing
    
    If IDPJ = 0 Then
        EnviarPaquete Paquetes.mensajeinfo, "El personaje no existe.", UserIndex, ToIndex
    ElseIf IDClan = 0 Then
        EnviarPaquete Paquetes.mensajeinfo, "El personaje no tiene clan.", UserIndex, ToIndex
    ElseIf EsLider Then
        EnviarPaquete Paquetes.mensajeinfo, "El personaje es lider del clan. No puede ser expulsado un lider", UserIndex, ToIndex
    Else
        Call SacarMiembroOffline(IDPJ, IDClan, False)
    End If
End If
    
End Sub


'Finalizo las elecciones que correspondan
Public Sub FinalizarElecciones()

Dim i As Integer
Dim clan As cClan

Call clanes.iteradorIniciar
    
For i = 1 To clanes.getCantidad()
    Set clan = clanes.iteradorObtener()
    
    If clan.getDiasProximaEleccion() < 0 Then
        Call clan.FinalizarEleccion
    End If
Next
    
End Sub

Public Sub DisolverAutomaticamente()

Dim i As Integer
Dim clan As cClan

Call clanes.iteradorIniciar

For i = 1 To clanes.getCantidad()
    Set clan = clanes.iteradorObtener()
    
    'Cumple con la condicion de disolucion de clan?
    If cumpleCondicionDisolucionAutomatica(clan) Then
        'Le aviso al clan que ya esta.. el mismo no existe más
        mdClanes.EnviarPaqueteAClan MensajeGuild, "El clan fue disuelto por no cumplir con las condiciones mínimimas de existencia.", clan
        
        LogDesarrollo ("El clan " & clan.getNombre & " fue disuelto automaticamente.")
        
        Call clanes.disolverclan(clan)
    End If
Next

End Sub

'Envia un paquete a todos los integrantes del clan.
'Es para cuando no se tiene la referencia a un clan
'La idea es para los paquetes de mensajaes
'TODO Esto quedo medio feo por el tema del iterador
Public Sub EnviarPaqueteAClan(paquete As Paquetes, mensaje As String, ByRef clan As cClan)

If clan.getCantidadOnline > 0 Then
 
    Dim integrantes As EstructurasLib.ColaConBloques
    Set integrantes = clan.getIntegrantesOnline
    
    integrantes.itIniciar
    
    EnviarPaquete paquete, mensaje, integrantes.itnext
End If
End Sub

'True: el clan cumple con la condicion para ser disuelto automaticamente
'False: el clan NO cumple
'Sino importa el valor sirve para actualizar la fecha de infraccion
'CONDICION ACTUAL:
'Minimo de miembros y tiempo maximo de clan en infraccion

Public Function cumpleCondicionDisolucionAutomatica(clan As cClan) As Boolean

cumpleCondicionDisolucionAutomatica = False

'El clan debe tener al menos X integrantes
    If clan.getCantidadMiembros < DISOLUCION_MINIMOS_INTEGRANTES Then
    'No tiene la cantidad suficiente de integrantes.
    'Tenemos que ver desde hace cuanto que no cumple con esto
        'Empezo hoy con este problema?
        If clan.getFechaInfraccion <> 0 Then
            'Si cumple hace másde Y tiempo, entonces finalmente cumple la condicion
            If Date > DateAdd("y", DISOLUCION_TIEMPO_PERMANENCIA_CONDICION, clan.getFechaInfraccion) Then
                cumpleCondicionDisolucionAutomatica = True
                Exit Function
            End If
        Else
            Call clan.establecerFechaInfraccion(Date)
        End If
    Else ' Estaba en infraccion y ahora dejo de estarlo?
        If clan.getFechaInfraccion <> 0 Then
            clan.establecerFechaInfraccion (0)
        End If
    End If

End Function

Public Sub revisarEstadoClanes()

    ' Clanes. Se ejecuta una vez por día
    If UltimaFechaProcesada < Date Then
        'Disolucion de clanes automaticos
        Call mdClanes.DisolverAutomaticamente
        'Fianlizo las elecciones de los clanes que corresponda
        Call mdClanes.FinalizarElecciones
       'Actualizo la fecha procesada
        UltimaFechaProcesada = Date
        Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "UltimaFechaProcesada", UltimaFechaProcesada)
    End If

End Sub

