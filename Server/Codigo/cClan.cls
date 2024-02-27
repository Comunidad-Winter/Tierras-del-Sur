VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cClan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'************
' Objetivo de la clase Clan:
'   Contiene la estructura del clan. NO guarda realiza validaciones (eso lo hace modClanes)
'   No envia información a los usuarios
'   Colaboradores: modUsuarios

Option Explicit

Public id As Long
Private nombre As String

Private Lider As String
Private Fundador As String

Private IDLider As Long
Private IDFundador As Long

Private FechaFundacion As Date
Private Eleccion As Date
Private descripcion As String
Private URL As String
Private Novedades As String

Private alineacion As eAlineaciones
'Estado del clan. Un clan puede estar activo o disuelto
'Si la fecha es distinta de null esta disuelto
Private fechaDisolucion As Date

'No es el mejor nombre.
'Es la fecha en la cual el clan empezo a no tener los requisitos fundamentales
'para poder mantenerse
Private fechaInfraccion As Date

'El nombre e ID de los integrantes
'El nombre esta para no tener que pregutarselo todo el tiempo
'a la base de datos cuando alguien quiere consultar los integrantes
Private integrantes As Collection
'El UserIndex de los integrantes Online
Private IntegrantesOnline As EstructurasLib.ColaConBloques
'Private IntegrantesOnline As Collection

'En esta lista se guarda todos los usuarios que votaron
Private Votantes As Collection
Private Solicitudes As Collection

Private Const DIAS_ELECCION = 40
Private Const MAX_CANTIDAD_SOLICITUDES = 30
Private Const CANTIDAD_CODECS = 8
'Esta variable esta repetidas en el modulo de clanes
Private Const MAX_CANTIDAD_MIEMBROS = 20

Private codecs(0 To CANTIDAD_CODECS - 1) As String


Private Sub Class_Initialize()
    Set integrantes = New Collection
    'Set IntegrantesOnline = New Collection
    Set IntegrantesOnline = New EstructurasLib.ColaConBloques
    Set Solicitudes = New Collection
    Set Votantes = New Collection
    
    IntegrantesOnline.setCantidadElementosNodo (MAX_CANTIDAD_MIEMBROS)
    
    Dim i As Byte
    
    For i = 1 To UBound(codecs)
        codecs(i) = ""
    Next
End Sub

'Devuelve el ID del clan
Public Function crear(IDFundador_ As Long, NombreFundador_, NombreClan_ As String, descripcion_ As String, Alineacion_ As Byte, Web_ As String, Codecs_() As String) As Integer
    Dim info As ADODB.Recordset
    Dim i As Byte
    
    nombre = NombreClan_
    descripcion = descripcion_

    IDLider = IDFundador_
    IDFundador = IDFundador_

    Lider = NombreFundador_
    Fundador = NombreFundador_

    alineacion = Alineacion_
    Novedades = "Clan iniciado."
    
    FechaFundacion = Now

    URL = Web_

    Eleccion = Date + DIAS_ELECCION
    'Agrego el clan a la base de datos
    sql = "INSERT INTO " & DB_NAME_PRINCIPAL & ".clanes(Nombre,IDLider,IDFundador,Descripcion,URL,Alineacion,Eleccion,News) values " & _
"('" & nombre & "','" & IDLider & "','" & IDFundador & "','" & mysql_real_escape_string(descripcion) & "','" & URL & "'," & alineacion & ",'" & Eleccion & "','" & Novedades & "')"

    conn.Execute sql, , adExecuteNoRecords
    
    sql = "SELECT last_insert_id() AS id FROM " & DB_NAME_PRINCIPAL & ".clanes"
    Set info = conn.Execute(sql, , adCmdText)

    'Obtengo el ID del clan
    id = info!id

    info.Close
    Set info = Nothing
    
    'Agrego los codecs
    sql = "INSERT INTO " & DB_NAME_PRINCIPAL & ".clanes_codecs(IDClan,IDCodec,Texto) values "
    
    For i = 1 To UBound(Codecs_)
        sql = sql & "(" & id & "," & i & ",'" & mysql_real_escape_string(Codecs_(i)) & "')"
        codecs(i - 1) = Codecs_(i)
        
        If i < UBound(Codecs_) Then sql = sql & ","
    Next i
    
    conn.Execute sql, , adExecuteNoRecords
    
    crear = id
End Function

'
Public Sub IniciarDesdeDB(IDClan As Integer)

    Dim sql As String
    Dim infoClan As ADODB.Recordset

    'Cargo la información general del clan
    sql = "SELECT clan.*, usr1.NickB AS Lider, usr2.NickB AS Fundador, info_dis.Fecha AS fechaDisolucion, info_infraccion.Fecha as fechaInfraccion FROM " & DB_NAME_PRINCIPAL & ".clanes AS clan LEFT JOIN " & DB_NAME_PRINCIPAL & _
    ".usuarios AS usr1 ON clan.IDLider = usr1.ID LEFT JOIN " & DB_NAME_PRINCIPAL & ".usuarios AS usr2 ON clan.IDFundador = usr2.ID " & _
    "LEFT JOIN " & DB_NAME_PRINCIPAL & ".clanes_disueltos AS info_dis ON info_dis.IDClan = clan.ID " & _
    "LEFT JOIN " & DB_NAME_PRINCIPAL & ".clanes_infraccion AS info_infraccion ON info_infraccion.IDClan = clan.ID " & _
    "WHERE clan.ID=" & IDClan

    Set infoClan = conn.Execute(sql)

    'Seteo los datos en la variable del clan
    id = infoClan!id
    nombre = infoClan!nombre
    
    IDFundador = 0
    Fundador = "Sin fundador."
    'Chequeo que el fundador y lider se hayan cargado correctamente. Por si fueron borrados o es lider es ID= 0. No tiene lider
    If Not IsNull(infoClan!Fundador) Then
        If infoClan!IDFundador > 0 Then
            IDFundador = infoClan!IDFundador
            Fundador = infoClan!Fundador
        End If
    End If
    
    IDLider = 0
    Lider = "Sin lider."
    
    If Not IsNull(infoClan!Lider) Then
        If infoClan!IDLider > 0 Then
            IDLider = infoClan!IDLider
            Lider = infoClan!Lider
        End If
    End If

    'Si no tiene fecha de disolución quiere decir que no fue disuelto
    If Not IsNull(infoClan!fechaDisolucion) Then
        fechaDisolucion = infoClan!fechaDisolucion
    Else
        fechaDisolucion = 0
    End If
    
    'Si no tiene la fecha de infraccion quiere decir que el clan esta ok y no correcte
    'peligro de ser disuelto
    If Not IsNull(infoClan!fechaInfraccion) Then
        fechaInfraccion = infoClan!fechaInfraccion
    Else
        fechaInfraccion = 0
    End If
    
    FechaFundacion = infoClan!fecha
    Eleccion = infoClan!Eleccion
    descripcion = infoClan!descripcion
    URL = infoClan!URL
    alineacion = infoClan!alineacion
    
    Novedades = infoClan!news
    
    'Libero
    infoClan.Close
    Set infoClan = Nothing
    
    CargarIntegrantes
    CargarSolicitudes
    CargarCodexs
    CargarVotos
    
End Sub

Private Sub CargarIntegrantes()

    Dim sql As String
    Dim infoIntegrantes As ADODB.Recordset
    Dim Nuevo As cIntegrante
    
    sql = "SELECT integrantes.IDPJ as ID, usr.NickB as Nombre FROM " & DB_NAME_PRINCIPAL & ".usuarios AS usr INNER JOIN " & DB_NAME_PRINCIPAL & ".clanes_integrantes AS integrantes ON integrantes.IDPJ = usr.ID  WHERE integrantes.IDClan = " & id
    
    Set infoIntegrantes = conn.Execute(sql)

    Do While (infoIntegrantes.EOF = False)
            
        Set Nuevo = New cIntegrante
        Call Nuevo.iniciar(infoIntegrantes!nombre, infoIntegrantes!id)
        Call integrantes.Add(Nuevo)
        
        infoIntegrantes.MoveNext
    Loop

    infoIntegrantes.Close
    Set infoIntegrantes = Nothing
    
End Sub

Private Sub CargarSolicitudes()

    Dim solicitud As cSolicitud
    Dim sql As String
    Dim infoSolicitudes As ADODB.Recordset

    sql = "SELECT  solicitud.IDPJ, usr.NickB as Nombre,solicitud.Texto AS Descripcion, solicitud.Fecha FROM " & DB_NAME_PRINCIPAL & ".usuarios AS usr INNER JOIN " & DB_NAME_PRINCIPAL & ".clanes_solicitudes AS solicitud ON solicitud.IDPJ = usr.ID  WHERE solicitud.IDClan = " & id

    Set infoSolicitudes = conn.Execute(sql)

    Do While (infoSolicitudes.EOF = False)
    
        Set solicitud = New cSolicitud
        Call solicitud.iniciar(infoSolicitudes!nombre, infoSolicitudes!IDPJ, infoSolicitudes!descripcion, infoSolicitudes!fecha)
        Call Solicitudes.Add(solicitud)
    
        infoSolicitudes.MoveNext
    Loop

    infoSolicitudes.Close
    Set infoSolicitudes = Nothing

End Sub

Private Sub CargarCodexs()

    Dim sql As String
    Dim infoCodexs As ADODB.Recordset
    Dim i As Byte

    sql = "SELECT IDCodec, Texto FROM " & DB_NAME_PRINCIPAL & ".clanes_codecs WHERE IDClan = " & id & " LIMIT 8"

    Set infoCodexs = conn.Execute(sql)

    Do While (infoCodexs.EOF = False)
        codecs(i) = infoCodexs!texto
        i = i + 1
        infoCodexs.MoveNext
    Loop

    infoCodexs.Close
    Set infoCodexs = Nothing
End Sub
'Carga las personas que ya votaron hasta este momento
Private Sub CargarVotos()
    Dim sql As String
    Dim info As ADODB.Recordset
   
    sql = "SELECT IDVotante FROM " & DB_NAME_PRINCIPAL & ".clanes_votos WHERE IDClan = " & id

    Set info = conn.Execute(sql)
    
    Do While (info.EOF = False)
        Votantes.Add (info!IDVotante)
        info.MoveNext
    Loop

    info.Close
    Set info = Nothing
       
End Sub
Private Sub ReCargarVotos()
    Set Votantes = New Collection
    CargarVotos
End Sub

Public Function agregarMiembro(nombre As String, IDPJ As Long) As Boolean
    Dim Nuevo As cIntegrante
    Set Nuevo = New cIntegrante
    
    Call Nuevo.iniciar(nombre, IDPJ)
    
    Call integrantes.Add(Nuevo)
    
    sql = "INSERT INTO " & DB_NAME_PRINCIPAL & ".clanes_integrantes(IDClan ,IDPJ) VALUES(" & id & "," & IDPJ & ")"

    conn.Execute (sql), , adExecuteNoRecords
End Function

Public Sub quitarMiembro(IDPJ As Long, UserIndex As Integer)
    'Lo quito de la lista de integrantes
    
    Dim loopC As Integer
    For loopC = 1 To integrantes.Count
            If integrantes(loopC).getIDPJ = IDPJ Then
                integrantes.Remove loopC
                Exit For
            End If
    Next loopC
    'Me fijo y lo quito de la lista de onlines
    
    If UserIndex > 0 Then
        IntegrantesOnline.eliminar (UserIndex)
    End If
    
    'Lo elimino de la base de datos de integrantes
    sql = "DELETE FROM " & DB_NAME_PRINCIPAL & ".clanes_integrantes WHERE IDClan=" & id & " and IDPJ=" & IDPJ

    conn.Execute (sql), , adExecuteNoRecords
    'Si estoy en elecciones, quito el voto que hizo el y todos los votos que recibio.
    'Esto para que el lider no meta gente, la haga votar por el y luego la saque.
    If Me.isElecciones Then
        Call quitarVotosDePersonaje(IDPJ)
    End If
    
    'Actualizo el personaje
    Call modUsuarios.quitarClanUsuario(IDPJ)
End Sub

'Quita los votos que recibio esta persona y los votos que realizo.
Public Sub quitarVotosDePersonaje(idPersonaje As Long)
    sql = "DELETE FROM " & DB_NAME_PRINCIPAL & ".clanes_votos WHERE IDClan=" & id & " and (IDVotado = " & idPersonaje & " OR IDVotante=" & idPersonaje & ")"
    conn.Execute sql, , adExecuteNoRecords
    'Re cargo los votos
    ReCargarVotos
End Sub
Public Sub agregarSolicitud(solicitud As cSolicitud)
    Call Solicitudes.Add(solicitud)
    If Solicitudes.Count > MAX_CANTIDAD_SOLICITUDES Then Solicitudes.Remove (1)
End Sub
'Remueve la solicutd generada por el personaje
Public Sub removerSolicitud(IDPJ As Long)
    Dim loopC As Integer
    For loopC = 1 To Solicitudes.Count
            If Solicitudes(loopC).getIDPJ = IDPJ Then
                Solicitudes.Remove (loopC)
                Exit Sub
            End If
    Next loopC
End Sub

'Devuelve True si el usuario ya habia mandado una solicitud a este clan
'False de lo contrario
Public Function existeSolicitud(IDPJ As Long) As Boolean
    Dim loopC As Integer
    For loopC = 1 To Solicitudes.Count
            If Solicitudes(loopC).getIDPJ = IDPJ Then
                existeSolicitud = True
                Exit Function
            End If
    Next loopC
    existeSolicitud = False
End Function

'Expulsa a todos los integrantes, incluido al lider
Private Sub ExpulsarTodosIntegrantes()
    Dim contador As Byte
    
    Do While integrantes.Count > 0
        Call quitarMiembro(integrantes(1).getIDPJ(), 0)
    Loop
    
    'No hay nadie online porque los eche a todos (esto no seria falta si en el quitar miembro se le pasaria el userindex)
    Call IntegrantesOnline.vaciar
End Sub
'Saco a todos menos a lider.
Public Sub ExpulsarTodosIntegrantesMenosLider()
    Dim contador As Byte
    Dim IDPJ As Long
    
    For contador = 1 To integrantes.Count
        IDPJ = integrantes(contador).getIDPJ()
        
        If IDPJ <> IDLider Then
            Call quitarMiembro(IDPJ, IDIndex(IDPJ))
        End If
    Next
End Sub
Public Function YaVoto(IDPJ As Long) As Boolean
    Dim loopC As Integer
    For loopC = 1 To Votantes.Count
            If Votantes.Item(loopC) = IDPJ Then
                YaVoto = True
                Exit Function
            End If
    Next loopC
    YaVoto = False
End Function
Public Sub ComputarVoto(IDVotante As Long, IDVotado)
Dim sql As String
'Inserto el voto en la base de datos
   
sql = "INSERT INTO " & DB_NAME_PRINCIPAL & ".clanes_votos(IDClan ,IDVotante,IDVotado) VALUES(" & id & "," & IDVotante & "," & IDVotado & ")"

conn.Execute (sql), , adExecuteNoRecords
'Agrego el usuario a la lista de usuarios que votaron

Call Votantes.Add(IDVotante)
End Sub
'Cambio el lider
Private Sub CambiarLider(IDNuevoLider As Long, NombreNuevoLider As String)

Dim IDViejoLider As Long

'Guardo el viejo lider
IDViejoLider = IDLider
Call setLider(IDNuevoLider, NombreNuevoLider)
'Actualizo la informacion del personaje
Call ActualizarLiderato(IDViejoLider, False)
Call ActualizarLiderato(IDNuevoLider, True)

End Sub
'Actualiza si un usuario es lider del clan o no
Private Sub ActualizarLiderato(idPersonaje As Long, EsLider As Boolean)

   Call modUsuarios.establecerLiderazgo(idPersonaje, EsLider)
   
End Sub

Public Sub FinalizarEleccion()
'Obtengo la lista de votados ordenado de mayor a menor cantidad de votos
'En caso de igualidad de cantidad de votos se ordena por fecha de ingreso al clna
Dim sql As String

Dim InfoVotos As ADODB.Recordset

Dim LiderElegido  As Boolean

Dim infoVotosLider As ADODB.Recordset
Dim VotosLider As Byte

Dim NombreVotado  As String
Dim IDVotado As Long
Dim Votos As Byte

Dim ViejoLider As Long

'Finalizar solo si hay integrantes en el clan. Por las dudas...
Dim razonLog As String
razonLog = ""

If integrantes.Count > 0 Then
    sql = "SELECT count(idvotado) AS Cant_votos, IDVotado, integrante.fecha,usr.NickB as Nick FROM " & DB_NAME_PRINCIPAL & ".clanes_votos AS votos INNER JOIN " & DB_NAME_PRINCIPAL & ".clanes_integrantes as integrante on integrante.IDPJ = idvotado INNER JOIN " & DB_NAME_PRINCIPAL & ".usuarios AS usr ON usr.ID = idvotado WHERE votos.IDClan=" & id & " GROUP BY IDVotado ORDER BY Cant_votos DESC, fecha ASC;"
    Set InfoVotos = conn.Execute(sql)

    LiderElegido = False
    'Guardo quien era el anterior lider
    ViejoLider = IDLider
    '¿Alguien voto?
    If Not InfoVotos.EOF Then
        'Obtengo los votos que tuvo el lider
        sql = "SELECT count(idvotado) as Cantidad_Votos FROM " & DB_NAME_PRINCIPAL & ".clanes_votos AS votos WHERE votos.IDClan=" & id & " AND votos.IDVotado=" & IDLider & " GROUP BY votos.IDVotado"

        Set infoVotosLider = conn.Execute(sql)
        'Recibio algún voto?
        If infoVotosLider.EOF Then VotosLider = 0 Else VotosLider = infoVotosLider!Cantidad_Votos
        
        'Al igual cantidad de votos la prioridad la tiene el lider
        Do While (InfoVotos.EOF = False And LiderElegido = False)
            IDVotado = InfoVotos!IDVotado
            NombreVotado = InfoVotos!nick
            Votos = InfoVotos!Cant_votos
            'Tiene prioridad el lider en caso de que otro usuario tenga igual votos que el
            If IDVotado <> IDLider And Votos > VotosLider Then
                If isMiembro(IDVotado) Then
                    'Ya tengo al nuevo lider.
                    Call CambiarLider(IDVotado, NombreVotado)
                    'Guardo el log
                    razonLog = "El mas votado fue " & NombreVotado & " con " & Votos
                    LiderElegido = True
                End If ' El ganador dejo de ser integrante del clan durante el dia de elecciones
            Else
                'El más votado fue el lider, no cambio lada
                razonLog = "El mas votado fue el lider con " & Votos
                LiderElegido = True
            End If
        
            InfoVotos.MoveNext 'Obtengo el siguiente integrante mas votado
        Loop
    Else
    razonLog = "Nadie voto, continua el lider"
    End If 'Nadie voto. Continua el lider
    
    'Borro todos los votos
    sql = "DELETE FROM " & DB_NAME_PRINCIPAL & ".clanes_votos WHERE IDClan=" & id
    conn.Execute sql, , adExecuteNoRecords

    'Actualizo la fecha de las proximas elecciones
    setEleccion (Eleccion + 40)
    'Actualizo las novedades
    If IDLider <> ViejoLider Then
        setNovedades (Lider & " ha ganado las elecciones y se ha convertido en el nuevo lider del clan.")
    Else
        setNovedades ("Finalizada las elecciones " & Lider & " mantiene su mandato como lider por un nuevo periodo.")
    End If
    
    
    LogDesarrollo ("Elecciones de " & nombre & " " & razonLog)
    
End If 'No tiene integrantes. Clan cerrado

End Sub
Public Sub disolver()

Dim sql As String
Dim loopC As Integer


Call ExpulsarTodosIntegrantes

fechaDisolucion = Now

'Agrego el registro a la base de datos que marca que el clan fue disuelto
sql = "INSERT INTO " & DB_NAME_PRINCIPAL & ".clanes_disueltos(IDClan, IDPJ) VALUES(" & id & "," & IDLider & ")"
conn.Execute (sql), , adExecuteNoRecords

'Si lo disuelvo dejo de estar en infraccion
Call establecerFechaInfraccion(0)

End Sub

Public Sub reanudar()

Dim diasDiferenciaEleccionDisolucion As Integer

'Elimino el registro del estado del clan en la base de datos
sql = "DELETE FROM " & DB_NAME_PRINCIPAL & ".clanes_disueltos WHERE IDClan=" & id
conn.Execute (sql), , adExecuteNoRecords

'¿Que pasa con la fecha de eleccion?
'¿Cuantos dias antes de la eleccion lo disolvi?
diasDiferenciaEleccionDisolucion = DateDiff("d", fechaDisolucion, Eleccion)

'Al día de hoy le agrego la cantidad de dias que faltaban para la proxima eleccion cuando
'disolvi el clan
Eleccion = DateAdd("d", diasDiferenciaEleccionDisolucion, Date)

fechaDisolucion = 0

End Sub

Public Function isMiembro(IDPJ As Long) As Boolean
    Dim loopC As Integer
    For loopC = 1 To integrantes.Count
            If integrantes(loopC).getIDPJ = IDPJ Then
                isMiembro = True
                Exit Function
            End If
    Next loopC
    isMiembro = False
End Function

Public Function isElecciones() As Boolean
    isElecciones = (Eleccion = Date)
End Function

Public Function isAlineacionCompatible(Alineacion_ As eAlineaciones) As Boolean

If alineacion = Alineacion_ Then
    isAlineacionCompatible = True
Else
    isAlineacionCompatible = False
End If

End Function



'***************************************************************************************************************
'                           SETTERS Y GETTERS
'***************************************************************************************************************
Public Sub setOnline(UserIndex As Integer)
    IntegrantesOnline.agregar (UserIndex)
End Sub

Public Sub setOffline(UserIndex As Integer)
   Call IntegrantesOnline.eliminar(UserIndex)
End Sub
Public Sub setNovedades(Novedades_ As String)
    Novedades = Novedades_
    sql = "UPDATE " & DB_NAME_PRINCIPAL & ".clanes SET News='" & mysql_real_escape_string(Novedades) & "' WHERE ID = " & id
    conn.Execute sql, , adExecuteNoRecords
End Sub
Public Sub setURL(URL_ As String)
    URL = URL_
    sql = "UPDATE " & DB_NAME_PRINCIPAL & ".clanes SET URL='" & mysql_real_escape_string(URL) & "' WHERE ID = " & id
    conn.Execute sql, , adExecuteNoRecords
End Sub
Public Sub setDescripcion(descripcion_ As String)
    Dim sql As String
    descripcion = descripcion_
    sql = "UPDATE " & DB_NAME_PRINCIPAL & ".clanes SET Descripcion='" & mysql_real_escape_string(descripcion) & "' WHERE ID = " & id
    conn.Execute sql, , adExecuteNoRecords
End Sub

Public Sub setNombre(NombreClan_ As String)
    Dim sql As String
    nombre = NombreClan_
    sql = "UPDATE " & DB_NAME_PRINCIPAL & ".clanes SET Nombre='" & mysql_real_escape_string(nombre) & "' WHERE ID = " & id
    conn.Execute sql, , adExecuteNoRecords
End Sub

Private Sub setLider(IDLider_ As Long, NombreLider_ As String)
    Dim sql As String
    'Actualizo el ID y el nombre
    IDLider = IDLider_
    Lider = NombreLider_
    'Actualizo en la base de datos
    sql = "UPDATE " & DB_NAME_PRINCIPAL & ".clanes SET IDLider=" & IDLider_ & " WHERE ID = " & id
    conn.Execute sql, , adExecuteNoRecords
End Sub

Public Sub setAlineacion(Alineacion_ As eAlineaciones)
    Dim sql As String
    alineacion = Alineacion_
    sql = "UPDATE " & DB_NAME_PRINCIPAL & ".clanes SET Alineacion='" & alineacion & "' WHERE ID = " & id
    conn.Execute sql, , adExecuteNoRecords
End Sub

Public Sub setCodecs(Codecs_() As String)
    Dim i As Byte
    
    For i = 0 To CANTIDAD_CODECS - 1
        codecs(i) = Codecs_(i)
        sql = "UPDATE " & DB_NAME_PRINCIPAL & ".clanes_codecs SET Texto='" & mysql_real_escape_string(Codecs_(i)) & "' WHERE IDClan=" & id & " AND IDCodec=" & i + 1
        conn.Execute sql, , adExecuteNoRecords
    Next i
End Sub
Public Sub setEleccion(Eleccion_ As Date)
    Dim sql As String
    Eleccion = Eleccion_
    sql = "UPDATE " & DB_NAME_PRINCIPAL & ".clanes SET Eleccion='" & Eleccion & "' WHERE ID = " & id
    conn.Execute sql, , adExecuteNoRecords
End Sub
Public Function getLider() As String
    getLider = Lider
End Function
Public Function getIDLider() As String
    getIDLider = IDLider
End Function
Public Function getIDFundador() As String
    getIDFundador = IDFundador
End Function
Public Function getWeb() As String
    getWeb = URL
End Function

Public Function getCantidadMiembros() As String
    getCantidadMiembros = integrantes.Count
End Function

Public Function getDiasProximaEleccion() As String
    getDiasProximaEleccion = DateDiff("d", Now, Eleccion)
End Function
Public Function getAlineacion() As eAlineaciones
    getAlineacion = alineacion
End Function
Public Function getCodec(id As Byte) As String
    getCodec = codecs(id)
End Function

Public Function getDescripcion() As String
    getDescripcion = descripcion
End Function
Public Function getNombre() As String
    getNombre = nombre
End Function

Public Function getFundador() As String
    getFundador = Fundador
End Function

Public Function getDiaFundacion() As String
    getDiaFundacion = FechaFundacion
End Function
Public Function getIntegrantesString() As String
    Dim i As Integer
    Dim cadena As String
    
    For i = 1 To integrantes.Count()
            cadena = cadena & integrantes(i).getNombre() & ","
    Next i
    
    getIntegrantesString = cadena
End Function

Public Function getNovedades() As String
    getNovedades = Novedades
End Function
Public Function getCantidadSolicitudes() As Integer
    getCantidadSolicitudes = Solicitudes.Count
End Function

Public Function getSolicitudesString() As String
    Dim i As Integer
    Dim cadena As String
    
    For i = 1 To Solicitudes.Count()
            cadena = cadena & Solicitudes(i).getNombre() & ","
    Next i
    
    getSolicitudesString = cadena
End Function
'Deuvleve la correspondiente solicitud del personaje
Public Function getSolicitudPorNombre(nombre As String) As cSolicitud
    Dim loopC As Integer
    nombre = UCase$(nombre)
    For loopC = 1 To Solicitudes.Count
            If Solicitudes(loopC).getNombre = nombre Then
                Set getSolicitudPorNombre = Solicitudes.Item(loopC)
                Exit Function
            End If
    Next loopC
    Set getSolicitudPorNombre = Nothing
End Function
Public Function getIDMiembro(nombre As String) As Long
    Dim loopC As Integer
    For loopC = 1 To integrantes.Count
            If integrantes(loopC).getNombre = UCase$(nombre) Then
                getIDMiembro = integrantes(loopC).getIDPJ()
                Exit Function
            End If
    Next loopC
    getIDMiembro = 0
End Function

Public Function getIntegrantesOnline() As EstructurasLib.ColaConBloques
    Set getIntegrantesOnline = IntegrantesOnline
End Function

Public Function getIntegranteID(numero As Integer) As Long
   getIntegranteID = integrantes(numero).getIDPJ()
End Function

Public Function getCantidadOnline() As Integer
    getCantidadOnline = IntegrantesOnline.getCantidadElementos
End Function

Public Function getEstado() As eEstadoClan
    If fechaDisolucion = 0 Then
        getEstado = eEstadoClan.Activo
    Else
        getEstado = eEstadoClan.Disuelto
    End If
End Function

Public Function getFechaDisolucion() As Date
    getFechaDisolucion = fechaDisolucion
End Function

Public Sub establecerFechaInfraccion(fecha As Date)
'Si es 0 quiere decir que no esta en infraccion
If fecha = 0 Then
    If fechaInfraccion <> 0 Then
        'Estaba en infraccion, actualizo el registro
        'En la base de datos
        conn.Execute "DELETE FROM " & DB_NAME_PRINCIPAL & ".clanes_infraccion WHERE IDClan = " & id, , adExecuteNoRecords
        fechaInfraccion = 0
        
        'Log
        LogDesarrollo ("El clan <" & nombre & "> dejo de estar en infraccion.")
    End If 'Ya era nula, no hacemos nada
Else ' Esta en infraccion
    If fechaInfraccion <> 0 Then
        'Ya tenia una... la actualizo
        conn.Execute sql, "UPDATE " & DB_NAME_PRINCIPAL & ".clanes_infraccion SET Fecha='" & Eleccion & "' WHERE IDClan = " & id, adExecuteNoRecords
    
        'Log
        LogDesarrollo ("Se actualizo la fecha de infraccion de <" & nombre & ">")
        
    Else 'No tenia fecha de infraccion
        'Establezco la fecha
        fechaInfraccion = fecha
        'Agrego el registro en la base de datos
        conn.Execute "INSERT INTO " & DB_NAME_PRINCIPAL & ".clanes_infraccion (IDClan, Fecha) VALUES (" & id & ",'" & fechaInfraccion & "')", adExecuteNoRecords
    
        'Log
        LogDesarrollo ("El clan <" & nombre & "> entro en infraccion.")
    End If
End If
End Sub

Public Function getFechaInfraccion() As Date
    getFechaInfraccion = fechaInfraccion
End Function
