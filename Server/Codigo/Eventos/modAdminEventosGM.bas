Attribute VB_Name = "modAdminEventosGM"
Option Explicit

Public Sub inscribirPersonajes(datos As String, GameMaster As User)
    ' Aux
    Dim longitudNombre As String
    Dim loopC As Integer
    Dim integrantesIndex() As Integer
    Dim tempbyte As Integer
    Dim tempString As String
    Dim error As String
    
    ' Parametros
    Dim testeo As Boolean
    Dim nombreEvento As String
    Dim equipos() As tEventoConfEquipo
    Dim cantidadOfflines As Integer
    
    ' Evento
    Dim evento As iEvento
    
    testeo = StringToByte(datos, 1) = 1
    longitudNombre = StringToByte(datos, 2)
    nombreEvento = mid$(datos, 3, longitudNombre)
    
    Set evento = modEventos.getEventoByNombre(nombreEvento)
        
    ' ¿Existe?
    If evento Is Nothing Then
        EnviarPaquete Paquetes.MensajeAdminEventos, "El evento no existe.", GameMaster.UserIndex, ToIndex
        Exit Sub
    End If
    
    ' ¿Puedo inscribir?
    If Not evento.getEstadoEvento = eEstadoEvento.esperandoConfirmacionInicio Then
        EnviarPaquete Paquetes.MensajeAdminEventos, "No se pueden inscribir personajes a este evento.", GameMaster.UserIndex, ToIndex
        Exit Sub
    End If
    
    ' Obtengo la informacion de los personajes
    cantidadOfflines = parsearListaEquipos(mid$(datos, longitudNombre + 3), equipos)

    error = ""
    If cantidadOfflines > 0 Then

        tempString = "Hay " & cantidadOfflines & " personajes offline: "
        For loopC = 1 To UBound(equipos)
            With equipos(loopC)
                For tempbyte = 1 To UBound(.participantes)
                    If .participantes(tempbyte).index = 0 Then
                        tempString = tempString & " " & .participantes(tempbyte).nick
                    End If
                Next
            End With
        Next

        'Le aviso al cliente.
        error = tempString

    ElseIf cantidadOfflines < 0 Then

        If cantidadOfflines > -10 Then 'Cadena mal enviada
            'Le aviso al cliente.
            error = "Parece que la lista de equipos este mal formada. Por favor, revisala."
        ElseIf cantidadOfflines > -20 Then
            'Le aviso al cliente.
            error = "Parece que la lista de equipos este mal formada. Por favor, revisala."
        End If
        
    End If
    
    ' ¿Error?
    If Len(error) > 0 Then
        EnviarPaquete Paquetes.MensajeAdminEventos, error, GameMaster.UserIndex, ToIndex
        Exit Sub
    End If
    
    ' Si es solo testo termino aca
    If testeo Then
        EnviarPaquete Paquetes.MensajeAdminEventos, "Al parecer todos los personajes están listos para ser sumoneados y comenzar a participar.", GameMaster.UserIndex, ToIndex
        Exit Sub
    End If
    
    For loopC = 1 To UBound(equipos)

        ' Genero la estructura necesaria para inscribir
        ReDim integrantesIndex(1 To UBound(equipos(loopC).participantes)) As Integer

        ' Copios los indices
        For tempbyte = 1 To UBound(equipos(loopC).participantes)
            integrantesIndex(tempbyte) = equipos(loopC).participantes(tempbyte).index
        Next

        ' Inscribo al equipo
        Call evento.agregarEquipo(integrantesIndex)

    Next loopC
    
End Sub
Public Sub obtenerInfoEvento(nombreEvento As String, GameMaster As User)
    Dim evento As iEvento
    Set evento = modEventos.getEventoByNombre(nombreEvento)
    
    If Not evento Is Nothing Then
        EnviarPaquete Paquetes.InfoEventoAdminEventos, ByteToString(evento.getEstadoEvento) & ByteToString(evento.getCantidadParticipantesEquipo) & ByteToString(evento.getCantidadEquiposMax) & ByteToString(Len(evento.getNombre)) & evento.getNombre & evento.obtenerInfoExtendida, GameMaster.UserIndex, ToIndex
    Else
        EnviarPaquete Paquetes.MensajeAdminEventos, "El evento no existe.", GameMaster.UserIndex, ToIndex
    End If
End Sub

Private Sub configReset(config As modEvento.tConfigEvento)
    Dim loopAux As Integer
    
    config.apuestas.activadas = False
    
    config.restriccionesPersonaje.alineacion.activada = False
    config.restriccionesPersonaje.inventario.activada = False
    config.restriccionesPersonaje.tipoCuenta = eCuenta.ninguna
    config.restriccionesPersonaje.clase.activada = False
    config.restriccionesPersonaje.Raza.activada = False
    config.restriccionesPersonaje.Nivel.activada = False
    
    config.restriccionesEquipo.grupoClases.activada = False
    config.restriccionesEquipo.repeticionClan.activada = False
    config.restriccionesEquipo.repeticionClase.activada = False
    config.restriccionesEquipo.repeticionRaza.activada = False
    
    ' Todas Permitidas
    config.restriccionesPersonaje.clase.clasesPermitidas = eClases.indefinido
    config.restriccionesPersonaje.Raza.razasPermitidas = eRazas.indefinido

    ' Todas permitidas
    For loopAux = 1 To 41
        config.reglas.hechizos(loopAux) = True
    Next
End Sub

Private Function cargarConfigDesdeBuffer(ByRef info As String, ByRef config As modEvento.tConfigEvento, ByRef error As String) As Boolean

' Para Parsear
Dim infoEvento() As String
Dim ingresoManual As Boolean
Dim cantidadPremios As Integer
Dim parametrosCondicion() As String
Dim offSet As Byte
Dim hojaDeRuta As String
Dim loopPremio As Byte

' Ingreso Manual de Participantes
Dim cantidadOfflines As Integer

Dim numeroCondicion As Byte

' Auxiliares
Dim tempLong As Long
Dim tempbyte As Byte
Dim TempInt As Integer
Dim tempString As String
Dim loopEquipo As Byte
Dim loopIntegrante As Byte

' Reset
Call configReset(config)

' Parseo la informacion
infoEvento = Split(info, "¦¦")

' Tipo de evento
config.automatico = StringToByte(infoEvento(1), 1) = 0

If Not config.automatico Then
    config.configManual.transportarInmediato = StringToByte(infoEvento(1), 2) = 1
    hojaDeRuta = "NADA"
Else
    
    config.configAutomatico.maxsRounds = StringToByte(infoEvento(1), 2) ' Rounds por Combate
    config.configAutomatico.objetosEnJuego.cuando = StringToByte(infoEvento(1), 3)

    If config.configAutomatico.objetosEnJuego.cuando = eEventoCaenItems.nunca Then
        config.configAutomatico.objetosEnJuego.activado = False
    Else
        config.configAutomatico.objetosEnJuego.activado = True
    End If
    
    config.configAutomatico.configCircular.activado = StringToByte(infoEvento(1), 4) = 1
    
    If config.configAutomatico.configCircular.activado Then
        config.configAutomatico.configCircular.cantidadAGanar = StringToByte(infoEvento(1), 4)
        config.configAutomatico.configCircular.eventosExcluido = StringToByte(infoEvento(1), 5)
        offSet = 7
    Else
        offSet = 5
    End If
    
    config.configAutomatico.tipo = StringToByte(infoEvento(1), offSet)
    
    ' Tipo de Evento Automatico
    If config.configAutomatico.tipo = eEventoTipoAutomatico.deathmatch Then
    ElseIf config.configAutomatico.tipo = eEventoTipoAutomatico.PlayOff Then
        config.configAutomatico.playOffConfig.parametros = mid$(infoEvento(1), offSet + 1, 1)
    ElseIf config.configAutomatico.tipo = eEventoTipoAutomatico.Liga Then
        config.configAutomatico.playOffConfig.parametros = mid$(infoEvento(1), offSet + 1, 1)
    Else
        ' Error!
        error = "No existe el tipo de evento que queres crear."
        Exit Function
    End If
End If


' Datos generales del evento
config.nombre = infoEvento(2)
config.importanciaEvento = StringToByte(infoEvento(3), 1)
config.descripcion = mid$(infoEvento(3), 2)

config.cantEquiposMinimo = val(infoEvento(4))
config.cantEquiposMaxima = val(infoEvento(5))

config.cantidadIntegrantesEquipo = val(infoEvento(6))

config.costoInscripcion = val(infoEvento(7))

config.tiempoAnuncio = val(infoEvento(8))
config.tiempoInscripcion = val(infoEvento(9))
config.tiempoTolerancia = val(infoEvento(10))

' Configuracion de los Pagos
config.premio.tipo = StringToByte(infoEvento(11), 1)
cantidadPremios = StringToByte(infoEvento(11), 2)

ReDim config.premio.valores(1 To cantidadPremios) As Long

For loopPremio = 0 To cantidadPremios - 1
    config.premio.valores(loopPremio + 1) = StringToLong(infoEvento(11), 3 + loopPremio * 4)
Next

' Tipo de Ring
config.tipoRing = StringToByte(infoEvento(12), 1)

' Tipo de descanso
config.tipoDescanso = StringToByte(infoEvento(13), 1)

' ¿Como se debe identificar a los equipos?
config.comoIdentificarEquipo = StringToByte(infoEvento(14), 1)

' ****************************************************************************
' Aqui voy a tener configuraciones especificas del Evento y la configuracion
' de las condiciones que deben cumplir los personajes para ingresar a este evento

For numeroCondicion = 15 To UBound(infoEvento)
    parametrosCondicion = Split(infoEvento(numeroCondicion), ";")
    
    Debug.Print parametrosCondicion(0)
    Select Case parametrosCondicion(0)
    
    '--------------------------------------------------------------------------------------------------'
        Case eEventoCondicion.nivelMinMax ' Nivel Minomo y Maximo
            config.restriccionesPersonaje.Nivel.minimo = StringToByte(parametrosCondicion(1), 1)
            config.restriccionesPersonaje.Nivel.maximo = StringToByte(parametrosCondicion(1), 2)
            config.restriccionesPersonaje.Nivel.activada = True
'--------------------------------------------------------------------------------------------------'
        Case eEventoCondicion.apuestasActivadas   'Apuestas
            
            config.apuestas.activadas = True
            config.apuestas.pozoInicial = StringToLong(parametrosCondicion(1), 1)
            config.apuestas.tiempoAbiertas = StringToByte(parametrosCondicion(1), 5)
   '--------------------------------------------------------------------------------------------------'

        Case eEventoCondicion.clanRepetir  'No repetir clan
        
            config.restriccionesEquipo.repeticionClan.activada = True
            config.restriccionesEquipo.repeticionClan.cantidad = 0

'--------------------------------------------------------------------------------------------------'
        Case eEventoCondicion.clasesPermitidas  'Restriccion de clases
        
            ' Desactivo las que no son posibles
            config.restriccionesPersonaje.clase.activada = True
        
            config.restriccionesPersonaje.clase.clasesPermitidas = 0
            For tempbyte = 1 To Len(parametrosCondicion(1))
                Dim claseConfigId As Byte
                claseConfigId = StringToByte(parametrosCondicion(1), tempbyte)
                config.restriccionesPersonaje.clase.clasesPermitidas = (config.restriccionesPersonaje.clase.clasesPermitidas Or modClases.claseConfigToEnum(claseConfigId))
            Next tempbyte
            config.restriccionesPersonaje.clase.clasesPermitidas = Not config.restriccionesPersonaje.clase.clasesPermitidas
'--------------------------------------------------------------------------------------------------'

        Case eEventoCondicion.claseRepetir  'No repetir clase
                
            tempLong = StringToByte(parametrosCondicion(1), 1)
            
            config.restriccionesEquipo.repeticionClase.activada = True
            config.restriccionesEquipo.repeticionClase.cantidad = tempLong
            '--------------------------------------------------------------------------------------------------'

        Case eEventoCondicion.razaRepetir  'No repetir raza
            ' Máxima cantidad de veces que se puede repetir
            tempLong = StringToByte(parametrosCondicion(1), 1)
            
            config.restriccionesEquipo.repeticionRaza.activada = True
            config.restriccionesEquipo.repeticionRaza.cantidad = tempLong
'--------------------------------------------------------------------------------------------------'

        Case eEventoCondicion.objetosPermitidos  'Limite de objetos (incluido oro en billetera)
        
            
            config.restriccionesPersonaje.inventario.activada = True
            
            'Exclusividad de llevar los items limitados?
            config.restriccionesPersonaje.inventario.restringir = StringToByte(parametrosCondicion(1), 1) = 1
            ' ¿Puede tener oro en la billetera?
            config.restriccionesPersonaje.inventario.BilleteraVacia = StringToByte(parametrosCondicion(1), 2) = 1
                        
            ' Luego tenemos los limites de objetos
            If UBound(parametrosCondicion) > 1 Then 'Hay items?
                ReDim config.restriccionesPersonaje.inventario.objetos(1 To UBound(parametrosCondicion) - 1) As tEventoObjetoRestringido
                
                With config.restriccionesPersonaje.inventario
                
                    For tempbyte = 2 To UBound(parametrosCondicion())
                        .objetos(tempbyte - 1).id = STI(parametrosCondicion(tempbyte), 1)
                        .objetos(tempbyte - 1).cantidad = STI(parametrosCondicion(tempbyte), 3)
                        .objetos(tempbyte - 1).tipo = StringToByte(parametrosCondicion(tempbyte), 5)
                    Next
                
                End With
                
            End If
'--------------------------------------------------------------------------------------------------'

        Case eEventoCondicion.personajesCuenta  'Solo personajes adheridos a cuentas
        
            If StringToByte(parametrosCondicion(1), 1) = 2 Then
                config.restriccionesPersonaje.tipoCuenta = eCuenta.Premium
            Else
                config.restriccionesPersonaje.tipoCuenta = eCuenta.todas
            End If
                       
  '--------------------------------------------------------------------------------------------------'
        Case eEventoCondicion.nivelesSumatoria  'Sumatoria de niveles
        
            config.restriccionesEquipo.limiteSumaDeNivel.activada = True
            config.restriccionesEquipo.limiteSumaDeNivel.cantidad = STI(parametrosCondicion(1), 1)
   '--------------------------------------------------------------------------------------------------'

        Case eEventoCondicion.clasesGrupo  ' Requisitos de las clases que debe tener el grupo
                
            config.restriccionesEquipo.grupoClases.activada = True
            
            config.restriccionesEquipo.grupoClases.magicas = StringToByte(parametrosCondicion(1), 1)
            config.restriccionesEquipo.grupoClases.semiMagicas = StringToByte(parametrosCondicion(1), 2)
            config.restriccionesEquipo.grupoClases.noMagicas = StringToByte(parametrosCondicion(1), 3)
            config.restriccionesEquipo.grupoClases.trabajadoras = StringToByte(parametrosCondicion(1), 4)
   '--------------------------------------------------------------------------------------------------'

        Case eEventoCondicion.hechizosPermitidos  'Hechizos que se permiten
        
            ' Obtengo las reglas
            For tempbyte = 1 To Len(parametrosCondicion(1))
                TempInt = StringToByte(parametrosCondicion(1), tempbyte)
                config.reglas.hechizos(TempInt) = False
            Next tempbyte
            
   '--------------------------------------------------------------------------------------------------'
        Case eEventoCondicion.razasPermitidas  'Razas permitidas
           
            config.restriccionesPersonaje.Raza.razasPermitidas = eClases.indefinido
            config.restriccionesPersonaje.Raza.activada = True
            
            For tempbyte = 1 To Len(parametrosCondicion(1))
                Dim razaConfigId As Byte
                razaConfigId = StringToByte(parametrosCondicion(1), tempbyte)
                config.restriccionesPersonaje.Raza.razasPermitidas = (config.restriccionesPersonaje.Raza.razasPermitidas Or razaConfigToEnum(razaConfigId))
            Next tempbyte
            
            config.restriccionesPersonaje.Raza.razasPermitidas = Not config.restriccionesPersonaje.Raza.razasPermitidas
                      
   '--------------------------------------------------------------------------------------------------'
        Case eEventoCondicion.alineacionesPermitidas ' Alineaciones permitidas
            ' Obtengo la informacion
            tempbyte = StringToByte(parametrosCondicion(1), 1) ' Flags de permitidos

            config.restriccionesPersonaje.alineacion.activada = True
            config.restriccionesPersonaje.alineacion.armada.cantidad = StringToByte(parametrosCondicion(1), 3)  ' Armada
            config.restriccionesPersonaje.alineacion.caos.cantidad = StringToByte(parametrosCondicion(1), 2) ' Caos
            
            config.restriccionesPersonaje.alineacion.armada.activada = tempbyte And eEventoPersonajesAlineacion.Armadas
            config.restriccionesPersonaje.alineacion.caos.activada = tempbyte And eEventoPersonajesAlineacion.Legionarios
            config.restriccionesPersonaje.alineacion.ciudadano = tempbyte And eEventoPersonajesAlineacion.Ciudadanos
            config.restriccionesPersonaje.alineacion.criminal = tempbyte And eEventoPersonajesAlineacion.criminales
    End Select
Next
  
cargarConfigDesdeBuffer = True
Exit Function

End Function
Public Sub parsearInfo(info As String, GameMaster As User)

' Evento Automatico
Dim evento As iEvento_DeathMach

'Condiciones
Dim condicionNoRepPJ As CondicionEventoNoRepPJ
Dim condicionNivel As CondicionEventoNivel
Dim condicionNoRepetirClan As CondicionEventoNoRepClan
Dim condicionClase As CondicionEventoClases
Dim condicionNoRepClase As CondicionEventoNoRepClase
Dim condicionNoRepraza As CondicionEventonoRepRaza
Dim condicionLimiteItems As CondicionEventoLimiteItem
Dim condicionCuenta As iCondicionEventoCuenta
Dim condicionMaxSumNiveles As CondicionEventoSumaNiveles
Dim condicionGrupoClases As CondicionEventoGrupoClases
Dim condicionRaza As CondicionEventoRazas
Dim condicionAlineacion As CondicionEventoAlineacion

' Configuracion del Evento
Dim config As modEvento.tConfigEvento
Dim error As String
Dim loopC As Integer
Dim integrantesIndex() As Integer ' Para el ingreso manual de participantes

' Auxilares
Dim tempbyte As Byte
Dim anunciarEvento As Boolean

' Version con el cual se creo esta configuracion
If Not StringToByte(info, 1) = 1 Then
    EnviarPaquete Paquetes.MensajeAdminEventos, "La versión del Editor de Eventos está vieja. Debes bajar un parche.", GameMaster.UserIndex, ToIndex
    Exit Sub
End If

' Cargo la configuracion
If Not cargarConfigDesdeBuffer(info, config, error) Then
    EnviarPaquete Paquetes.MensajeAdminEventos, error, GameMaster.UserIndex, ToIndex
    Exit Sub
End If

' Validamos!
If Not modEventos.getEventoByNombre(config.nombre) Is Nothing Then
    EnviarPaquete Paquetes.MensajeAdminEventos, "El evento " & config.nombre & " ya existe.", GameMaster.UserIndex, ToIndex
    Exit Sub
End If

' Validaciones ok. Creamos el evento
Set evento = New iEvento_DeathMach

'Configuracion general
Call evento.setNombre(config.nombre)
Call evento.setDescripcion(config.descripcion)
Call evento.setCantidadMaxMinEquipos(config.cantEquiposMinimo, config.cantEquiposMaxima)
Call evento.setPrecioInscripcion(config.costoInscripcion)
Call evento.setCantidadParticipantesEquipo(config.cantidadIntegrantesEquipo)
Call evento.setComoIdentificarEquipos(config.comoIdentificarEquipo)
Call evento.setTiempos(config.tiempoAnuncio, config.tiempoInscripcion, config.tiempoTolerancia)

If config.automatico Then
    Call evento.setCantidadRoundGanadosGanador(config.configAutomatico.maxsRounds)
    Call evento.iEvento_setTiporing(config.tipoRing)
    Call evento.iEvento_setTipoDescanso(config.tipoDescanso)
ElseIf config.automatico = False And config.configManual.transportarInmediato = True Then
    Call evento.iEvento_setTipoDescanso(config.tipoDescanso)
End If

'Configuracion hoja de ruta
If Not config.automatico Then
    Call evento.iEvento_setHojaDeRuta("NADA", "")
Else
    If config.configAutomatico.tipo = eEventoTipoAutomatico.deathmatch Then
        Call evento.iEvento_setHojaDeRuta("DEATHMATCH", "")
    ElseIf config.configAutomatico.tipo = eEventoTipoAutomatico.PlayOff Then
        Call evento.iEvento_setHojaDeRuta("PLAYOFF", config.configAutomatico.playOffConfig.parametros)
    ElseIf config.configAutomatico.tipo = eEventoTipoAutomatico.Liga Then
        Call evento.iEvento_setHojaDeRuta("LIGA", config.configAutomatico.playOffConfig.parametros)
    End If
End If

' Pagos
Call evento.iEvento_establecerTablaDePagos(config.premio.valores, config.premio.tipo)

' Reglas durante el Evento
Call evento.iEvento_setHechizosPermitidos(config.reglas.hechizos)
Call LogTorneos(config.nombre & "-> Configurado hechizos permitidos.")

' Apuestas
If config.apuestas.activadas Then
    Call evento.iEvento_configurarApuestas(True, config.apuestas.pozoInicial, config.apuestas.tiempoAbiertas)
    Call LogTorneos(config.nombre & "-> Configuradas Apuestas.")
End If

' ¿SoloInscripcion?
If config.automatico = False And config.configManual.transportarInmediato = False Then
    Set condicionNoRepPJ = New CondicionEventoNoRepPJ
    Call condicionNoRepPJ.iCondicionEvento_setMaximaMemoria(config.cantEquiposMaxima * config.cantidadIntegrantesEquipo)
    Call evento.iEvento_agregarCondicionIngreso(condicionNoRepPJ)
    Call evento.setSoloInscripcion(True)
End If

' ****************************************************************************
' Aqui voy a tener configuraciones especificas del Evento y la configuracion
' de las condiciones que deben cumplir los personajes para ingresar a este evento

' - Nivel
If config.restriccionesPersonaje.Nivel.activada Then
    
    Set condicionNivel = New CondicionEventoNivel
    Call condicionNivel.setParametros(config.restriccionesPersonaje.Nivel.minimo, config.restriccionesPersonaje.Nivel.maximo)
    
    Call evento.iEvento_agregarCondicionIngreso(condicionNivel)
End If

' - Clase
If config.restriccionesPersonaje.clase.activada Then

    Set condicionClase = New CondicionEventoClases
    Call condicionClase.setParametros(config.restriccionesPersonaje.clase.clasesPermitidas)
    Call evento.iEvento_agregarCondicionIngreso(condicionClase)
    
    Call LogTorneos(config.nombre & "-> Agregada restriccion de clases.")
End If

' - Tipo de Cuenta
If Not config.restriccionesPersonaje.tipoCuenta = eCuenta.ninguna Then  '
    
    Set condicionCuenta = New iCondicionEventoCuenta
    Call condicionCuenta.setParametros(config.restriccionesPersonaje.tipoCuenta = eCuenta.Premium)
    Call evento.iEvento_agregarCondicionIngreso(condicionCuenta)
    
    Call LogTorneos(config.nombre & "-> Agregada restriccion de personajes en cuentas.")
End If

' - Raza
If config.restriccionesPersonaje.Raza.activada Then

    Set condicionRaza = New CondicionEventoRazas
    Call condicionRaza.setParametros(config.restriccionesPersonaje.Raza.razasPermitidas)
    Call evento.iEvento_agregarCondicionIngreso(condicionRaza)

    Call LogTorneos(config.nombre & "-> Agregada restriccion de razas.")
End If

' - Alineacion
If config.restriccionesPersonaje.alineacion.activada Then
    tempbyte = 0
    
    If config.restriccionesPersonaje.alineacion.caos.activada Then tempbyte = tempbyte Or eEventoPersonajesAlineacion.Legionarios
    If config.restriccionesPersonaje.alineacion.armada.activada Then tempbyte = tempbyte Or eEventoPersonajesAlineacion.Armadas
    If config.restriccionesPersonaje.alineacion.ciudadano Then tempbyte = tempbyte Or eEventoPersonajesAlineacion.Ciudadanos
    If config.restriccionesPersonaje.alineacion.criminal Then tempbyte = tempbyte Or eEventoPersonajesAlineacion.criminales
    
    Set condicionAlineacion = New CondicionEventoAlineacion
    Call condicionAlineacion.setParametros(tempbyte, config.restriccionesPersonaje.alineacion.caos.cantidad, config.restriccionesPersonaje.alineacion.armada.cantidad)
    Call evento.iEvento_agregarCondicionIngreso(condicionAlineacion)

    Call LogTorneos(config.nombre & "-> Agregada restriccion de alineaciones.")
End If

' - Inventario
If config.restriccionesPersonaje.inventario.activada Then
    
    Set condicionLimiteItems = New CondicionEventoLimiteItem
    Call condicionLimiteItems.setParametros(config.restriccionesPersonaje.inventario.objetos, config.restriccionesPersonaje.inventario.restringir, config.restriccionesPersonaje.inventario.BilleteraVacia)
    Call evento.iEvento_agregarCondicionIngreso(condicionLimiteItems)
    
    Call LogTorneos(config.nombre & "-> Agregada restriccion de objetos.")
End If

' - Sumatoria de niveles
If config.restriccionesEquipo.limiteSumaDeNivel.activada Then

    Set condicionMaxSumNiveles = New CondicionEventoSumaNiveles
    Call condicionMaxSumNiveles.setParametros(config.restriccionesEquipo.limiteSumaDeNivel.cantidad)
    Call evento.iEvento_agregarCondicionIngreso(condicionMaxSumNiveles)

    Call LogTorneos(config.nombre & "-> Agregada restriccion suma maxima de niveles.")
End If

' - Torneo de Clanes
If config.restriccionesEquipo.repeticionClan.activada Then

    Set condicionNoRepetirClan = New CondicionEventoNoRepClan
    Call condicionNoRepetirClan.iCondicionEvento_setMaximaMemoria(config.cantEquiposMaxima)
    Call evento.iEvento_agregarCondicionIngreso(condicionNoRepetirClan)
    
    Call LogTorneos(config.nombre & "-> Agregada restriccion de clanes.")
End If

' - Repetir clase
If config.restriccionesEquipo.repeticionClase.activada Then

    Set condicionNoRepClase = New CondicionEventoNoRepClase
    Call condicionNoRepClase.setParametros(config.restriccionesEquipo.repeticionClase.cantidad)
    Call evento.iEvento_agregarCondicionIngreso(condicionNoRepClase)

    Call LogTorneos(config.nombre & "-> Agregada restriccion de maximo repeticion de clases.")
End If

' - Repetir raza
If config.restriccionesEquipo.repeticionRaza.activada Then

    Set condicionNoRepraza = New CondicionEventonoRepRaza
    Call condicionNoRepraza.setParametros(config.restriccionesEquipo.repeticionRaza.cantidad)
    Call evento.iEvento_agregarCondicionIngreso(condicionNoRepraza)
    
    Call LogTorneos(config.nombre & "-> Agregada restriccion de maximo repeticion de Raza.")
End If

' - Obligatoriedad de Clases
If config.restriccionesEquipo.grupoClases.activada Then
    Set condicionGrupoClases = New CondicionEventoGrupoClases
    
    Call condicionGrupoClases.setParametros(1, config.restriccionesEquipo.grupoClases.magicas)
    Call condicionGrupoClases.setParametros(2, config.restriccionesEquipo.grupoClases.semiMagicas)
    Call condicionGrupoClases.setParametros(3, config.restriccionesEquipo.grupoClases.noMagicas)
    Call condicionGrupoClases.setParametros(4, config.restriccionesEquipo.grupoClases.trabajadoras)
        
    Call evento.iEvento_agregarCondicionIngreso(condicionGrupoClases)
    
    Call LogTorneos(config.nombre & "-> Agregada restriccion de Grupo de Clases")
End If

' ****************************************************************************'
' ************* Configure al evento                                           '
' ****************************************************************************'
'Lo inicio
If evento.iniciar() Then
    'Lo agrego a la lista de eventos activos
    Call modEventos.agregarEvento(evento)
    ' Avisamos
    EnviarPaquete Paquetes.MensajeAdminEventos, "El evento fue creado correctamente.", GameMaster.UserIndex, ToIndex
Else
    EnviarPaquete Paquetes.MensajeAdminEventos, "Surgió un error al intentar crear el evento. Esto puede ser porque no se lograron obtener los rings o descansos necesarios.", GameMaster.UserIndex, ToIndex
End If

End Sub

' Devuelve la cantidad de Offlines
Public Function parsearListaEquipos(data As String, equipos() As modEvento.tEventoConfEquipo) As Integer
    
Dim infoEquipo() As String
Dim infoIntegrantes() As String
Dim loopEquipo As Byte
Dim loopIntegrante As Byte
Dim UserIndex As Integer

Dim cantidadEquipos As Integer
Dim cantidadIntegrantesOffline As Integer

' Los equipos están separados por "-" y los integrantes por ","
' Hago algunos chequeos basicos

data = Trim$(data)
infoEquipo = Split(data, "-")

' ¿Ok?
If UBound(infoEquipo) = -1 Then
    parsearListaEquipos = -1
    Exit Function
End If

cantidadEquipos = UBound(infoEquipo) + 1
cantidadIntegrantesOffline = 0 ' Voy contando la cantidad de integrantes Offline

' Redimensionamos para guardar la informacion de los equipos
ReDim equipos(1 To cantidadEquipos) As modEvento.tEventoConfEquipo

' Parseamos cada equipos, obteniendo informacion de los integrantes
For loopEquipo = 0 To cantidadEquipos - 1
    
    ' ¿Hay info del equipo?
    If (infoEquipo(loopEquipo) = "") Then
        parsearListaEquipos = -10
        Exit Function
    End If
    
    ' Los participantes vienen separados con coma ,
    infoIntegrantes = Split(infoEquipo(loopEquipo), ",")
    
    ReDim equipos(loopEquipo + 1).participantes(1 To UBound(infoIntegrantes) + 1) As modEvento.tEventoConfParticipante
    
    ' Analizo cada integrante
    For loopIntegrante = 0 To UBound(infoIntegrantes)
    
        ' Guardamos el nombre
        equipos(loopEquipo + 1).participantes(loopIntegrante + 1).nick = Trim$(infoIntegrantes(loopIntegrante))
        
        ' ¿Esta online?
        UserIndex = NameIndex(Trim$(infoIntegrantes(loopIntegrante)))
                 
        If UserIndex > 0 Then
            equipos(loopEquipo + 1).participantes(loopIntegrante + 1).index = UserIndex
        Else
            equipos(loopEquipo + 1).participantes(loopIntegrante + 1).index = 0
            cantidadIntegrantesOffline = cantidadIntegrantesOffline + 1
        End If
        
    Next loopIntegrante
        
Next loopEquipo

parsearListaEquipos = cantidadIntegrantesOffline

End Function

