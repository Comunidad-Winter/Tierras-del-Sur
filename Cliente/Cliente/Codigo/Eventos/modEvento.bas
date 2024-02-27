Attribute VB_Name = "modEvento"
Option Explicit

Public Enum eEstadoEvento
    esperandoConfirmacionInicio = 0         ' Se espere que se confirme la ejecucion de este evento
    Preparacion = 1                         ' Se esta anunciando, todavia no se abren las inscripciones
    Desarrollandose = 2                     ' Se esta desarrollando el evento.
    Terminado = 3                           ' El evento termino su operatoria
End Enum

Public Enum eEventoPremio                      ' ¿Cómo está expresado el premio?
    monedasDeOro = 1                            ' En monedas de oro.
    porcentajeSobreAcumulado = 2                ' Porcentaje sobre el premio acumulado
End Enum

Public Enum eRangoLimite
    minimo = 1
    maximo = 2
End Enum

Public Enum eCuenta
    ninguna = 0
    estandar = 1
    premium = 2
    todas = 3
End Enum

Public Enum eEventoRing
    ringCualquiera = 1
    ringReto = 2
    ringTorneo = 4
    ringPlantado = 8
    ringAcuatico = 16
End Enum

Public Enum eEventoCaenItems
    nunca = 0
    alFinalizarEvento = 1
    alFinalizarCombate = 2
End Enum

Public Enum eEventoDescanso
    torneo = 1
    reto = 2
    conBoveda = 4
End Enum

' Formas de identificar a un equipo
Public Enum eEventoIdentificacionEquipo
    identificaPersonajes = 1
    identificaClan = 2
    identificaFaccion = 3
End Enum

' Condiciones que se pueden configurar en los eventos
Public Enum eEventoCondicion
    apuestasActivadas = 1
    clanRepetir = 2
    clasesPermitidas = 3
    razaRepetir = 4
    claseRepetir = 5
    objetosPermitidos = 6
    personajesCuenta = 7
    nivelesSumatoria = 8
    clasesGrupo = 9
    hechizosPermitidos = 10
    razasPermitidas = 11
    alineacionesPermitidas = 12
    nivelMinMax = 13
End Enum

' Alineaciones
Public Enum eEventoPersonajesAlineacion
    Ciudadanos = 1
    criminales = 2
    Legionarios = 4
    Armadas = 8
End Enum

Public Enum eEventoTipoAutomatico
    deathmatch = 1
    playoff = 2
    liga = 3
End Enum

' ---------------------------------------------------------------------------- '

Public Type tEventoPlayOffConfig
    clasificacionCompleta As Boolean        ' Todos juegan la misma cantidad de rounds
End Type

Public Type tEventoLigaConfig
    conVuelta As Boolean                    ' Cada equipo lucha dos veces contra el otro
End Type

Public Type tEventoCircular
    activado As Boolean
    cantidadAGanar As Byte                  ' Cantidad de eventos que debe ganar
    eventosExcluido As Byte                 ' Cantidad de eventos que debe esperar en caso de perder
End Type

Private Type tEventoPorObjetos              ' ¿En el evento se disputan los objetos del usuario?
    activado As Boolean
    cuando As eEventoCaenItems              ' ¿Cuando caen los items?
End Type

Public Type tEventoAutomaticoConfig
    tipo As eEventoTipoAutomatico
    maxsRounds  As Byte                     ' Máxima cantidad de Rounds por Combate
    playOffConfig As tEventoPlayOffConfig   ' Configuracion para PlayOff
    ligaConfig As tEventoLigaConfig         ' Configuración para DeathMatch
    configCircular As tEventoCircular       ' ¿Evento circular?
    objetosEnJuego As tEventoPorObjetos     ' ¿Se disputan los objetos en el evento?
End Type

Public Type tEventoManualConfig
    transportarInmediato As Boolean         ' Una vez que termina la inscripcón ¿inmediatamente lleva a los users al descanso?
End Type

Public Type tEventoPremio                        ' Premio
    tipo As eEventoPremio
    valores() As Long
End Type

Public Type tEventoApuestas                ' Sistema de apuestas
    activadas As Boolean
    pozoInicial As Long                     ' Cuanta plara pone TDS de pozo
    tiempoAbiertas As Byte                  ' Minutos que las apuestas estan abiertas
                                            ' (tiempo entre que se cierra la inscripcion y comienzan
                                            ' los combates)
End Type

Public Type tEventoRestriccionCantidad     ' Estructura Auxuliar
    activada As Boolean
    cantidad As Integer
End Type

Public Type tEventoRestriccionRango         ' Estructura Auxuliar
    activada As Boolean
    minimo As Integer
    maximo As Integer
End Type

Private Type tEventoRestriccionGrupoClases
    activada As Boolean
    magicas As Byte                           ' ¿Cuantas clases debe tener el equipo del grupo de "magicas"?
    semiMagicas As Byte                       ' ¿Cuantas clases debe tener el equipo del grupo de "semi-magicas"?
    noMagicas As Byte                         ' ¿Cuantas clases debe tener el equipo del grupo de "no-magicas"?
    trabajadoras As Byte                      ' ¿Cuantas clases debe tener el equipo del grupo de "trabajadoras"?
End Type

Public Type tEventoRestriccionEquipo       ' Restricciones que aplican al equipo
    
    repeticionClan As tEventoRestriccionCantidad      ' ¿Cuantas veces puede repetir clan?
    repeticionClase As tEventoRestriccionCantidad     ' ¿Cuantas veces puede repetir clase?
    repeticionRaza As tEventoRestriccionCantidad      ' ¿Cuantas veces puede repetir raza?
    grupoClases As tEventoRestriccionGrupoClases      ' Conformacion de las Clases
    limiteSumaDeNivel As tEventoRestriccionCantidad   ' ¿Cuanto puede ser el maximo que sume la cantidad de niveles del equipo?

End Type



Public Type tEventoObjetoRestringido
    id As Integer
    cantidad As Long
    tipo As eRangoLimite
End Type

Public Type tEventoRestriccionObjetos       ' Restrincciones al inventario del usuario
    activada As Boolean
    BilleteraVacia    As Boolean            ' ¿No puede tener oro en la billetera?
    restringir As Boolean                   ' ¿Se le obliga a llevar objetos listados?
    objetos() As tEventoObjetoRestringido   ' Restriccion a objetos en el inventario
End Type

Public Type tEventoRestriccionAlineacion    ' Restricciones que aplican al personaje
    activada As Boolean                     ' ¿Esta restriccion se activa?
    ciudadano As Boolean                    ' ¿Pueden los ciudas?
    criminal As Boolean                     ' ¿Pueden los criminales?
    armada As tEventoRestriccionCantidad    ' ¿Pueden los armadas? ¿A partir de que rango?
    caos As tEventoRestriccionCantidad      ' ¿Pueden los caos? ¿A partir de que rango?
End Type

Public Type tEventoRestriccionClases
    activada As Boolean
    clasesPermitidas(1 To 15) As Boolean    ' Debe ser alguna de estas clases
End Type

Public Type tEventoRestriccionRazas
    activada As Boolean
    razasPermitidas(1 To 5) As Boolean      ' Debe ser alguna de estas alineaciones
End Type

Public Type tEventoRestriccionPersonaje     ' Restricciones que aplican al personaje                ' Nivel MinMax del Personaje
    Nivel As tEventoRestriccionRango        ' Sobre el Nivel
    tipoCuenta As eCuenta                   ' Tipo de cuenta a la cual debe estar adherido
    Clase As tEventoRestriccionClases       ' Sobre la Clase
    Raza As tEventoRestriccionRazas         ' Restriccion sobre la raza
    alineacion  As tEventoRestriccionAlineacion ' Debe seguir estas reglas de su alineacion
    inventario As tEventoRestriccionObjetos ' Objetos que puede o debe tener el personaje en el inventario
End Type

Public Type tEventoReglas                   ' Reglas del evento
    hechizos(1 To 41) As Boolean            ' Con respecto a los hechizos. ¿Se puede lanzar o no?
End Type

Public Type tConfigEvento

    Nombre As String                        ' Nombre del Evento
    descripcion As String                   ' Descripcion del evento
    
    costoInscripcion As Long                ' Precio por Personaje que cuesta para ingresar
    
    cantEquiposMinimo As Byte               ' Cantidad Mínima de equipos que se pueden inscribir
    cantEquiposMaxima As Byte              ' Cantidad Maxima de equipos que se pueden inscribir
    
    cantidadIntegrantesEquipo As Byte       ' Cantidad de personajes por equipo
    
    comoIdentificarEquipo As eEventoIdentificacionEquipo ' Como se deben mostrar en pantalla los eventos
    
    importanciaEvento As Byte               ' Importante del evento de 1 a 5
    
    tiempoAnuncio As Integer                ' Cantidad de minutos el cual se anuncia por consola
    tiempoInscripcion As Integer            ' Cantidad de minutos durante el cual está abierta la inscripcion
    tiempoTolerancia As Byte                ' Cantidad de minutos que el usuario puede estar offline cuando se lo llama para jugar
    
    automatico As Boolean                                   ' ¿El evento tiene una parte automatica?
    
    configAutomatico As tEventoAutomaticoConfig             ' Informacion del evento automatico
    configManual As tEventoManualConfig                     ' o manual. Excluyente
        
    tipoRing As eEventoRing
    tipoDescanso As eEventoDescanso
    
    apuestas As tEventoApuestas                           ' Sistema de apuestas

    premio As tEventoPremio                                     ' Informacion del premio que ganará

    
    restriccionesEquipo As tEventoRestriccionEquipo       ' Restricciones a la conformacion del equipo
    restriccionesPersonaje As tEventoRestriccionPersonaje ' Restricciones a los personajes
    
    reglas As tEventoReglas
End Type
