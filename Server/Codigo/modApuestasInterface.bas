Attribute VB_Name = "modApuestasInterface"
Option Explicit

Private Const TABLA_APUESTAS = "tds_principal"

#If TDSFacil = 0 Then
    Private Const PREFIJO_SERVER = "(TDS)"
#Else
    Private Const PREFIJO_SERVER = "(TDSF)"
#End If

Private Enum eApuestasApiPaquetes
    eCrearApuesta = 1
    eCerrarApuesta = 2
    eCancelarApuesta = 3
    eEstablecerGanadorApuesta = 4
End Enum

Public Function crearApuesta(evento As iEvento, pozoBase As Long, cantidadEstrellas As Byte, tablaEquipos() As tEquipoTablaTorneo) As Integer

Dim mensaje As String
Dim equipos As String
Dim loopEquipo As Byte

'String de equipos
For loopEquipo = 1 To UBound(tablaEquipos())
    If loopEquipo = UBound(tablaEquipos) Then
        equipos = equipos & mid(modTorneos.obtenerStringEquipo(tablaEquipos(loopEquipo), False, eFormatoDisplayEquipo.completo), 1, 30)
    Else
        equipos = equipos & mid(modTorneos.obtenerStringEquipo(tablaEquipos(loopEquipo), False, eFormatoDisplayEquipo.completo), 1, 30) & "-"
    End If
Next loopEquipo

'Mensjae completo
mensaje = Chr$(eApuestasApiPaquetes.eCrearApuesta) & PREFIJO_SERVER & " " & Replace(evento.getNombre, "|", " ") & "|" & _
         Replace(evento.getDescripcion, "|", " ") & "|" & _
         pozoBase & "|" & cantidadEstrellas & "|" & equipos

'Devuelve el id de apuesta
crearApuesta = API_Manager.enviarMensaje(eManagerPaquetes.eApuestas, mensaje)

End Function
                            
Public Sub cerrarApuesta(idApuesta As Integer)
    
Dim mensaje As String

mensaje = Chr$(eApuestasApiPaquetes.eCerrarApuesta) & idApuesta

Call API_Manager.enviarMensaje(eManagerPaquetes.eApuestas, mensaje)

End Sub

Public Sub cancelarApuesta(idApuesta As Integer)

Dim mensaje As String

mensaje = Chr$(eApuestasApiPaquetes.eCancelarApuesta) & idApuesta

Call API_Manager.enviarMensaje(eManagerPaquetes.eApuestas, mensaje)

End Sub

Public Sub establecerGanador(idApuesta As Integer, equipoGanador As Byte)
Dim mensaje As String

mensaje = Chr$(eApuestasApiPaquetes.eEstablecerGanadorApuesta) & idApuesta & "|" & equipoGanador

Call API_Manager.enviarMensaje(eManagerPaquetes.eApuestas, mensaje)

End Sub
