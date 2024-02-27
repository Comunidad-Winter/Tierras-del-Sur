VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CondicionEventoAlineacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iCondicionEvento

Private permitidos As Byte
Private rangoMinimoCaos As Byte
Private rangoMinimoArmada As Byte

Private Enum eEventoPersonajesAlineacion
    Ciudadanos = 1
    criminales = 2
    Legionarios = 4
    Armadas = 8
End Enum

Public Sub setParametros(ByVal permitidos_ As Byte, ByVal rangoMinimoCaos_ As Byte, ByVal rangoMinimoArmada_ As Byte)

    permitidos = permitidos_
    rangoMinimoCaos = rangoMinimoCaos_
    rangoMinimoArmada = rangoMinimoArmada_

End Sub

Private Sub iCondicionEvento_asentarIngresoEquipo(integrantesIndex() As Integer)
End Sub

Private Function iCondicionEvento_obtenerDescripcion() As String
    Dim temp As String
    
    temp = "Alineaciones permitidas:"
    
    If (permitidos And eEventoPersonajesAlineacion.Ciudadanos) Then
        temp = temp & " - Ejército Índigo"
    End If
    
    If (permitidos And eEventoPersonajesAlineacion.criminales) Then
        temp = temp & " - Ejército Escarlata"
    End If
    
    If (permitidos And eEventoPersonajesAlineacion.Legionarios) Then
        temp = temp & " - Legionarios (rango " & rangoMinimoCaos & " o superior)"
    End If
    
    If (permitidos And eEventoPersonajesAlineacion.Armadas) Then
        temp = temp & " - Armadas (rango " & rangoMinimoArmada & " o superior)"
    End If
    
    iCondicionEvento_obtenerDescripcion = temp
End Function

Private Function cumpleCondicionPersonaje(personaje As User) As Boolean

    Dim esCrimi As Boolean
    
    cumpleCondicionPersonaje = False

    '¿Es amrada?
    If personaje.faccion.ArmadaReal = 1 Then
        
        If (permitidos And eEventoPersonajesAlineacion.Armadas) Then cumpleCondicionPersonaje = True

        ' Salimos si o si
        Exit Function
    End If
    
    '¿Es del caos?
    If personaje.faccion.FuerzasCaos = 1 Then
        
        If (permitidos And eEventoPersonajesAlineacion.Legionarios) Then cumpleCondicionPersonaje = True
    
        Exit Function
    End If
        
    ' ¿Es criminal?
    If personaje.faccion.alineacion = eAlineaciones.caos Then
        If (permitidos And eEventoPersonajesAlineacion.criminales) Then cumpleCondicionPersonaje = True
        Exit Function
    End If
    
    ' ¿Es ciudadano?
    If personaje.faccion.alineacion = eAlineaciones.Real Then
        If (permitidos And eEventoPersonajesAlineacion.Ciudadanos) Then cumpleCondicionPersonaje = True
        Exit Function
    End If

End Function


Private Function iCondicionEvento_puedeIngresarEquipo(integrantesIndex() As Integer) As Boolean
    
    Dim loopIntegrante As Byte

    For loopIntegrante = 1 To UBound(integrantesIndex)
        
        If Not cumpleCondicionPersonaje(UserList(integrantesIndex(loopIntegrante))) Then
            iCondicionEvento_puedeIngresarEquipo = False
            Exit Function
        End If
        
    Next loopIntegrante

'Todos los integrantes tiene las clases permitidas
iCondicionEvento_puedeIngresarEquipo = True
        
        
End Function

Private Sub iCondicionEvento_removerEquipoInscripto(integrantes() As Long)
    ' No lo necesito
End Sub

Private Sub iCondicionEvento_setMaximaMemoria(slots As Byte)
    ' No lo necesito
End Sub
