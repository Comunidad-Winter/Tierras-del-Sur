VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CondicionEventoGrupoClases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iCondicionEvento

Private Const CANTIDAD_GRUPOS As Byte = 4

Private Enum eGrupoClases
    indefinido = 0
    magica = 1
    semiMagica = 2
    noMagica = 3
    trabajadora = 4
End Enum

Private claseDeGrupo(1 To CANTIDAD_CLASES) As Byte
Private cantidadMinima(1 To CANTIDAD_GRUPOS) As Byte
Private tempContador(1 To CANTIDAD_GRUPOS) As Byte


'Trabajadoras
'    Ladron = 5
'    Pescador = 10
'    Herrero = 11
'    Leñador = 12
'    Minero = 13
'    Carpintero = 14
'    Pirata = 15
'
'Magica
'    Mago = 1
'        Clerigo = 2
'        Bardo = 6
'        Druida = 7
'
'SemiMagica
'        Paladin = 8
'        Asesino = 4
'
'NoMagica
'        Cazador = 9
'        Guerrero = 3


'tipoGrupoClases: el grupo de clases a la cual aplica la regla
' 1 Magicas
' 2 Semi Magicas
' 3 No magicas
' 4 Trabajadora

' 0. Sin limite
' 1 o más: Cantidad mínima.
Public Sub setParametros(ByVal grupo As Integer, ByVal cantidadMinima_ As Byte)
    cantidadMinima(grupo) = cantidadMinima_
End Sub

Private Sub Class_Initialize()
    claseDeGrupo(eClases.asesino) = eGrupoClases.semiMagica
    claseDeGrupo(eClases.Bardo) = eGrupoClases.magica
    claseDeGrupo(eClases.Carpintero) = eGrupoClases.trabajadora
    claseDeGrupo(eClases.Cazador) = eGrupoClases.noMagica
    claseDeGrupo(eClases.Clerigo) = eGrupoClases.magica
    claseDeGrupo(eClases.Druida) = eGrupoClases.magica
    claseDeGrupo(eClases.Guerrero) = eGrupoClases.noMagica
    claseDeGrupo(eClases.Herrero) = eGrupoClases.trabajadora
    claseDeGrupo(eClases.Ladron) = eGrupoClases.trabajadora
    claseDeGrupo(eClases.Leñador) = eGrupoClases.trabajadora
    claseDeGrupo(eClases.Mago) = eGrupoClases.magica
    claseDeGrupo(eClases.Minero) = eGrupoClases.trabajadora
    claseDeGrupo(eClases.Paladin) = eGrupoClases.semiMagica
    claseDeGrupo(eClases.Pescador) = eGrupoClases.trabajadora
    claseDeGrupo(eClases.Pirata) = eGrupoClases.trabajadora
    
    cantidadMinima(eGrupoClases.magica) = 0 ' Sin limite
    cantidadMinima(eGrupoClases.noMagica) = 0 ' Sin limite
    cantidadMinima(eGrupoClases.semiMagica) = 0 ' Sin limite
    cantidadMinima(eGrupoClases.trabajadora) = 0 'Sin limite
End Sub

Private Function obtenerGrupoDeClase(clase As eClases) As Byte
    If LBound(claseDeGrupo) <= clase And UBound(claseDeGrupo) >= clase Then
        obtenerGrupoDeClase = claseDeGrupo(clase)
    Else
        obtenerGrupoDeClase = eGrupoClases.indefinido
    End If
    
End Function
Private Sub iCondicionEvento_asentarIngresoEquipo(integrantesIndex() As Integer)
End Sub

Private Function iCondicionEvento_obtenerDescripcion() As String
    Dim loopGrupo As Byte
    
    iCondicionEvento_obtenerDescripcion = "Cada equipo debe tener al menos"
    
    For loopGrupo = 1 To CANTIDAD_GRUPOS
        If cantidadMinima(loopGrupo) > 0 Then
            If loopGrupo = eGrupoClases.magica Then
                If cantidadMinima(loopGrupo) = 1 Then
                    iCondicionEvento_obtenerDescripcion = iCondicionEvento_obtenerDescripcion & " una clase magica"
                Else
                    iCondicionEvento_obtenerDescripcion = iCondicionEvento_obtenerDescripcion & " " & cantidadMinima(loopGrupo) & " clases magicas"
                End If
            ElseIf loopGrupo = eGrupoClases.semiMagica Then
                If cantidadMinima(loopGrupo) = 1 Then
                    iCondicionEvento_obtenerDescripcion = iCondicionEvento_obtenerDescripcion & " una clase semi magica."
                Else
                    iCondicionEvento_obtenerDescripcion = iCondicionEvento_obtenerDescripcion & " " & cantidadMinima(loopGrupo) & " clases semi magicas"
                End If
            ElseIf loopGrupo = eGrupoClases.noMagica Then
                If cantidadMinima(loopGrupo) = 1 Then
                    iCondicionEvento_obtenerDescripcion = iCondicionEvento_obtenerDescripcion & " una clase no magica."
                Else
                    iCondicionEvento_obtenerDescripcion = iCondicionEvento_obtenerDescripcion & " " & cantidadMinima(loopGrupo) & " clases no magicas"
                End If
            ElseIf loopGrupo = eGrupoClases.trabajadora Then
                If cantidadMinima(loopGrupo) = 1 Then
                    iCondicionEvento_obtenerDescripcion = iCondicionEvento_obtenerDescripcion & " una clase trabajadora."
                Else
                    iCondicionEvento_obtenerDescripcion = iCondicionEvento_obtenerDescripcion & " " & cantidadMinima(loopGrupo) & " clases trabajadoras"
                End If
            End If
        End If
    Next
End Function

Private Function iCondicionEvento_puedeIngresarEquipo(integrantesIndex() As Integer) As Boolean
    Dim loopGrupo As Byte
    Dim loopC As Byte
    Dim UserIndex As Integer
    Dim grupoClase As eGrupoClases
    
    ' Reseteo el contador
    For loopGrupo = 1 To CANTIDAD_GRUPOS
        tempContador(loopGrupo) = 0
    Next
    
    ' Cuanto a los personajes
    For loopC = 1 To UBound(integrantesIndex)
        UserIndex = integrantesIndex(loopC)
        
        grupoClase = obtenerGrupoDeClase(UserList(UserIndex).clase)
        
        If Not grupoClase = eGrupoClases.indefinido Then
            tempContador(grupoClase) = tempContador(grupoClase) + 1
        Else
            Call LogError("Clase indefinida. Usuario " & UserList(UserIndex).Name & "en eventos.")
        End If
    Next loopC
    
    'Chequeo que se cumnplan las condiciones
    For loopGrupo = 1 To CANTIDAD_GRUPOS
        If cantidadMinima(loopGrupo) > 0 Then 'No es ilimitado
            If tempContador(loopGrupo) < cantidadMinima(loopGrupo) Then
                iCondicionEvento_puedeIngresarEquipo = False
                Exit Function
            End If
        End If
    Next
    
    iCondicionEvento_puedeIngresarEquipo = True
End Function

Private Sub iCondicionEvento_removerEquipoInscripto(integrantes() As Long)
End Sub

Private Sub iCondicionEvento_setMaximaMemoria(slots As Byte)
End Sub
