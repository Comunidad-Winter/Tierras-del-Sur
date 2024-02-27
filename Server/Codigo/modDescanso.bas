Attribute VB_Name = "modDescansos"
Option Explicit

Public Type tZonaDescanso
    id As Integer 'Identificador del descanso
    capacidad As Byte 'Capacidad de personas del repositorio
    centro As WorldPos
    superiorIzquierda As Position ' Estos dos estan como dato para saber desde donde
    inferiorDerecha As Position '   hasta donde tengo que borrar. No se me ocurrio nada mejor.
    tipo As eDescansoTipo
End Type

Private Type tStock
    zona As tZonaDescanso
    disponible As Boolean
End Type

Public Enum eDescansoTipo
    torneo = 1
    reto = 2
    conBoveda = 4
End Enum

Private zonas() As tStock


Private Sub cargarZonasDescanso()

Dim zona As tZonaDescanso
Dim ruta As String
Dim loopDescanso As Integer
Dim comienzoInfoCoordenada As Integer
Dim infoDescanso As String
Dim cantidadDescansos As Integer

ruta = App.Path & "\Dat\Descansos.dat"
    
cantidadDescansos = val(GetVar(ruta, "INIT", "NumDescansos"))

ReDim Preserve zonas(1 To cantidadDescansos) As tStock

For loopDescanso = 1 To cantidadDescansos

    On Error GoTo hayError
    ' Generales
    zona.id = loopDescanso
    zona.capacidad = val(GetVar(ruta, "DESCANSO" & loopDescanso, "capacidad"))
    zona.tipo = val(GetVar(ruta, "DESCANSO" & loopDescanso, "tipo"))
    
    'Cargo el centro del mapa
    zona.centro.map = val(GetVar(ruta, "DESCANSO" & loopDescanso, "mapa"))
        
    infoDescanso = Trim(GetVar(ruta, "DESCANSO" & loopDescanso, "centro"))
        
    zona.centro.x = obtenerCoordenadas(infoDescanso)(0)
    zona.centro.y = obtenerCoordenadas(infoDescanso)(1)

    'Cargo el vertice superir izquierda
    infoDescanso = Trim(GetVar(ruta, "DESCANSO" & loopDescanso, "superiorIzquierda"))
        
    zona.superiorIzquierda.x = obtenerCoordenadas(infoDescanso)(0)
    zona.superiorIzquierda.y = obtenerCoordenadas(infoDescanso)(1)
    
    'Cargo el vertice inferior derecho
    infoDescanso = Trim(GetVar(ruta, "DESCANSO" & loopDescanso, "inferiorDerecha"))

    zona.inferiorDerecha.x = obtenerCoordenadas(infoDescanso)(0)
    zona.inferiorDerecha.y = obtenerCoordenadas(infoDescanso)(1)
    
    'Cargo definitivamente la zona
    zonas(loopDescanso).zona = zona
    GoTo sinError
hayError:
    Call LogError("Fallo la carga del descanso número " & loopDescanso & ".")
    
sinError:
Next loopDescanso

End Sub

Public Sub iniciarZonasDescanso()

    Dim cantidadZonasDescanso As Integer
    Dim loopZona As Integer
    
    Call cargarZonasDescanso
    
    cantidadZonasDescanso = UBound(zonas)
    
    'Marco a todas como disponibles
    For loopZona = 1 To cantidadZonasDescanso
        
        zonas(loopZona).disponible = True
    
    Next loopZona
    
End Sub

Public Sub reCargarZonasDescanso()
    'La re-carga no pone a los descansos como disponibles
    'Esto es para no hacer caos en caso de que haya descansos que se estan utilizado
    'Si marco a todos como liberados puede que otro evento pueda pedir un ring
    'y que se le de uno que se esta utilizando.
    Dim cantidadActual As Integer
    Dim loopZona As Integer
    
    'Obtengo la cantidad de zonas que hay antes de la recarga
    cantidadActual = UBound(zonas)
    
    'Las cargo
    Call cargarZonasDescanso
    
    'Se agregaron nuevos descansos?
    If cantidadActual < UBound(zonas) Then
        'Las marco como disponibles
        For loopZona = cantidadActual + 1 To UBound(zonas)
            zonas(loopZona).disponible = True
        Next
    End If
    
End Sub

'Devuelve un descanso que cumpla con la necesidad de capacidad de personas
Public Function getZonaDescanso(capacidad As Byte, tipo As eDescansoTipo) As tZonaDescanso

Dim loopC As Byte

loopC = 1

Do While loopC <= UBound(zonas)

    ' ¿Esta disponible?
    If zonas(loopC).disponible Then
        ' ¿Esta zona es del tipo que necesito?
        If zonas(loopC).zona.tipo = tipo Then
            'Esta zona tiene una capacidad igual o mayor a la que necesito?
            If zonas(loopC).zona.capacidad >= capacidad Then
                'La guardo
                getZonaDescanso = zonas(loopC).zona
                'La marco como utilizada
                zonas(loopC).disponible = False
                Exit Function
            End If
        End If
    End If

    'A ver la siguiente..
    loopC = loopC + 1
Loop

'Si llegue hasta acá es porque no hay una zona
getZonaDescanso.id = 0

End Function


'Libera la zona para que otro evento lo pueda utilizar
Public Sub liberarZonaDescanso(ByRef zona As tZonaDescanso)
    Dim liberada As Boolean
    Dim loopC As Byte
    
    Call limpiarZonaDescanso(zona)
    
    'La libero
    liberada = False
    
    loopC = 1
    Do While Not liberada
        If zonas(loopC).zona.id = zona.id Then
            zonas(loopC).disponible = True
            liberada = True
        Else
            loopC = loopC + 1
        End If
    Loop

    Call LogNuevosRetos("Libere el descanso " & loopC)

End Sub

'Borra los items que estan en la zona
Private Sub limpiarZonaDescanso(ByRef zona As tZonaDescanso)
    Dim mapa As Integer
    
    Dim y As Byte
    Dim x As Byte
    
    mapa = zona.centro.map
    
    For y = zona.superiorIzquierda.y To zona.inferiorDerecha.y
        For x = zona.superiorIzquierda.x To zona.inferiorDerecha.x
            If MapData(mapa, x, y).OBJInfo.ObjIndex > 0 Then
                If ItemNoEsDeMapa(MapData(mapa, x, y).OBJInfo.ObjIndex) Then
                    Call EraseObj(ToMap, 0, mapa, 10000, mapa, x, y)
                End If
            End If
        Next x
    Next y
    
End Sub

Public Sub verEstado(ByRef lista As VB.ListBox)
    Dim loopZona As Integer
    Dim cantidadDisponibles As Integer
    Dim linea1 As String
    Dim linea2 As String
    
    cantidadDisponibles = 0
    
    For loopZona = 1 To UBound(zonas)
        If zonas(loopZona).disponible = True Then
            linea1 = loopZona & ":" & " Disponible."
            linea2 = getDescTipoDescanso(zonas(loopZona).zona.tipo) & ". Map " & zonas(loopZona).zona.centro.map & "(" & zonas(loopZona).zona.centro.x & "," & zonas(loopZona).zona.centro.y & "). Capacidad: " & zonas(loopZona).zona.capacidad
            Call lista.AddItem(linea1)
            Call lista.AddItem(linea2)
            Call LogDesarrollo(linea1)
            Call LogDesarrollo(linea2)
            cantidadDisponibles = cantidadDisponibles + 1
        Else
            Call lista.AddItem(loopZona & ":" & " NO disponible")
        End If
    Next loopZona
    
    Call lista.AddItem("Disponibles: " & cantidadDisponibles)
End Sub

Public Function getDescTipoDescanso(tipoDescanso As eDescansoTipo) As String

    Dim temp As String
    
    temp = ""
    
    If (tipoDescanso And torneo) Then temp = temp & " torneo"
    
    If (tipoDescanso And reto) Then temp = temp & " reto"
    
    If (tipoDescanso And conBoveda) Then temp = temp & " con bóveda"

    getDescTipoDescanso = temp

End Function

