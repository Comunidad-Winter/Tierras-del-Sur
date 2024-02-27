Attribute VB_Name = "modTriggers"

Option Explicit

Public Enum eTriggers

    NoCaminable = 1
    
    BloqueoEste = 2
    BloqueoOeste = 4
    BloqueoNorte = 8
    BloqueoSur = 16
    
    Navegable = 32
    BajoTecho = 64
    
    AntiRespawnNpc = &H80
    PosicionInvalidaNpc = &H100
    PosicionSegura = &H200
    
    AntiPiquete = &H400
    CombateSeguro = &H800
    RevivirAutomatico = &H1000
    NoDragAndDrop = &H2000
    NoTirarItem = &H4000 'Lo cambié a heza asi queda más elegante...
    Transparentar = &H800 '¿Se debe transpoarentar el objeto en la capa 3?
    
    Puente = &H1600 '¿Se debe transpoarentar el objeto en la capa 3?
    
    'Triggers en funcion de otros
    TodosBordesBloqueados = eTriggers.BloqueoEste Or eTriggers.BloqueoOeste Or eTriggers.BloqueoNorte Or eTriggers.BloqueoSur
    
    ' Esto es para que siempre se trate como un long (!)
    e_triggers_force_dword = &H7FFFFFFF
End Enum


Public Function PuedoCaminar(ByVal x As Byte, ByVal y As Byte, Optional ByVal heading As E_Heading = E_Heading.None, Optional ByVal EstaNavegando As Boolean = False, Optional ByVal EstaMuerto As Boolean = False) As Boolean
    '*****************************************************************
    'Menduz
    '*****************************************************************

    Dim tTrigger As Long, nuevoX As Byte, nuevoY As Byte
    Dim bloqueoInteresado As eTriggers
    
    'Me fijo la viabilidad del tile HACIA donde voy.
    'Sino me queda a la mitad, algunas cosas chequeandolas donde estoy y otras hacia donde voy
    'Y la idea es que mantenga una homogeneanidad logica para que sea mas facil de pensar
    
    nuevoX = x
    nuevoY = y
        
    Select Case heading
        
        Case E_Heading.EAST
            
            nuevoX = x + 1
            bloqueoInteresado = eTriggers.BloqueoOeste
                
        Case E_Heading.WEST
            
            nuevoX = x - 1
            bloqueoInteresado = eTriggers.BloqueoEste
                
        Case E_Heading.NORTH
                
            nuevoY = y - 1
            bloqueoInteresado = eTriggers.BloqueoSur
                
        Case E_Heading.SOUTH
                
            nuevoY = y + 1
            bloqueoInteresado = eTriggers.BloqueoNorte
  
        Case E_Heading.None
            
            PuedoCaminar = True
            
    End Select
    
    '¿En esta posicion el usuario puede jugar?
    If CLI_PosicionesLegales.esPosicionJugable(nuevoX, nuevoY) Then
    
        tTrigger = mapdata(nuevoX, nuevoY).trigger 'Tomo el triger
        
        'La parte que me intersa de tile no esta bloqueada
        If Not CBool(tTrigger And bloqueoInteresado) Then
            
            If EstaNavegando Then
                PuedoCaminar = CBool(mapdata(nuevoX, nuevoY).trigger And eTriggers.Navegable)
            Else
                PuedoCaminar = Not CBool(mapdata(nuevoX, nuevoY).trigger And eTriggers.NoCaminable)
            End If
    
            Exit Function
        End If
    End If

    PuedoCaminar = False
End Function

Public Sub BloquearTile(ByVal x As Byte, ByVal y As Byte, Optional ByVal HardBlock As Boolean = True)
    '*****************************************************************
    'Menduz
    '*****************************************************************
    mapdata(x, y).trigger = mapdata(x, y).trigger Or eTriggers.TodosBordesBloqueados
        
'    If HardBlock Then
'        Call BloquearLinea(X - 1, Y, EAST, False)
'        Call BloquearLinea(X + 1, Y, WEST, False)
'        Call BloquearLinea(X, Y - 1, SOUTH, False)
'        Call BloquearLinea(X, Y + 1, NORTH, False)
'    End If
End Sub

Public Sub DesBloquearTile(ByVal x As Byte, ByVal y As Byte)
    '*****************************************************************
    'Menduz
    '*****************************************************************
    
    If InMapBounds(x, y) Then
        mapdata(x, y).trigger = mapdata(x, y).trigger And Not eTriggers.TodosBordesBloqueados
    End If
End Sub

Public Sub BloquearLinea(ByVal x As Byte, ByVal y As Byte, Optional ByVal heading As E_Heading = E_Heading.None, Optional ByVal HagoSegundaLlamada As Boolean = True)
    '*****************************************************************
    'Menduz
    '*****************************************************************

    Dim tTrigger As Long
        
    
    If InMapBounds(x, y) Then
    
        If heading = E_Heading.None Then
            Call BloquearTile(x, y)
        Else
            tTrigger = mapdata(x, y).trigger
            
            Select Case heading
                Case E_Heading.EAST
                    tTrigger = (tTrigger Or eTriggers.BloqueoEste)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then _
                        Call BloquearLinea(x + 1, y, WEST, False)
                        
                Case E_Heading.WEST
                    tTrigger = (tTrigger Or eTriggers.BloqueoOeste)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then _
                        Call BloquearLinea(x - 1, y, EAST, False)
                        
                Case E_Heading.NORTH
                    tTrigger = (tTrigger Or eTriggers.BloqueoNorte)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then _
                        Call BloquearLinea(x, y - 1, SOUTH, False)
                        
                Case E_Heading.SOUTH
                    tTrigger = (tTrigger Or eTriggers.BloqueoSur)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then _
                        Call BloquearLinea(x, y + 1, NORTH, False)
                        
            End Select
            
            mapdata(x, y).trigger = tTrigger
        End If
    End If
End Sub

Public Sub DesBloquearLinea(ByVal x As Byte, ByVal y As Byte, Optional ByVal heading As E_Heading = E_Heading.None, Optional ByVal HagoSegundaLlamada As Boolean = True)
    '*****************************************************************
    'Menduz
    '*****************************************************************

    Dim tTrigger As Long
    
    If InMapBounds(x, y) Then
        If heading = E_Heading.None Then
            Call DesBloquearTile(x, y)
        Else
            tTrigger = mapdata(x, y).trigger

            Select Case heading
                Case E_Heading.EAST
                    tTrigger = (tTrigger And Not eTriggers.BloqueoEste)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then _
                        Call DesBloquearLinea(x + 1, y, WEST, False)
                        
                Case E_Heading.WEST
                    tTrigger = (tTrigger And Not eTriggers.BloqueoOeste)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then _
                        Call DesBloquearLinea(x - 1, y, EAST, False)
                        
                Case E_Heading.NORTH
                    tTrigger = (tTrigger And Not eTriggers.BloqueoNorte)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then _
                        Call DesBloquearLinea(x, y - 1, SOUTH, False)
                        
                Case E_Heading.SOUTH
                    tTrigger = (tTrigger And Not eTriggers.BloqueoSur)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then _
                        Call DesBloquearLinea(x, y + 1, NORTH, False)
                        
            End Select
            
            
            mapdata(x, y).trigger = tTrigger
        End If
    End If
End Sub



Public Function EstaBloqueado(ByVal x As Byte, ByVal y As Byte, Optional ByVal heading As E_Heading = E_Heading.None) As Boolean
    Dim tTrigger As Long
    If InMapBounds(x, y) Then
        tTrigger = mapdata(x, y).trigger
        Select Case heading
            Case E_Heading.EAST
                EstaBloqueado = CBool(tTrigger And eTriggers.BloqueoEste)
            Case E_Heading.WEST
                EstaBloqueado = CBool(tTrigger And eTriggers.BloqueoOeste)
            Case E_Heading.NORTH
                EstaBloqueado = CBool(tTrigger And eTriggers.BloqueoNorte)
            Case E_Heading.SOUTH
                EstaBloqueado = CBool(tTrigger And eTriggers.BloqueoSur)
            Case E_Heading.None
                EstaBloqueado = (tTrigger And eTriggers.TodosBordesBloqueados) = eTriggers.TodosBordesBloqueados
        End Select
    End If
End Function


Public Sub ToggleBloqueo(ByVal x As Byte, ByVal y As Byte, Optional ByVal heading As E_Heading = E_Heading.None)
    If EstaBloqueado(x, y, heading) Then
        DesBloquearLinea x, y, heading
    Else
        BloquearLinea x, y, heading
    End If
End Sub

