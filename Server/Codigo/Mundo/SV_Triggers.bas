Attribute VB_Name = "SV_Triggers"
Option Explicit

Public Enum eTriggers

    NoCaminable = 1
    
    BloqueoEste = 2
    BloqueoOeste = 4
    BloqueoNorte = 8
    BloqueoSur = 16
    
    Navegable = 32
    BajoTecho = 64
    
    AntiRespawnNpc = 128
    PosicionInvalidaNpc = 256
    PosicionSegura = 512
    
    AntiPiquete = 1024
    CombateSeguro = 2048
    RevivirAutomatico = 4096
    NoDragAndDrop = 8192
    NoTirarItem = 16384
    
    
    'Triggers en funcion de otros
    TodosBordesBloqueados = eTriggers.BloqueoEste Or eTriggers.BloqueoOeste Or eTriggers.BloqueoNorte Or eTriggers.BloqueoSur
    
    ' Esto es para que siempre se trate como un long (!)
    e_triggers_force_dword = &H7FFFFFFF
End Enum

Public Sub BloquearTile(ByVal mapa As Integer, ByVal x As Byte, ByVal y As Byte, Optional ByVal HardBlock As Boolean = True)
    '*****************************************************************
    'Menduz
    '*****************************************************************
    MapData(x, y).Trigger = MapData(x, y).Trigger Or eTriggers.TodosBordesBloqueados
        
    If HardBlock Then
        Call BloquearLinea(mapa, x - 1, y, EAST, False)
        Call BloquearLinea(mapa, x + 1, y, WEST, False)
        Call BloquearLinea(mapa, x, y - 1, SOUTH, False)
        Call BloquearLinea(mapa, x, y + 1, NORTH, False)
    End If
End Sub

Public Sub DesBloquearTile(ByVal mapa As Integer, ByVal x As Byte, ByVal y As Byte)
    If SV_PosicionesValidas.esPosicionJugable(x, y) Then
        MapData(mapa, x, y).Trigger = MapData(mapa, x, y).Trigger And Not eTriggers.TodosBordesBloqueados
    End If
End Sub

Public Sub BloquearLinea(ByVal mapa As Integer, ByVal x As Byte, ByVal y As Byte, Optional ByVal heading As eHeading = eHeading.Ninguno, Optional ByVal HagoSegundaLlamada As Boolean = True)
    '*****************************************************************
    'Menduz
    '*****************************************************************

    Dim tTrigger As Long
        
    
    If SV_PosicionesValidas.existePosicionMundo(mapa, x, y) Then
    
        If heading = eHeading.Ninguno Then
            Call BloquearTile(mapa, x, y)
        Else
            tTrigger = MapData(x, y).Trigger
            
            Select Case heading
                Case eHeading.EAST
                    tTrigger = (tTrigger Or eTriggers.BloqueoEste)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then _
                        Call BloquearLinea(mapa, x + 1, y, WEST, False)
                        
                Case eHeading.WEST
                    tTrigger = (tTrigger Or eTriggers.BloqueoOeste)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then _
                        Call BloquearLinea(mapa, x - 1, y, EAST, False)
                        
                Case eHeading.NORTH
                    tTrigger = (tTrigger Or eTriggers.BloqueoNorte)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then _
                        Call BloquearLinea(mapa, x, y - 1, SOUTH, False)
                        
                Case eHeading.SOUTH
                    tTrigger = (tTrigger Or eTriggers.BloqueoSur)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then _
                        Call BloquearLinea(mapa, x, y + 1, NORTH, False)
                        
            End Select
            
            MapData(mapa, x, y).Trigger = tTrigger
        End If
    End If
End Sub

Public Sub DesBloquearLinea(ByVal mapa As Integer, ByVal x As Byte, ByVal y As Byte, Optional ByVal heading As eHeading = eHeading.Ninguno, Optional ByVal HagoSegundaLlamada As Boolean = True)
    '*****************************************************************
    'Menduz
    '*****************************************************************

    Dim tTrigger As Long
    
    If SV_PosicionesValidas.existePosicionMundo(mapa, x, y) Then
        If heading = eHeading.Ninguno Then
            Call DesBloquearTile(mapa, x, y)
        Else
            tTrigger = MapData(mapa, x, y).Trigger

            Select Case heading
                Case eHeading.EAST
                    tTrigger = (tTrigger And Not eTriggers.BloqueoEste)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then _
                        Call DesBloquearLinea(mapa, x + 1, y, WEST, False)
                        
                Case eHeading.WEST
                    tTrigger = (tTrigger And Not eTriggers.BloqueoOeste)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then _
                        Call DesBloquearLinea(mapa, x - 1, y, EAST, False)
                        
                Case eHeading.NORTH
                    tTrigger = (tTrigger And Not eTriggers.BloqueoNorte)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then _
                        Call DesBloquearLinea(mapa, x, y - 1, SOUTH, False)
                        
                Case eHeading.SOUTH
                    tTrigger = (tTrigger And Not eTriggers.BloqueoSur)
                    
                    'Bloqueo el opuesto
                    If HagoSegundaLlamada Then _
                        Call DesBloquearLinea(mapa, x, y + 1, NORTH, False)
                        
            End Select
            
            
            MapData(mapa, x, y).Trigger = tTrigger
        End If
    End If
End Sub



Public Function EstaBloqueado(ByVal mapa As Integer, ByVal x As Byte, ByVal y As Byte, Optional ByVal heading As eHeading = eHeading.Ninguno) As Boolean
    Dim tTrigger As Long
    If SV_PosicionesValidas.existePosicionMundo(mapa, x, y) Then
        tTrigger = MapData(mapa, x, y).Trigger
        Select Case heading
            Case eHeading.EAST
                EstaBloqueado = CBool(tTrigger And eTriggers.BloqueoEste)
            Case eHeading.WEST
                EstaBloqueado = CBool(tTrigger And eTriggers.BloqueoOeste)
            Case eHeading.NORTH
                EstaBloqueado = CBool(tTrigger And eTriggers.BloqueoNorte)
            Case eHeading.SOUTH
                EstaBloqueado = CBool(tTrigger And eTriggers.BloqueoSur)
            Case eHeading.Ninguno
                EstaBloqueado = (tTrigger And eTriggers.TodosBordesBloqueados) = eTriggers.TodosBordesBloqueados
        End Select
    End If
End Function


Public Sub ToggleBloqueo(ByVal mapa As Integer, ByVal x As Byte, ByVal y As Byte, Optional ByVal heading As eHeading = eHeading.Ninguno)
    If EstaBloqueado(mapa, x, y, heading) Then
        DesBloquearLinea mapa, x, y, heading
    Else
        BloquearLinea mapa, x, y, heading
    End If
End Sub


