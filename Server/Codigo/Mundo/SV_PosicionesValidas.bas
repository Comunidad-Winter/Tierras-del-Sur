Attribute VB_Name = "SV_PosicionesValidas"
Option Explicit

Function LegalPos(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByRef personaje As User) As Boolean

'¿Es un mapa valido?
If esPosicionJugable(x, y) = False Then
    LegalPos = False
    Exit Function
End If

If esPosicionUsablePersonaje(personaje, MapData(map, x, y)) = False Then
    LegalPos = False
    Exit Function
End If

If (MapData(map, x, y).UserIndex > 0) Or (MapData(map, x, y).npcIndex > 0) Then
    LegalPos = False
    Exit Function
End If

LegalPos = True
End Function

Public Function esPosicionNavegable(mapa As Integer, x As Byte, y As Byte) As Boolean
    
    If SV_PosicionesValidas.esPosicionJugable(x, y) Then
        esPosicionNavegable = (MapData(mapa, x, y).Trigger And eTriggers.Navegable)
    Else
        esPosicionNavegable = False
    End If
   
End Function

Public Function esPosicionCaminable(mapa As Integer, x As Byte, y As Byte) As Boolean

    If SV_PosicionesValidas.esPosicionJugable(x, y) Then
        esPosicionCaminable = Not CBool(MapData(mapa, x, y).Trigger And eTriggers.NoCaminable)
    Else
        esPosicionCaminable = False
    End If
    
End Function
'*********************************************
Public Function personajePuedeIngresarMapa(Usuario As User, mapa As MapInfo) As Boolean
Dim esCrimi As Boolean

'Mapa para newbies?
If Usuario.flags.Privilegios > 0 Then
    personajePuedeIngresarMapa = True
    Exit Function
End If

If (mapa.restringir = 1 And EsNewbie(Usuario.UserIndex)) Or Not mapa.restringir = 1 Then
    'Tiene el nivel suficiente para entrar al mapa
    If (Usuario.Stats.ELV >= mapa.Nivel And Usuario.Stats.ELV <= mapa.MaxLevel) Then
        'Cantidad de usuarios maximo por mapa
        If mapa.usuarios.getCantidadElementos() < mapa.UsuariosMaximo Then
            'Mapa solo para una faccion=?
            If Not ((mapa.SoloCiudas = 1 And Usuario.faccion.alineacion = eAlineaciones.Real) Or (mapa.SoloCrimis = 1 And Usuario.faccion.alineacion = eAlineaciones.caos)) Then
                If (mapa.SoloCaos = 1 And Usuario.faccion.FuerzasCaos = 1) Or (mapa.SoloArmada = 1 And Usuario.faccion.ArmadaReal = 1) Or (mapa.SoloArmada = 0 And mapa.SoloCaos = 0) Then
                    'Todo ok
                    personajePuedeIngresarMapa = True
                    Exit Function
                Else
                    If mapa.SoloArmada = 1 And mapa.SoloCaos = 1 Then
                        EnviarPaquete Paquetes.mensajeinfo, "Sólo integrantes de la Armada Real o del Caos pueden entrar a este mapa.", Usuario.UserIndex, ToIndex
                    ElseIf mapa.SoloCaos = 1 Then
                        EnviarPaquete Paquetes.mensajeinfo, "Sólo integrantes del Ejercito del Caos pueden entrar a este mapa.", Usuario.UserIndex, ToIndex
                    ElseIf mapa.SoloArmada = 1 Then
                        EnviarPaquete Paquetes.mensajeinfo, "Sólo integrantes de la Armada Real pueden entrar a este mapa.", Usuario.UserIndex, ToIndex
                    End If
                End If
            Else
                If mapa.SoloCiudas = 1 Then
                    EnviarPaquete Paquetes.mensajeinfo, "Sólo ciudadanos pueden entrar a este mapa.", Usuario.UserIndex, ToIndex
                ElseIf mapa.SoloCiudas = 1 Then
                    EnviarPaquete Paquetes.mensajeinfo, "Sólo criminales pueden entrar a este mapa.", Usuario.UserIndex, ToIndex
                End If
            End If
        Else
            EnviarPaquete Paquetes.mensajeinfo, "No puedes ingresar a este mapa. El mapa no tiene más capacidad.", Usuario.UserIndex, ToIndex
        End If
    Else
    EnviarPaquete Paquetes.MensajeSimple2, Chr$(21), Usuario.UserIndex
    End If
Else
    EnviarPaquete Paquetes.MensajeSimple, Chr$(38), Usuario.UserIndex
End If

'Si llegue hasta acá, todo mal
personajePuedeIngresarMapa = False

End Function

'*******************************************
'El personaje puede posicionar en este tile o no
Public Function esPosicionUsablePersonaje(personaje As User, posicionMapa As MapBlock) As Boolean

    If isTileBloqueado(posicionMapa) Then
        esPosicionUsablePersonaje = False
        Exit Function
    End If
    
    If personaje.flags.Privilegios = 0 Then
        If personaje.flags.Navegando = 1 Then 'Esta navegando
            esPosicionUsablePersonaje = (eTriggers.Navegable And posicionMapa.Trigger) > 0
        Else 'Esta caminando
            esPosicionUsablePersonaje = (eTriggers.NoCaminable And posicionMapa.Trigger) = 0
        End If
    Else
        esPosicionUsablePersonaje = True
    End If

End Function

'******************************************************
'Nuevo inmapbounds
'Devuelve true si el mapa existe y la posicion es jugable
Public Function existePosicionMundo(mapa As Integer, ByVal x As Byte, ByVal y As Byte) As Boolean
     If mapa > 0 And mapa <= SV_Mundo.NumMaps Then 'Numero valido?
        If MapInfo(mapa).Existe Then 'Mapa cargado?
            If x >= SV_Constantes.X_MINIMO_JUGABLE And x <= SV_Constantes.X_MAXIMO_JUGABLE Then
                If y >= SV_Constantes.Y_MINIMO_JUGABLE And y <= SV_Constantes.Y_MAXIMO_JUGABLE Then
                    existePosicionMundo = True
                    Exit Function
                End If
            End If
       End If
    End If
End Function
    
'Devuelve True si la posicion es jugable
Public Function esPosicionJugable(ByVal x As Integer, ByVal y As Integer) As Boolean

    If x >= SV_Constantes.X_MINIMO_JUGABLE And x <= SV_Constantes.X_MAXIMO_JUGABLE Then
            If y >= SV_Constantes.Y_MINIMO_JUGABLE And y <= SV_Constantes.Y_MAXIMO_JUGABLE Then
                esPosicionJugable = True
                Exit Function
            End If
    End If

    esPosicionJugable = False

End Function

Public Function existeMapa(mapa As Integer) As Boolean

    If mapa > 0 And mapa <= SV_Mundo.NumMaps Then 'Numero valido?
        If MapInfo(mapa).Existe Then
            existeMapa = True
            Exit Function
        End If
    End If
    
    existeMapa = False
End Function

Function esPosicionUsableNPC(infoPos As MapBlock, criatura As npc) As Boolean

'Mapa y coordenadas correctas
If infoPos.UserIndex > 0 Or infoPos.npcIndex > 0 Then
    esPosicionUsableNPC = False
    Exit Function
End If

If isTileBloqueado(infoPos) Then
    esPosicionUsableNPC = False
    Exit Function
End If

If Not CBool(infoPos.Trigger And eTriggers.PosicionInvalidaNpc) Then
    If criatura.flags.Terreno = eTerrenoNPC.AguayTierra Then 'Todo terreno
        esPosicionUsableNPC = True
    ElseIf criatura.flags.Terreno = eTerrenoNPC.Tierra Then  'Solo pude estar tierra firme
        esPosicionUsableNPC = Not CBool(infoPos.Trigger And eTriggers.NoCaminable)
    ElseIf criatura.flags.Terreno = eTerrenoNPC.Agua Then  'Solo puede estar en el agua
        esPosicionUsableNPC = infoPos.Trigger And eTriggers.Navegable
    End If
End If


End Function

