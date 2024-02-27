Attribute VB_Name = "modSumoneoAutomatico"
'Option Explicit
'
'Private parseoActual() As tipoParserEquipo
'Private descansos() As tZonaDescanso
'Private sumoneandoActualmente As Boolean
'
'
'Public Function reintentarSumonear(nombresNoSumoneados As String) As Integer
'    Dim loopEquipo As Byte
'    Dim loopIntegrante As Byte
'    Dim UserINdex As Integer
'    Dim nPos As WorldPos
'
'    reintentarSumonear = 0
'    nombresNoSumoneados = ""
'    For loopEquipo = 1 To UBound(parseoActual)
'        With parseoActual(loopEquipo)
'            For loopIntegrante = 1 To UBound(.intgerantesNick)
'                If .auxiliarIndividual(loopIntegrante) = False Then
'                    UserINdex = NameIndex(.intgerantesNick(loopIntegrante))
'                    If UserINdex > 0 Then
'                        Call ClosestLegalPos(descansos(loopEquipo).centro, nPos)
'                        Call WarpUserChar(UserINdex, nPos.Map, nPos.x, nPos.y, True)
'                        .auxiliarIndividual(loopIntegrante) = True
'                    Else
'                        reintentarSumonear = reintentarSumonear + 1
'                        nombresNoSumoneados = nombresNoSumoneados & ", " & .intgerantesNick(loopIntegrante)
'                    End If
'                End If
'            Next loopIntegrante
'        End With
'    Next loopEquipo
'
'
'End Function
'
'Public Function sumonearParseados(parseoActual_() As tipoParserEquipo, listaOffline As String) As Integer
'    Dim cantidadDescansosNecesarios As Byte
'    Dim loopDescanso As Byte
'    Dim loopEquipo As Byte
'    Dim loopIntegrante As Byte
'
'    Call resetSumonedos
'
'    'Tomo el nuevo
'    parseoActual = parseoActual_
'    sumoneandoActualmente = True
'
'    'Pido los descansos
'    cantidadDescansosNecesarios = UBound(parseoActual_)
'
'    ReDim descansos(1 To cantidadDescansosNecesarios)
'
'    For loopDescanso = 1 To cantidadDescansosNecesarios
'        descansos(loopDescanso) = modDescansos.getZonaDescanso(UBound(parseoActual(loopEquipo).integrantesIndex))
'    Next
'
'    'Reseteo el estado
'    For loopEquipo = 1 To UBound(parseoActual)
'        With parseoActual(loopEquipo)
'            .auxiliar = False
'            For loopIntegrante = 1 To UBound(.intgerantesNick)
'                .auxiliarIndividual(loopIntegrante) = False
'            Next loopIntegrante
'        End With
'    Next loopEquipo
'
'    'Intento de sumonear
'    sumonearParseados = reintentarSumonear(listaOffline)
'End Function
'
'Public Sub resetSumonedos()
'    Dim loopDescanso As Byte
'
'    If sumoneandoActualmente Then
'        sumoneandoActualmente = False
'
'        For loopDescanso = 1 To UBound(parseoActual)
'            Call modDescansos.liberarZonaDescanso(descansos(loopDescanso))
'        Next loopDescanso
'    End If
'End Sub
