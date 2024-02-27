Attribute VB_Name = "UsUaRiOs"
Option Explicit

Private Const MAX_NIVEL_VIDA_PROMEDIO = 16


'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios

'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Sub UsuarioMataAUsuario(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)

Dim DaExp As Integer

DaExp = UserList(VictimIndex).Stats.ELV * 2

Call modUsuarios.agregarExperiencia(AttackerIndex, DaExp)

'Lo mata
EnviarPaquete Paquetes.MensajeFight, "Has matado " & UserList(VictimIndex).Name & "!", AttackerIndex
EnviarPaquete Paquetes.MensajeCompuesto, Chr$(3) & DaExp, AttackerIndex
EnviarPaquete Paquetes.MensajeFight, UserList(AttackerIndex).Name & " te ha matado!", VictimIndex

'If EsPosicionParaAtacarSinPenalidad(UserList(VictimIndex).pos) And EsPosicionParaAtacarSinPenalidad(UserList(AttackerIndex).pos) Then
'End If

Call UserDie(VictimIndex, False, AttackerIndex)

Call AddtoVar(UserList(AttackerIndex).Stats.UsuariosMatados, 1, 31000)

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChangeUserChar
' DateTime  : 18/02/2007 19:51
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub ChangeUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal UserIndex As Integer, _
ByVal Body As Integer, ByVal Head As Integer, ByVal heading As Byte, _
ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

UserList(UserIndex).Char.Body = Body
UserList(UserIndex).Char.Head = Head
UserList(UserIndex).Char.heading = heading
UserList(UserIndex).Char.WeaponAnim = Arma
UserList(UserIndex).Char.ShieldAnim = Escudo
UserList(UserIndex).Char.CascoAnim = Casco
EnviarPaquete Paquetes.pChangeUserChar, ITS(UserList(UserIndex).Char.charIndex) & ITS(Body) & ITS(Head) & ByteToString(heading) & ByteToString(Arma) & ByteToString(Escudo) & ByteToString(UserList(UserIndex).Char.FX) & ITS(UserList(UserIndex).Char.loops) & ByteToString(Casco), UserIndex, ToArea

End Sub

'---------------------------------------------------------------------------------------
' Procedure : EnviarSubirNivel
' DateTime  : 18/02/2007 19:52
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub EnviarSubirNivel(ByVal UserIndex As Integer, ByVal Puntos As Integer)
EnviarPaquete leveLUp, ITS(Puntos), UserIndex
End Sub

'---------------------------------------------------------------------------------------
' Procedure : EnviarSkills
' DateTime  : 18/02/2007 19:52
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub EnviarSkills(ByVal UserIndex As Integer)
Dim i As Integer
Dim cad$

For i = 1 To NUMSKILLS
   cad$ = cad$ & ByteToString(UserList(UserIndex).Stats.UserSkills(i))
Next

EnviarPaquete SendSkills, cad$, UserIndex

End Sub

'---------------------------------------------------------------------------------------
' Procedure : EnviarFama
' DateTime  : 18/02/2007 19:52
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub EnviarFama(ByVal UserIndex As Integer)
Dim cad$

cad$ = LongToString(UserList(UserIndex).Reputacion.AsesinoRep)
cad$ = cad$ & LongToString(UserList(UserIndex).Reputacion.BandidoRep)
cad$ = cad$ & LongToString(UserList(UserIndex).Reputacion.BurguesRep)
cad$ = cad$ & LongToString(UserList(UserIndex).Reputacion.LadronesRep)
cad$ = cad$ & LongToString(UserList(UserIndex).Reputacion.NobleRep)
cad$ = cad$ & LongToString(UserList(UserIndex).Reputacion.PlebeRep)
cad$ = cad$ & LongToString(UserList(UserIndex).Reputacion.PlebeRep)

EnviarPaquete Paquetes.SendFama, cad$, UserIndex

End Sub

'---------------------------------------------------------------------------------------
' Procedure : EnviarAtrib
' DateTime  : 18/02/2007 19:52
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub EnviarAtrib(ByVal UserIndex As Integer)
Dim i As Integer
Dim cad$

For i = 1 To NUMATRIBUTOS
  cad$ = cad$ & Chr(UserList(UserIndex).Stats.UserAtributos(i))
Next

EnviarPaquete Paquetes.SendAtributos, cad$, UserIndex


End Sub

'---------------------------------------------------------------------------------------
' Procedure : EnviarMiniEstadisticas
' DateTime  : 18/02/2007 19:52
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub EnviarMiniEstadisticas(ByVal UserIndex As Integer)
'TODO no es necesario enviar el string de la clase
With UserList(UserIndex)
EnviarPaquete Paquetes.mest, ITS(.faccion.CiudadanosMatados) & ITS(.faccion.CriminalesMatados) & ITS(.faccion.NeutralesMatados) & ITS(.Stats.UsuariosMatados) & ByteToString(.faccion.alineacion) & LongToString(.Stats.NPCsMuertos) & ByteToString(.Counters.Pena) & byteToClase(.clase), UserIndex, ToIndex
End With
End Sub

Sub EraseUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, UserIndex As Integer)

    ' Libero el
    CharList(UserList(UserIndex).Char.charIndex) = 0

    If UserList(UserIndex).Char.charIndex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar <= 1 Then Exit Do
        Loop
    End If

    MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).UserIndex = 0
    'Le mandamos el mensaje para que borre el personaje a los clientes que este en el mismo mapa

    EnviarPaquete Paquetes.BorrarUser, ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToMap
    UserList(UserIndex).Char.charIndex = 0

    NumChars = NumChars - 1
End Sub

Public Function getPromedioAumentoVida(personaje As User) As Single
    Dim Rango As tRango
    
    Rango = getRangoAumentoVida(personaje)
    
    getPromedioAumentoVida = (Rango.maximo + Rango.minimo) / 2
End Function

' Retrona el minimo/maximo de puntos de vida que pude subir este usuario por nivel.
Public Function getRangoAumentoVida(personaje As User) As tRango

getRangoAumentoVida.maximo = 0
getRangoAumentoVida.minimo = 0

Select Case personaje.clase

Case eClases.Guerrero

    Select Case personaje.Stats.UserAtributos(eAtributos.constitucion)
        Case 21
            getRangoAumentoVida.minimo = 9
            getRangoAumentoVida.maximo = 12
        Case 20
            getRangoAumentoVida.minimo = 8
            getRangoAumentoVida.maximo = 12
        Case 19
            getRangoAumentoVida.minimo = 8
            getRangoAumentoVida.maximo = 11
        Case 18
            getRangoAumentoVida.minimo = 7
            getRangoAumentoVida.maximo = 11
        Case Else
            getRangoAumentoVida.minimo = 6 + AdicionalHPCazador
            getRangoAumentoVida.maximo = personaje.Stats.UserAtributos(eAtributos.constitucion) \ 2 + AdicionalHPCazador
    End Select

Case eClases.Cazador
    
    Select Case personaje.Stats.UserAtributos(eAtributos.constitucion)
        Case 21
            getRangoAumentoVida.minimo = 8
            getRangoAumentoVida.maximo = 11
        Case 20
            getRangoAumentoVida.minimo = 7
            getRangoAumentoVida.maximo = 11
        Case 19
            getRangoAumentoVida.minimo = 6
            getRangoAumentoVida.maximo = 11
        Case 18
            getRangoAumentoVida.minimo = 6
            getRangoAumentoVida.maximo = 10
        Case Else
            getRangoAumentoVida.minimo = 5
            getRangoAumentoVida.maximo = personaje.Stats.UserAtributos(eAtributos.constitucion) \ 2 + AdicionalHPCazador
    End Select

    Case eClases.Pirata

        Select Case personaje.Stats.UserAtributos(eAtributos.constitucion)
            Case 21
                getRangoAumentoVida.minimo = 9
                getRangoAumentoVida.maximo = 11
            Case 20
                getRangoAumentoVida.minimo = 8
                getRangoAumentoVida.maximo = 11
            Case 19
                getRangoAumentoVida.minimo = 7
                getRangoAumentoVida.maximo = 11
            Case 18
                getRangoAumentoVida.minimo = 6
                getRangoAumentoVida.maximo = 11
            Case Else
                getRangoAumentoVida.minimo = 5
                getRangoAumentoVida.maximo = personaje.Stats.UserAtributos(eAtributos.constitucion) \ 2 + AdicionalHPCazador
        End Select

    Case eClases.Paladin

        Select Case personaje.Stats.UserAtributos(eAtributos.constitucion)
            Case 21
                getRangoAumentoVida.minimo = 9
                getRangoAumentoVida.maximo = 11
            Case 20
                getRangoAumentoVida.minimo = 8
                getRangoAumentoVida.maximo = 11
            Case 19
                getRangoAumentoVida.minimo = 7
                getRangoAumentoVida.maximo = 11
            Case 18
                getRangoAumentoVida.minimo = 6
                getRangoAumentoVida.maximo = 11
            Case Else
                getRangoAumentoVida.minimo = 5
                getRangoAumentoVida.maximo = personaje.Stats.UserAtributos(eAtributos.constitucion) \ 2 + AdicionalHPCazador
        End Select

    Case eClases.Ladron

        Select Case personaje.Stats.UserAtributos(eAtributos.constitucion)
            Case 21
                getRangoAumentoVida.minimo = 6
                getRangoAumentoVida.maximo = 9
            Case 20
                getRangoAumentoVida.minimo = 5
                getRangoAumentoVida.maximo = 9
            Case 19
                getRangoAumentoVida.minimo = 4
                getRangoAumentoVida.maximo = 9
            Case 18
                getRangoAumentoVida.minimo = 4
                getRangoAumentoVida.maximo = 8
            Case 16, 17
                getRangoAumentoVida.minimo = 3
                getRangoAumentoVida.maximo = 7
            Case 16
                getRangoAumentoVida.minimo = 3
                getRangoAumentoVida.maximo = 6
            Case 14
                getRangoAumentoVida.minimo = 2
                getRangoAumentoVida.maximo = 6
            Case 13
                getRangoAumentoVida.minimo = 2
                getRangoAumentoVida.maximo = 5
            Case 12
                getRangoAumentoVida.minimo = 1
                getRangoAumentoVida.maximo = 5
            Case 11
                getRangoAumentoVida.minimo = 1
                getRangoAumentoVida.maximo = 4
            Case 10
                getRangoAumentoVida.minimo = 0
                getRangoAumentoVida.maximo = 4
            Case Else
                getRangoAumentoVida.minimo = 3
                getRangoAumentoVida.maximo = personaje.Stats.UserAtributos(eAtributos.constitucion) \ 2 - AdicionalHPGuerrero
        End Select
    
    Case eClases.Mago

        Select Case personaje.Stats.UserAtributos(eAtributos.constitucion)
            Case 21
                getRangoAumentoVida.minimo = 6
                getRangoAumentoVida.maximo = 8
            Case 20
                getRangoAumentoVida.minimo = 5
                getRangoAumentoVida.maximo = 8
            Case 19
                getRangoAumentoVida.minimo = 4
                getRangoAumentoVida.maximo = 8
            Case 18
                getRangoAumentoVida.minimo = 3
                getRangoAumentoVida.maximo = 8
            Case Else
                getRangoAumentoVida.minimo = 3
                getRangoAumentoVida.maximo = personaje.Stats.UserAtributos(eAtributos.constitucion) \ 2 - AdicionalHPGuerrero
        End Select

    Case eClases.Leñador

        Select Case personaje.Stats.UserAtributos(eAtributos.constitucion)
            Case 21
                getRangoAumentoVida.minimo = 9
                getRangoAumentoVida.maximo = 12
            Case 20
                getRangoAumentoVida.minimo = 8
                getRangoAumentoVida.maximo = 12
            Case 19
                getRangoAumentoVida.minimo = 8
                getRangoAumentoVida.maximo = 11
            Case 18
                getRangoAumentoVida.minimo = 7
                getRangoAumentoVida.maximo = 11
            Case Else
                getRangoAumentoVida.minimo = 6
                getRangoAumentoVida.maximo = personaje.Stats.UserAtributos(eAtributos.constitucion) \ 2 - AdicionalHPGuerrero
        End Select

    Case eClases.Minero

        Select Case personaje.Stats.UserAtributos(eAtributos.constitucion)
            Case 21
                getRangoAumentoVida.minimo = 9
                getRangoAumentoVida.maximo = 12
            Case 20
                getRangoAumentoVida.minimo = 8
                getRangoAumentoVida.maximo = 12
            Case 19
                getRangoAumentoVida.minimo = 8
                getRangoAumentoVida.maximo = 11
            Case 18
                getRangoAumentoVida.minimo = 7
                getRangoAumentoVida.maximo = 11
            Case Else
                getRangoAumentoVida.minimo = 6
                getRangoAumentoVida.maximo = personaje.Stats.UserAtributos(eAtributos.constitucion) \ 2 - AdicionalHPGuerrero
        End Select

    Case eClases.Pescador

        Select Case personaje.Stats.UserAtributos(eAtributos.constitucion)
            Case 21
                getRangoAumentoVida.minimo = 9
                getRangoAumentoVida.maximo = 12
            Case 20
                getRangoAumentoVida.minimo = 8
                getRangoAumentoVida.maximo = 12
            Case 19
                getRangoAumentoVida.minimo = 8
                getRangoAumentoVida.maximo = 11
            Case 18
                getRangoAumentoVida.minimo = 7
                getRangoAumentoVida.maximo = 11
            Case Else
                getRangoAumentoVida.minimo = 6
                getRangoAumentoVida.maximo = personaje.Stats.UserAtributos(eAtributos.constitucion) \ 2 - AdicionalHPGuerrero
        End Select

    Case eClases.Clerigo

        Select Case personaje.Stats.UserAtributos(eAtributos.constitucion)
            Case 21
                getRangoAumentoVida.minimo = 7
                getRangoAumentoVida.maximo = 10
            Case 20
                getRangoAumentoVida.minimo = 6
                getRangoAumentoVida.maximo = 10
            Case 19
                getRangoAumentoVida.minimo = 6
                getRangoAumentoVida.maximo = 9
            Case 18
                getRangoAumentoVida.minimo = 5
                getRangoAumentoVida.maximo = 9
            Case Else
                getRangoAumentoVida.minimo = 4
                getRangoAumentoVida.maximo = personaje.Stats.UserAtributos(eAtributos.constitucion) \ 2 - AdicionalHPCazador
         End Select

    Case eClases.Druida

        Select Case personaje.Stats.UserAtributos(eAtributos.constitucion)
             Case 21
                getRangoAumentoVida.minimo = 7
                getRangoAumentoVida.maximo = 10
            Case 20
                getRangoAumentoVida.minimo = 6
                getRangoAumentoVida.maximo = 10
            Case 19
                getRangoAumentoVida.minimo = 6
                getRangoAumentoVida.maximo = 9
            Case 18
                getRangoAumentoVida.minimo = 5
                getRangoAumentoVida.maximo = 9
            Case Else
                getRangoAumentoVida.minimo = 4
                getRangoAumentoVida.maximo = personaje.Stats.UserAtributos(eAtributos.constitucion) \ 2 - AdicionalHPCazador
         End Select
        
    Case eClases.asesino

        Select Case personaje.Stats.UserAtributos(eAtributos.constitucion)
             Case 21
                getRangoAumentoVida.minimo = 7
                getRangoAumentoVida.maximo = 10
            Case 20
                getRangoAumentoVida.minimo = 6
                getRangoAumentoVida.maximo = 10
            Case 19
                getRangoAumentoVida.minimo = 6
                getRangoAumentoVida.maximo = 9
            Case 18
                getRangoAumentoVida.minimo = 5
                getRangoAumentoVida.maximo = 9
            Case Else
                getRangoAumentoVida.minimo = 4
                getRangoAumentoVida.maximo = personaje.Stats.UserAtributos(eAtributos.constitucion) \ 2 - AdicionalHPCazador
         End Select

    Case eClases.Bardo

        Select Case personaje.Stats.UserAtributos(eAtributos.constitucion)
             Case 21
                getRangoAumentoVida.minimo = 7
                getRangoAumentoVida.maximo = 10
            Case 20
                getRangoAumentoVida.minimo = 6
                getRangoAumentoVida.maximo = 10
            Case 19
                getRangoAumentoVida.minimo = 6
                getRangoAumentoVida.maximo = 9
            Case 18
                getRangoAumentoVida.minimo = 5
                getRangoAumentoVida.maximo = 9
            Case Else
                getRangoAumentoVida.minimo = 4
                getRangoAumentoVida.maximo = personaje.Stats.UserAtributos(eAtributos.constitucion) \ 2 - AdicionalHPCazador
         End Select

    Case eClases.Herrero

        Select Case personaje.Stats.UserAtributos(eAtributos.constitucion)
             Case 21
                getRangoAumentoVida.minimo = 7
                getRangoAumentoVida.maximo = 9
            Case 20
                getRangoAumentoVida.minimo = 7
                getRangoAumentoVida.maximo = 8
            Case 19
                getRangoAumentoVida.minimo = 6
                getRangoAumentoVida.maximo = 8
            Case 18
                getRangoAumentoVida.minimo = 6
                getRangoAumentoVida.maximo = 7
            Case Else
                getRangoAumentoVida.minimo = 5
                getRangoAumentoVida.maximo = personaje.Stats.UserAtributos(eAtributos.constitucion) \ 2 + AdicionalHPCazador
         End Select

    Case eClases.Carpintero
    
        Select Case personaje.Stats.UserAtributos(eAtributos.constitucion)
             Case 21
                getRangoAumentoVida.minimo = 7
                getRangoAumentoVida.maximo = 9
            Case 20
                getRangoAumentoVida.minimo = 7
                getRangoAumentoVida.maximo = 8
            Case 19
                getRangoAumentoVida.minimo = 6
                getRangoAumentoVida.maximo = 8
            Case 18
                getRangoAumentoVida.minimo = 6
                getRangoAumentoVida.maximo = 7
            Case Else
                getRangoAumentoVida.minimo = 5
                getRangoAumentoVida.maximo = (personaje.Stats.UserAtributos(eAtributos.constitucion) \ 2) - AdicionalHPCazador
         End Select
        
    Case Else

         Select Case personaje.Stats.UserAtributos(eAtributos.constitucion)
             Case 21
                getRangoAumentoVida.minimo = 6
                getRangoAumentoVida.maximo = 9
            Case 20
                getRangoAumentoVida.minimo = 5
                getRangoAumentoVida.maximo = 9
            Case 19
                getRangoAumentoVida.minimo = 4
                getRangoAumentoVida.maximo = 8
            Case Else
                getRangoAumentoVida.minimo = 5
                getRangoAumentoVida.maximo = personaje.Stats.UserAtributos(eAtributos.constitucion) \ 2 - AdicionalHPCazador
         End Select

 End Select

End Function

' Retorna la vida ideal que deberia tener el personaje para su nivel
Public Function getVidaIdeal(ByRef personaje As User) As Single
    Dim promedio As Single
    Dim vidaBase As Integer
    Dim rangoAumento As tRango
    
    If Not personaje.clase = eClases.asesino Then
        vidaBase = 15 + Int(getPromedioAumentoVida(personaje) + 0.5)
    Else
        vidaBase = 30
    End If
    
    rangoAumento = getRangoAumentoVida(personaje)
    promedio = (rangoAumento.minimo + rangoAumento.maximo) / 2
    
    getVidaIdeal = vidaBase + (personaje.Stats.ELV - 1) * promedio
End Function

' Esto se aplica cuando el prsonaje ya paso de nivel.
Public Function obtenerAumentoHp(ByRef personaje As User) As Byte
    ' Calculo de vida
    Dim vidaPromedio As Integer
    Dim promedio As Single
    Dim vidaIdeal As Single
        
    Dim minimoAumento As Integer
    Dim maximoAumento As Integer
    Dim aumentoHp As Byte
    
    Dim rangoAumento As tRango
    Dim vidaBase As Integer
    
    If Not personaje.clase = eClases.asesino Then
        vidaBase = 15 + Int(getPromedioAumentoVida(personaje) + 0.5)
    Else
        vidaBase = 30
    End If
    
    
    rangoAumento = getRangoAumentoVida(personaje)
    promedio = (rangoAumento.minimo + rangoAumento.maximo) / 2
    vidaIdeal = vidaBase + (personaje.Stats.ELV - 2) * promedio
    
    If personaje.Stats.ELV <= MAX_NIVEL_VIDA_PROMEDIO Then
        ' Sube exactamente el promedio, pero si es decimal tenemos que hacer algo más
        If Int(promedio) = promedio Then
             aumentoHp = promedio
        Else
            If vidaIdeal > personaje.Stats.MaxHP Then
                ' Si es 8.5, se transforma a 8 y luego le sumo 1
                aumentoHp = Int(promedio) + 1
            Else
                aumentoHp = Int(promedio)
            End If
        End If
    Else
        Dim puntosAumento As Integer
        
        puntosAumento = Int(RandomNumberByte(rangoAumento.minimo, rangoAumento.maximo))
                
        If personaje.Stats.MaxHP < vidaIdeal - 10 Then
            ' Si esta por debajo del promedio menos 10: El valor será el máximo entre el promedio y lo que salio.
            aumentoHp = maxi(puntosAumento, Int(0.5 + promedio))
        ElseIf personaje.Stats.MaxHP > vidaIdeal + 15 Then
            ' Si el usuario esta por encima del promedio mas 15. El valor será el mínimo entre el promedio y lo que le salió.
            aumentoHp = mini(puntosAumento, Int(promedio))
        Else
            ' Si el usuario esta dentro del rango [-10, 15]: Sube lo que salió que es un random
            aumentoHp = puntosAumento
        End If
    End If
    
    obtenerAumentoHp = aumentoHp
End Function

Public Sub getPremioSubaNivel(ByRef personaje As User, ByRef aumentoST As Integer, ByRef aumentoHit As Integer, ByRef aumentoMana As Integer)
     Select Case personaje.clase
    
        Case eClases.asesino
        
            aumentoST = 15

            If personaje.Stats.ELV < 35 Then
                aumentoHit = 3
            Else
                aumentoHit = 1
            End If

            aumentoMana = personaje.Stats.UserAtributos(eAtributos.Inteligencia)

        Case eClases.Bardo
                    
            aumentoST = 15
            aumentoHit = 2
            aumentoMana = 2 * personaje.Stats.UserAtributos(eAtributos.Inteligencia)
    
        Case eClases.Cazador
                      
            'Aumento de la energia
            aumentoST = 15
                
            'Aumento del golpe
            If personaje.Stats.ELV < 35 Then
                aumentoHit = 3
            Else
                aumentoHit = 2
            End If
                
        Case eClases.Clerigo
                    
            aumentoST = 15
            aumentoHit = 2
            aumentoMana = 2 * personaje.Stats.UserAtributos(eAtributos.Inteligencia)
            
        Case eClases.Druida
                    
            aumentoST = 15
            aumentoHit = 2
            aumentoMana = 2 * personaje.Stats.UserAtributos(eAtributos.Inteligencia)
            
        Case eClases.Guerrero
                        
            'Aumento de la energia
            aumentoST = 15
            aumentoMana = 0
                
            'Aumenoto del golpe
            If personaje.Stats.ELV < 35 Then
                aumentoHit = 3
            Else
                aumentoHit = 2
            End If
              
         Case eClases.Herrero
                    
            aumentoST = 14
            aumentoHit = 1
            
        Case eClases.Ladron

            aumentoST = 15 + AdicionalSTLadron
            aumentoHit = 1

        Case eClases.Mago
                                
            aumentoST = 15 - AdicionalSTLadron / 2
            If aumentoST < 1 Then aumentoST = 5
            
            aumentoHit = 1
            
            If personaje.Stats.MaxMAN < 2000 Then
                aumentoMana = 3 * personaje.Stats.UserAtributos(eAtributos.Inteligencia)
            Else
                aumentoMana = 1.98 * personaje.Stats.UserAtributos(eAtributos.Inteligencia)
            End If

        ' Clases Trabajadoras
        Case eClases.Minero
                    
            aumentoST = 14
            aumentoHit = 2
                           
        Case eClases.Leñador
                    
            aumentoST = 14
            aumentoHit = 2
            
        Case eClases.Paladin
        
            aumentoST = 15

            If personaje.Stats.ELV < 35 Then
                aumentoHit = 3
            Else
                aumentoHit = 1
            End If
 
            aumentoMana = personaje.Stats.UserAtributos(eAtributos.Inteligencia)

        Case eClases.Pescador
                    
            aumentoST = 14
            aumentoHit = 1
        
        Case eClases.Pirata
                    
            aumentoST = 15
            aumentoHit = 3
                        
        Case Else
                    
            aumentoST = 15
            aumentoHit = 2
            
    End Select
    
End Sub
Sub CheckUserLevel(ByRef Usuario As User)

Dim Pts As Integer

Dim aumentoHit As Integer
Dim aumentoST As Integer
Dim aumentoMana As Integer
Dim aumentoHp As Integer
    
Dim WasNewbie As Boolean
Dim pasoDeNivel As Boolean
    
'¿Alcanzo el maximo nivel?
If Usuario.Stats.ELV = STAT_MAXELV Then
    Usuario.Stats.Exp = 0
    Usuario.Stats.ELU = 0
    Exit Sub
End If

'Variable que me indica si el usuario paso almenos un nivel en este chequeo
pasoDeNivel = False

WasNewbie = EsNewbie(Usuario.UserIndex)

'Si exp >= then Exp para subir de nivel entonce subimos el nivel
'If Usuario.Stats.Exp >= Usuario.Stats.ELU Then
Do While Usuario.Stats.Exp >= Usuario.Stats.ELU And Usuario.Stats.ELV < STAT_MAXELV
    
    pasoDeNivel = True 'Activo este flag para que el personaje se guarde.
                       'Con esto no hay rieso de que pasen de nivel y se pierda la información
    
    aumentoHit = 0
    aumentoST = 0
    aumentoMana = 0
    aumentoHp = 0
        
    If Usuario.Stats.ELV = 1 Then
      Pts = 10
    Else
      Pts = 5
    End If
    
    Usuario.Stats.SkillPts = Usuario.Stats.SkillPts + Pts
        
    Usuario.Stats.ELV = Usuario.Stats.ELV + 1
    
    'En esta linea se guarda la experiencia que me sobra.
    'Pero si pase al nivel máximo me tiene que quedar 0 de sobra
    If Usuario.Stats.ELV = STAT_MAXELV Then
        Usuario.Stats.Exp = 0
    Else
        Usuario.Stats.Exp = Usuario.Stats.Exp - Usuario.Stats.ELU
    End If
    
    If Not EsNewbie(Usuario.UserIndex) And WasNewbie Then
        Call QuitarNewbieObj(Usuario.UserIndex)
    End If
    
    ' Obtenemos la experiencia para el proximo nivel
    Usuario.Stats.ELU = obtenerExperienciaNecesaria(Usuario.Stats.ELV)

    aumentoHp = obtenerAumentoHp(Usuario)
        
    Call getPremioSubaNivel(Usuario, aumentoST, aumentoHit, aumentoMana)
    
    'Actualizo sus stats
    AddtoVar Usuario.Stats.MaxHP, aumentoHp, STAT_MAXHP ' Vida
    AddtoVar Usuario.Stats.MaxSta, aumentoST, STAT_MAXSTA ' Energia
    AddtoVar Usuario.Stats.MaxMAN, aumentoMana, STAT_MAXMAN 'Mana
    AddtoVar Usuario.Stats.MaxHIT, aumentoHit, STAT_MAXHIT 'Golpe Minimo
    AddtoVar Usuario.Stats.MinHIT, aumentoHit, STAT_MAXHIT 'Golpe Maximo
    
    EnviarPaquete Paquetes.WavSnd, Chr$(SOUND_NIVEL), Usuario.UserIndex, ToPCArea
    EnviarPaquete Paquetes.MensajeSimple, Chr$(47), Usuario.UserIndex
    EnviarPaquete Paquetes.MensajeCompuesto, Chr$(34) & Pts, Usuario.UserIndex
    
    'Empieza con la vida bien arriba.
    Usuario.Stats.minHP = Usuario.Stats.MaxHP
    
    'Le infomo y le actualizo sus stats
    If aumentoST > 0 Then EnviarPaquete Paquetes.MensajeCompuesto, Chr$(15) & aumentoHp, Usuario.UserIndex, ToIndex
    If aumentoMana > 0 Then EnviarPaquete Paquetes.MensajeCompuesto, Chr$(16) & aumentoMana, Usuario.UserIndex, ToIndex
    
    If aumentoHit > 0 Then
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(17) & aumentoHit, Usuario.UserIndex, ToIndex
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(18) & aumentoHit, Usuario.UserIndex, ToIndex
    End If
    
    Call LogSubeNivel(Usuario.id, aumentoHp, Usuario.Stats.MaxHP, Usuario.Stats.ELV)
    
    Call EnviarSkills(Usuario.UserIndex)
    Call EnviarSubirNivel(Usuario.UserIndex, Pts)
    
    Call EnviarPaquete(Paquetes.EnviarST, Codify(Usuario.Stats.MinSta), Usuario.UserIndex, ToIndex)
    Call SendUserStatsBox(Usuario.UserIndex)
    
    If Usuario.Stats.ELV = Constantes_Generales.STAT_MAXELV Then
        EnviarPaquete Paquetes.MensajeGuild, "¡¡¡" & Usuario.Name & " alcanzó el máximo poder!!!", Usuario.UserIndex, ToAll
    End If
Loop

'Si paso de nivel guardo el char por las dudas de que se caiga el server o algo
If pasoDeNivel Then Call SaveUser(Usuario.UserIndex, 1)

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChangeUserInv
' DateTime  : 18/02/2007 19:53
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub ChangeUserInv(UserIndex As Integer, slot As Byte, Object As UserOBJ)
 
UserList(UserIndex).Invent.Object(slot) = Object

If Object.ObjIndex > 0 Then
    EnviarPaquete 122, Chr$(slot) & ITS(Object.ObjIndex) & ITS(Object.Amount) & Object.Equipped & ITS(ObjData(Object.ObjIndex).GrhIndex) & Chr$(ObjData(Object.ObjIndex).ObjType) & ITS(ObjData(Object.ObjIndex).MaxHIT) & ITS(ObjData(Object.ObjIndex).MinHIT) & ITS(ObjData(Object.ObjIndex).MaxDef) & Codify((ObjData(Object.ObjIndex).valor \ 3)), UserIndex
Else
    EnviarPaquete 122, Chr$(slot), UserIndex
End If

End Sub

Function NextOpenCharIndex() As Integer

Dim loopC As Integer

For loopC = 1 To LastChar + 1
    If CharList(loopC) = 0 Then
        NextOpenCharIndex = loopC
        NumChars = NumChars + 1
        If loopC > LastChar Then LastChar = loopC
        Exit Function
    End If
Next loopC

End Function

'---------------------------------------------------------------------------------------
' Procedure : SendUserStatsBox
' DateTime  : 18/02/2007 19:53
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub SendUserStatsBox(ByVal UserIndex As Integer)
EnviarPaquete EnviarStat, ITS(UserList(UserIndex).Stats.MaxHP) & ITS(UserList(UserIndex).Stats.minHP) & ITS(UserList(UserIndex).Stats.MaxMAN) & ITS(UserList(UserIndex).Stats.MinMAN) & ITS(UserList(UserIndex).Stats.MaxSta) & ITS(UserList(UserIndex).Stats.MinSta) & LongToString(UserList(UserIndex).Stats.GLD) & Chr$(UserList(UserIndex).Stats.ELV) & WriteString(FormatNumber(UserList(UserIndex).Stats.ELU, 0, vbTrue, vbFalse, vbFalse)) & WriteString(FormatNumber(UserList(UserIndex).Stats.Exp, 0, vbTrue, vbFalse, vbFalse)), UserIndex
End Sub

'---------------------------------------------------------------------------------------
' Procedure : EnviarHambreYsed
' DateTime  : 18/02/2007 19:53
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub EnviarHambreYsed(ByVal UserIndex As Integer)
EnviarPaquete EnviarHYS, ByteToString(UserList(UserIndex).Stats.minAgu) & ByteToString(UserList(UserIndex).Stats.minham), UserIndex
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SendUserStatsTxt
' DateTime  : 18/02/2007 19:53
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)

EnviarPaquete Paquetes.mensajeinfo, "Estadisticas de: " & UserList(UserIndex).Name, sendIndex, ToIndex

EnviarPaquete Paquetes.mensajeinfo, "Clase: " & byteToClase(UserList(UserIndex).clase) & " Nivel: " & UserList(UserIndex).Stats.ELV & "  EXP: " & FormatNumber(UserList(UserIndex).Stats.Exp, 0, vbTrue, vbFalse, vbTrue) & "/" & FormatNumber(UserList(UserIndex).Stats.ELU, 0, vbTrue, vbFalse, vbTrue), sendIndex, ToIndex
EnviarPaquete Paquetes.mensajeinfo, "Salud: " & UserList(UserIndex).Stats.minHP & "/" & UserList(UserIndex).Stats.MaxHP & "  Mana: " & UserList(UserIndex).Stats.MinMAN & "/" & UserList(UserIndex).Stats.MaxMAN & "  Vitalidad: " & UserList(UserIndex).Stats.MinSta & "/" & UserList(UserIndex).Stats.MaxSta, sendIndex

If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
    EnviarPaquete Paquetes.mensajeinfo, "Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHIT & " (" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MaxHIT & ")", sendIndex, ToIndex
Else
    EnviarPaquete Paquetes.mensajeinfo, "Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHIT, sendIndex
End If

If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
    EnviarPaquete Paquetes.mensajeinfo, "(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef, sendIndex, ToIndex
Else
    EnviarPaquete Paquetes.mensajeinfo, "(CUERPO) Min Def/Max Def: 0", sendIndex, ToIndex
End If
If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
    EnviarPaquete Paquetes.mensajeinfo, "(CABEZA) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MaxDef, sendIndex, ToIndex
Else
    EnviarPaquete Paquetes.mensajeinfo, "(CABEZA) Min Def/Max Def: 0", sendIndex, ToIndex
End If
If UserList(UserIndex).GuildInfo.id > 0 Then
    EnviarPaquete Paquetes.mensajeinfo, "Clan: " & UserList(UserIndex).GuildInfo.GuildName, sendIndex, ToIndex
    If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
       If UserList(UserIndex).GuildInfo.ClanFundadoID = UserList(UserIndex).GuildInfo.id Then
            EnviarPaquete Paquetes.mensajeinfo, "Status:" & "Fundador/Lider", sendIndex, ToIndex
       Else
            EnviarPaquete Paquetes.mensajeinfo, "Status:" & "Lider", sendIndex, ToIndex
       End If
    Else
        EnviarPaquete Paquetes.mensajeinfo, "Status:" & UserList(UserIndex).GuildInfo.GuildPoints, sendIndex, ToIndex
    End If
    EnviarPaquete Paquetes.mensajeinfo, "User GuildPoints: " & UserList(UserIndex).GuildInfo.GuildPoints, sendIndex, ToIndex
End If
EnviarPaquete Paquetes.mensajeinfo, "Oro: " & UserList(UserIndex).Stats.GLD & "  Posicion: " & UserList(UserIndex).pos.x & "," & UserList(UserIndex).pos.y & " en mapa " & UserList(UserIndex).pos.map, sendIndex, ToIndex
EnviarPaquete Paquetes.mensajeinfo, "Dados: " & UserList(UserIndex).Stats.UserAtributos(1) & ", " & UserList(UserIndex).Stats.UserAtributos(2) & ", " & UserList(UserIndex).Stats.UserAtributos(3) & ", " & UserList(UserIndex).Stats.UserAtributos(4) & ", " & UserList(UserIndex).Stats.UserAtributos(5), sendIndex

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SendUserInvTxt
' DateTime  : 18/02/2007 19:54
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)

Dim j As Integer

EnviarPaquete Paquetes.mensajeinfo, "Inventario de: " & UserList(UserIndex).Name, sendIndex, ToIndex
EnviarPaquete Paquetes.mensajeinfo, "Tiene " & UserList(UserIndex).Invent.NroItems & " objetos.", sendIndex, ToIndex

For j = 1 To UserList(UserIndex).Stats.MaxItems
    If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
        EnviarPaquete Paquetes.mensajeinfo, "Objeto " & j & " " & ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Name & " Cantidad:" & UserList(UserIndex).Invent.Object(j).Amount, sendIndex, ToIndex
    End If
Next

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SendUserSkillsTxt
' DateTime  : 18/02/2007 19:54
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)

Dim j As Integer

EnviarPaquete Paquetes.MensajeClan1, "Skills de " & UserList(UserIndex).Name, sendIndex, ToIndex

For j = 1 To NUMSKILLS
    EnviarPaquete Paquetes.MensajeClan1, SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j), sendIndex, ToIndex
Next

End Sub

'---------------------------------------------------------------------------------------
' Procedure : UpdateUserMap
' DateTime  : 18/02/2007 19:54
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub UpdateUserMap(ByVal UserIndex As Integer)
' MODIFICADO

Dim map As Integer
Dim x As Integer
Dim y As Integer
Dim obj As obj
Dim cadena As String
Dim cantidad As Integer
Dim OtroUserIndex As Integer

map = UserList(UserIndex).pos.map

enviarNoche UserList(UserIndex)

For y = Y_MINIMO_USABLE To Y_MAXIMO_USABLE
    For x = X_MINIMO_USABLE To X_MAXIMO_USABLE
    
        OtroUserIndex = MapData(map, x, y).UserIndex
        If OtroUserIndex > 0 And UserIndex <> OtroUserIndex Then
        
            Call modPersonaje_TCP.MakeUserChar(UserList(OtroUserIndex), UserIndex, ToIndex)
            'Call MakeUserChar(ToIndex, UserIndex, 0, MapData(Map, x, y).UserIndex, Map, x, y)
            
             If UserList(OtroUserIndex).flags.Oculto = 1 Then
                EnviarPaquete Paquetes.ocultar, ITS(UserList(OtroUserIndex).Char.charIndex), UserIndex, ToIndex
            ElseIf UserList(OtroUserIndex).flags.Invisible = 1 Then
                EnviarPaquete Paquetes.Invisible, ITS(UserList(OtroUserIndex).Char.charIndex) & ByteToString(getInvisibilidadTipo(UserList(OtroUserIndex))), UserIndex, ToIndex
            End If
        End If

        If MapData(map, x, y).npcIndex > 0 Then
            Call EnviarNPCChar(ToIndex, UserIndex, 0, MapData(map, x, y).npcIndex, x, y)
        End If

        If MapData(map, x, y).OBJInfo.ObjIndex > 0 Then
                

                If obj.ObjIndex <= UBound(ObjData) And Not ObjData(MapData(map, x, y).OBJInfo.ObjIndex).ObjType = OBJTYPE_ARBOLES Then
                cadena = cadena & ITS(ObjData(MapData(map, x, y).OBJInfo.ObjIndex).GrhIndex) & Chr(x) & Chr(y)
                cantidad = cantidad + 1
                End If
 

            If ObjData(MapData(map, x, y).OBJInfo.ObjIndex).ObjType = OBJTYPE_PUERTAS Then
                ' TODO
                ' Call Bloquear(ToIndex, UserIndex, 0, map, x, y, MapData(map, x, y).Blocked)
                ' Call Bloquear(ToIndex, UserIndex, 0, map, x - 1, y, MapData(map, x - 1, y).Blocked)
            End If
        End If
    Next x
Next y

EnviarPaquete Paquetes.CrearObjetoInicio, ITS(cantidad) & cadena, UserIndex

End Sub

'---------------------------------------------------------------------------------------
' Procedure : NpcAtacado
' DateTime  : 18/02/2007 19:55
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub NpcAtacado(ByVal npcIndex As Integer, ByVal UserIndex As Integer)

If NpcList(npcIndex).MaestroUser > 0 Then
    Call AllMascotasAtacanUser(UserIndex, NpcList(npcIndex).MaestroUser)
End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ContarMuerte
' DateTime  : 18/02/2007 19:55
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub ContarMuerte(ByRef asesinado As User, ByRef atacante As User)

If EsNewbie(asesinado.UserIndex) Then Exit Sub

If EsPosicionParaAtacarSinPenalidad(asesinado.pos) And EsPosicionParaAtacarSinPenalidad(atacante.pos) Then Exit Sub

If asesinado.faccion.alineacion = eAlineaciones.caos Then
    If atacante.flags.LastCrimMatado <> asesinado.Name Then
        atacante.flags.LastCrimMatado = asesinado.Name
        Call AddtoVar(atacante.faccion.CriminalesMatados, 1, MAXUSERMATADOS)
    End If
ElseIf asesinado.faccion.alineacion = eAlineaciones.Neutro Then
    If atacante.flags.LastNeutralMatado <> asesinado.Name Then
        atacante.flags.LastNeutralMatado = asesinado.Name
        Call AddtoVar(atacante.faccion.NeutralesMatados, 1, MAXUSERMATADOS)
    End If
ElseIf asesinado.faccion.alineacion = eAlineaciones.Real Then
    If atacante.flags.LastCiudMatado <> asesinado.Name Then
        atacante.flags.LastCiudMatado = asesinado.Name
        Call AddtoVar(atacante.faccion.CiudadanosMatados, 1, MAXUSERMATADOS)
    End If
End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : WarpUserChar
' DateTime  : 18/02/2007 19:55
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub WarpUserChar(ByVal UserIndex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, Optional ByVal FX As Boolean = False)

Dim OldMap As Integer
Dim OldX As Integer
Dim OldY As Integer


EnviarPaquete Paquetes.QDL, ITS(UserList(UserIndex).Char.charIndex), UserIndex, ToMap
EnviarPaquete Paquetes.QTDL, "", UserIndex, ToIndex
EnviarPaquete Paquetes.BorrarArea, "", UserIndex, ToIndex

OldMap = UserList(UserIndex).pos.map
OldX = UserList(UserIndex).pos.x
OldY = UserList(UserIndex).pos.y

Call EraseUserChar(ToMap, 0, OldMap, UserIndex)

UserList(UserIndex).pos.x = x
UserList(UserIndex).pos.y = y
UserList(UserIndex).pos.map = map

If OldMap <> map Then
    Call personajeCambioDeMapa(UserList(UserIndex), OldMap)
Else
    Call modPersonaje_TCP.MakeUserChar(UserList(UserIndex), 0, ToMap)
  
    EnviarPaquete Paquetes.IndiceChar, ITS(UserList(UserIndex).Char.charIndex), UserIndex
    
    UserList(UserIndex).Counters.Invisibilidad = 0
    UserList(UserIndex).flags.Invisible = 0
    UserList(UserIndex).flags.Oculto = 0
End If

Call ActualizarTodaArea(UserIndex)

If FX And UserList(UserIndex).flags.AdminInvisible = 0 Then 'FX
    If MapInfo(map).Name = "Dungeon Magma" Then
        EnviarPaquete Paquetes.HechizoFX, ITS(UserList(UserIndex).Char.charIndex) & ByteToString(19) & ITS(0) & Chr$(SND_WARP), UserIndex, ToPCArea, UserList(UserIndex).pos.map
    Else
        EnviarPaquete Paquetes.HechizoFX, ITS(UserList(UserIndex).Char.charIndex) & ByteToString(FXWARP) & ITS(0) & Chr$(SND_WARP), UserIndex, ToPCArea, UserList(UserIndex).pos.map
    End If
End If

' ¿Esta Mimetizado? Si pasa a Zona Segura se le va el efecto
If UserList(UserIndex).flags.Mimetizado = 1 Then
    If MapInfo(map).Pk = False Then
        Call modMimetismo.finalizarEfecto(UserList(UserIndex))
    End If
End If

Call WarpMascotas(UserIndex)


End Sub


Private Sub personajeCambioDeMapa(ByRef personaje As User, ByVal anteriorMapa As Integer)
    Dim mapa As Integer
    
    mapa = personaje.pos.map
    
    EnviarPaquete ChangeMap, ITS(mapa) & ITS(MapInfo(mapa).climaActual) & MapInfo(mapa).Terreno & "," & MapInfo(mapa).zona & "," & MapInfo(mapa).Name, personaje.UserIndex

    If personaje.LuchandoNPC > 0 Then
        ' Si antes le estaba pegando a otro npc, libero a ese npc
        Call AntiRoboNpc.resetearLuchador(NpcList(personaje.LuchandoNPC))
    End If

    If val(MapInfo(mapa).Music) > 1 Then
        If MapInfo(mapa).Music <> MapInfo(anteriorMapa).Music Then
            EnviarPaquete Paquetes.ChangeMusic, Chr$(Left(val(MapInfo(mapa).Music), 1)), personaje.UserIndex
        End If
    End If

    'Agrego el personaje al mapa donde paso
    MapInfo(mapa).usuarios.agregar (personaje.UserIndex)

    'Saco del mapa viejo la referencia al usuario
    MapInfo(anteriorMapa).usuarios.eliminar (personaje.UserIndex)

    Call modPersonaje_TCP.MakeUserChar(personaje, 0, ToMap)
   
    EnviarPaquete Paquetes.IndiceChar, ITS(personaje.Char.charIndex), personaje.UserIndex

    Call UpdateUserMap(personaje.UserIndex)
    
    If MapInfo(mapa).AntiHechizosPts = 1 Then
        If personaje.flags.Oculto = 1 Then
            Call quitarOcultamiento(personaje)
        End If
        
        If personaje.flags.Invisible = 1 Then
            Call quitarInvisibilidad(personaje)
        End If
    Else
        'Seguis invisible al pasar de mapa. Marche, esto estaba al final del sub lo traje aca por que sino cambia de mapa no tiene por que irse la invi!
        If personaje.flags.Oculto = 1 Then
            EnviarPaquete Paquetes.ocultar, ITS(personaje.Char.charIndex), personaje.UserIndex, ToMap
        ElseIf personaje.flags.Invisible = 1 And (Not personaje.flags.Mimetizado = 1) Then
            EnviarPaquete Paquetes.Invisible, ITS(personaje.Char.charIndex) & ByteToString(getInvisibilidadTipo(personaje)), personaje.UserIndex, ToMap
        End If
    End If
End Sub


'Elimina las mascotas en la posicion donde estan y las trae donde esta su dueño.
Sub WarpMascotas(ByVal UserIndex As Integer, Optional noMatar As Boolean)

Dim i As Integer
Dim InvocadosMatados As Byte
Dim donde As WorldPos
Dim nDonde As WorldPos
Dim npcIndex As Integer

InvocadosMatados = 0

' Donde voy a poner a las mascotas
donde.map = UserList(UserIndex).pos.map
donde.x = UserList(UserIndex).pos.x + 1
donde.y = UserList(UserIndex).pos.y + 1

'Matamos los invocados
For i = 1 To MAXMASCOTAS

    npcIndex = UserList(UserIndex).MascotasIndex(i)

    If npcIndex > 0 Then

        If NpcList(npcIndex).Contadores.TiempoExistencia > 0 And Not noMatar Then
            'Si la mascota tiene tiempo de vida > 0 significa q fue invocada.
            'A menos que se haya indicado lo contrario (nomatar =true), la mato
            Call QuitarNPC(npcIndex)
            
            InvocadosMatados = InvocadosMatados + 1
            'Una mascota menos. Esto esta incluido en el quitarNPC
        Else
            'La quito del mapa
            Call EraseNPCChar(NpcList(npcIndex).pos.map, npcIndex)

            'Obtengo una posicion donde puda respawenarlo
            Call ClosestLegalPosNPC(donde, nDonde, NpcList(npcIndex))

            'Encontro una posicion valida
            If nDonde.map > 0 Then
                'Lo pongo en el mapa nuevo
                Call MakeNPCChar(ToMap, UserIndex, nDonde.map, npcIndex, nDonde.map, nDonde.x, nDonde.y)
                'A seguir al amo
                Call FollowAmo(npcIndex)
            Else ' Si no encontro una posicion válida, lo siento. Se queda sin el npc.
                Call QuitarNPC(npcIndex)
                InvocadosMatados = InvocadosMatados + 1
            End If
            
        End If
    End If
Next i

'Si alguna de sus mascotas murio, le aviso.
If InvocadosMatados > 0 Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(49), UserIndex
End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SendUserMana
' DateTime  : 18/02/2007 19:57
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub SendUserMana(ByVal UserIndex As Integer)
EnviarPaquete Paquetes.EnviarMP, Codify(UserList(UserIndex).Stats.MinMAN), UserIndex
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SendUserVida
' DateTime  : 18/02/2007 19:57
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub SendUserVida(ByVal UserIndex As Integer)
EnviarPaquete Paquetes.EnviarHP, Codify(UserList(UserIndex).Stats.minHP), UserIndex
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SendUserEsta
' DateTime  : 18/02/2007 19:57
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub SendUserEsta(ByVal UserIndex As Integer)
EnviarPaquete Paquetes.EnviarST, Codify(UserList(UserIndex).Stats.MinSta), UserIndex
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SendUserStatsBoxBasicas
' DateTime  : 18/02/2007 19:57
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub SendUserStatsBoxBasicas(ByVal UserIndex As Integer)

EnviarPaquete Paquetes.EnviarStatsBasicas, ITS(UserList(UserIndex).Stats.minHP) & ITS(UserList(UserIndex).Stats.MinMAN) & ITS(UserList(UserIndex).Stats.MinSta), UserIndex

End Sub

'---------------------------------------------------------------------------------------
' Procedure : WarpUserCharEspecial
' DateTime  : 18/02/2007 19:57
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub WarpUserCharEspecial(ByVal UserIndex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, donde As Byte)

UserList(UserIndex).pos.x = x
UserList(UserIndex).pos.y = y
UserList(UserIndex).pos.map = map
    
'Call MakeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
MapData(map, x, y).UserIndex = UserIndex
'EnviarPaquete Paquetes.IndiceChar, ITS(UserList(UserIndex).Char.charindex), UserIndex
EnviarPaquete Paquetes.MoverMuerto, ITS(UserList(UserIndex).Char.charIndex) & donde, 0, ToMap, map

End Sub

' Un personaje baneado siempre va a estar offline
Public Sub UnbanearUsuario(nombreGM As String, nombreUsuario As String, GmOnline As Boolean)

Dim infoPersonaje As ADODB.Recordset

Call cargarAtributosPersonajeOffline(nombreUsuario, infoPersonaje, "BANB, UNBAN, BANRAZB", True)

If Not infoPersonaje.EOF Then

    infoPersonaje!banb = 0
    infoPersonaje!Unban = ""
    infoPersonaje!banrazb = infoPersonaje!banrazb & vbCrLf & "Unban por " & nombreGM & ". " & Date
    infoPersonaje.Update

    If GmOnline Then EnviarPaquete Paquetes.mensajeinfo, nombreUsuario & " ha sido desbaneado.", NameIndex(nombreGM), ToIndex
Else
    If GmOnline Then EnviarPaquete Paquetes.mensajeinfo, "El personaje " & nombreUsuario & " no existe.", NameIndex(nombreGM), ToIndex
End If

'Libero memoria
infoPersonaje.Close
Set infoPersonaje = Nothing
       
End Sub


Public Function getInvisibilidadTipo(ByRef personaje As User) As Byte
    If personaje.clase = eClases.asesino Then
        getInvisibilidadTipo = 1
    Else
        getInvisibilidadTipo = 0
    End If
End Function
