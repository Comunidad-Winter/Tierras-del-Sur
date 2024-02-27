Attribute VB_Name = "Anticheat"
Option Explicit

Public Enum eAnticheat
    memCheck = 1
    cheatEngine = 2
    macro = 3
    intervalos = 4
End Enum

Const TOLERANCIA_CHEAT_ENGINE As Integer = 8

Public Sub anticheatCliente(ByRef personaje As User, datos As String)
    Dim tipo As String
    Dim velocidad As Long
    
    ' Obtenemos el tipo de alerta del Cliente
    tipo = mid$(datos, 1, 1)
        
    If tipo = "2" Then
    
        velocidad = StringToLong(datos, 2)
        
        EnviarPaquete Paquetes.MensajeFight, personaje.Name & " posible uso de CHEAT ENGINE (" & Round(velocidad / 60000, 3) & ").", 0, ToAdmins
        
        Call LogAnticheat(personaje, eAnticheat.cheatEngine, "Velocidad: " & Round(velocidad / 60000, 3) & ". Mapa " & personaje.pos.map)
            
        ' Sistema de Baneo Automatico
        personaje.controlCheat.vecesCheatEngine = personaje.controlCheat.vecesCheatEngine + 1
            
        If personaje.controlCheat.vecesCheatEngine = TOLERANCIA_CHEAT_ENGINE Then
            Call BanearUsuario("Anticheat", personaje.Name, "Uso de aceleradores", 0, False)
            
            EnviarPaquete Paquetes.MensajeFight, personaje.Name & " baneado por sistema anticheat por uso de aceleradores.", 0, ToAdmins
        End If
            
    ElseIf tipo = "1" Then
    
        EnviarPaquete Paquetes.mensajeinfo, personaje.Name & " posible uso de auto hechizo. Gravedad: " & StringToByte(velocidad, 2) & "-" & StringToByte(datos, 3) & "-" & StringToByte(datos, 4) & "-" & StringToByte(datos, 5) & "-" & StringToByte(datos, 6), 0, ToAdmins
        
        Call LogAnticheat(personaje, macro, "Posible uso de auto hechizo. Gravedad: " & StringToByte(datos, 2) & "-" & StringToByte(datos, 3) & "-" & StringToByte(datos, 4) & "-" & StringToByte(datos, 5) & "-" & StringToByte(datos, 6))
            
    ElseIf tipo = "0" Then
    
        EnviarPaquete Paquetes.mensajeinfo, personaje.Name & " posible uso de auto remover.", 0, ToAdmins
        
        Call LogAnticheat(personaje, macro, "Posible uso de auto remover.")
            
    End If
        
End Sub
' Este anticheat sirve para verificar las marcas de tiempo con las que genera los paquetes el Cliente
' Si es menor a lo que deberia ser, significa que de alguna manera el cliente esta genernado paquetes mas rapido
' de lo que deberia
Public Sub chequeoIntervaloCliente(ByRef personaje As User, ByRef counterIntervalo As Single, ByVal intervalo As Single, ByVal timeStamp As Single, ByVal tipo As String)
    Dim tempSingle As Single
    
    ' Obtengo la diferencia del CLIENTE que tardo en chupar
    tempSingle = timeStamp - counterIntervalo
            
    ' ¿Es menor al intervalo? Nos damos cuenta si lo está rompiendo
    If tempSingle < (intervalo / 1000) And tempSingle >= 0 Then
    
        ' Si lo rompio, lo guardamos
        personaje.controlCheat.rompeIntervalo = personaje.controlCheat.rompeIntervalo + 1
        
        ' Logueamos siempre.
        Call LogAnticheat(personaje, eAnticheat.intervalos, "Rompe Intervalos " & tipo & " . Corte: " & tempSingle)
        
        If personaje.controlCheat.rompeIntervalo > 5 Then
        
            personaje.controlCheat.rompeIntervalo = 0
            
             ' Call BanearUsuario("Anticheat", personaje.Name, "Uso de cheat", 0, False)
            Call Anticheat_MemCheck.chequearPersonaje(personaje)
            
            ' EnviarPaquete Paquetes.MensajeFight, Personaje.Name & " baneado por speed para " & tipo & ".", 0, ToAdmins
            EnviarPaquete Paquetes.MensajeFight, personaje.Name & "Rompe Intervalos " & tipo & ".", 0, ToAdmins
        End If
    End If
            
    ' Guardamos el momento del ultimo tick
    counterIntervalo = timeStamp
     
    ' Si el tiempo que tardo medido en milisegundos es igual... sospechoso
    ' If tempSingle = UserList(UserIndex).Counters.ultimaDifClic Then
        ' Damos un alerta
    '    UserList(UserIndex).VecesAtack = UserList(UserIndex).VecesAtack + 1
    'Else
        ' No es igual, reseteamos las alertas y guardamos este intervalo
    '    UserList(UserIndex).VecesAtack = 0
    '    UserList(UserIndex).Counters.ultimaDifClic = tempSingle
    'End If
End Sub


Public Sub LogAnticheat(personaje As User, anticheat As eAnticheat, descripcion As String)
    Dim sql As String
    
    sql = "INSERT DELAYED " & DB_NAME_PRINCIPAL & ".juego_logs_anticheat(personajeId, anticheatId, descripcion) VALUES(" & personaje.id & ",'" & anticheat & "','" & descripcion & "')"
    conn.Execute sql, , adExecuteNoRecords
End Sub
