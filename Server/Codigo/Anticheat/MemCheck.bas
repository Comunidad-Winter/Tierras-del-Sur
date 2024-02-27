Attribute VB_Name = "Anticheat_MemCheck"
' Voy a revisar un Long que me va a indicar la posicion de memoria a Leer
' Un Byte me va a indicar cuantos tengo que leer.
' Lo voy a leer y se lo voy a enviar.
        
' En el Servidor voy a tener una Lista posible de checks
' memCheck(1).posicion = X
' memCheck(1).bytes = Y
' memCheck(1).resultado(...) = Byte
        
' En cada usuario voy a gaurdar la info de la consulta
' .memConsulta.tipo = 1
' .memConsulta.fecha = fecha
'
' Si no tiene chequeo va a ser igual a 0.
' Si tiene chequeo y pasaron más de 30 segundos, se considera que contesto mal.
        
' Si contesta mal... se lo banea directamente.
        
' Si cierra justo cuando estaba un chequeo, lo guardo como altera del anticheat.
        
' La forma de cargar esto es haciendo un Hook sobre ReadProcessMemori
' Si la posición a leer cae dentro de lo que el cheat modifica, entonces envio una copia de seguridad.
' Sino, mando lo correcto.

Option Explicit
        
Private Type MemCheckDataType
    nombre As String                ' Nombre que identifica al chequeo
    direccion As Long               ' Dirección de memoria que se va a leer
    cantidadBytes As Byte           ' Cantidad de Bytes a leer
    resultadoStr As String          ' El resultado que debe dar la evaluacion en el cliente
End Type

Private MemCheckData() As MemCheckDataType      ' Posibles chequeos que se pueden hacer

Private chequeosActivos As Collection           ' Chequeos que se estan realizados

Private cantidadPosiblesChequeos As Integer     ' Cantidad de distintos tipos de chequeos

Private Const TOLERANCIA_BANEO As Long = 180000
Private Const PROBABILIDAD_CHEQUEO_MEM As Integer = 150 ' Mientras mas bajo sea este numero, más veces por hora hara el chequeo

' El personaje cierra
' El personaje responde
' Pasa el tiempo admitido
' Se crea un chequeo para el Personaje

Private Function obtenerChequeo(idPersonaje As Long) As clsChequeoMemoriaPersonaje
    
    Dim chequeo As clsChequeoMemoriaPersonaje
    
    If chequeosActivos.Count = 0 Then
        Set obtenerChequeo = Nothing
        Exit Function
    End If
    
    For Each chequeo In chequeosActivos
        If chequeo.personajeID = idPersonaje Then
            Set obtenerChequeo = chequeo
            Exit Function
        End If
    Next

    Set obtenerChequeo = Nothing
End Function

Public Function existeChequeoPara(personaje As User) As Boolean
    
    Dim chequeo As clsChequeoMemoriaPersonaje
    
    If chequeosActivos.Count = 0 Then
        existeChequeoPara = False
        Exit Function
    End If
    
    For Each chequeo In chequeosActivos
        If chequeo.personajeID = personaje.id Then
            existeChequeoPara = True
            Exit Function
        End If
    Next

    existeChequeoPara = False

End Function

Private Sub generarNuevosChequeos()
      
    Dim i As Integer
    Dim numeroAzar As Integer
    Dim chequeo As clsChequeoMemoriaPersonaje
    Dim ahora As Long
    Dim numeroChequeo As Long
    
    If cantidadPosiblesChequeos = 0 Then
        Exit Sub
    End If
    
    ahora = GetTickCount()
    numeroAzar = HelperRandom.RandomIntNumber(0, PROBABILIDAD_CHEQUEO_MEM)
    numeroChequeo = HelperRandom.RandomIntNumber(0, UBound(MemCheckData))
    
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            
            ' La probabilidad de tener que hacer un chequeo depende del nivel.
            ' Mientras más nivel sea, más probabilidad tiene de ser chequeado
            If UserList(i).flags.Privilegios = 0 Then
                If (UserList(i).Stats.ELV <= 35 And UserList(i).Stats.ELV >= numeroAzar) Or _
                    (UserList(i).Stats.ELV > 35 And UserList(i).Stats.ELV * 1.5 >= numeroAzar) Or _
                    (UserList(i).Stats.ELV > 35 And Not UserList(i).evento Is Nothing) Then
                
                       If Not existeChequeoPara(UserList(i)) Then
                           Set chequeo = New clsChequeoMemoriaPersonaje
                           
                           chequeo.fecha = ahora
                           chequeo.personajeID = UserList(i).id
                           chequeo.rtaEsperada = MemCheckData(numeroChequeo).resultadoStr
                           chequeo.nombreChequeo = MemCheckData(numeroChequeo).nombre
                           
                           Call chequeosActivos.Add(chequeo)
                           
                           Call LogDesarrollo("Chequeo sobre " & UserList(i).Name)
    
                           ' Enviamos la consulta
                           EnviarPaquete Paquetes.checkMem, LongToString(MemCheckData(numeroChequeo).direccion) & Chr$(MemCheckData(numeroChequeo).cantidadBytes), UserList(i).UserIndex, ToIndex
        
                       End If
                End If
            End If
            
        End If
    Next i

End Sub

Private Function removerChequeo_Personaje(personaje As User) As clsChequeoMemoriaPersonaje
    Dim loopChequeo As Integer
    Dim chequeo As clsChequeoMemoriaPersonaje
    
    If chequeosActivos.Count = 0 Then
        Set removerChequeo_Personaje = Nothing
        Exit Function
    End If
    
    For loopChequeo = 1 To chequeosActivos.Count
        Set chequeo = chequeosActivos.Item(loopChequeo)
        
        If chequeo.personajeID = personaje.id Then
            Set removerChequeo_Personaje = chequeo
            chequeosActivos.Remove (loopChequeo)
            Exit Function
        End If
    Next

    Set removerChequeo_Personaje = Nothing
End Function

Public Sub hook_cierraPersonaje(personaje As User)

    Dim chequeo As clsChequeoMemoriaPersonaje
    
    Set chequeo = removerChequeo_Personaje(personaje)
    
    If chequeo Is Nothing Then Exit Sub

    Call LogAnticheat(personaje, memCheck, "Cierra al momento del chequeo")

End Sub
Public Sub respuestaPersonaje(personaje As User, respuesta As String)

    Dim chequeo As clsChequeoMemoriaPersonaje
    
    Set chequeo = removerChequeo_Personaje(personaje)
    
    If chequeo Is Nothing Then
        LogError ("MemCheck. Respuesta Personaje. Se intenta dar una respuesta a un chequeo que no existe.")
        Exit Sub
    End If

    If Not (chequeo.rtaEsperada = respuesta) Then
        Call LogAnticheat(personaje, eAnticheat.memCheck, "respuesta incorrecta: " & chequeo.nombreChequeo)
        EnviarPaquete Paquetes.MensajeFight, personaje.Name & "  poseé cliente editado.", 0, ToAdmins
        EnviarPaquete Paquetes.MensajeFight, "Se banea a " & personaje.Name & " por uso de cheat.", 0, ToAdmins
        Call BanearUsuario("Anticheat", personaje.Name, "Cliente Editado", 0, False)
    Else
        Call LogDesarrollo(personaje.Name & " contesta correctamente.")
    End If

End Sub
Public Sub chequearPersonaje(personaje As User)
    Dim chequeo As clsChequeoMemoriaPersonaje
    Dim numeroChequeo As Long
    Dim ahora As Long

    If cantidadPosiblesChequeos = 0 Then
        Exit Sub
    End If
    
    Set chequeo = New clsChequeoMemoriaPersonaje
    
    ahora = GetTickCount()
    numeroChequeo = HelperRandom.RandomIntNumber(0, UBound(MemCheckData))
                        
    chequeo.fecha = ahora
    chequeo.personajeID = personaje.id
    chequeo.rtaEsperada = MemCheckData(numeroChequeo).resultadoStr
    chequeo.nombreChequeo = MemCheckData(numeroChequeo).nombre
    
    Call chequeosActivos.Add(chequeo)
                    
    ' Enviamos la consulta
    EnviarPaquete Paquetes.checkMem, LongToString(MemCheckData(numeroChequeo).direccion) & Chr$(MemCheckData(numeroChequeo).cantidadBytes), personaje.UserIndex, ToIndex
End Sub

Private Sub revisarChequeosActuales()
     
    Dim cotaTiempo As Long      ' Tiempo maximo para contestar
    Dim loopChequeo As Integer
    Dim chequeo As clsChequeoMemoriaPersonaje
    Dim UserIndex As Integer
    
    
    If chequeosActivos.Count = 0 Then Exit Sub

    loopChequeo = 1
    cotaTiempo = GetTickCount - TOLERANCIA_BANEO
    
    Do While loopChequeo <= chequeosActivos.Count
                
        Set chequeo = chequeosActivos.Item(loopChequeo)
        
        If chequeo.fecha < cotaTiempo Then
            Call chequeosActivos.Remove(loopChequeo)
            
            UserIndex = IDIndex(chequeo.personajeID)
            
            If (UserIndex = 0) Then
                ' Esto no deberia suceder
                Call LogError("MemCheck. Revisar Chequeos Actuales: se revisa un chequeo de un personaje que cerro.")
                ' Proximo
                GoTo continue
            End If

            Call LogAnticheat(UserList(UserIndex), eAnticheat.memCheck, "no contesto: " & chequeo.nombreChequeo)
            EnviarPaquete Paquetes.MensajeFight, UserList(UserIndex).Name & " poseé cliente Editado", 0, ToAdmins
            
            ' EnviarPaquete Paquetes.mensajeinfo, "Se banea a " & UserList(userIndex).Name & " por uso de cheat.", 0, ToAdmins
        
            ' Call BanearUsuario("Sistema", UserList(userIndex).Name, "Uso de cheat.", 0, False)
        Else
            loopChequeo = loopChequeo + 1
        End If
        
continue:
    Loop
     
End Sub
' Reviso los chequeos actuales
' Creo nuevo chequeos
Public Sub hook_pasar_Minuto()
    Call revisarChequeosActuales
    
    Call generarNuevosChequeos
End Sub

Public Sub iniciarEstructuras()
    Set chequeosActivos = New Collection
    
    Call cargarChequeosData
End Sub

Public Function cargarChequeosData() As Boolean

    Dim Leer As clsLeerInis
    Dim loopData As Integer
    Dim cantidadBytes As Integer
    Dim loopResultado As Integer
    
    cargarChequeosData = False
    
    Set Leer = New clsLeerInis

    Call Leer.Abrir(DatPath & "MemChecks.dat")

    cantidadPosiblesChequeos = CInt(Leer.DarValor("INIT", "CANTIDAD"))

    If cantidadPosiblesChequeos = 0 Then
        cargarChequeosData = True
        Exit Function
    End If
    
    ReDim MemCheckData(0 To cantidadPosiblesChequeos - 1) As MemCheckDataType

    For loopData = 0 To cantidadPosiblesChequeos - 1
        cantidadBytes = CInt(Leer.DarValor(CStr(loopData), "CANTIDAD_BYTES"))
        
        MemCheckData(loopData).nombre = Leer.DarValor(CStr(loopData), "NOMBRE")
        MemCheckData(loopData).cantidadBytes = cantidadBytes
        MemCheckData(loopData).direccion = CLng(Leer.DarValor(CStr(loopData), "DIRECCION"))
        
        For loopResultado = 0 To cantidadBytes - 1
            MemCheckData(loopData).resultadoStr = MemCheckData(loopData).resultadoStr & Chr$(CByte(Leer.DarValor(CStr(loopData), "RESULTADO_" & loopResultado)))
        Next
        
    Next
   
    cargarChequeosData = True
End Function

