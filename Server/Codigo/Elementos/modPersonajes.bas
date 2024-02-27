Attribute VB_Name = "modPersonajes"
Option Explicit

Private Const SEGUNDOS_PENA_DESCONEXION = 30 * 1000
Private Const SEGUNDOS_SALIR = 10 * 1000

' Usuario jugando
Public UserList() As User
Private UserIndexLibres As ColaConBloques

Private BackupUserIndex() As Boolean

' Mayor UserIndex utilizado
Public LastUser As Integer

Public Function obtenerSlotsLibres() As Integer
    obtenerSlotsLibres = UserIndexLibres.getCantidadElementos
End Function

Public Sub iniciarEstructuras()
    Dim i As Integer
    
    Set UserIndexLibres = New ColaConBloques
    
    ' Establecemos la cantidad
    Call UserIndexLibres.setCantidadElementosNodo(MaxUsers)

    ReDim BackupUserIndex(1 To MaxUsers) As Boolean
    
    ' Agrego los Index
    For i = MaxUsers To 1 Step -1
        Call UserIndexLibres.agregar(i)
        BackupUserIndex(i) = False
    Next
End Sub

Public Function obtenerUserIndexLibre() As Integer
    
    If UserIndexLibres.getCantidadElementos > 0 Then
        obtenerUserIndexLibre = UserIndexLibres.sacar
        
        BackupUserIndex(obtenerUserIndexLibre) = True ' Marco
        
        If obtenerUserIndexLibre > LastUser Then LastUser = obtenerUserIndexLibre
    Else
        obtenerUserIndexLibre = -1
    End If
    
    Debug.Print "TomoUser Index" & obtenerUserIndexLibre
    
End Function

Public Function liberarUserIndex(ByVal UserIndex As Integer) As Boolean

    Debug.Print "Libero User Index " & UserIndex
    ' ¿Es el mas alto?
    If UserIndex = LastUser Then
        ' Buscamos el Anterior
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser < 1 Then Exit Do
        Loop
    End If
    
    ' Liberamos
    If BackupUserIndex(UserIndex) = True Then
        Call UserIndexLibres.agregar(UserIndex)
        BackupUserIndex(UserIndex) = False
        liberarUserIndex = True
    Else
        LogError ("OJO!!! Estoy agregando por segunda vez un UserIndex:" & UserIndex)
        liberarUserIndex = False
    End If
    
End Function

Public Sub Cerrar_Usuario_Forzadamente(ByRef personaje As User)
 
    personaje.flags.Saliendo = eTipoSalida.SaliendoForsozamente
      
    ' Si el personaje es Game Master o esta en Zona Segura, no lo penalizo
    If personaje.flags.Privilegios > 0 Or MapInfo(personaje.pos.map).Pk = False Then
        personaje.Counters.Salir = 0
    Else
        personaje.Counters.Salir = GetTickCount + SEGUNDOS_PENA_DESCONEXION
    End If
        
End Sub

Public Sub Cerrar_Usuario(ByRef personaje As User)
 
    ' No se puede cerrar paralizado
    If personaje.flags.Paralizado = 1 Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(255), personaje.UserIndex, ToIndex:
        Exit Sub
    End If
            
    ' Mato los comercios seguros
    If personaje.ComUsu.DestUsu > 0 Then
        If UserList(personaje.ComUsu.DestUsu).flags.UserLogged Then
            If UserList(personaje.ComUsu.DestUsu).ComUsu.DestUsu = personaje.UserIndex Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(129), personaje.ComUsu.DestUsu, ToIndex, 0
                Call FinComerciarUsu(personaje.ComUsu.DestUsu)
            End If
        End If
        
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(1), personaje.UserIndex, ToIndex
        Call FinComerciarUsu(personaje.UserIndex)
    End If

    ' ¿Ya está saliendo?
    If Not personaje.flags.Saliendo = eTipoSalida.NoSaliendo Then Exit Sub
    
    ' Esta Todo Ok. Comienza el procedimiento de cerrado
    personaje.flags.Saliendo = eTipoSalida.SaliendoNaturalmente
        
    If personaje.flags.Privilegios > 0 Or MapInfo(personaje.pos.map).Pk = False Then
        personaje.Counters.Salir = 0
    ElseIf personaje.flags.Privilegios = 0 Or MapInfo(personaje.pos.map).Pk = True Then
        personaje.Counters.Salir = GetTickCount + SEGUNDOS_SALIR
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(23) & (SEGUNDOS_SALIR / 1000), personaje.UserIndex
    End If
    
End Sub
Function NameIndex(ByVal Name As String) As Integer
Dim UserIndex As Integer

If Name = "" Then
    NameIndex = 0
    Exit Function
End If
Name = UCase$(Replace(Name, "+", " "))
UserIndex = 1
Do Until UCase$(UserList(UserIndex).Name) = Name
    UserIndex = UserIndex + 1
    If UserIndex > LastUser Then
        NameIndex = 0
        Exit Function
    End If
Loop
NameIndex = UserIndex

End Function

Function IDIndex(ByVal id As Long) As Integer
Dim UserIndex As Integer

UserIndex = 1

Do Until UserList(UserIndex).id = id
    UserIndex = UserIndex + 1
    If UserIndex > LastUser Then
        IDIndex = 0
        Exit Function
    End If
Loop

IDIndex = UserIndex

End Function

Function ObtengoIndex_CharIndex(ByVal charIndex As Long) As Integer
Dim UserIndex As Integer
Dim loopUser As Integer


For loopUser = 1 To LastUser

    If UserList(loopUser).Char.charIndex = charIndex Then
        ObtengoIndex_CharIndex = loopUser
        Exit Function
    End If
 
Next loopUser

ObtengoIndex_CharIndex = 0

End Function

Function CheckForSameName(ByVal Name As String) As Boolean
'Controlo que no existan usuarios con el mismo nombre
Dim loopC As Integer
For loopC = 1 To LastUser
    If UserList(loopC).flags.UserLogged Then
        If UCase$(UserList(loopC).Name) = UCase$(Name) Then
            CheckForSameName = True
            Exit Function
        End If
    End If
Next loopC
CheckForSameName = False
End Function

Public Function personajeYaEstaLogueado(ByVal idPersonaje As Long) As Integer
'Controlo que no existan usuarios con el mismo nombre
Dim loopC As Integer
For loopC = 1 To LastUser
    If UserList(loopC).flags.UserLogged Then
        If UserList(loopC).id = idPersonaje Then
            personajeYaEstaLogueado = loopC
            Exit Function
        End If
    End If
Next loopC
personajeYaEstaLogueado = 0
End Function

