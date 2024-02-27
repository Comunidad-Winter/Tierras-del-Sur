Attribute VB_Name = "modNuevoTimer"
Option Explicit
'Funciones utilitarias

'Used for the 64-bit timer
Private GetSystemTimeOffset As Currency
Private Declare Sub GetSystemTime Lib "kernel32.dll" Alias "GetSystemTimeAsFileTime" (ByRef lpSystemTimeAsFileTime As Currency)

Public Sub InitTimeGetTime()
'*****************************************************************
'Gets the offset time for the timer so we can start at 0 instead of
'the returned system time, allowing us to not have a time roll-over until
'the program is running for 25 days
'More info: http://www.vbgore.com/GameServer.General.InitTimeGetTime
'*****************************************************************

    'Get the initial time
    GetSystemTime GetSystemTimeOffset
    
    'Subtract to the loeHeading.WEST value so we start at -(2^31) instead of 0
    'Doubles the time we can run the program (50 days instead of 25) before a rollover
    'This isn't done because of the way vbGORE already handles some timers
    'Not like it is needed for the client, but some servers may last 25 days without a reset
    'Maybe if it ever becomes a problem, then some day, someone may fix it... ;)
    'GetSystemTimeOffset = GetSystemTimeOffset + (2 ^ 31) - 1

End Sub

Private Function timeGetTime() As Long
'*****************************************************************
'Grabs the time from the 64-bit system timer and returns it in 32-bit
'after calculating it with the offset - allows us to have the
'"no roll-over" advantage of 64-bit timers with the RAM usage of 32-bit
'though we limit things slightly, so the rollover still happens, but after 25 days
'More info: http://www.vbgore.com/GameServer.General.timeGetTime
'*****************************************************************
Dim CurrentTime As Currency

    'Grab the current time (we have to pass a variable ByRef instead of a function return like the other timers)
    GetSystemTime CurrentTime
    
    'Calculate the difference between the 64-bit times, return as a 32-bit time
    timeGetTime = CurrentTime - GetSystemTimeOffset

End Function

'Establece los intervalos que le corresponde al usuario segun su clase

Public Function establecerIntervalos(ByRef Usuario As User)

Dim multiplicador As Integer
'Paso de segundos a milisegundos con esta variables
'Si es menor a 1000, se reduce el intervalo. Esto para dar un poco de tolerancia si hay lag
'y llegan dos paquetes mas juntos de lo que deberian

multiplicador = 900

Select Case Usuario.clase

    Case eClases.Guerrero
        Usuario.intervalos.UsarClick = IntervaloClickG * multiplicador
        Usuario.intervalos.Flecha = IntervaloFlechaG * multiplicador
        Usuario.intervalos.Golpe = IntervaloGolpeG * multiplicador
        Usuario.intervalos.Magia = IntervaloMagiaG * multiplicador
        Usuario.intervalos.UsarU = IntervaloUG * multiplicador
    Case eClases.Cazador
        Usuario.intervalos.UsarClick = IntervaloClickC * multiplicador
        Usuario.intervalos.Flecha = IntervaloFlechaC * multiplicador
        Usuario.intervalos.Golpe = IntervaloGolpeC * multiplicador
        Usuario.intervalos.Magia = IntervaloMagiaC * multiplicador
        Usuario.intervalos.UsarU = IntervalouC * multiplicador
    Case Else ' EL resto de las clases
        Usuario.intervalos.UsarClick = IntervaloClick * multiplicador
        Usuario.intervalos.Flecha = IntervaloFlecha * multiplicador
        Usuario.intervalos.Golpe = IntervaloGolpe * multiplicador
        Usuario.intervalos.Magia = IntervaloMagia * multiplicador
        Usuario.intervalos.UsarU = IntervaloU * multiplicador
        
End Select


End Function



' Las siguientes funciones devuelven TRUE o FALSE si el intervalo
' permite hacerlo. Si devuelve TRUE, setean automaticamente el
' timer para que no se pueda hacer la accion hasta el nuevo ciclo.
'
' Lanzamientos de magia
Public Function IntervaloPermiteLanzarSpell(ByRef Usuario As User, Optional ByVal Actualizar As Boolean = True) As Boolean

Dim TActual As Long

TActual = timeGetTime()

If Usuario.Counters.TimerLanzarSpell + Usuario.intervalos.Magia < TActual Then

    If Actualizar Then
       Usuario.Counters.TimerLanzarSpell = TActual
    End If
    
    IntervaloPermiteLanzarSpell = True
Else
    IntervaloPermiteLanzarSpell = False
End If
End Function

' ATAQUE CUERPO A CUERPO
Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long
TActual = GetTickCount() And &H7FFFFFFF
If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= 40 * IntervaloUserPuedeAtacar Then
    If Actualizar Then UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
    IntervaloPermiteAtacar = True
Else
    IntervaloPermiteAtacar = False
End If
End Function

' USAR OBJETOS
Public Function IntervaloPermiteUsarClic(ByVal UserIndex As Integer) As Boolean
Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - UserList(UserIndex).Counters.TimerUsarClic >= UserList(UserIndex).intervalos.UsarClick Then
    UserList(UserIndex).Counters.TimerUsarClic = TActual
    IntervaloPermiteUsarClic = True
Else
    IntervaloPermiteUsarClic = False
End If

End Function

Public Function IntervaloPermiteUsarU(ByVal UserIndex As Integer) As Boolean
Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - UserList(UserIndex).Counters.TimerUsarU >= UserList(UserIndex).intervalos.UsarU Then
    UserList(UserIndex).Counters.TimerUsarU = TActual
    IntervaloPermiteUsarU = True
Else
    IntervaloPermiteUsarU = False
End If
End Function

