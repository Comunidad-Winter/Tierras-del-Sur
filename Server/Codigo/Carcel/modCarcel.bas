Attribute VB_Name = "modCarcel"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : PurgarPenas
' DateTime  : 18/02/2007 19:03
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub PurgarPenas()
Dim i As Integer
 
For i = 1 To LastUser
    If UserList(i).flags.UserLogged Then
        UserList(i).Counters.FotoDenuncia = 0
        
        If UserList(i).Counters.Pena > 0 Then
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                
                Dim carcel As carcel
                
                carcel = getCarcel(UserList(i))
                
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    Call SV_Desplazamientos.trasportarUsuarioOnline(UserList(i), carcel.salida.map, carcel.salida.x, carcel.salida.y, 3)
                    EnviarPaquete Paquetes.MensajeSimple, Chr$(32), i
                End If
        End If
    End If
Next i

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Encarcelar
' DateTime  : 18/02/2007 19:02
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Encarcelar(ByRef personaje As User, ByVal minutos As Long, Optional ByVal GmName As String = "")

    personaje.Counters.Pena = minutos
    
    Dim carcel As carcel
    Dim Posicion As Byte
    
    carcel = getCarcel(personaje)
    Posicion = RandomNumberInt(1, 3)
    
    Call SV_Desplazamientos.trasportarUsuarioOnline(personaje, carcel.posiciones(Posicion).map, carcel.posiciones(Posicion).x, carcel.posiciones(Posicion).y, 3)
    
    If GmName = "" Then
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(27) & minutos, personaje.UserIndex
    Else
        EnviarPaquete Paquetes.MensajeCompuesto, Chr$(28) & GmName & "," & minutos, personaje.UserIndex
    End If
    
    If personaje.flags.Trabajando = True Then
        Call DejarDeTrabajar(personaje)
    End If
 
End Sub
