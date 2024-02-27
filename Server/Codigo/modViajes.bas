Attribute VB_Name = "modViajes"
Public Type Pasaje

Posicion As WorldPos
Continente As Integer
        
End Type

Public Pasajes() As Pasaje

Public Sub LoadPasajes()

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Pasajes."

Dim Pasaje As Integer
Dim Leer As New clsLeerInis
Dim NumeroPasajes As Integer
Dim ln As String

Leer.Abrir DatPath & "Pasajes.dat"

'obtiene el numero de Pasajes

NumeroPasajes = val(Leer.DarValor("INIT", "NumeroPasajes"))
ReDim Pasajes(1 To NumeroPasajes) As Pasaje

frmCargando.cargar.min = 0
frmCargando.cargar.max = NumeroPasajes
frmCargando.cargar.value = 0

'Llena la lista
For Pasaje = 1 To NumeroPasajes
    Pasajes(Pasaje).Posicion.map = Leer.DarValor("PASAJE" & Pasaje, "MAPA")
    Pasajes(Pasaje).Posicion.x = Leer.DarValor("PASAJE" & Pasaje, "X")
    Pasajes(Pasaje).Posicion.y = Leer.DarValor("PASAJE" & Pasaje, "Y")
    Pasajes(Pasaje).Continente = Leer.DarValor("PASAJE" & Pasaje, "CONTINENTE")
Next
 
End Sub

Public Sub Viajar(ByVal UserIndex As Integer, ByVal slot As Byte, obj As ObjData)

Dim numeroPasaje As Integer

numeroPasaje = obj.subTipo

If UserList(UserIndex).flags.Muerto = 1 Then
    EnviarPaquete Paquetes.mensajeinfo, "Solo los vivos pueden utilizar el Sistema de Viajes.", UserIndex, ToIndex
    Exit Sub
End If

If Not UserList(UserIndex).evento Is Nothing Then
    EnviarPaquete Paquetes.mensajeinfo, "No puedes usar un pasaje durante un evento.", UserIndex, ToIndex
    Exit Sub
End If

If UserList(UserIndex).pos.map = Pasajes(numeroPasaje).Posicion.map Then
    EnviarPaquete Paquetes.mensajeinfo, "¡¡Ya te encuentras en esa ciudad!!.", UserIndex, ToIndex
    Exit Sub
End If

If MapInfo(UserList(UserIndex).pos.map).Pk = 1 Then
    EnviarPaquete Paquetes.mensajeinfo, "Debes estar en zona segura para utilizar el Sistema de Viajes", UserIndex, ToIndex
    Exit Sub
Else
    WarpUserChar UserIndex, Pasajes(numeroPasaje).Posicion.map, Pasajes(numeroPasaje).Posicion.x, Pasajes(numeroPasaje).Posicion.y, True
    ' Quitamos el Pasaje
    Call InvUsuario.QuitarUserInvItem(UserIndex, slot, 1)
    Call InvUsuario.UpdateUserInv(False, UserIndex, slot)
End If

End Sub
