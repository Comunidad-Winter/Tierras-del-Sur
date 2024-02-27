Attribute VB_Name = "modPergamino"
Public Type Pergamino

Posicion As WorldPos
Continente As Integer
        
End Type

Public Pergaminos() As Pergamino

Public Sub LoadPergaminos()

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Pergaminos."

Dim Pergamino As Integer
Dim Leer As New clsLeerInis
Dim NumeroPergaminos As Integer
Dim ln As String

Leer.Abrir DatPath & "Pergaminos.dat"

'obtiene el numero de Pergaminos

NumeroPergaminos = val(Leer.DarValor("INIT", "NumeroPergaminos"))
ReDim Pergaminos(1 To NumeroPergaminos) As Pergamino

frmCargando.cargar.min = 0
frmCargando.cargar.max = NumeroPergaminos
frmCargando.cargar.value = 0

'Llena la lista
For Pergamino = 1 To NumeroPergaminos
    Pergaminos(Pergamino).Posicion.map = Leer.DarValor("PERGAMINO" & Pergamino, "MAPA")
    Pergaminos(Pergamino).Posicion.x = Leer.DarValor("PERGAMINO" & Pergamino, "X")
    Pergaminos(Pergamino).Posicion.y = Leer.DarValor("PERGAMINO" & Pergamino, "Y")
    Pergaminos(Pergamino).Continente = Leer.DarValor("PERGAMINO" & Pergamino, "CONTINENTE")
Next
 
End Sub

Public Sub Transportarse(ByVal UserIndex As Integer, ByVal slot As Byte, obj As ObjData)

Dim numeroPergamino As Integer

numeroPergamino = obj.subTipo

If UserList(UserIndex).flags.Muerto = 0 Then
    EnviarPaquete Paquetes.mensajeinfo, "Solo los muertos pueden utilizar el sistema de Pergaminos.", UserIndex, ToIndex
    Exit Sub
End If

If Not UserList(UserIndex).evento Is Nothing Then
    EnviarPaquete Paquetes.mensajeinfo, "No puedes usar un pergamino durante un evento.", UserIndex, ToIndex
    Exit Sub
End If

If MapInfo(UserList(UserIndex).pos.map).Pk = 0 Then
    EnviarPaquete Paquetes.mensajeinfo, "No puedes usar un pergamino dentro de una zona segura.", UserIndex, ToIndex
    Exit Sub
End If

If Pergaminos(numeroPergamino).Continente > 0 And MapInfo(UserList(UserIndex).pos.map).Continente <> Pergaminos(numeroPergamino).Continente Then
    EnviarPaquete Paquetes.mensajeinfo, "Los pergaminos solo pueden ser utilizados en el mismo contienente en el que se encuentra la ciudad a la que se desea transportar.", UserIndex, ToIndex
    Exit Sub
Else
    WarpUserChar UserIndex, Pergaminos(numeroPergamino).Posicion.map, Pergaminos(numeroPergamino).Posicion.x, Pergaminos(numeroPergamino).Posicion.y, True
    ' Quitamos el Pergamino
    Call InvUsuario.QuitarUserInvItem(UserIndex, slot, 1)
    Call InvUsuario.UpdateUserInv(False, UserIndex, slot)
End If

End Sub
