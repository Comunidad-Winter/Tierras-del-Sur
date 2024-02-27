Attribute VB_Name = "modCofres"
Option Explicit

Type CofreDrop
ObjIndex As Integer
Amount As Integer
Probability As Integer
    
End Type

Public Type Cofre
DropItem(1 To 5) As CofreDrop
NroObjetos As Integer
        
End Type


Public Cofres() As Cofre

Public Sub LoadCofres()

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Cofres."

Dim Cofre As Integer
Dim Leer As New clsLeerInis
Dim i As Integer
Dim loopA As Integer
Dim NumeroCofres As Integer
Dim CantidadDrop As Integer
Dim ln As String

Leer.Abrir DatPath & "Cofres.dat"

'obtiene el numero de Cofres

NumeroCofres = val(Leer.DarValor("INIT", "NumeroCofres"))
ReDim Cofres(1 To NumeroCofres) As Cofre

frmCargando.cargar.min = 0
frmCargando.cargar.max = NumeroCofres
frmCargando.cargar.value = 0

'Llena la lista
For Cofre = 1 To NumeroCofres
        Cofres(Cofre).NroObjetos = val(Leer.DarValor("COFRE" & Cofre, "NroObjetos"))
        CantidadDrop = Cofres(Cofre).NroObjetos
        ReDim DropItem(1 To CantidadDrop) As CofreDrop
        For loopA = 1 To Cofres(Cofre).NroObjetos
            ln = Leer.DarValor("COFRE" & Cofre, "Obj" & loopA)
            Cofres(Cofre).DropItem(loopA).ObjIndex = val(ReadField(1, ln, 45))
            Cofres(Cofre).DropItem(loopA).Amount = val(ReadField(2, ln, 45))
            Cofres(Cofre).DropItem(loopA).Probability = val(ReadField(3, ln, 45))
        Next loopA
        
Next
 
End Sub

Sub DropeoCofre(ByVal UserIndex As Integer, ByVal slot As Byte, obj As ObjData, MiObj As obj)

Dim numeroCofre As Integer
Dim entregoItem As Boolean
Dim i As Integer

numeroCofre = obj.subTipo
entregoItem = False
i = 1

' Loop para que entregue 1 solo item.
Do
    If Cofres(numeroCofre).DropItem(i).Probability >= RandomNumber(1, 100) Then
        MiObj.Amount = Cofres(numeroCofre).DropItem(i).Amount
        MiObj.ObjIndex = Cofres(numeroCofre).DropItem(i).ObjIndex
        entregoItem = True
    End If
i = i + 1
Loop While entregoItem = False And i <= Cofres(numeroCofre).NroObjetos

If (i > Cofres(numeroCofre).NroObjetos And entregoItem = False) Then
    EnviarPaquete Paquetes.mensajeinfo, "Ocurrio un problema, comunicate con un Administrador.", UserIndex, ToIndex
    Call LogError("No se le entregó ningún item al usuario " & UserList(UserIndex).Name & " al abrir el cofre " & numeroCofre & ".")
    Exit Sub
End If

' Quitamos el cofre
Call InvUsuario.QuitarUserInvItem(UserIndex, slot, 1)
Call InvUsuario.UpdateUserInv(False, UserIndex, slot)

' Metemos el item en el inventario. En caso de no tener lugar en el inventario, lo arroja al piso.
If Not MeterItemEnInventario(UserIndex, MiObj) Then
    Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)
End If

EnviarPaquete Paquetes.mensajeinfo, "El cofre te ha entregado un item!", UserIndex, ToIndex

End Sub


