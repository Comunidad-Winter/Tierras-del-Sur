Attribute VB_Name = "modBanco"
Option Explicit


Sub IniciarDeposito(ByVal UserIndex As Integer)
'Hacemos un Update del inventario del usuario
Call UpdateBanUserInv(True, UserIndex, 0)
'Atcualizamos el dinero
EnviarPaquete Paquetes.EnviarOro, Codify(UserList(UserIndex).Stats.GLD), UserIndex, ToIndex
'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
EnviarPaquete Paquetes.InitBanco, "", UserIndex, ToIndex
UserList(UserIndex).flags.Comerciando = True
End Sub

Sub SendBanObj(UserIndex As Integer, slot As Byte, Object As UserOBJ)
UserList(UserIndex).BancoInvent.Object(slot) = Object
If Object.ObjIndex > 0 Then
    EnviarPaquete Paquetes.EnviarBancoObj, Chr$(slot) & ITS(Object.ObjIndex) & ITS(Object.Amount) & ITS(ObjData(Object.ObjIndex).GrhIndex) & Chr$(ObjData(Object.ObjIndex).ObjType) & ITS(ObjData(Object.ObjIndex).MaxHIT) & ITS(ObjData(Object.ObjIndex).MinHIT) & ByteToString(ObjData(Object.ObjIndex).MinDef) & ByteToString(ObjData(Object.ObjIndex).MaxDef), UserIndex
Else
    EnviarPaquete Paquetes.EnviarBancoObj, Chr$(slot), UserIndex, ToIndex
End If
End Sub

Sub UpdateBanUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal slot As Byte)
Dim NullObj As UserOBJ
Dim loopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then
    'Actualiza el inventario
    If UserList(UserIndex).BancoInvent.Object(slot).ObjIndex > 0 Then
        Call SendBanObj(UserIndex, slot, UserList(UserIndex).BancoInvent.Object(slot))
    Else
        Call SendBanObj(UserIndex, slot, NullObj)
    End If
Else
'Actualiza todos los slots
    For loopC = 1 To MAX_BANCOINVENTORY_SLOTS
        'Actualiza el inventario
        If UserList(UserIndex).BancoInvent.Object(loopC).ObjIndex > 0 And UserList(UserIndex).BancoInvent.Object(loopC).Amount > 0 Then
            Call SendBanObj(UserIndex, loopC, UserList(UserIndex).BancoInvent.Object(loopC))
        Else
            Call SendBanObj(UserIndex, loopC, NullObj)
        End If
    Next loopC
End If
End Sub

Sub UserRetiraItem(ByVal UserIndex As Integer, ByVal i As Integer, ByVal cantidad As Integer)

If cantidad < 1 Then Exit Sub
'Call SendUserStatsBox(UserIndex) [Marce - Quien hizo esto?? Para que quiero saber
'el nivel de mi pj y cuanta vida tengo cuando quiero sacar un items del banco?]
If UserList(UserIndex).BancoInvent.Object(i).Amount > 0 Then
      If cantidad > UserList(UserIndex).BancoInvent.Object(i).Amount Then cantidad = UserList(UserIndex).BancoInvent.Object(i).Amount
      'Agregamos el obj que compro al inventario
      Call UserReciveObj(UserIndex, CInt(i), cantidad)
      'Actualizamos el banco
      Call UpdateBanUserInv(False, UserIndex, i)
      'Actualizamos la ventana de comercio
      Call UpdateVentanaBanco(i, 0, UserIndex)
End If

End Sub

Sub UserReciveObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal cantidad As Integer)
Dim slot As Integer
Dim obji As Integer


If UserList(UserIndex).BancoInvent.Object(ObjIndex).Amount <= 0 Then Exit Sub
obji = UserList(UserIndex).BancoInvent.Object(ObjIndex).ObjIndex
'¿Ya tiene un objeto de este tipo?
slot = 1
Do Until UserList(UserIndex).Invent.Object(slot).ObjIndex = obji And _
   UserList(UserIndex).Invent.Object(slot).Amount + cantidad <= MAX_INVENTORY_OBJS
    
    slot = slot + 1
    If slot > UserList(UserIndex).Stats.MaxItems Then
        Exit Do
    End If
Loop
'Sino se fija por un slot vacio
If slot > UserList(UserIndex).Stats.MaxItems Then
        slot = 1
        Do Until UserList(UserIndex).Invent.Object(slot).ObjIndex = 0
            slot = slot + 1

            If slot > UserList(UserIndex).Stats.MaxItems Then
                EnviarPaquete Paquetes.MensajeBoveda, Chr$(21), UserIndex, ToIndex, 0
                Exit Sub
            End If
        Loop
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If
'Mete el obj en el slot
If UserList(UserIndex).Invent.Object(slot).Amount + cantidad <= MAX_INVENTORY_OBJS Then
    'Menor que MAX_INV_OBJS
    UserList(UserIndex).Invent.Object(slot).ObjIndex = obji
    UserList(UserIndex).Invent.Object(slot).Amount = UserList(UserIndex).Invent.Object(slot).Amount + cantidad
    Call QuitarBancoInvItem(UserIndex, CByte(ObjIndex), cantidad)
    Call UpdateUserInv(False, UserIndex, slot)
Else
    EnviarPaquete Paquetes.MensajeBoveda, Chr$(21), UserIndex, ToIndex, 0
End If
End Sub

Sub QuitarBancoInvItem(ByVal UserIndex As Integer, ByVal slot As Byte, ByVal cantidad As Integer)
Dim ObjIndex As Integer
ObjIndex = UserList(UserIndex).BancoInvent.Object(slot).ObjIndex
    'Quita un Obj
       UserList(UserIndex).BancoInvent.Object(slot).Amount = UserList(UserIndex).BancoInvent.Object(slot).Amount - cantidad
        If UserList(UserIndex).BancoInvent.Object(slot).Amount <= 0 Then
            UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems - 1
            UserList(UserIndex).BancoInvent.Object(slot).ObjIndex = 0
            UserList(UserIndex).BancoInvent.Object(slot).Amount = 0
        End If
End Sub

Sub UpdateVentanaBanco(ByVal slot As Integer, ByVal NpcInv As Byte, ByVal UserIndex As Integer)
EnviarPaquete Paquetes.BancoOk, NpcInv, UserIndex, ToIndex
End Sub

Sub UserDepositaItem(ByVal UserIndex As Integer, ByVal Item As Integer, ByVal cantidad As Integer)

'El usuario deposita un item
' Call SendUserStatsBox(UserIndex)
If UserList(UserIndex).Invent.Object(Item).Amount > 0 And UserList(UserIndex).Invent.Object(Item).Equipped = 0 Then
    If cantidad > 0 And cantidad > UserList(UserIndex).Invent.Object(Item).Amount Then cantidad = UserList(UserIndex).Invent.Object(Item).Amount
    'Agregamos el obj que compro al inventario [ Hay que aflojar con el copiar y pegar hermano..]
    Call UserDejaObj(UserIndex, CInt(Item), cantidad)
    'Actualizamos el inventario del usuario
    Call UpdateUserInv(False, UserIndex, Item)
    'Actualizamos el inventario del banco [Esto lo saque de aca y lo puse en el userdejaobj]
    ' Call UpdateBanUserInv(True, UserIndex, 0)
    'Actualizamos la ventana del banco
    Call UpdateVentanaBanco(Item, 1, UserIndex)
End If
End Sub

Sub UserDejaObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal cantidad As Integer)
Dim slot As Integer
Dim obji As Integer

If cantidad < 1 Then Exit Sub
obji = UserList(UserIndex).Invent.Object(ObjIndex).ObjIndex
'¿Ya tiene un objeto de este tipo?
slot = 1
Do Until UserList(UserIndex).BancoInvent.Object(slot).ObjIndex = obji And _
         UserList(UserIndex).BancoInvent.Object(slot).Amount + cantidad <= MAX_INVENTORY_OBJS
            slot = slot + 1
            If slot > MAX_BANCOINVENTORY_SLOTS Then
                Exit Do
            End If
Loop
'Sino se fija por un slot vacio antes del slot devuelto
If slot > MAX_BANCOINVENTORY_SLOTS Then
        slot = 1
        Do Until UserList(UserIndex).BancoInvent.Object(slot).ObjIndex = 0
            slot = slot + 1
            If slot > MAX_BANCOINVENTORY_SLOTS Then
                EnviarPaquete Paquetes.MensajeBoveda, Chr$(211), UserIndex, ToIndex, 0
                Exit Sub
                Exit Do
            End If
        Loop
        If slot <= MAX_BANCOINVENTORY_SLOTS Then UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems + 1
End If
If slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido
    'Mete el obj en el slot
    If UserList(UserIndex).BancoInvent.Object(slot).Amount + cantidad <= MAX_INVENTORY_OBJS Then
        'Menor que MAX_INV_OBJS
        UserList(UserIndex).BancoInvent.Object(slot).ObjIndex = obji
        UserList(UserIndex).BancoInvent.Object(slot).Amount = UserList(UserIndex).BancoInvent.Object(slot).Amount + cantidad
        Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), cantidad)
        Call UpdateBanUserInv(False, UserIndex, slot)
    Else
        EnviarPaquete Paquetes.MensajeBoveda, Chr$(212), UserIndex, ToIndex, 0
    End If
Else
    Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), cantidad)
End If
End Sub

Sub SendUserBovedaTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
Dim j As Integer

EnviarPaquete Paquetes.mensajeinfo, UserList(UserIndex).Name, sendIndex, ToIndex
EnviarPaquete Paquetes.mensajeinfo, "Tiene " & UserList(UserIndex).BancoInvent.NroItems & " objetos.", sendIndex, ToIndex

For j = 1 To MAX_BANCOINVENTORY_SLOTS
    If UserList(UserIndex).BancoInvent.Object(j).ObjIndex > 0 Then
        EnviarPaquete Paquetes.mensajeinfo, "Objeto " & j & " " & ObjData(UserList(UserIndex).BancoInvent.Object(j).ObjIndex).Name & " Cantidad:" & UserList(UserIndex).BancoInvent.Object(j).Amount, sendIndex, ToIndex
    End If
Next

End Sub
