Attribute VB_Name = "Comercio"
Option Explicit
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%          MODULO DE COMERCIO NPC-USER              %%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


Public Sub personajeVendeObjetoACriatura(ByRef personaje As User, anexo As String)

    Dim npcIndex As Integer
    Dim cantidad As Long
    Dim slot As Integer
    Dim ObjetoIndex As Integer
    
    ' Datos del anexo
    slot = Asc(Left$(anexo, 1))
    cantidad = DeCodify(Right$(anexo, Len(anexo) - 1))
    npcIndex = personaje.flags.TargetNPC
    
    If personaje.flags.Muerto Then Exit Sub
    
    ' ¿Tiene a una criatura seleccionada?
    If npcIndex = 0 Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(4), personaje.UserIndex
        Exit Sub
    End If
        
    ' ¿Es un comerciante?
    If NpcList(npcIndex).Comercia = 0 Then
        EnviarPaquete Paquetes.DescNpc2, Chr$(3) & ITS(NpcList(npcIndex).Char.charIndex), personaje.UserIndex
        Exit Sub
    End If
        
    '¿Esta demasiado lejos?
    If distancia(personaje.pos, NpcList(npcIndex).pos) > 5 Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(7), personaje.UserIndex
        Exit Sub
    End If
    
    ' ¿Slot Valido?
    If slot < 0 Or slot > personaje.Stats.MaxItems Then
        LogError ("Error en personaje vende objeto a criatura. Slot invalido: " & slot)
        Exit Sub
    End If
    
    ' ¿Hay algo en el Slot?
    ObjetoIndex = personaje.Invent.Object(slot).ObjIndex

    If ObjetoIndex = 0 Then
        EnviarPaquete Paquetes.mensajeinfo, "Debés seleccionar el objeto que deseas vender.", personaje.UserIndex, ToIndex
        Exit Sub
    End If

    ' No se pueden vender monedas de oro
    If ObjetoIndex = iORO Then
        EnviarPaquete Paquetes.mensajeinfo, "No se pueden vender monedas de oro.", personaje.UserIndex, ToIndex
        Exit Sub
    End If
    
    ' No se pueden vender objetos equipados
    If personaje.Invent.Object(slot).Equipped = 1 Then
        EnviarPaquete Paquetes.mensajeinfo, "Debes desequiparte el objeto antes de venderlo.", personaje.UserIndex, ToIndex
        Exit Sub
    End If
    
    ' No objetos newbies
    If ObjData(ObjetoIndex).Newbie = 1 Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(34), personaje.UserIndex
        Exit Sub
    End If

    ' No se pueden vender objetos faccionarios
    If modObjeto.isFaccionario(ObjData(ObjetoIndex)) Then
        EnviarPaquete Paquetes.mensajeinfo, "No puedes vender items faccionarios.", personaje.UserIndex, ToIndex
        Exit Sub
    End If
    
    'No se puede tirar la armadura del dragon
    If ObjetoIndex = Objetos_Constantes.ARMADURA_DRAGON_E Or _
            ObjetoIndex = Objetos_Constantes.ARMADURA_DRAGON_H Or _
            ObjetoIndex = Objetos_Constantes.ARMADURA_DRAGON_M Then
        EnviarPaquete Paquetes.mensajeinfo, "No puedes tirar la Armadura del Dragón, el enfado del Dragón te desterrará de estas tierras.", personaje.UserIndex
        Exit Sub
    End If
    
    ' Cantidad logica?
    If cantidad < 0 Or cantidad > InvUsuario.MAX_OBJETOS_X_SLOT Then Exit Sub

    ' Ajustamos la cantidad
    If cantidad > personaje.Invent.Object(slot).Amount Then cantidad = personaje.Invent.Object(slot).Amount

    ' La criatura compra efectivamente el objeto
    Call NpcCompraObj(personaje, slot, cantidad, 1)

    ' Actualizaciones

    ' - Inventario y oro
    Call UpdateUserInv(False, personaje.UserIndex, slot)
    
    EnviarPaquete Paquetes.EnviarOro, Codify(personaje.Stats.GLD), personaje.UserIndex, ToIndex
End Sub
'---------------------------------------------------------------------------------------
' Procedure : UserCompraObj
' DateTime  : 18/02/2007 19:08
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub UserCompraObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal npcIndex As Integer, ByVal cantidad As Integer)
Dim infla As Long
Dim Descuento As Single
Dim unidad As Long, monto As Long
Dim slot As Integer
Dim obji As Integer

If (NpcList(UserList(UserIndex).flags.TargetNPC).Invent.Object(ObjIndex).Amount <= 0) Then Exit Sub
obji = NpcList(UserList(UserIndex).flags.TargetNPC).Invent.Object(ObjIndex).ObjIndex

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
                EnviarPaquete Paquetes.MensajeSimple, Chr$(21), UserIndex
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
    'Le sustraemos el valor en oro del obj comprado
    infla = (NpcList(npcIndex).Inflacion * ObjData(obji).valor) / 100
    Descuento = Comercio.Descuento(UserList(UserIndex))

    unidad = ((ObjData(NpcList(npcIndex).Invent.Object(ObjIndex).ObjIndex).valor + infla) * Descuento)
    
    If unidad = 0 Then
        unidad = 1
    End If
    
    monto = unidad * cantidad
    
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - monto
    'Actualizaamos solo el slot , no todo el inventario como antes 'Marche
     Call UpdateUserInv(False, UserIndex, slot)
    'tal vez suba el skill comerciar ;-)
    Call SubirSkill(UserIndex, Comerciar)
    If ObjData(obji).ObjType = OBJTYPE_LLAVES Then Call logVentaCasa(UserList(UserIndex).Name & " compro " & ObjData(obji).Name)
'    If UserList(UserIndex).Stats.GLD < 0 Then UserList(UserIndex).Stats.GLD = 0
    Call QuitarNpcInvItem(UserList(UserIndex).flags.TargetNPC, CByte(ObjIndex), cantidad)
Else
    EnviarPaquete Paquetes.MensajeSimple, Chr$(21), UserIndex
End If
End Sub

Sub NpcCompraObj(personaje As User, ByVal SlotUsuario As Byte, ByVal cantidad As Integer, Optional Actualizar As Byte)

Dim SlotGuardado As Integer
Dim ObjetoIndex As Integer
Dim npcIndex As Integer
Dim monto As Long

npcIndex = personaje.flags.TargetNPC
ObjetoIndex = personaje.Invent.Object(SlotUsuario).ObjIndex

'¿Son los items con los que comercia el npc?
If NpcList(npcIndex).TipoItems <> OBJTYPE_CUALQUIERA Then
    If NpcList(npcIndex).TipoItems <> ObjData(ObjetoIndex).ObjType Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(35), personaje.UserIndex, ToIndex
        Exit Sub
    End If
End If

'¿Ya tiene un objeto de este tipo?
SlotGuardado = 1
Do Until NpcList(npcIndex).Invent.Object(SlotGuardado).ObjIndex = ObjetoIndex And _
         NpcList(npcIndex).Invent.Object(SlotGuardado).Amount + cantidad <= MAX_INVENTORY_OBJS
        
        SlotGuardado = SlotGuardado + 1
        
        If SlotGuardado > MAX_INVENTORY_SLOTS Then Exit Do
Loop


'Sino se fija por un slot vacio antes del slot devuelto
If SlotGuardado > MAX_INVENTORY_SLOTS Then
    SlotGuardado = 1
        
    Do Until NpcList(npcIndex).Invent.Object(SlotGuardado).ObjIndex = 0
        SlotGuardado = SlotGuardado + 1
            
        If SlotGuardado > MAX_INVENTORY_SLOTS Then Exit Do
    Loop
        
    If SlotGuardado <= MAX_INVENTORY_SLOTS Then NpcList(npcIndex).Invent.NroItems = NpcList(npcIndex).Invent.NroItems + 1
End If

' Le quitamos el objeto la usuario
Call QuitarUserInvItem(personaje.UserIndex, SlotUsuario, cantidad)

' Si encontre un slot válido lo guardo
If SlotGuardado <= MAX_INVENTORY_SLOTS Then
    NpcList(npcIndex).Invent.Object(SlotGuardado).ObjIndex = ObjetoIndex
    NpcList(npcIndex).Invent.Object(SlotGuardado).Amount = NpcList(npcIndex).Invent.Object(SlotGuardado).Amount + cantidad
End If

'Le sumamos al user el valor en oro del obj vendido
monto = (ObjData(ObjetoIndex).valor \ 3) * cantidad
Call AddtoVar(personaje.Stats.GLD, monto, MAXORO)

'Tal vez suba el skill comerciar ;-)
Call SubirSkill(personaje.UserIndex, Comerciar)

'Actualizamos ventana del inventario
If Actualizar = 1 And SlotGuardado <= MAX_INVENTORY_SLOTS Then
    Call EnviarNpcInv(personaje.UserIndex, npcIndex, SlotGuardado)
End If

End Sub

Sub IniciarCOmercioNPC(ByVal UserIndex As Integer)
'Mandamos el Inventario
Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
'Hacemos un Update del inventario del usuario
Call UpdateUserInv(True, UserIndex, 0)
'Atcualizamos el dinero
EnviarPaquete Paquetes.EnviarOro, Codify(UserList(UserIndex).Stats.GLD), UserIndex, ToIndex
'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
UserList(UserIndex).flags.Comerciando = True
EnviarPaquete Paquetes.pIniciarComercioNPC, "", UserIndex, ToIndex, 0
End Sub

Sub NPCVentaItem(ByVal UserIndex As Integer, ByVal i As Integer, ByVal cantidad As Integer, ByVal npcIndex As Integer)

Dim infla As Long
Dim val As Long
Dim desc As Single

'Quiere comprar menos de una unidad?
If cantidad < 1 Then Exit Sub

If i > MAX_INVENTORY_SLOTS Then
    EnviarPaquete Paquetes.MensajeFight, "Posible intento de romper el sistema de comercio. Usuario: " & UserList(UserIndex).Name, UserIndex, ToAdmins
    Exit Sub
End If

If cantidad > MAX_INVENTORY_OBJS Then
    EnviarPaquete Paquetes.MensajeFight, "Posible intento de romper el sistema de comercio. Usuario: " & UserList(UserIndex).Name, UserIndex, ToAdmins
    Exit Sub
End If

'El npc tiene algo en el slot indicado?
If NpcList(npcIndex).Invent.Object(i).ObjIndex = 0 Then Exit Sub

'Calculamos el valor unitario
infla = (NpcList(npcIndex).Inflacion * ObjData(NpcList(npcIndex).Invent.Object(i).ObjIndex).valor) / 100

desc = Comercio.Descuento(UserList(UserIndex))

val = (ObjData(NpcList(npcIndex).Invent.Object(i).ObjIndex).valor + infla) * desc

If val = 0 Then
    val = 1
End If

If UserList(UserIndex).Stats.GLD >= (val * cantidad) Then

       If NpcList(UserList(UserIndex).flags.TargetNPC).Invent.Object(i).Amount > 0 Then
            If cantidad > NpcList(UserList(UserIndex).flags.TargetNPC).Invent.Object(i).Amount Then cantidad = NpcList(UserList(UserIndex).flags.TargetNPC).Invent.Object(i).Amount
            
            'Agregamos el obj que compro al inventario
            Call UserCompraObj(UserIndex, CInt(i), UserList(UserIndex).flags.TargetNPC, cantidad)
            
            'Actualizamos el oro
            EnviarPaquete Paquetes.EnviarOro, Codify(UserList(UserIndex).Stats.GLD), UserIndex, ToIndex, 0
            
            'Actualizamos la ventana de comercio
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC, i)
       End If
Else
    EnviarPaquete Paquetes.MensajeSimple, Chr$(37), UserIndex
    Exit Sub
End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Descuento
' DateTime  : 18/02/2007 19:10
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Descuento(personaje As User) As Single
'Establece el descuento en funcion del skill comercio
Dim PtsComercio As Integer

PtsComercio = personaje.Stats.UserSkills(Comerciar)

Select Case PtsComercio
    Case 0:
        Descuento = 1
    Case 1 To 30
        Descuento = 0.9
    Case 31 To 60
        Descuento = 0.8
    Case 61 To 90
        Descuento = 0.7
    Case 91 To 99
        Descuento = 0.6
    Case 100
        Descuento = 0.5
End Select

End Function

'---------------------------------------------------------------------------------------
' Procedure : EnviarNpcInv
' DateTime  : 18/02/2007 19:10
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub EnviarNpcInv(ByVal UserIndex As Integer, ByVal npcIndex As Integer, Optional slot As Integer)
'Enviamos el inventario del npc con el cual el user va a comerciar...
Dim i As Integer
Dim infla As Long
Dim desc As Single
Dim val As Long

desc = Descuento(UserList(UserIndex))

If slot > 0 Then
      If NpcList(npcIndex).Invent.Object(slot).ObjIndex > 0 Then
        'Calculamos el porc de inflacion del npc
        infla = (NpcList(npcIndex).Inflacion * ObjData(NpcList(npcIndex).Invent.Object(slot).ObjIndex).valor) / 100
        
        val = (ObjData(NpcList(npcIndex).Invent.Object(slot).ObjIndex).valor + infla) * desc
        
        If val = 0 Then
            val = 1
        End If
        
        EnviarPaquete Paquetes.pEnviarNpcInvBySlot, _
        ByteToString(slot) & Chr$(ObjData(NpcList(npcIndex).Invent.Object(slot).ObjIndex).ObjType) & _
        ITS(NpcList(npcIndex).Invent.Object(slot).Amount) & _
        ITS(ObjData(NpcList(npcIndex).Invent.Object(slot).ObjIndex).GrhIndex) & _
        ITS(NpcList(npcIndex).Invent.Object(slot).ObjIndex) & _
        ITS(ObjData(NpcList(npcIndex).Invent.Object(slot).ObjIndex).MaxHIT) & _
        ITS(ObjData(NpcList(npcIndex).Invent.Object(slot).ObjIndex).MinHIT) & _
        ByteToString(ObjData(NpcList(npcIndex).Invent.Object(slot).ObjIndex).MinDef) & _
        ByteToString(ObjData(NpcList(npcIndex).Invent.Object(slot).ObjIndex).MaxDef) & _
        Codify(val), UserIndex
  Else
        EnviarPaquete Paquetes.pEnviarNpcInvBySlot, ByteToString(slot) & "", UserIndex
  End If
Else
    For i = 1 To MAX_INVENTORY_SLOTS
    Dim cadena As String
      If NpcList(npcIndex).Invent.Object(i).ObjIndex > 0 Then
            'Calculamos el porc de inflacion del npc
            infla = (NpcList(npcIndex).Inflacion * ObjData(NpcList(npcIndex).Invent.Object(i).ObjIndex).valor) / 100
            val = (ObjData(NpcList(npcIndex).Invent.Object(i).ObjIndex).valor + infla) * desc
            
            If val = 0 Then
                val = 1
            End If
            
            cadena = cadena & _
            Chr$(ObjData(NpcList(npcIndex).Invent.Object(i).ObjIndex).ObjType) & _
            ITS(NpcList(npcIndex).Invent.Object(i).Amount) & _
            ITS(ObjData(NpcList(npcIndex).Invent.Object(i).ObjIndex).GrhIndex) & _
            ITS(NpcList(npcIndex).Invent.Object(i).ObjIndex) & _
            ITS(ObjData(NpcList(npcIndex).Invent.Object(i).ObjIndex).MaxHIT) & _
            ITS(ObjData(NpcList(npcIndex).Invent.Object(i).ObjIndex).MinHIT) & _
            ByteToString(ObjData(NpcList(npcIndex).Invent.Object(i).ObjIndex).MinDef) & _
            ByteToString(ObjData(NpcList(npcIndex).Invent.Object(i).ObjIndex).MaxDef) & _
            LongToString(val)
      Else
            cadena = cadena & "ÿ"
      End If
    Next
    EnviarPaquete Paquetes.pNpcInventory, cadena, UserIndex
End If

End Sub



Public Sub ActualizarPrecios(UserIndex As Integer, npcIndex As Integer)
'Marche
'Actualiza los descuentos cuando estas comerciando
Dim infla, val As Long
Dim desc As Single
Dim i As Byte

desc = Descuento(UserList(UserIndex))

For i = 1 To MAX_INVENTORY_SLOTS
Dim cadena As String
  If NpcList(npcIndex).Invent.Object(i).ObjIndex > 0 Then
        'Calculamos el porc de inflacion del npc
        infla = (NpcList(npcIndex).Inflacion * ObjData(NpcList(npcIndex).Invent.Object(i).ObjIndex).valor) / 100
        val = (ObjData(NpcList(npcIndex).Invent.Object(i).ObjIndex).valor + infla) * desc
        cadena = cadena & LongToString(val)
  Else
        cadena = cadena & "X"
  End If
Next

EnviarPaquete Paquetes.pNpcActualizarPrecios, cadena, UserIndex

End Sub
