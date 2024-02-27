Attribute VB_Name = "InvUsuario"
Option Explicit

Public Const MAX_OBJETOS_X_SLOT = 10000

Function TieneObjetos(ByVal ItemIndex As Integer, ByVal cant As Long, personaje As User) As Boolean

Dim i As Integer
Dim Total As Long

If personaje.Invent.BarcoSlot > 1 And personaje.clase = eClases.Pirata Then

    For i = 1 To 25
        If personaje.Invent.Object(i).ObjIndex = ItemIndex Then
            Total = Total + personaje.Invent.Object(i).Amount
        End If
    Next i

Else

    For i = 1 To MAX_INVENTORY_SLOTS
        If personaje.Invent.Object(i).ObjIndex = ItemIndex Then
            Total = Total + personaje.Invent.Object(i).Amount
        End If
    Next i

End If

If cant <= Total Then
    TieneObjetos = True
    Exit Function
End If

End Function

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean
'17/09/02
'Agregue que la función se asegure que el objeto no es un barco
Dim i As Integer
Dim ObjIndex As Integer


For i = 1 To UserList(UserIndex).Stats.MaxItems
    ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    If ObjIndex > 0 Then
            If (ObjData(ObjIndex).ObjType <> OBJTYPE_LLAVES And _
                ObjData(ObjIndex).ObjType <> OBJTYPE_BARCOS) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
    End If
Next i
End Function

Function ClasePuedeUsarItem(ByRef personaje As User, ByRef objeto As ObjData) As Boolean
If objeto.clasesPermitidas = 0 Then
    ClasePuedeUsarItem = True
    Exit Function
End If

If personaje.flags.Privilegios > 0 Then
    ClasePuedeUsarItem = True
    Exit Function
End If

ClasePuedeUsarItem = (objeto.clasesPermitidas And personaje.clase)
End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
Dim j As Integer

'Del inventario
For j = 1 To UserList(UserIndex).Stats.MaxItems
        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
             If ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Newbie = 1 Then
                    Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, UserIndex, j)
             End If
        
        End If
Next

'De la boveda.
For j = 1 To MAX_BANCOINVENTORY_SLOTS
        If UserList(UserIndex).BancoInvent.Object(j).ObjIndex > 0 Then
             If ObjData(UserList(UserIndex).BancoInvent.Object(j).ObjIndex).Newbie = 1 Then
                UserList(UserIndex).BancoInvent.Object(j).ObjIndex = 0
                UserList(UserIndex).BancoInvent.Object(j).Amount = 0
            End If
        End If
Next


If UserList(UserIndex).pos.map = 37 Or _
UserList(UserIndex).pos.map = 167 Or _
UserList(UserIndex).pos.map = 168 Then
    Dim DeDonde As WorldPos
    
    DeDonde = GetCiudad(UserList(UserIndex))
    
    Call WarpUserChar(UserIndex, DeDonde.map, DeDonde.x, DeDonde.y, True)
End If
'[/Barrin]
End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)
Dim j As Integer


With UserList(UserIndex).Invent

    For j = 1 To 30
        .Object(j).ObjIndex = 0
        .Object(j).Amount = 0
        .Object(j).Equipped = 0
    Next
    
    .NroItems = 0
    .ArmourEqpObjIndex = 0
    .ArmourEqpSlot = 0
    .WeaponEqpObjIndex = 0
    .WeaponEqpSlot = 0
    .CascoEqpObjIndex = 0
    .CascoEqpSlot = 0
    .EscudoEqpObjIndex = 0
    .EscudoEqpSlot = 0
    .HerramientaEqpObjIndex = 0
    .HerramientaEqpSlot = 0
    .MunicionEqpObjIndex = 0
    .MunicionEqpSlot = 0
    .BarcoObjIndex = 0
    .BarcoSlot = 0

    UserList(UserIndex).Stats.MaxItems = 0
    
End With

End Sub


Public Function TirarOro(ByVal cantidad As Long, ByRef personaje As User, Optional ByVal duenio As Integer = 0) As Long

Dim MiObj As obj
Dim loops As Integer
Dim cantidadOriginal As Long

cantidadOriginal = cantidad

If cantidad > 100000 Then Exit Function

'SI EL NPC TIENE ORO LO TIRAMOS
If (cantidad = 0) Or (cantidad > personaje.Stats.GLD) Then
    Exit Function
End If

Do While (cantidad > 0) And (personaje.Stats.GLD > 0)
    If cantidad > MAX_INVENTORY_OBJS And personaje.Stats.GLD > MAX_INVENTORY_OBJS Then
        MiObj.Amount = MAX_INVENTORY_OBJS
        personaje.Stats.GLD = personaje.Stats.GLD - MAX_INVENTORY_OBJS
        cantidad = cantidad - MiObj.Amount
    Else
        MiObj.Amount = cantidad
        personaje.Stats.GLD = personaje.Stats.GLD - cantidad
        cantidad = cantidad - MiObj.Amount
    End If
    
    MiObj.ObjIndex = iORO
            
    If personaje.flags.Privilegios > 0 Then
        Call LogGM(personaje.id, "Tiro cantidad:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name)
    End If
    
    Call TirarItemAlPisoConDuenio(personaje.pos, MiObj, duenio)
    
    'info debug
    loops = loops + 1
    If loops > 100 Then
        LogError ("Error en tiraroro")
        TirarOro = cantidadOriginal - cantidad
        Exit Function
    End If
Loop

TirarOro = cantidadOriginal

End Function

Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal slot As Byte, ByVal cantidad As Integer)
'Desequipa

If slot < 1 Or slot > UserList(UserIndex).Stats.MaxItems Then Exit Sub
'Quita un objeto
UserList(UserIndex).Invent.Object(slot).Amount = UserList(UserIndex).Invent.Object(slot).Amount - cantidad
'¿Quedan mas?
If UserList(UserIndex).Invent.Object(slot).Amount <= 0 Then

    UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
    
    If UserList(UserIndex).Invent.Object(slot).Equipped = 1 Then
        Desequipar UserIndex, slot
    End If
    
    UserList(UserIndex).Invent.Object(slot).ObjIndex = 0
    UserList(UserIndex).Invent.Object(slot).Amount = 0
End If
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal slot As Byte)
Dim NullObj As UserOBJ
Dim loopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then
    'Actualiza el inventario
    If UserList(UserIndex).Invent.Object(slot).ObjIndex > 0 Then
        Call ChangeUserInv(UserIndex, slot, UserList(UserIndex).Invent.Object(slot))
    Else
        Call ChangeUserInv(UserIndex, slot, NullObj)
    End If
Else


'Actualiza todos los slots
    For loopC = 1 To UserList(UserIndex).Stats.MaxItems
        'Actualiza el inventario
        If UserList(UserIndex).Invent.Object(loopC).ObjIndex > 0 Then
            Call ChangeUserInv(UserIndex, loopC, UserList(UserIndex).Invent.Object(loopC))
        Else
            Call ChangeUserInv(UserIndex, loopC, NullObj)
        End If
    Next loopC
End If
End Sub

' Le saca un objeto determinado al usuario del inventario y lo pone en el suelo
' Devuelve Verdadero si se pudo hacer con exito
Public Function QuitarOBjetoYPonerEnSuelo(Usuario As User, ByVal slot As Byte, ByVal cantidad As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, Optional ByVal duenio As Integer = 0) As Boolean
Dim obj As obj

QuitarOBjetoYPonerEnSuelo = False

' Chequeo minimo  y existencia de objeto
If cantidad = 0 Then
    QuitarOBjetoYPonerEnSuelo = True
    Exit Function
End If

If Usuario.Invent.Object(slot).ObjIndex = 0 Then
    QuitarOBjetoYPonerEnSuelo = True
    Exit Function
End If

' Reviso maximo
If cantidad > Usuario.Invent.Object(slot).Amount Then cantidad = Usuario.Invent.Object(slot).Amount

' Creamos el Objeto
obj.ObjIndex = Usuario.Invent.Object(slot).ObjIndex
obj.Amount = cantidad

' ¿Hay lugar en el tile?
If Not (MapData(map, x, y).OBJInfo.ObjIndex = 0 Or MapData(map, x, y).OBJInfo.ObjIndex = obj.ObjIndex And obj.Amount + MapData(map, x, y).OBJInfo.Amount <= 10000) Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(152), Usuario.UserIndex
    Exit Function
End If
    
' Si es un objeto que tiene equipado y se va a quedar sin ese objeto, lo desequipo
If Usuario.Invent.Object(slot).Equipped = 1 And Usuario.Invent.Object(slot).Amount - obj.Amount <= 0 Then
    Call Desequipar(Usuario.UserIndex, slot)
End If
    
' ¿Ya está este tile?. Sabemos que va a entrar por el IF de mas arriba
If MapData(map, x, y).OBJInfo.ObjIndex = obj.ObjIndex And obj.Amount + MapData(map, x, y).OBJInfo.Amount <= 10000 Then
    obj.Amount = obj.Amount + MapData(map, x, y).OBJInfo.Amount    ' La cantidad sería lo que ya habia más lo que yo quiero tirar
End If
    
' Creamos el objeto
Call MakeObjDuenio(ToMap, 0, map, obj, map, x, y, duenio)
    
' Lo quitamos del personaje
Call QuitarUserInvItem(Usuario.UserIndex, slot, cantidad)
    
' Actualizamos el inventario
Call UpdateUserInv(False, Usuario.UserIndex, slot)

End Function


' Le saca un objeto determinado al usuario del inventario y lo pone en el suelo
' Devuelve Verdadero si se pudo hacer con exito
Public Function QuitarOBjetoSlot(personaje As User, ByVal slot As Byte, ByVal cantidad As Integer) As Boolean

' Lo quitamos del personaje
Call QuitarUserInvItem(personaje.UserIndex, slot, cantidad)
    
' Actualizamos el inventario
Call UpdateUserInv(False, personaje.UserIndex, slot)


End Function

Sub DropObj(personaje As User, ByVal slot As Byte, ByVal cantidad As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer)

Dim obj As obj
Dim aux As Boolean

' Chequeo minimo y máximo, y existencia de objeto
If cantidad = 0 Then Exit Sub
If cantidad > personaje.Invent.Object(slot).Amount Then cantidad = personaje.Invent.Object(slot).Amount
If personaje.Invent.Object(slot).ObjIndex = 0 Then Exit Sub

' Creamos el Objeto
obj.ObjIndex = personaje.Invent.Object(slot).ObjIndex
obj.Amount = cantidad

' ¿Esta en un evento? ¿El evento le permite lanzar objetos?
If Not personaje.evento Is Nothing Then
    If Not personaje.evento.puedeTirarObjeto(personaje.UserIndex, obj.ObjIndex, obj.Amount, suelo, 0) Then
       Exit Sub
    End If
End If

'   No puede tirar objetos Newbies
If ObjData(obj.ObjIndex).Newbie = 1 And EsNewbie(personaje.UserIndex) Then
    EnviarPaquete Paquetes.MensajeSimple2, Chr$(123), personaje.UserIndex
    Exit Sub
End If

' Restricciones que sólo aplica para personajes comunes y no Game Masters

If personaje.flags.Privilegios = 0 Then

    '   No pude tirar objetos faccionarios.
    If (ObjData(obj.ObjIndex).alineacion And eAlineaciones.caos) And personaje.faccion.FuerzasCaos = 1 Then
        EnviarPaquete Paquetes.mensajeinfo, "No puedes tirar un objeto faccionario.", personaje.UserIndex
        Exit Sub
    End If
    
    If (ObjData(obj.ObjIndex).alineacion And eAlineaciones.Real) And personaje.faccion.ArmadaReal = 1 Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(151), personaje.UserIndex
        Exit Sub
    End If
    
    ' No se puede tirar la armadura del dragon
    If obj.ObjIndex = Objetos_Constantes.ARMADURA_DRAGON_E Or _
            obj.ObjIndex = Objetos_Constantes.ARMADURA_DRAGON_H Or _
            obj.ObjIndex = Objetos_Constantes.ARMADURA_DRAGON_M Then
        EnviarPaquete Paquetes.mensajeinfo, "No puedes tirar la Armadura del Dragón, el enfado del Dragón te desterrará de estas tierras.", personaje.UserIndex
        Exit Sub
    End If
End If

' ¿Hay lugar en el tile?
If Not (MapData(map, x, y).OBJInfo.ObjIndex = 0 Or MapData(map, x, y).OBJInfo.ObjIndex = obj.ObjIndex And obj.Amount + MapData(map, x, y).OBJInfo.Amount <= 10000) Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(152), personaje.UserIndex
    Exit Sub
End If

' Si es un objeto que tiene equipado y se va a quedar sin ese objeto, lo desequipo
If personaje.Invent.Object(slot).Equipped = 1 And personaje.Invent.Object(slot).Amount - obj.Amount <= 0 Then
    Call Desequipar(personaje.UserIndex, slot)
End If

' ¡¡¡¡ OJO !!! Vuelvo a chequear nuevamente por si paso algo al desequipar
If Not personaje.Invent.Object(slot).ObjIndex = obj.ObjIndex Or obj.Amount > personaje.Invent.Object(slot).Amount Then Exit Sub

' ¿Ya está este tile?. Sabemos que va a entrar
If MapData(map, x, y).OBJInfo.ObjIndex = obj.ObjIndex And obj.Amount + MapData(map, x, y).OBJInfo.Amount <= 10000 Then
    ' La cantidad sería lo que ya habia más lo que yo quiero tirar
    obj.Amount = obj.Amount + MapData(map, x, y).OBJInfo.Amount
End If
    
' Creamos el objeto
Call MakeObj(ToMap, 0, map, obj, map, x, y)
    
' Lo quitamos del personaje
Call QuitarUserInvItem(personaje.UserIndex, slot, cantidad)
    
' Actualizamos el inventario
Call UpdateUserInv(False, personaje.UserIndex, slot)
                   
' Si tiro la barca le aviso! Warning!
If ObjData(obj.ObjIndex).ObjType = OBJTYPE_BARCOS Then EnviarPaquete Paquetes.MensajeSimple, Chr$(150), personaje.UserIndex

' Si es un Game Master Guardo Logs
If personaje.flags.Privilegios > 0 Then
    Call LogGM(personaje.id, "Tiro cantidad: " & cantidad & " Objeto:" & ObjData(obj.ObjIndex).Name & "PJs: " & modMapa.listarPersonajesOnline(MapInfo(personaje.pos.map)), "TIRAR")
End If

End Sub

Sub EraseObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal num As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer)
MapData(map, x, y).OBJInfo.Amount = MapData(map, x, y).OBJInfo.Amount - num

If MapData(map, x, y).OBJInfo.Amount <= 0 Then
    MapData(map, x, y).OBJInfo.ObjIndex = 0
    MapData(map, x, y).OBJInfo.Amount = 0
    
    EnviarPaquete Paquetes.BorrarObj, ITS(x) & ITS(y), sndIndex, sndRoute, sndMap
End If

End Sub

Sub MakeObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, obj As obj, map As Integer, ByVal x As Integer, ByVal y As Integer)
    Call MakeObjDuenio(sndRoute, sndIndex, sndMap, obj, map, x, y, 0)
End Sub

Sub MakeObjDuenio(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, obj As obj, map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal duenio As Integer)
    If obj.ObjIndex > 0 And obj.ObjIndex <= UBound(ObjData) Then
        MapData(map, x, y).OBJInfo = obj
        
        If duenio > 0 Then
            MapData(map, x, y).ObjInfoPoseedor.UserIndex = duenio
            MapData(map, x, y).ObjInfoPoseedor.fecha = GetTickCount
        End If
        
        EnviarPaquete Paquetes.CrearObjeto, ITS(ObjData(obj.ObjIndex).GrhIndex) & ITS(x) & ITS(y), sndIndex, sndRoute, sndMap
    End If
End Sub

Public Function tieneLugar(Usuario As User, objeto As obj) As Boolean

Dim slot As Byte

'¿El user ya tiene un objeto del mismo tipo?
slot = 1

Do Until Usuario.Invent.Object(slot).ObjIndex = objeto.ObjIndex And _
         Usuario.Invent.Object(slot).Amount + objeto.Amount <= MAX_INVENTORY_OBJS
         
   slot = slot + 1
   
   If slot > Usuario.Stats.MaxItems Then Exit Do
Loop

'Sino busca un slot vacio
If slot > Usuario.Stats.MaxItems Then

   slot = 1
   
   Do Until Usuario.Invent.Object(slot).ObjIndex = 0
   
       slot = slot + 1
       
       If slot > Usuario.Stats.MaxItems Then Exit Do
   Loop

End If

' ¿Encontre?
If slot > Usuario.Stats.MaxItems Then
    tieneLugar = False
Else
    tieneLugar = True
End If

End Function




Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As obj) As Boolean
'Call LogTarea("MeterItemEnInventario")
Dim slot As Byte

'¿el user ya tiene un objeto del mismo tipo?
slot = 1
Do Until UserList(UserIndex).Invent.Object(slot).ObjIndex = MiObj.ObjIndex And _
         UserList(UserIndex).Invent.Object(slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
   slot = slot + 1
   If slot > UserList(UserIndex).Stats.MaxItems Then
         Exit Do
   End If
Loop
'Sino busca un slot vacio
If slot > UserList(UserIndex).Stats.MaxItems Then
   slot = 1
   Do Until UserList(UserIndex).Invent.Object(slot).ObjIndex = 0
       slot = slot + 1
       If slot > UserList(UserIndex).Stats.MaxItems Then
           EnviarPaquete Paquetes.MensajeSimple, Chr$(153), UserIndex
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If
'Mete el objeto
If UserList(UserIndex).Invent.Object(slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
   'Menor que MAX_INV_OBJS
   UserList(UserIndex).Invent.Object(slot).ObjIndex = MiObj.ObjIndex
   UserList(UserIndex).Invent.Object(slot).Amount = UserList(UserIndex).Invent.Object(slot).Amount + MiObj.Amount
Else
   UserList(UserIndex).Invent.Object(slot).Amount = MAX_INVENTORY_OBJS
End If
MeterItemEnInventario = True
Call UpdateUserInv(False, UserIndex, slot)

End Function

Sub GetObj(ByVal UserIndex As Integer)
Dim obj As ObjData
Dim MiObj As obj

Dim x As Integer
Dim y As Integer
Dim mapa As Integer
Dim ObjIndex As Integer

x = UserList(UserIndex).pos.x
y = UserList(UserIndex).pos.y
mapa = UserList(UserIndex).pos.map

'¿Hay algun obj?
If MapData(mapa, x, y).OBJInfo.ObjIndex = 0 Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(155), UserIndex
    Exit Sub
End If

ObjIndex = MapData(mapa, x, y).OBJInfo.ObjIndex

'¿Esta permitido agarrar este obj?
If ObjData(ObjIndex).Agarrable = 0 Then
    Exit Sub
End If


' Alguien es el dueño?
'If MapData(mapa, x, y).ObjInfoPoseedor.UserIndex > 0 Then
'    ' Si no es el dueño y tampoco paso el tiempo
'    If (Not MapData(mapa, x, y).ObjInfoPoseedor.UserIndex = UserIndex) And MapData(mapa, x, y).ObjInfoPoseedor.fecha + 5000 > GetTickCount Then
'        ' Un aopcion mas. Si estan en la misma party, no cancelo.
'        If UserList(UserIndex).PartyIndex = 0 Or Not UserList(MapData(mapa, x, y).ObjInfoPoseedor.UserIndex).PartyIndex = UserList(UserIndex).PartyIndex Then
'            EnviarPaquete Paquetes.mensajeinfo, "Debes esperar 5 segundos para agarrar este objeto.", UserIndex, ToIndex
'            Exit Sub
'        End If
'    End If
'End If
'
'' Si esta invisible solo puede agarrar sus items
'If UserList(UserIndex).flags.Invisible = 1 And Not MapData(mapa, x, y).ObjInfoPoseedor.UserIndex = UserIndex Then
'    If UserList(UserIndex).PartyIndex = 0 Or Not UserList(MapData(mapa, x, y).ObjInfoPoseedor.UserIndex).PartyIndex = UserList(UserIndex).PartyIndex Then
'        EnviarPaquete Paquetes.mensajeinfo, "No puedes agarrar objetos estando invisible a menos que sean conseguidos por tí.", UserIndex, ToIndex
'        Exit Sub
'    End If
'End If

MapData(mapa, x, y).ObjInfoPoseedor.UserIndex = 0

MiObj.Amount = MapData(mapa, x, y).OBJInfo.Amount
MiObj.ObjIndex = ObjIndex

If MiObj.Amount = 0 Then Exit Sub

If obj.Ubicable = 1 And UserList(UserIndex).flags.Privilegios = 0 Then
    EnviarPaquete Paquetes.MensajeCompuesto, Chr$(14) & UserList(UserIndex).Name & " tiene la " & ObjData(ObjIndex).Name & ". Se encuentra en el mapa " & mapa & ".~250~250~250~1~0", 0, ToAll
End If

If Not MeterItemEnInventario(UserIndex, MiObj) Then
  '  EnviarPaquete Paquetes.MensajeSimple, Chr$(154), UserINdex
  ' Mensaje redudante
Else
    'Quitamos el objeto
    Call EraseObj(ToMap, 0, mapa, MapData(mapa, x, y).OBJInfo.Amount, mapa, x, y)
    If UserList(UserIndex).flags.Privilegios > 0 Then Call LogGM(UserList(UserIndex).id, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(ObjIndex).Name)
End If

End Sub

Public Sub desequiparByItem(ByRef personaje As User, ByVal ObjIndex As Integer)
    Dim slot As Byte
    
    For slot = 1 To personaje.Stats.MaxItems
        If personaje.Invent.Object(slot).Equipped = 1 And personaje.Invent.Object(slot).ObjIndex = ObjIndex Then
            Call Desequipar(personaje.UserIndex, slot)
        End If
    Next slot

End Sub
Sub Desequipar(ByVal UserIndex As Integer, ByVal slot As Byte, Optional NoPasar As Boolean)
'Desequipa el item slot del inventario
Dim obj As ObjData
Dim tempbyte As Byte
Dim Obj2 As obj

With UserList(UserIndex)
If (slot < LBound(.Invent.Object)) Or (slot > UBound(.Invent.Object)) Then
    Exit Sub
ElseIf .Invent.Object(slot).ObjIndex = 0 Then
    Exit Sub
End If

obj = ObjData(.Invent.Object(slot).ObjIndex)

Select Case obj.ObjType
    Case OBJTYPE_WEAPON
                .Invent.Object(slot).Equipped = 0
                .Invent.WeaponEqpObjIndex = 0
                .Invent.WeaponEqpSlot = 0
                If .flags.Mimetizado = 0 Then
                    .Char.WeaponAnim = NingunArma
                    EnviarPaquete Paquetes.pChangeUserCharArma, ITS(.Char.charIndex) & ByteToString(.Char.WeaponAnim), UserIndex, ToArea
                End If
    Case OBJTYPE_ANILLOS
                .Invent.Object(slot).Equipped = 0
                .Invent.AnilloEqpObjIndex = 0
                .Invent.AnilloEqpSlot = 0
    Case OBJTYPE_FLECHAS
                .Invent.Object(slot).Equipped = 0
                .Invent.MunicionEqpObjIndex = 0
                .Invent.MunicionEqpSlot = 0
    Case OBJTYPE_HERRAMIENTAS, OBJTYPE_MINERALES
                .Invent.Object(slot).Equipped = 0
                .Invent.HerramientaEqpObjIndex = 0
                .Invent.HerramientaEqpSlot = 0
                If UserList(UserIndex).flags.Trabajando Then
                    Call DejarDeTrabajar(UserList(UserIndex))
                End If
    Case OBJTYPE_COLLAR
            .Invent.CollarObjIndex = 0
            .Invent.Object(slot).Equipped = 0
    Case OBJTYPE_BRASALETE
            .Invent.BrasaleteEqpObjIndex = 0
            .Invent.Object(slot).Equipped = 0
    Case OBJTYPE_ARMOUR
        Select Case obj.subTipo
            Case OBJTYPE_ARMADURA
                        .Invent.Object(slot).Equipped = 0
                        .Invent.ArmourEqpObjIndex = 0
                        .Invent.ArmourEqpSlot = 0
                
                        If Not NoPasar Then
                            Call modPersonaje.DarAparienciaCorrespondiente(UserList(UserIndex))
                            Call modPersonaje_TCP.ActualizarEstetica(UserList(UserIndex))
                            'If .flags.Mimetizado = 0 And .flags.Navegando = 0 Then
                            '    Call DarCuerpoDesnudo(UserList(UserIndex))
                            '    EnviarPaquete Paquetes.pChangeUserCharArmadura, ITS(.Char.charIndex) & ITS(.Char.Body), UserIndex, ToArea
                            'End If
                        End If
            Case OBJTYPE_CASCO
                        .Invent.Object(slot).Equipped = 0
                        .Invent.CascoEqpObjIndex = 0
                        .Invent.CascoEqpSlot = 0
                        
                        If .flags.Mimetizado = 0 Then
                            .Char.CascoAnim = NingunCasco
                            EnviarPaquete Paquetes.pChangeUserCharCasco, ITS(.Char.charIndex) & Chr$(.Char.CascoAnim), UserIndex, ToArea
                        End If
            Case OBJTYPE_ESCUDO
                        .Invent.Object(slot).Equipped = 0
                        .Invent.EscudoEqpObjIndex = 0
                        .Invent.EscudoEqpSlot = 0
                
                        If .flags.Mimetizado = 0 Then
                            .Char.ShieldAnim = NingunEscudo
                            EnviarPaquete Paquetes.pChangeUserCharEscudo, ITS(.Char.charIndex) & Chr$(.Char.ShieldAnim), UserIndex, ToArea
                        End If
        End Select
    Case OBJTYPE_BARCOS
                .Invent.Object(slot).Equipped = 0
                .Invent.BarcoEqpSlot = 0
                
                If .Stats.MaxItems > 20 Then
                    For tempbyte = 21 To .Stats.MaxItems
                        If .Invent.Object(tempbyte).ObjIndex <> 0 Then
                            Obj2.Amount = .Invent.Object(tempbyte).Amount
                            Obj2.ObjIndex = .Invent.Object(tempbyte).ObjIndex
                            Call TirarItemAlPiso(.pos, Obj2)
                            Call QuitarUserInvItem(UserIndex, tempbyte, .Invent.Object(tempbyte).Amount)
                        End If
                    Next
                    Call UpdateUserInv(True, UserIndex, 0)
                End If
                .Stats.MaxItems = 20
End Select
End With
Call UpdateUserInv(False, UserIndex, slot)
End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean

If ObjData(ObjIndex).Genero = eGeneros.indefinido Then
    SexoPuedeUsarItem = True
    Exit Function
End If

SexoPuedeUsarItem = (ObjData(ObjIndex).Genero And UserList(UserIndex).Genero)

End Function

Function FaccionPuedeUsarItem(ByRef personaje As User, ByRef objeto As ObjData) As Boolean

If modObjeto.isFaccionario(objeto) = False Then
    FaccionPuedeUsarItem = True
    Exit Function
End If

FaccionPuedeUsarItem = (objeto.alineacion = personaje.faccion.alineacion)

End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal slot As Byte)
'Equipa un item del inventario
Dim obj As ObjData
Dim ObjIndex As Integer

With UserList(UserIndex)

ObjIndex = .Invent.Object(slot).ObjIndex
obj = ObjData(ObjIndex)
If obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(156), UserIndex
     Exit Sub
End If
       

Select Case obj.ObjType
    Case OBJTYPE_WEAPON
    
        If ClasePuedeUsarItem(UserList(UserIndex), ObjData(ObjIndex)) = False Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(157), UserIndex
            Exit Sub
        End If
       
        '[Wizard] como no se que objetos puede neceitar va para todos:D
        If obj.SkillM > .Stats.UserSkills(eSkills.Magia) Then
            EnviarPaquete Paquetes.MensajeCompuesto, Chr$(24) & obj.SkillM, UserIndex
            Exit Sub
        End If

       If FaccionPuedeUsarItem(UserList(UserIndex), ObjData(ObjIndex)) = False Then
            EnviarPaquete Paquetes.MensajeSimple, Chr$(157), UserIndex
            Exit Sub
       End If
       
       If ObjData(ObjIndex).proyectil = 1 Then
            If ObjData(ObjIndex).SkillCombate > .Stats.UserSkills(proyectiles) Then
            EnviarPaquete Paquetes.mensajeinfo, "Para usar este arma necesitas " & ObjData(ObjIndex).SkillCombate & " skills en armas de proyectiles.", UserIndex
            Exit Sub
            End If
        ElseIf ObjData(ObjIndex).Apuñala = 1 Then
            If ObjData(ObjIndex).SkillCombate > .Stats.UserSkills(Apuñalar) Then
            EnviarPaquete Paquetes.mensajeinfo, "Para usar este arma necesitas " & ObjData(ObjIndex).SkillCombate & " skills en Apuñalar.", UserIndex
            Exit Sub
            End If
        Else
            If ObjData(ObjIndex).SkillCombate > .Stats.UserSkills(Armas) Then
            EnviarPaquete Paquetes.mensajeinfo, "Para usar este arma necesitas " & ObjData(ObjIndex).SkillCombate & " skills en Combate con armas.", UserIndex
            Exit Sub
            End If
        End If
         
        'Si esta equipado lo quita
        If .Invent.Object(slot).Equipped Then
            Call Desequipar(UserIndex, slot)
            Exit Sub
        End If
                

        If .Invent.WeaponEqpObjIndex > 0 Or .Invent.HerramientaEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
        End If
                
        Select Case .clase
        
            Case eClases.asesino
            
                If UCase(obj.Name) = "KATANA" Or UCase(obj.Name) = "SABLE" Then
                    If .Invent.EscudoEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
                    End If
                End If
                
            Case eClases.Cazador
            
                If UCase(obj.Name) = "ESPADA DOS MANOS" Then
                    If .Invent.EscudoEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
                    End If
                End If
                
            Case eClases.Guerrero
            
                If UCase(obj.Name) = "ESPADA DOS MANOS" Then
                    If .Invent.EscudoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
                    End If
                End If
                
        End Select
    
                
        .Invent.Object(slot).Equipped = 1
        .Invent.WeaponEqpObjIndex = .Invent.Object(slot).ObjIndex
        .Invent.WeaponEqpSlot = slot
 
        'Sonido
        If InStr(1, ObjData(.Invent.WeaponEqpObjIndex).Name, "Espada") > 0 Then
            EnviarPaquete Paquetes.WavSnd, Chr$(SOUND_SACARESPADA), UserIndex, ToPCArea, .pos.map
        Else
            EnviarPaquete Paquetes.WavSnd, Chr$(SOUND_SACARARMA), UserIndex, ToPCArea, .pos.map
        End If
            
        If .flags.Mimetizado = 0 Then
            .Char.WeaponAnim = obj.WeaponAnim
            EnviarPaquete Paquetes.pChangeUserCharArma, ITS(.Char.charIndex) & ByteToString(.Char.WeaponAnim), UserIndex, ToArea
        End If
            
        Call UpdateUserInv(False, UserIndex, slot)

   Case OBJTYPE_ANILLOS
   
        If ClasePuedeUsarItem(UserList(UserIndex), ObjData(ObjIndex)) = False Then
          EnviarPaquete Paquetes.MensajeSimple2, Chr$(31), UserIndex
          Exit Sub
        End If
        
        '[Wizard] como no se que objetos puede neceitar va para todos:D
        If obj.SkillM > .Stats.UserSkills(eSkills.Magia) Then
            EnviarPaquete Paquetes.MensajeCompuesto, Chr$(24) & obj.SkillM, UserIndex
            Exit Sub
        End If
        
        'Si esta equipado lo quita
        If .Invent.Object(slot).Equipped Then
            'Quitamos del inv el item
            Call Desequipar(UserIndex, slot)
            Exit Sub
        End If
        
        'Quitamos el elemento anterior
         If .Invent.AnilloEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)
        End If
        
        .Invent.Object(slot).Equipped = 1
        .Invent.AnilloEqpObjIndex = ObjIndex
        .Invent.AnilloEqpSlot = slot
        Call UpdateUserInv(False, UserIndex, slot)
        

    Case OBJTYPE_HERRAMIENTAS, OBJTYPE_MINERALES
    
       If ClasePuedeUsarItem(UserList(UserIndex), ObjData(ObjIndex)) = False Then
            EnviarPaquete Paquetes.MensajeSimple2, Chr$(31), UserIndex
            Exit Sub
        End If
        
        '[Wizard] como no se que objetos puede neceitar va para todos:D
        If obj.SkillM > .Stats.UserSkills(eSkills.Magia) Then
            EnviarPaquete Paquetes.MensajeCompuesto, Chr$(24) & obj.SkillM, UserIndex
            Exit Sub
        End If
        
        If FaccionPuedeUsarItem(UserList(UserIndex), ObjData(ObjIndex)) = False Then
            EnviarPaquete Paquetes.MensajeSimple2, Chr$(31), UserIndex
            Exit Sub
        End If
        
        If obj.SkillMin > .Stats.UserSkills(eSkills.Mineria) Then
            EnviarPaquete Paquetes.MensajeCompuesto, Chr$(41) & obj.SkillMin, UserIndex
            Exit Sub
        End If
          
        If .Invent.Object(slot).ObjIndex = RED_PESCA Then
            If .Stats.UserSkills(eSkills.Pesca) < 100 Or .Invent.BarcoObjIndex <> 475 Then
                EnviarPaquete Paquetes.mensajeinfo, "Para equipar la red de pesca debes tener 100 skills en pesca y estar dentro de una galera.", UserIndex, ToIndex
                Exit Sub
            End If
        End If
        
         'Si esta equipado lo quita
        If .Invent.Object(slot).Equipped Then
            'Quitamos del inv el item
            Call Desequipar(UserIndex, slot)
            Exit Sub
        End If
        'Quitamos el elemento anterior
        If .Invent.HerramientaEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.HerramientaEqpSlot)
        End If
        
        .Invent.Object(slot).Equipped = 1
        .Invent.HerramientaEqpObjIndex = ObjIndex
        .Invent.HerramientaEqpSlot = slot

        Call UpdateUserInv(False, UserIndex, slot)
    Case OBJTYPE_FLECHAS
            
        If ClasePuedeUsarItem(UserList(UserIndex), ObjData(.Invent.Object(slot).ObjIndex)) = False Then
            EnviarPaquete Paquetes.MensajeSimple2, Chr$(31), UserIndex
            Exit Sub
        End If
        
         '[Wizard] como no se que objetos puede neceitar va para todos:D
        If obj.SkillM > .Stats.UserSkills(eSkills.Magia) Then
            EnviarPaquete Paquetes.MensajeCompuesto, Chr$(24) & obj.SkillM, UserIndex
            Exit Sub
        End If
        
        If FaccionPuedeUsarItem(UserList(UserIndex), ObjData(.Invent.Object(slot).ObjIndex)) = False Then
            EnviarPaquete Paquetes.MensajeSimple2, Chr$(31), UserIndex
            Exit Sub
        End If
          
        'Si esta equipado lo quita
        If .Invent.Object(slot).Equipped Then
            'Quitamos del inv el item
            Call Desequipar(UserIndex, slot)
            Exit Sub
        End If
         
        'Quitamos el elemento anterior
        If .Invent.MunicionEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
        End If
         
        .Invent.Object(slot).Equipped = 1
        .Invent.MunicionEqpObjIndex = .Invent.Object(slot).ObjIndex
        .Invent.MunicionEqpSlot = slot
         
        Call UpdateUserInv(False, UserIndex, slot)

    Case OBJTYPE_ARMOUR

        If .flags.Navegando = 1 Then Exit Sub
        
         
         Select Case obj.subTipo
            Case OBJTYPE_ARMADURA

                'Nos aseguramos que puede usarla
                If ClasePuedeUsarItem(UserList(UserIndex), ObjData(.Invent.Object(slot).ObjIndex)) = False Then
                    EnviarPaquete Paquetes.MensajeSimple, Chr$(158), UserIndex
                    Exit Sub
                End If
                
                If SexoPuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex) = False Then
                    EnviarPaquete Paquetes.MensajeSimple, Chr$(158), UserIndex
                    Exit Sub
                End If
                
                If CheckRazaUsaRopa(UserList(UserIndex), ObjData(.Invent.Object(slot).ObjIndex)) = False Then
                    EnviarPaquete Paquetes.MensajeSimple, Chr$(158), UserIndex
                    Exit Sub
                End If
                
                If FaccionPuedeUsarItem(UserList(UserIndex), ObjData(.Invent.Object(slot).ObjIndex)) = False Then
                    EnviarPaquete Paquetes.MensajeSimple2, Chr$(31), UserIndex
                    Exit Sub
                End If
                   
                '[Wizard] como no se que objetos puede neceitar va para todos:D
                If obj.SkillM > .Stats.UserSkills(eSkills.Magia) Then
                    EnviarPaquete Paquetes.MensajeCompuesto, Chr$(24) & obj.SkillM, UserIndex
                    Exit Sub
                End If
                   
                'Si esta equipado lo quita
                If .Invent.Object(slot).Equipped Then
                    Call Desequipar(UserIndex, slot)
                    Call DarCuerpoDesnudo(UserList(UserIndex))
                    If Not .flags.Mimetizado = 1 Then
                        Call ChangeUserChar(ToMap, 0, .pos.map, UserIndex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                    Exit Sub
                End If
                    
                If ObjData(ObjIndex).SkillTacticass > .Stats.UserSkills(tacticas) Then
                    EnviarPaquete Paquetes.mensajeinfo, "Necesitas " & ObjData(ObjIndex).SkillTacticass & " skills en Tacticas de combate para usar esta Armadura.", UserIndex
                    Exit Sub
                End If
                    
                'Quita el anterior
                If .Invent.ArmourEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, .Invent.ArmourEqpSlot, True)
                End If
                
                'Lo equipa
                .Invent.Object(slot).Equipped = 1
                .Invent.ArmourEqpObjIndex = .Invent.Object(slot).ObjIndex
                .Invent.ArmourEqpSlot = slot
                        
                If .flags.Mimetizado = 0 Then
                    .Char.Body = obj.Ropaje
                    EnviarPaquete Paquetes.pChangeUserCharArmadura, ITS(.Char.charIndex) & ITS(.Char.Body), UserIndex, ToArea
                End If
                
                .flags.Desnudo = 0
                Call UpdateUserInv(False, UserIndex, slot)
                     
            Case OBJTYPE_CASCO
                If ClasePuedeUsarItem(UserList(UserIndex), ObjData(.Invent.Object(slot).ObjIndex)) Then
                 
                '[Wizard] como no se que objetos puede neceitar va para todos:D
                If obj.SkillM > .Stats.UserSkills(eSkills.Magia) Then
                    EnviarPaquete Paquetes.MensajeCompuesto, Chr$(24) & obj.SkillM, UserIndex
                    Exit Sub
                End If
        
                If ObjData(ObjIndex).SkillTacticass > .Stats.UserSkills(tacticas) Then
                    EnviarPaquete Paquetes.mensajeinfo, "Necesitas " & ObjData(ObjIndex).SkillTacticass & " skills en Tacticas de combate para usar este casco.", UserIndex
                    Exit Sub
                End If
            
            'Si esta equipado lo quita
                    If .Invent.Object(slot).Equipped Then
                        Call Desequipar(UserIndex, slot)
                        If .flags.Mimetizado = 0 Then
                            .Char.CascoAnim = NingunCasco
                            Call ChangeUserChar(ToMap, 0, .pos.map, UserIndex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If
                        Exit Sub
                    End If

                    If .Invent.CascoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.CascoEqpSlot)
                    End If
                    
                    'Lo equipa
                    .Invent.Object(slot).Equipped = 1
                    .Invent.CascoEqpObjIndex = .Invent.Object(slot).ObjIndex
                    .Invent.CascoEqpSlot = slot
                    .Char.CascoAnim = obj.CascoAnim
                    EnviarPaquete Paquetes.pChangeUserCharCasco, ITS(.Char.charIndex) & Chr$(.Char.CascoAnim), UserIndex, ToArea
                Else
                    EnviarPaquete Paquetes.MensajeSimple2, Chr$(31), UserIndex, ToIndex, 0
                End If
                Call UpdateUserInv(False, UserIndex, slot)
                Exit Sub
            Case OBJTYPE_ESCUDO
                If ClasePuedeUsarItem(UserList(UserIndex), ObjData(.Invent.Object(slot).ObjIndex)) = False Then
                    EnviarPaquete Paquetes.MensajeSimple2, Chr$(31), UserIndex
                    Exit Sub
                End If
                
                '[Wizard] como no se que objetos puede neceitar va para todos:D
                If obj.SkillM > .Stats.UserSkills(eSkills.Magia) Then
                    EnviarPaquete Paquetes.MensajeCompuesto, Chr$(24) & obj.SkillM, UserIndex
                    Exit Sub
                End If
        
                If ObjData(ObjIndex).SkillDefe > .Stats.UserSkills(Defensa) Then
                    EnviarPaquete Paquetes.mensajeinfo, "Necesitas " & ObjData(ObjIndex).SkillDefe & " skills en Defensa con escudos para usar este escudo.", UserIndex, ToIndex, 0
                    Exit Sub
                End If
            
                'Si esta equipado lo quita
                If .Invent.Object(slot).Equipped Then
                    Call Desequipar(UserIndex, slot)
                    If .flags.Mimetizado = 0 Then
                        .Char.ShieldAnim = NingunEscudo
                        Call ChangeUserChar(ToMap, 0, .pos.map, UserIndex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                    Exit Sub
                End If
                    
                'Quita el anterior
                If .Invent.EscudoEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
                End If
                    
                'marche
                 Select Case .clase
                 
                     Case eClases.asesino
                     
                         If UCase(obj.Name) = UCase("Escudo de Tortuga") Then
                              If .Invent.WeaponEqpObjIndex > 0 Then
                                 ObjIndex = .Invent.Object(.Invent.WeaponEqpSlot).ObjIndex
                                 If ObjData(ObjIndex).Name = "Katana" Or ObjData(ObjIndex).Name = "Sable" Then
                                 Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                                 End If
                             End If
                         End If
                         
                     Case eClases.Cazador
                     
                          If UCase(obj.Name) = UCase("Escudo de Tortuga") Then
                              If .Invent.WeaponEqpObjIndex > 0 Then
                                 ObjIndex = .Invent.Object(.Invent.WeaponEqpSlot).ObjIndex
                                 If UCase(ObjData(ObjIndex).Name) = UCase("Espada dos Manos") Then
                                     Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                                 End If
                             End If
                         End If
                         
                     Case eClases.Guerrero
                     
                         If .Invent.WeaponEqpObjIndex > 0 Then
                             ObjIndex = .Invent.Object(.Invent.WeaponEqpSlot).ObjIndex
                             If UCase(ObjData(ObjIndex).Name) = "ESPADA DOS MANOS" Then
                                 Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                             End If
                         End If
                         
                 End Select
                 
                 'Lo equipa
                 .Invent.Object(slot).Equipped = 1
                 .Invent.EscudoEqpObjIndex = .Invent.Object(slot).ObjIndex
                 .Invent.EscudoEqpSlot = slot
                 
                 If .flags.Mimetizado = 0 Then
                     .Char.ShieldAnim = obj.ShieldAnim
                      EnviarPaquete Paquetes.pChangeUserCharEscudo, ITS(.Char.charIndex) & Chr$(.Char.ShieldAnim), UserIndex, ToArea
                  End If
                 
                 Call UpdateUserInv(False, UserIndex, slot)

         End Select

    Case OBJTYPE_COLLAR
            
            If ClasePuedeUsarItem(UserList(UserIndex), ObjData(ObjIndex)) = False Then
                EnviarPaquete Paquetes.mensajeinfo, "Tu clase no puede utilizar este objeto.", UserIndex, ToIndex
                 Exit Sub
             End If
            
            If .Invent.Object(slot).Equipped = 1 Then
                Call Desequipar(UserIndex, slot)
             Else
                .Invent.Object(slot).Equipped = 1
                .Invent.CollarObjIndex = ObjIndex
                Call UpdateUserInv(False, UserIndex, slot)
             End If

    Case OBJTYPE_BRASALETE
            
            If ClasePuedeUsarItem(UserList(UserIndex), ObjData(ObjIndex)) = False Then
                EnviarPaquete Paquetes.mensajeinfo, "Tu clase no puede utilizar este objeto.", UserIndex, ToIndex
                 Exit Sub
             End If
            
            If .Invent.Object(slot).Equipped = 1 Then
                Call Desequipar(UserIndex, slot)
             Else
            
                If .Invent.BrasaleteEqpObjIndex > 0 Then
                    Call desequiparByItem(UserList(UserIndex), .Invent.BrasaleteEqpObjIndex)
                 End If
                
                .Invent.Object(slot).Equipped = 1
                .Invent.BrasaleteEqpObjIndex = ObjIndex
                Call UpdateUserInv(False, UserIndex, slot)
             End If
            
            
    Case OBJTYPE_BARCOS
            If ClasePuedeUsarItem(UserList(UserIndex), ObjData(ObjIndex)) = False Then
                 Exit Sub
             End If
            
             '[Wizard] como no se que objetos puede neceitar va para todos:D
            If obj.SkillM > .Stats.UserSkills(eSkills.Magia) Then
                EnviarPaquete Paquetes.MensajeCompuesto, Chr$(24) & obj.SkillM, UserIndex
                 Exit Sub
             End If
        
            If .Stats.UserSkills(Navegacion) / ModNavegacion(.clase) < ObjData(ObjIndex).MinSkill Then
                EnviarPaquete Paquetes.mensajeinfo, "Necesitas " & ObjData(ObjIndex).MinSkill * ModNavegacion(.clase) & " skills en navegar para poder usar esta barca.", UserIndex
                 Exit Sub
             End If
            
            If slot > 20 Then
                EnviarPaquete Paquetes.mensajeinfo, "No puedes equipar la barco en el espacio de otra barco.", UserIndex
                 Exit Sub
             End If
                
             'Si esta equipado lo quita
            If .Invent.Object(slot).Equipped = 1 Then
                Call Desequipar(UserIndex, slot)
             Else 'Lo equipa
                If .Invent.BarcoEqpSlot > 0 Then Call Desequipar(UserIndex, .Invent.BarcoEqpSlot)
                
                .Invent.Object(slot).Equipped = 1
                .Invent.BarcoEqpSlot = slot
                
                 'Nuevo Plus para los pescadores
                If .clase = eClases.Pirata Then
                    If UCase(obj.Name) = UCase("Galera") Then
                        .Stats.MaxItems = 25
                    ElseIf UCase(obj.Name) = UCase("Galeon") Then
                        .Stats.MaxItems = 30
                      Else
                        .Stats.MaxItems = 20
                      End If
                  End If
              End If
        
            Call UpdateUserInv(False, UserIndex, slot)
         
      End Select
  End With
End Sub

Private Function CheckRazaUsaRopa(ByRef personaje As User, ByRef objeto As ObjData) As Boolean

If objeto.razas = eRazas.indefinido Then
    CheckRazaUsaRopa = True
    Exit Function
End If

CheckRazaUsaRopa = (personaje.Raza And objeto.razas)

End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal slot As Byte, ByVal Origen As Byte, ByVal timeStamp As Single)
'Usa un item del inventario
Dim obj As ObjData
Dim ObjIndex As Integer
Dim TargObj As ObjData
Dim MiObj As obj
Dim tempSingle As Single
Dim maximoIncremento As Integer
Dim i As Integer

With UserList(UserIndex)

    If .Invent.Object(slot).Amount = 0 Then Exit Sub

    obj = ObjData(.Invent.Object(slot).ObjIndex)
    
    If obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
        EnviarPaquete Paquetes.MensajeSimple2, Chr$(32), UserIndex
        Exit Sub
    End If

    If obj.ObjType = OBJTYPE_PERGAMINOS Then
        If Not ClasePuedeUsarItem(UserList(UserIndex), ObjData(.Invent.Object(slot).ObjIndex)) Then
            EnviarPaquete Paquetes.MensajeSimple2, Chr$(95), UserIndex
            Exit Sub
        End If
    End If

    ObjIndex = .Invent.Object(slot).ObjIndex

    Select Case obj.ObjType
        Case OBJTYPE_USEONCE
        
            If .flags.Muerto = 1 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(26), UserIndex
                Exit Sub
            End If
            
            'Usa el item
            Call AddtoVar(.Stats.minham, obj.minham, .Stats.MaxHam)
            
            .flags.Hambre = 0
        
            Call EnviarHambreYsed(UserIndex)
            
            'Sonido
            EnviarPaquete Paquetes.WavSnd, Chr$(SOUND_COMIDA), UserIndex, ToPCArea, .pos.map
            
            'Quitamos del inv el item
            Call QuitarUserInvItem(UserIndex, slot, 1)
            Call UpdateUserInv(False, UserIndex, slot)

        Case OBJTYPE_GUITA
            If .flags.Muerto = 1 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(26), UserIndex
                Exit Sub
            End If
            
            .Stats.GLD = .Stats.GLD + .Invent.Object(slot).Amount
            .Invent.Object(slot).Amount = 0
            .Invent.Object(slot).ObjIndex = 0
            .Invent.NroItems = .Invent.NroItems - 1
            
            Call UpdateUserInv(False, UserIndex, slot)
            
            EnviarPaquete Paquetes.EnviarOro, Codify(.Stats.GLD), UserIndex, ToIndex
        
        Case OBJTYPE_COFRES

            If .flags.Muerto = 1 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(26), UserIndex
                Exit Sub
            End If
            
            Call DropeoCofre(UserIndex, slot, obj, MiObj)
            
        Case OBJTYPE_TRANSLADO
        
            Call Transportarse(UserIndex, slot, obj)
            
        Case OBJTYPE_VIAJES
        
            Call Viajar(UserIndex, slot, obj)

        Case OBJTYPE_WEAPON
            If .flags.Muerto = 1 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(26), UserIndex
                Exit Sub
            End If
            
            If ObjData(ObjIndex).proyectil = 1 Then
                EnviarPaquete Paquetes.ApuntarProyectil, "", UserIndex
            Else
                If obj.Name = "Laúd Mágico" Then
                    EnviarPaquete Paquetes.WavSnd, Chr$(obj.Snd1), UserIndex, ToPCArea
                End If
            
                If .flags.TargetObj = 0 Then Exit Sub
                
                TargObj = ObjData(.flags.TargetObj)
            
                '¿El target-objeto es leña?
                If TargObj.ObjType = OBJTYPE_LEÑA Then
                    If .Invent.Object(slot).ObjIndex = DAGA Then
                        Call TratarDeHacerFogata(.flags.TargetObjMap _
                             , .flags.TargetObjX, .flags.TargetObjY, UserIndex)
                    End If
                End If
            End If

        Case OBJTYPE_ANILLOS
            If .flags.Muerto = 1 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(26), UserIndex
                Exit Sub
            End If
            
            If obj.Snd1 > 1 Then
                EnviarPaquete Paquetes.WavSnd, Chr$(obj.Snd1), UserIndex, ToPCArea, .pos.map
            End If
            
        Case OBJTYPE_POCIONES
    
            If .flags.Muerto = 1 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(26), UserIndex
                Exit Sub
            End If

            ' Chequeamos que no este rompiendo el intervalo
            If Origen = 1 Then ' Pulsa la U
                Call anticheat.chequeoIntervaloCliente(UserList(UserIndex), UserList(UserIndex).Counters.ultimoTickU, UserList(UserIndex).intervalos.UsarU, timeStamp, "usar pocion")
            Else ' Doble clic
                Call anticheat.chequeoIntervaloCliente(UserList(UserIndex), UserList(UserIndex).Counters.ultimoTickClicUsar, UserList(UserIndex).intervalos.UsarClick, timeStamp, "clic pocion")
            End If
        
            If Not IntervaloPermiteAtacar(UserIndex, False) Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(159), UserIndex
                Exit Sub
            End If
                
            Select Case obj.TipoPocion
                Case ePociones.Agilidad
                    Call modPersonaje.incrementarAgilidad(UserList(UserIndex), RandomNumber(obj.MinModificador, obj.MaxModificador), obj.DuracionEfecto)
                
                    Call QuitarUserInvItem(UserIndex, slot, 1)
                
                    EnviarPaquete Paquetes.SonidoTomarPociones, "", UserIndex, ToPCArea
                Case ePociones.Fuerza
    
                    Call modPersonaje.incrementarFuerza(UserList(UserIndex), RandomNumber(obj.MinModificador, obj.MaxModificador), obj.DuracionEfecto)
                
                    Call QuitarUserInvItem(UserIndex, slot, 1)
                
                    EnviarPaquete Paquetes.SonidoTomarPociones, "", UserIndex, ToPCArea
                Case ePociones.Roja  'Pocion roja, restaura HP
            
                    'Usa el item
                    If .Stats.minHP = .Stats.MaxHP Then
                        Call QuitarUserInvItem(UserIndex, slot, 1)
                        EnviarPaquete Paquetes.WavSnd, Chr$(SND_BEBER), UserIndex, ToPCArea
                    Else
                
                        Dim cantidad As Integer
                    
                        ' El 10% de la vida del personaje, con un mínimo de 25 y un máximo de 35
                        'cantidad = Round(.Stats.MaxHP * 0.1)
                        'cantidad = maxi(cantidad, 25)
                        'cantidad = mini(cantidad, 35)
                        cantidad = 32
                        
                        AddtoVar .Stats.minHP, cantidad, .Stats.MaxHP
                    
                        Call QuitarUserInvItem(UserIndex, slot, 1)
                    
                        EnviarPaquete Paquetes.SonidoTomarPociones, "", UserIndex, ToPCArea
                    
                        Call SendUserVida(UserIndex)
                    End If
                Case ePociones.Azul
                    'Usa el item
                    Call AddtoVar(.Stats.MinMAN, Porcentaje(.Stats.MaxMAN, 5), .Stats.MaxMAN)
                    'Quitamos del inv el item
                    Call QuitarUserInvItem(UserIndex, slot, 1)
                    EnviarPaquete Paquetes.SonidoTomarPociones, "", UserIndex, ToPCArea
                    'EnviarPaquete Paquetes.WavSnd, Chr$(SND_BEBER), UserIndex, ToPCArea
                    Call SendUserMana(UserIndex)
                Case ePociones.Violeta
                    If .flags.Envenenado = 1 Then
                        .flags.Envenenado = 0
                        EnviarPaquete Paquetes.MensajeSimple, Chr$(160), UserIndex
                        EnviarPaquete Paquetes.EstaEnvenenado, "", UserIndex, ToIndex
                    End If
                
                    'Quitamos del inv el item
                    Call QuitarUserInvItem(UserIndex, slot, 1)
                
                    EnviarPaquete Paquetes.SonidoTomarPociones, "", UserIndex, ToPCArea
                    'EnviarPaquete Paquetes.WavSnd, Chr$(SND_BEBER), UserIndex, ToPCArea
                Case ePociones.Negra
                    If .flags.Privilegios = 0 Then 'Los gms no se pueden explotar
                        Call QuitarUserInvItem(UserIndex, slot, 1)
                        Call UserDie(UserIndex, True)
                        EnviarPaquete Paquetes.MensajeSimple, Chr$(161), UserIndex
                    End If
                Case ePociones.Energia
                    Call AddtoVar(.Stats.MinSta, .Stats.MaxSta * 0.1, .Stats.MaxSta)
                    Call QuitarUserInvItem(UserIndex, slot, 1)
                    EnviarPaquete Paquetes.SonidoTomarPociones, "", UserIndex, ToPCArea
                    'EnviarPaquete Paquetes.WavSnd, Chr$(SND_BEBER), UserIndex, ToPCArea
                    Call SendUserEsta(UserIndex)
            End Select
            
            EnviarPaquete Paquetes.ActualizaCantidadItem, Chr$(slot) & Codify(.Invent.Object(slot).Amount), UserIndex, ToIndex
        
        Case OBJTYPE_BEBIDA
            
            If .flags.Muerto = 1 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(26), UserIndex
                Exit Sub
            End If
            
            AddtoVar .Stats.minAgu, obj.MinSed, .Stats.MaxAGU
            
            .flags.Sed = 0
        
            Call EnviarHambreYsed(UserIndex)
            
            'Quitamos del inv el item
            Call QuitarUserInvItem(UserIndex, slot, 1)
            EnviarPaquete Paquetes.WavSnd, Chr$(SND_BEBER), UserIndex, ToPCArea
            
            Call UpdateUserInv(False, UserIndex, slot)

        Case OBJTYPE_LLAVES
            If .flags.Muerto = 1 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(26), UserIndex
                Exit Sub
            End If
        
            If .flags.TargetObj = 0 Then Exit Sub
            
            TargObj = ObjData(.flags.TargetObj)
            
            '¿El objeto clickeado es una puerta?
            If Not TargObj.ObjType = OBJTYPE_PUERTAS Then
                Exit Sub
            End If
                
            '¿Esta cerrada?
            If TargObj.Cerrada = 0 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(165), UserIndex
                Exit Sub
            End If
            
            '¿Cerrada con llave?
            If TargObj.Llave > 0 Then
               'Meto la llave y doy vuelta, permitiendo que se pueda abrir la puerta luego
               If TargObj.clave = obj.clave Then
                  MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).OBJInfo.ObjIndex _
                  = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).OBJInfo.ObjIndex).IndexCerrada
                  .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).OBJInfo.ObjIndex
                  EnviarPaquete Paquetes.MensajeSimple, Chr$(162), UserIndex
                  Exit Sub
               Else
                  EnviarPaquete Paquetes.MensajeSimple, Chr$(163), UserIndex
                  Exit Sub
               End If
            Else
               'Meto la llave y doy vuelta. Cierro al puerta
               If TargObj.clave = obj.clave Then
                  MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).OBJInfo.ObjIndex _
                  = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).OBJInfo.ObjIndex).IndexCerradaLlave
                  EnviarPaquete Paquetes.MensajeSimple, Chr$(164), UserIndex
                  .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).OBJInfo.ObjIndex
                  Exit Sub
               Else
                  EnviarPaquete Paquetes.MensajeSimple, Chr$(163), UserIndex
                  Exit Sub
               End If
            End If
    
        Case OBJTYPE_BOTELLAVACIA
        
            If .flags.Muerto = 1 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(26), UserIndex
                Exit Sub
            End If
            
            If Not HayAgua(.pos.map, .flags.TargetX, .flags.TargetY) Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(166), UserIndex
                Exit Sub
            End If
            
            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(.Invent.Object(slot).ObjIndex).IndexAbierta
             
            If MeterItemEnInventario(UserIndex, MiObj) Then
                Call QuitarUserInvItem(UserIndex, slot, 1) 'Call TirarItemAlPiso(.Pos, MiObj)
            Else
                EnviarPaquete Paquetes.mensajeinfo, "No tienes más espacio en el inventario.", UserIndex, ToIndex
            End If
            
            Call UpdateUserInv(False, UserIndex, slot)
    
        Case OBJTYPE_BOTELLALLENA
            If .flags.Muerto = 1 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(26), UserIndex
                Exit Sub
            End If
            
            ' Aumento la cantidad de agua
            AddtoVar .Stats.minAgu, obj.MinSed, .Stats.MaxAGU
            
            ' Desmarco que tiene sed
            .flags.Sed = 0
            ' Actualizo el Hambre y la Sed
            Call EnviarHambreYsed(UserIndex)
            
            ' Creamos la botella vacia
            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(.Invent.Object(slot).ObjIndex).IndexCerrada
            
            ' Quitamos la botella vacia
            Call InvUsuario.QuitarUserInvItem(UserIndex, slot, 1)
            Call InvUsuario.UpdateUserInv(False, UserIndex, slot)
            
            ' Mandamos el sonido de que la tomo
            If obj.Snd1 > 0 Then EnviarPaquete Paquetes.WavSnd, Chr$(obj.Snd1), UserIndex, ToPCArea
            
            ' La metemos en el inventario
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                 Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)
            End If
            
        Case OBJTYPE_HERRAMIENTAS
            If .flags.Muerto = 1 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(26), UserIndex
                Exit Sub
            End If
            If Not .Stats.MinSta > 0 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(167), UserIndex
                Exit Sub
            End If
            If .Invent.Object(slot).Equipped = 0 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(168), UserIndex
                Exit Sub
            End If
            
            Select Case ObjIndex
                Case MARTILLO_HERRERO
                    'Call Senddata(ToIndex, UserIndex, 0, "T01" & Herreria)
                Case SERRUCHO_CARPINTERO
                    Call EnivarObjConstruibles(UserIndex)
                    EnviarPaquete Paquetes.ShowCarp, "", UserIndex, ToIndex
            End Select
        
        Case OBJTYPE_PERGAMINOS
            If .flags.Muerto = 1 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(26), UserIndex
                Exit Sub
            End If
            
            If .flags.Hambre = 1 Or .flags.Sed = 1 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(169), UserIndex
                Exit Sub
            End If
            
            Call AgregarHechizo(UserIndex, slot)
            Call UpdateUserInv(False, UserIndex, slot)
            
       Case OBJTYPE_MINERALES
       
           If .flags.Muerto = 1 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(26), UserIndex
                Exit Sub
           End If
           
       Case OBJTYPE_INSTRUMENTOS
            If .flags.Muerto = 1 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(26), UserIndex
                Exit Sub
            End If
            EnviarPaquete Paquetes.WavSnd, Chr$(obj.Snd1), UserIndex, ToPCArea
       
       Case OBJTYPE_BARCOS
 
            Call DoNavega(UserList(UserIndex), obj, slot)
    
End Select

End With
'Actualiza
'Call SendUserStatsBox(UserIndex)
'Call UpdateUserInv(False, UserIndex, Slot)
End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)
Dim i As Integer, cad$

For i = 1 To UBound(ArmasHerrero)
    If ObjData(ArmasHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(eSkills.Herreria) \ ModHerreriA(UserList(UserIndex).clase) Then
            cad$ = cad$ & ITS(ObjData(ArmasHerrero(i)).MinHIT) & ITS(ObjData(ArmasHerrero(i)).MaxHIT) & ITS(ArmasHerrero(i))
    End If
Next i
EnviarPaquete Paquetes.EnviarArmasConstruibles, cad$, UserIndex
End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)
Dim i As Integer
Dim cad As String
Dim loopObjeto As Integer

cad = ""

For i = 1 To UBound(ObjCarpintero)
    If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(UserIndex).Stats.UserSkills(eSkills.Carpinteria) / ModCarpinteria(UserList(UserIndex).clase) Then
    
        For loopObjeto = LBound(ObjData(ObjCarpintero(i)).recursosNecesarios) To UBound(ObjData(ObjCarpintero(i)).recursosNecesarios)
            cad = cad & LongToString(ObjData(ObjCarpintero(i)).recursosNecesarios(loopObjeto).cantidad) & ITS(ObjCarpintero(i))
        Next loopObjeto

    End If

Next i

EnviarPaquete Paquetes.EnviarObjConstruibles, cad, UserIndex
End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)
Dim i As Integer, cad$
cad$ = ""
For i = 1 To UBound(ArmadurasHerrero)
    If ObjData(ArmadurasHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(eSkills.Herreria) / ModHerreriA(UserList(UserIndex).clase) Then _
        cad$ = cad$ & ITS(ObjData(ArmadurasHerrero(i)).MinDef) & ITS(ObjData(ArmadurasHerrero(i)).MaxDef) & ITS(ArmadurasHerrero(i))
     '   Debug.Print Len(cad$)
       ' Debug.Print ArmadurasHerrero(i)
Next i
EnviarPaquete Paquetes.EnviarArmadurasConstruibles, cad$, UserIndex, ToIndex
'Debug.Print Timer & "AA"
End Sub

Public Sub TirarTodo(personaje As User, Optional ByVal duenio As Integer = 0)
    Call TirarTodosLosItems(personaje, duenio)
    Call TirarOro(personaje.Stats.GLD, personaje, duenio)
End Sub

Public Function ItemSeCae(ByVal index As Integer) As Boolean
ItemSeCae = modObjeto.isFaccionario(ObjData(index)) = False And _
            ObjData(index).ObjType <> OBJTYPE_LLAVES And _
            ObjData(index).ObjType <> OBJTYPE_BARCOS And _
            ObjData(index).SeCae = 1
End Function

' Tiramos todos Los Objetos que tenga el Usuario
Sub TirarTodosLosItems(personaje As User, Optional ByVal duenio As Integer = 0)

Dim loopSlot As Byte
Dim NuevaPos As WorldPos
Dim ItemIndex As Integer
Dim obj As obj


For loopSlot = 1 To personaje.Stats.MaxItems

  ItemIndex = personaje.Invent.Object(loopSlot).ObjIndex

  obj.ObjIndex = ItemIndex
  obj.Amount = personaje.Invent.Object(loopSlot).Amount
    
  If ItemIndex > 0 Then
  
         If ItemSeCae(ItemIndex) Then
            NuevaPos.x = 0
            NuevaPos.y = 0
            NuevaPos.map = 0
            
            ' Si el personaje esta en el agua, los items se hunden.
            If HayAgua(personaje.pos.map, personaje.pos.x, personaje.pos.y) = False Then
                ' Buscamos una posición libre para poner el objeto cercana al usuario
                TileLibreParaObjeto personaje.pos, NuevaPos, obj
            End If
            
            ' ¿Encontre una posicion?
            If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then
                ' Tiramos todo la cantidad que puede llegar a tener
                Call InvUsuario.QuitarOBjetoYPonerEnSuelo(personaje, loopSlot, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.x, NuevaPos.y, duenio)
            Else
                Call InvUsuario.QuitarOBjetoSlot(personaje, loopSlot, MAX_INVENTORY_OBJS)
            End If
            
         End If
         
  End If
  
Next loopSlot
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean

ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function

Sub TirarTodosLosItemsNoNewbies(personaje As User, Optional ByVal duenio As Integer = 0)

Dim loopSlot As Byte
Dim NuevaPos As WorldPos
Dim ItemIndex As Integer
Dim obj As obj
If MapData(personaje.pos.map, personaje.pos.x, personaje.pos.y).Trigger = 6 Then
    Exit Sub
End If

Call TirarOro(personaje.Stats.GLD, personaje, duenio)

For loopSlot = 1 To MAX_INVENTORY_SLOTS

  ItemIndex = personaje.Invent.Object(loopSlot).ObjIndex
  obj.ObjIndex = personaje.Invent.Object(loopSlot).ObjIndex
  obj.Amount = personaje.Invent.Object(loopSlot).Amount
  
  If ItemIndex > 0 Then
         If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
                NuevaPos.x = 0
                NuevaPos.y = 0
                NuevaPos.map = 0
                
                TileLibreParaObjeto personaje.pos, NuevaPos, obj
                
                If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then
                    Call InvUsuario.QuitarOBjetoYPonerEnSuelo(personaje, loopSlot, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.x, NuevaPos.y, duenio)
                End If
         End If
  End If
  
Next loopSlot
End Sub
'Devuelve el ID (Identificador unico) del usuario
Public Function ObtenerIDUsuario(nombre As String) As Long
'¿Existe el personaje?
Dim infoPersonaje As ADODB.Recordset

sql = "SELECT ID FROM " & DB_NAME_PRINCIPAL & ".usuarios WHERE NickB = '" & mysql_real_escape_string(nombre) & "'"

Set infoPersonaje = conn.Execute(sql, , adCmdText)

If infoPersonaje.EOF Then
    ' El personaje no existe
    ObtenerIDUsuario = 0
Else
    ObtenerIDUsuario = val(infoPersonaje!id)
End If

'Cierro
infoPersonaje.Close
Set infoPersonaje = Nothing
End Function
