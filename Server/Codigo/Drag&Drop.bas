Attribute VB_Name = "DragAndDrop"
Option Explicit

Public Enum eDestinoObjeto
    suelo = 1
    criatura = 2
    Usuario = 3
End Enum

'En este modulo pongo las cosas y bugueo y asi marce
'cuando ve un bug, ya sabe en q modulo buscar
'^^
'ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
Sub DraguedClick(personaje As User, ByVal x As Integer, ByVal y As Integer, ByVal slot As Byte, ByVal cant As Integer)
    '[SISTEMA DE DRAG&DROP] Desarrollado por Wizard(Leandro)
    'Primero, buscamos un usuario en el tile exacto que
    'clikeo el jugador, sino hay ningun user, buscamos un npc
    'Sino encontramos nada, deberiamos buscar users o npcs en el Tile Y +1
    'No habiendo encontrado nada solamente tiramos el objeto
    'Variables
     Dim objeto As obj
     Dim Adonde As WorldPos
     Dim usuarioDestino As Integer
     Dim destino As eDestinoObjeto
     Dim destinoIndex As Integer
    '*****************************************
    'Chequeos
    If personaje.flags.Muerto = 1 Then Exit Sub
    If personaje.flags.Trabajando = True Then Exit Sub
    If personaje.flags.Comerciando = True Then Exit Sub
    If personaje.flags.Meditando = True Then Exit Sub
        
    If personaje.flags.Privilegios = 1 Then
        EnviarPaquete Paquetes.mensajeinfo, "Los consejeros no pueden arrojar items.", personaje.UserIndex, ToIndex
        Exit Sub
    End If
        
    If cant = 0 Then Exit Sub
    If personaje.Invent.Object(slot).ObjIndex = 0 Then Exit Sub
    If personaje.Invent.Object(slot).Amount < cant Then cant = personaje.Invent.Object(slot).Amount
        
    ' Seteamos el objeto que vamos a tirar
    objeto.Amount = cant
    objeto.ObjIndex = personaje.Invent.Object(slot).ObjIndex
    
    ' No se puede tirar la barca desequipada
    If ObjData(objeto.ObjIndex).ObjType = OBJTYPE_BARCOS And personaje.Invent.Object(slot).Equipped = 1 Then
        EnviarPaquete Paquetes.mensajeinfo, "Debes desequipar el barco antes de lanzarlo.", personaje.UserIndex, ToIndex
        Exit Sub
    End If
    
    ' No se puede tirar objetos newbies
    If ObjData(objeto.ObjIndex).Newbie = 1 Then
        EnviarPaquete Paquetes.mensajeinfo, "No puedes arrojar items newbies.", personaje.UserIndex, ToIndex
        Exit Sub
    End If
    
    ' No se puede tirar la armadura del dragon
    If objeto.ObjIndex = Objetos_Constantes.ARMADURA_DRAGON_E Or _
            objeto.ObjIndex = Objetos_Constantes.ARMADURA_DRAGON_H Or _
            objeto.ObjIndex = Objetos_Constantes.ARMADURA_DRAGON_M Then
        EnviarPaquete Paquetes.mensajeinfo, "No puedes tirar la Armadura del Dragón, el enfado del Dragón te desterrará de estas tierras.", personaje.UserIndex
        Exit Sub
    End If
    
    'No puede hacer un drag de un objeto faccionario
    If modObjeto.isFaccionario(ObjData(objeto.ObjIndex)) Then
        EnviarPaquete Paquetes.mensajeinfo, "No puedes lanzar un item faccionario.", personaje.UserIndex, ToIndex
        Exit Sub
    End If
        
    'Chequeos de propiedades del objeto
    If personaje.flags.Navegando = 1 Then
        If ObjData(objeto.ObjIndex).ObjType = OBJTYPE_BARCOS Then
            EnviarPaquete Paquetes.mensajeinfo, "No puedes arrojar tu barca si estas navegando!!.", personaje.UserIndex, ToIndex
            Exit Sub
        ElseIf personaje.Invent.Object(slot).Equipped = 1 Then
            EnviarPaquete Paquetes.mensajeinfo, "No puedes arrojar un item si lo tienes equipado y estas navegando!!.", personaje.UserIndex, ToIndex
            Exit Sub
        End If
    End If
        
    '********* Busqueda del destino del objeto
    
    ' ¿Posicion valida?
    If Not SV_PosicionesValidas.existePosicionMundo(personaje.pos.map, x, y) Then
        Exit Sub
    End If
    
    usuarioDestino = MapData(personaje.pos.map, x, y).UserIndex
        
    If usuarioDestino > 0 Then
        ' Si es un GM, no vale :p
        If UserList(usuarioDestino).flags.Privilegios > 0 Then usuarioDestino = 0
    End If
    
    ' Por defecto suelo
    destino = eDestinoObjeto.suelo
    destinoIndex = 0
        
    If usuarioDestino > 0 And usuarioDestino <> personaje.UserIndex Then ' Usuario
        destino = eDestinoObjeto.Usuario
        destinoIndex = usuarioDestino
    ElseIf MapData(personaje.pos.map, x, y).npcIndex > 0 Then  ' Criatura
        destino = eDestinoObjeto.criatura
        destinoIndex = MapData(personaje.pos.map, x, y).npcIndex
    ElseIf SV_PosicionesValidas.existePosicionMundo(personaje.pos.map, x, y + 1) Then ' Chequeo si hay posicion valida por aproximacion simple
    
        If MapData(personaje.pos.map, x, y + 1).npcIndex > 0 Then ' ¿Hay una criatura?
            destino = eDestinoObjeto.criatura
            destinoIndex = MapData(personaje.pos.map, x, y + 1).npcIndex
        End If
    End If
        
    ' ¿El evento lo permite?
   If Not personaje.evento Is Nothing Then
        If Not personaje.evento.puedeTirarObjeto(personaje.UserIndex, objeto.ObjIndex, objeto.Amount, destino, destinoIndex) Then
            Exit Sub
       End If
   End If
        
    If destino = eDestinoObjeto.Usuario Then
            
            ' Chequeamos el destino
            If Not UserList(destinoIndex).flags.PermitirDragAndDrop Then
                EnviarPaquete Paquetes.mensajeinfo, "El usuario no quiere tus objetos.", personaje.UserIndex, ToIndex
                Exit Sub
            End If
            
            If UserList(destinoIndex).flags.Muerto Then
                EnviarPaquete Paquetes.mensajeinfo, "¡No puedes darle un objeto a un muerto!.", personaje.UserIndex, ToIndex
                Exit Sub
            End If
            
            If UserList(destinoIndex).flags.Comerciando = True Then
                EnviarPaquete Paquetes.mensajeinfo, "¡No puedes darle un objeto al usuario en este momento!.", personaje.UserIndex, ToIndex
                Exit Sub
            End If
            
            ' ¿ Tiene lugar el usuario?
            If Not InvUsuario.tieneLugar(UserList(destinoIndex), objeto) Then
                EnviarPaquete Paquetes.mensajeinfo, UserList(destinoIndex).Name & " no tiene lugar para guardar el objeto que le quieras dar.", personaje.UserIndex, ToIndex
                Exit Sub
            End If
            
            ' Por si tiene el objeto equipado
            Call chequeo(personaje.UserIndex, slot, personaje.Invent.Object(slot).Amount - objeto.Amount)
            
            ' OJO ¿Luego del chequeo lo tiene todvia?
            If Not UserList(personaje.UserIndex).Invent.Object(slot).ObjIndex = objeto.ObjIndex Or UserList(personaje.UserIndex).Invent.Object(slot).Amount < objeto.Amount Then
                Exit Sub
            End If
                        
            ' Primero lo quito
            QuitarUserInvItem personaje.UserIndex, slot, objeto.Amount
                                    
            ' Actualizo el inventario
            UpdateUserInv False, personaje.UserIndex, slot
            
            'Me fijo si puedo meter el item en el inventario
            If Not MeterItemEnInventario(destinoIndex, objeto) Then
                Call LogError("No se le pudo guardar un objeto al usuario, pero previamente se habia chequeado si tenia lugar! Objeto: " & ObjData(objeto.ObjIndex).Name & " Cantidad " & objeto.Amount & " Usuario: " & personaje.Name)
                Exit Sub
            End If
            
            ' Si es un Game Master Guardo Logs
            If personaje.flags.Privilegios > 0 Then
                Call LogGM(personaje.id, "Cantidad:" & objeto.Amount & " Objeto:" & ObjData(objeto.ObjIndex).Name & " Usuario: " & UserList(destinoIndex).Name & " Mapa: " & personaje.pos.map & " X: " & personaje.pos.x & " Y: " & personaje.pos.y, "DRAGPJ")
            End If
        
            ' Aviso
            If cant = 1 Then
                EnviarPaquete Paquetes.mensajeinfo, "¡" & personaje.Name & " te ha dado su " & ObjData(objeto.ObjIndex).Name & "!.", destinoIndex, ToIndex
                EnviarPaquete Paquetes.mensajeinfo, "¡Le has arrojado tu " & ObjData(objeto.ObjIndex).Name & " a " & UserList(usuarioDestino).Name & "!.", personaje.UserIndex, ToIndex
            Else
                EnviarPaquete Paquetes.mensajeinfo, "¡Has recibido " & cant & " " & ObjData(objeto.ObjIndex).Name & " de " & personaje.Name & "!.", destinoIndex, ToIndex
                EnviarPaquete Paquetes.mensajeinfo, "¡Le has arrojado " & cant & " " & ObjData(objeto.ObjIndex).Name & " a " & UserList(usuarioDestino).Name & "!.", personaje.UserIndex, ToIndex
            End If
    
    ElseIf destino = eDestinoObjeto.criatura Then
    
        ' Por si tiene el objeto equipado
         Call chequeo(personaje.UserIndex, slot, personaje.Invent.Object(slot).Amount - objeto.Amount)
         
        ' OJO ¿Luego del chequeo lo tiene todvia?
        If Not UserList(personaje.UserIndex).Invent.Object(slot).ObjIndex = objeto.ObjIndex Or UserList(personaje.UserIndex).Invent.Object(slot).Amount < objeto.Amount Then
            Exit Sub
        End If
            
        'Destino = NPC Banquero o comerciante
        If NpcList(destinoIndex).NPCtype = NPCTYPE_BANQUERO Then
                          
              UserDejaObj personaje.UserIndex, slot, objeto.Amount
              
              UpdateUserInv False, personaje.UserIndex, slot

              EnviarPaquete Paquetes.mensajeinfo, "Has depositado un objeto.", personaje.UserIndex
              
        ElseIf NpcList(destinoIndex).Comercia = 1 Then
              
            personaje.flags.TargetNPC = destinoIndex
                            
            NpcCompraObj personaje, slot, objeto.Amount
              
            Call UpdateUserInv(False, personaje.UserIndex, slot)
              
            ' Si es un Game Master Guardo Logs
            If personaje.flags.Privilegios > 0 Then
                Call LogGM(personaje.id, "Cantidad:" & objeto.Amount & " Objeto:" & ObjData(objeto.ObjIndex).Name & " NPC: " & NpcList(destinoIndex).Name & " Mapa: " & personaje.pos.map & " X: " & personaje.pos.x & " Y: " & personaje.pos.y, "DRAGNPC")
            End If
        
            EnviarPaquete Paquetes.EnviarOro, Codify(personaje.Stats.GLD), personaje.UserIndex, ToIndex
    
        End If
        
    ElseIf destino = eDestinoObjeto.suelo Then
    
        Adonde.map = personaje.pos.map
        Adonde.y = y
        Adonde.x = x
        
        If MapInfo(Adonde.map).Pk = False Then
            EnviarPaquete Paquetes.mensajeinfo, "No se permite arrojar objetos en zonas seguras", personaje.UserIndex, ToIndex
            Exit Sub
        End If

        If distancia(personaje.pos, Adonde) >= 6 Then
            If personaje.Stats.UserSkills(Supervivencia) < 90 Then
                
                ' Pongo Min y Maximo para no pasarme de los limites validos
                If personaje.Stats.UserSkills(Supervivencia) >= 80 Then
                    x = RandomNumber(maxi(x - 1, SV_Constantes.X_MINIMO_JUGABLE), mini(x + 1, SV_Constantes.X_MAXIMO_JUGABLE))
                    y = RandomNumber(maxi(y - 1, SV_Constantes.Y_MINIMO_JUGABLE), mini(y + 1, SV_Constantes.Y_MAXIMO_JUGABLE))
                ElseIf personaje.Stats.UserSkills(Supervivencia) > 70 Then
                    x = RandomNumber(maxi(x - 2, SV_Constantes.X_MINIMO_JUGABLE), mini(x + 1, SV_Constantes.X_MAXIMO_JUGABLE))
                    y = RandomNumber(maxi(y - 2, SV_Constantes.Y_MINIMO_JUGABLE), mini(y + 1, SV_Constantes.Y_MAXIMO_JUGABLE))
                ElseIf personaje.Stats.UserSkills(Supervivencia) > 60 Then
                    x = RandomNumber(maxi(x - 2, SV_Constantes.X_MINIMO_JUGABLE), mini(x + 2, SV_Constantes.X_MAXIMO_JUGABLE))
                    y = RandomNumber(maxi(y - 2, SV_Constantes.Y_MINIMO_JUGABLE), mini(y + 2, SV_Constantes.Y_MAXIMO_JUGABLE))
                Else
                    x = RandomNumber(maxi(x - 2, SV_Constantes.X_MINIMO_JUGABLE), mini(x + 2, SV_Constantes.X_MAXIMO_JUGABLE))
                    y = RandomNumber(maxi(y - 2, SV_Constantes.Y_MINIMO_JUGABLE), mini(y + 2, SV_Constantes.Y_MAXIMO_JUGABLE))
                End If
                
                EnviarPaquete Paquetes.mensajeinfo, "¡Lanzas impresisamente!.", personaje.UserIndex, ToIndex
           End If
        End If
        
        If HayAgua(personaje.pos.map, x, y) Then
            EnviarPaquete Paquetes.mensajeinfo, "No puedes arrojar un objeto al agua!.", personaje.UserIndex, ToIndex
            Exit Sub
        End If
        
        'Chequeo para no traspasar paredes
        If ((MapData(personaje.pos.map, personaje.pos.x, personaje.pos.y).Trigger And eTriggers.BajoTecho) And Not (MapData(personaje.pos.map, x, y).Trigger And eTriggers.BajoTecho)) _
                Or (Not (MapData(personaje.pos.map, personaje.pos.x, personaje.pos.y).Trigger And eTriggers.BajoTecho) And (MapData(personaje.pos.map, x, y).Trigger And eTriggers.BajoTecho)) Then
            
            EnviarPaquete Paquetes.mensajeinfo, "No puedes arrojar un objeto ahí, no puedes traspasar la pared!.", personaje.UserIndex, ToIndex
            Exit Sub
         End If
         
        ' Chequeamos el tema del equipamiento
        Call chequeo(personaje.UserIndex, slot, personaje.Invent.Object(slot).Amount - objeto.Amount)
        
        ' OJO ¿Luego del chequeo lo tiene todvia?
        If Not UserList(personaje.UserIndex).Invent.Object(slot).ObjIndex = objeto.ObjIndex Or UserList(personaje.UserIndex).Invent.Object(slot).Amount < objeto.Amount Then
            Exit Sub
        End If
            
        ' Quitamos el objeto del usuario
        QuitarUserInvItem personaje.UserIndex, slot, objeto.Amount
            
        ' Actualizamos el nventario
        Call UpdateUserInv(False, personaje.UserIndex, slot)
        
        ' Tiramos al Piso
        TirarItemAlPisoConDragAndDrop personaje.pos.map, objeto, x, y
        
        ' Si es un Game Master Guardo Logs
        If personaje.flags.Privilegios > 0 Then
             Call LogGM(personaje.id, "Cantidad:" & objeto.Amount & " Objeto:" & ObjData(objeto.ObjIndex).Name & " Mapa: " & personaje.pos.map & " X: " & Adonde.x & " Y: " & Adonde.y & " PJs:" & modMapa.listarPersonajesOnline(MapInfo(personaje.pos.map)), "DRAGPISO")
        End If
    End If
    
End Sub

Public Function TirarItemAlPisoConDragAndDrop(mapa As Integer, obj As obj, x As Integer, y As Integer) As WorldPos
    Dim NuevaPos As WorldPos
    Dim Posicion As WorldPos
    
    NuevaPos.x = 0
    NuevaPos.y = 0
    Posicion.map = mapa
    Posicion.x = x
    Posicion.y = y
    
    If MapData(mapa, Posicion.x, Posicion.y).OBJInfo.ObjIndex > 0 And MapData(mapa, Posicion.x, Posicion.y).OBJInfo.ObjIndex = obj.ObjIndex And MapData(mapa, Posicion.x, Posicion.y).OBJInfo.Amount + obj.Amount <= 10000 Then
     
        obj.Amount = MapData(mapa, Posicion.x, Posicion.y).OBJInfo.Amount + obj.Amount
     
        Call MakeObj(ToMap, 0, Posicion.map, _
             obj, Posicion.map, Posicion.x, Posicion.y)
             TirarItemAlPisoConDragAndDrop = NuevaPos
    Else
        Call TileLibreParaObjeto(Posicion, NuevaPos, obj)
    
        If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then
              Call MakeObj(ToMap, 0, Posicion.map, _
              obj, Posicion.map, NuevaPos.x, NuevaPos.y)
              TirarItemAlPisoConDragAndDrop = NuevaPos
        End If
    End If
    
End Function



'---------------------------------------------------------------------------------------
' Procedure : ChangeItemSlot
' DateTime  : 18/02/2007 19:11
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub ChangeItemSlot(Item1 As Byte, Item2 As Byte, ByVal UserIndex As Integer)
   'El item 1 es el que la persona agarro y movio
    Dim Obj2 As obj
    Dim TempInt As Integer

    If Item2 = Item1 Or Item2 <= 0 Or Item1 < 0 Then Exit Sub
    'tiene los slots que quiere cambiar habilitados?
    
    With UserList(UserIndex)
    
        If Item2 > .Stats.MaxItems Or Item1 > .Stats.MaxItems Then Exit Sub
        ''If UserList(UserIndex).Invent.Object(Item1).Equipped = 1 And ObjData(UserList(UserIndex).Invent.Object(Item1).ObjIndex).ObjType = OBJTYPE_BARCOS Then
        ' EnviarPaquete Paquetes.Mensajeinfo, "Debes desequipar la barca antes de moverla.", UserIndex, ToIndex
        ' Exit Sub
        ' End If
        Obj2.Amount = .Invent.Object(Item2).Amount
        Obj2.ObjIndex = .Invent.Object(Item2).ObjIndex
        TempInt = .Invent.Object(Item2).Equipped
     
        If .Invent.Object(Item1).ObjIndex = Obj2.ObjIndex And Obj2.Amount + .Invent.Object(Item1).Amount <= 10000 Then
            If TempInt = 1 Or .Invent.Object(Item1).Equipped = 1 Then .Invent.Object(Item2).Equipped = 1
            .Invent.Object(Item2).Amount = Obj2.Amount + .Invent.Object(Item1).Amount
            .Invent.Object(Item1).Amount = 0
            .Invent.Object(Item1).ObjIndex = 0
            .Invent.Object(Item1).Equipped = 0
        Else
            .Invent.Object(Item2).Equipped = .Invent.Object(Item1).Equipped
            .Invent.Object(Item2).Amount = .Invent.Object(Item1).Amount
            .Invent.Object(Item2).ObjIndex = .Invent.Object(Item1).ObjIndex
    
            .Invent.Object(Item1).Equipped = TempInt
            .Invent.Object(Item1).ObjIndex = Obj2.ObjIndex
            .Invent.Object(Item1).Amount = Obj2.Amount
        End If
    
        If .Invent.Object(Item2).Equipped = 1 Then
            Call ActualizarObjSlots(UserIndex, Item2)
        End If
    
        If .Invent.Object(Item1).Equipped = 1 Then
            Call ActualizarObjSlots(UserIndex, Item1)
        End If
        
        If .Invent.BarcoSlot = Item2 Then
            .Invent.BarcoSlot = Item1
        ElseIf .Invent.BarcoSlot = Item1 Then
            .Invent.BarcoSlot = Item2
        End If
            
    
    End With
    
    Call UpdateUserInv(True, UserIndex, 0)

End Sub

Private Sub ActualizarObjSlots(ByVal UserIndex As Integer, ByVal slot As Byte)
Dim obj As ObjData
Dim index As Integer

obj = ObjData(UserList(UserIndex).Invent.Object(slot).ObjIndex)
index = UserList(UserIndex).Invent.Object(slot).ObjIndex

Select Case obj.ObjType
    Case OBJTYPE_WEAPON
        UserList(UserIndex).Invent.WeaponEqpSlot = slot
   Case OBJTYPE_ANILLOS
        UserList(UserIndex).Invent.AnilloEqpSlot = slot
    Case OBJTYPE_HERRAMIENTAS, OBJTYPE_MINERALES
        UserList(UserIndex).Invent.HerramientaEqpSlot = slot
        UserList(UserIndex).Invent.HerramientaEqpObjIndex = index
    Case OBJTYPE_FLECHAS
        UserList(UserIndex).Invent.MunicionEqpSlot = slot
    Case OBJTYPE_ARMOUR
         Select Case obj.subTipo
            Case OBJTYPE_ARMADURA
                    UserList(UserIndex).Invent.ArmourEqpSlot = slot
            Case OBJTYPE_CASCO
                    UserList(UserIndex).Invent.CascoEqpSlot = slot
            Case OBJTYPE_ESCUDO
                    UserList(UserIndex).Invent.EscudoEqpSlot = slot
     End Select
    Case OBJTYPE_BARCOS
        UserList(UserIndex).Invent.BarcoEqpSlot = slot
End Select
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChangeItemSlotBoveda
' DateTime  : 18/02/2007 19:11
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub ChangeItemSlotBoveda(Item1 As Byte, Item2 As Byte, ByVal UserIndex As Integer)
   'El item 1 es el que la persona agarro y movio
    Dim Obj2 As obj

    If Item2 = Item1 Or Item2 <= 0 Or Item1 < 0 Then Exit Sub
    'tiene los slots que quiere cambiar habilitados?
    If Item2 > MAX_BANCOINVENTORY_SLOTS Or Item1 > MAX_BANCOINVENTORY_SLOTS Then Exit Sub

    Obj2.Amount = UserList(UserIndex).BancoInvent.Object(Item2).Amount
    Obj2.ObjIndex = UserList(UserIndex).BancoInvent.Object(Item2).ObjIndex
    
    If UserList(UserIndex).BancoInvent.Object(Item1).ObjIndex = Obj2.ObjIndex And Obj2.Amount + UserList(UserIndex).BancoInvent.Object(Item1).Amount <= 10000 Then
       UserList(UserIndex).BancoInvent.Object(Item2).Amount = Obj2.Amount + UserList(UserIndex).BancoInvent.Object(Item1).Amount
       UserList(UserIndex).BancoInvent.Object(Item1).Amount = 0
       UserList(UserIndex).BancoInvent.Object(Item1).ObjIndex = 0
    Else
        UserList(UserIndex).BancoInvent.Object(Item2).Amount = UserList(UserIndex).BancoInvent.Object(Item1).Amount
        UserList(UserIndex).BancoInvent.Object(Item2).ObjIndex = UserList(UserIndex).BancoInvent.Object(Item1).ObjIndex
    
        UserList(UserIndex).BancoInvent.Object(Item1).ObjIndex = Obj2.ObjIndex
        UserList(UserIndex).BancoInvent.Object(Item1).Amount = Obj2.Amount
    End If
    
    Call UpdateBanUserInv(False, UserIndex, Item1)
    Call UpdateBanUserInv(False, UserIndex, Item2)
End Sub

Private Sub chequeo(UserIndex As Integer, slot As Byte, cant As Integer)
If UserList(UserIndex).Invent.Object(slot).Equipped = 1 And cant <= 0 Then
    Desequipar UserIndex, slot
End If
End Sub



