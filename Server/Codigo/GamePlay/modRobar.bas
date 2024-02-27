Attribute VB_Name = "modRobar"
Option Explicit

Private Sub RobarObjeto(ByRef Ladron As User, Victima As User)
    
    Dim i As Integer
    Dim MiObj As obj
    Dim cantidad As Integer
    
    Dim encontre As Boolean
    
    ' ¿Encontre objeto robable?
    encontre = False
    
    ' Vamos a buscar un objeto robable. Comenzamos por el principio o el final del inventario?
    If RandomNumber(1, 12) < 6 Then
        i = 1
        Do While Not encontre And i <= Victima.Stats.MaxItems
            'Hay objeto en este slot?
            If Victima.Invent.Object(i).ObjIndex > 0 Then
                If Victima.Invent.Object(i).Equipped = 0 Then
                    If ObjEsRobable(ObjData(Victima.Invent.Object(i).ObjIndex)) Then
                        If RandomNumber(1, 10) < 4 Then encontre = True
                    End If
               End If
            End If
            If Not encontre Then i = i + 1
        Loop
    Else
        i = 20
        Do While Not encontre And i > 0
          'Hay objeto en este slot?
          If Victima.Invent.Object(i).ObjIndex > 0 Then
            If Victima.Invent.Object(i).Equipped = 0 Then
                If ObjEsRobable(ObjData(Victima.Invent.Object(i).ObjIndex)) Then
                    If RandomNumber(1, 10) < 4 Then encontre = True
                End If
             End If
          End If
          If Not encontre Then i = i - 1
        Loop
    End If
    
    ' ¿Encontre algo?
    If Not encontre Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(122), Ladron.UserIndex
        Exit Sub
    End If
    
    ' Creamos el objeto
    MiObj.ObjIndex = Victima.Invent.Object(i).ObjIndex
    
    ' obtemos la cantidad que le vamos a sacar
    Select Case Ladron.Stats.UserSkills(eSkills.Robar)
        Case Is <= 60
            If EsMineral(MiObj.ObjIndex) Then
                cantidad = 100
            Else
                cantidad = RandomNumber(5, 10)
            End If
        Case Is <= 70
            If EsMineral(MiObj.ObjIndex) Then
                cantidad = 100
            Else
                cantidad = RandomNumber(5, 10)
            End If
        Case Is <= 80
            If EsMineral(MiObj.ObjIndex) Then
                cantidad = 200
            Else
                cantidad = RandomNumber(20, 25)
            End If
        Case Is <= 90
            If EsMineral(MiObj.ObjIndex) Then
                cantidad = 200
            Else
                cantidad = RandomNumber(20, 25)
            End If
        Case Is < 100
            If EsMineral(MiObj.ObjIndex) Then
                cantidad = 250
            Else
                cantidad = RandomNumber(30, 35)
            End If
        Case 100
            If EsMineral(MiObj.ObjIndex) Then
                cantidad = 300
            Else
                cantidad = RandomNumber(35, 40)
            End If
        Case Else
            cantidad = 1
    End Select
    
    ' Si la cantidad es mayor a lo qeu tiene en el inventario...
    If cantidad > Victima.Invent.Object(i).Amount Then
         cantidad = Victima.Invent.Object(i).Amount
    End If
    
    ' Seteamos la cantidad
    MiObj.Amount = cantidad
    
    ' Le quitamos
    Victima.Invent.Object(i).Amount = Victima.Invent.Object(i).Amount - cantidad
    
    If Victima.Invent.Object(i).Amount <= 0 Then
          Call QuitarUserInvItem(Victima.UserIndex, CByte(i), 1)
    End If
    Call UpdateUserInv(False, Victima.UserIndex, CByte(i))
    
    ' Se lo damos al ladron
    If Not MeterItemEnInventario(Ladron.UserIndex, MiObj) Then
        Call TirarItemAlPiso(Ladron.pos, MiObj)
    End If
    
    ' Informamos
    EnviarPaquete Paquetes.MensajeCompuesto, Chr$(31) & MiObj.Amount & "," & ObjData(MiObj.ObjIndex).Name, Ladron.UserIndex

End Sub

Public Sub Robar(ByRef Ladron As User, ByVal x As Integer, y As Integer)
    Dim wpaux As WorldPos
    
    ' ¿El mapa lo Admite?
    If MapInfo(Ladron.pos.map).AntiHechizosPts = 1 Then Exit Sub
            
    ' ¿Zona segura?
    If Not MapInfo(Ladron.pos.map).Pk Then Exit Sub

    ' Intervalo?
    If Not IntervaloPermiteAtacar(Ladron.UserIndex) Then Exit Sub
    
    ' Buscamos Personaje Clickeado
    Call LookatTile(Ladron.UserIndex, Ladron.pos.map, x, y)
    
    ' No selecciono a nadie o se selecciono a el mismo?
    If Ladron.flags.TargetUser = 0 Or Ladron.flags.TargetUser = Ladron.UserIndex Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(237), Ladron.UserIndex, ToIndex
        Exit Sub
    End If
    
    ' No se le puede robar a personajes muertos
    If UserList(Ladron.flags.TargetUser).flags.Muerto = 1 Then
        Exit Sub
    End If
    
    wpaux.map = Ladron.pos.map
    wpaux.x = x
    wpaux.y = y
            
    ' ¿Puede robar tan lejos?
    If Ladron.clase = eClases.Ladron Then
            If Ladron.Stats.ELV >= 25 Then
                    If distancia(wpaux, Ladron.pos) > 5 Then
                        EnviarPaquete Paquetes.MensajeSimple, Chr$(5), Ladron.UserIndex, ToIndex
                        Exit Sub
                    End If
            Else
                    If distancia(wpaux, Ladron.pos) > 2 Then
                        EnviarPaquete Paquetes.MensajeSimple, Chr$(5), Ladron.UserIndex, ToIndex
                        Exit Sub
                    End If
            End If
    Else
            If distancia(wpaux, Ladron.pos) > 2 Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(5), Ladron.UserIndex, ToIndex
                Exit Sub
            End If
    End If
                                
    
    'No aseguramos que el trigger le permite robar
    If (MapData(UserList(Ladron.flags.TargetUser).pos.map, UserList(Ladron.flags.TargetUser).pos.x, UserList(Ladron.flags.TargetUser).pos.y).Trigger And eTriggers.PosicionSegura) Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(236), Ladron.UserIndex, ToIndex
        Exit Sub
    End If
    
    If (MapData(Ladron.pos.map, Ladron.pos.x, Ladron.pos.y).Trigger And eTriggers.PosicionSegura) Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(236), Ladron.UserIndex, ToIndex
        Exit Sub
    End If
        
    Call DoRobar(Ladron, UserList(Ladron.flags.TargetUser))

End Sub

Private Sub DoRobar(ByRef Ladron As User, ByRef Victima As User)

Dim N As Integer

' Zona Seguro
If Ladron.flags.Seguro Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(120), Ladron.UserIndex
    Exit Sub
End If

' Energia
If Ladron.Stats.MinSta < 25 Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(167), Ladron.UserIndex
    Exit Sub
Else
    Call QuitarSta(Ladron.UserIndex, 25)
End If

' Solo pueden robar oculto
If Ladron.flags.Oculto = 0 Then Exit Sub

' No pueden robar en el Medio del Mar
If HayAgua(Ladron.pos.map, Ladron.pos.x, Ladron.pos.y) Then
    EnviarPaquete Paquetes.mensajeinfo, "No puedes robar en el medio del mar.", Ladron.UserIndex, ToIndex
    Exit Sub
End If

' No se pueden robar a Game Masters
If Victima.flags.Privilegios > 0 Then Exit Sub

' No pueden robar a personajes sin energia
If Victima.Stats.minham = 0 Or Victima.Stats.minAgu = 0 Then Exit Sub
If Victima.Stats.MinSta = 0 Then Exit Sub

' Robos entre la misma alineacion
If Ladron.faccion.ArmadaReal = 1 Then
    Call ExpulsarFaccionReal(Ladron.UserIndex)
ElseIf Victima.faccion.FuerzasCaos = 1 And Ladron.faccion.FuerzasCaos = 1 Then
    Call ExpulsarFaccionCaos(Ladron.UserIndex)
End If


N = 0

If Ladron.clase = eClases.Ladron Then
    
    If RandomNumber(1, 200 - Ladron.Stats.UserSkills(eSkills.Robar)) < 80 Then 'probabilida de robar
    
    Select Case Ladron.Stats.UserSkills(eSkills.Robar)
    
    Case Is <= 10
        N = RandomNumber(20, 70)
        If Victima.Stats.GLD = 0 Then EnviarPaquete Paquetes.MensajeCompuesto, Chr$(33) & Victima.Name, Ladron.UserIndex: Exit Sub
        If N > Victima.Stats.GLD Then N = Victima.Stats.GLD
        Victima.Stats.GLD = Victima.Stats.GLD - N
        Call AddtoVar(Ladron.Stats.GLD, N, MAXORO)
    Case Is <= 20
        N = RandomNumber(120, 220)
        If Victima.Stats.GLD = 0 Then EnviarPaquete Paquetes.MensajeCompuesto, Chr$(33) & Victima.Name, Ladron.UserIndex: Exit Sub
        If N > Victima.Stats.GLD Then N = Victima.Stats.GLD
        Victima.Stats.GLD = Victima.Stats.GLD - N
        Call AddtoVar(Ladron.Stats.GLD, N, MAXORO)
    Case Is <= 30
        N = RandomNumber(250, 370)
        If Victima.Stats.GLD = 0 Then EnviarPaquete Paquetes.MensajeCompuesto, Chr$(33) & Victima.Name, Ladron.UserIndex: Exit Sub
        If N > Victima.Stats.GLD Then N = Victima.Stats.GLD
        Victima.Stats.GLD = Victima.Stats.GLD - N
        Call AddtoVar(Ladron.Stats.GLD, N, MAXORO)
    Case Is <= 40
        N = RandomNumber(400, 520)
        If Victima.Stats.GLD = 0 Then EnviarPaquete Paquetes.MensajeCompuesto, Chr$(33) & Victima.Name, Ladron.UserIndex: Exit Sub
        If N > Victima.Stats.GLD Then N = Victima.Stats.GLD
        Victima.Stats.GLD = Victima.Stats.GLD - N
        Call AddtoVar(Ladron.Stats.GLD, N, MAXORO)
    Case Is <= 50
        N = RandomNumber(550, 670)
        If Victima.Stats.GLD = 0 Then EnviarPaquete Paquetes.MensajeCompuesto, Chr$(33) & Victima.Name, Ladron.UserIndex: Exit Sub
        If N > Victima.Stats.GLD Then N = Victima.Stats.GLD
        Victima.Stats.GLD = Victima.Stats.GLD - N
        Call AddtoVar(Ladron.Stats.GLD, N, MAXORO)
    Case Is <= 60
        If Victima.Stats.GLD = 0 Then
            EnviarPaquete Paquetes.MensajeCompuesto, Chr$(33) & Victima.Name, Ladron.UserIndex
        Else
            N = RandomNumber(700, 820)
            If N > Victima.Stats.GLD Then N = Victima.Stats.GLD
            Victima.Stats.GLD = Victima.Stats.GLD - N
            Call AddtoVar(Ladron.Stats.GLD, N, MAXORO)
        End If
        
            If Int(RandomNumber(0, 10)) <= 1 Then
                If TieneObjetosRobables(Victima.UserIndex) Then
                Call RobarObjeto(Ladron, Victima)
                Else
                EnviarPaquete Paquetes.mensajeinfo, Victima.Name & " no tiene objetos.", Ladron.UserIndex
                End If
            End If
    Case Is <= 70
        
        If Victima.Stats.GLD = 0 Then
            EnviarPaquete Paquetes.MensajeCompuesto, Chr$(33) & Victima.Name, Ladron.UserIndex
        Else
            N = RandomNumber(850, 970)
            If N > Victima.Stats.GLD Then N = Victima.Stats.GLD
            Victima.Stats.GLD = Victima.Stats.GLD - N
            Call AddtoVar(Ladron.Stats.GLD, N, MAXORO)
        End If
        If Int(RandomNumber(0, 10)) <= 2 Then
            If TieneObjetosRobables(Victima.UserIndex) Then
            Call RobarObjeto(Ladron, Victima)
            Else
            EnviarPaquete Paquetes.mensajeinfo, Victima.Name & " no tiene objetos.", Ladron.UserIndex
            End If
        End If
    
    Case Is <= 80
        If Victima.Stats.GLD = 0 Then
            EnviarPaquete Paquetes.MensajeCompuesto, Chr$(33) & Victima.Name, Ladron.UserIndex
        Else
            N = RandomNumber(1020, 1100)
            If N > Victima.Stats.GLD Then N = Victima.Stats.GLD
            Victima.Stats.GLD = Victima.Stats.GLD - N
            Call AddtoVar(Ladron.Stats.GLD, N, MAXORO)
        End If
        
        If Int(RandomNumber(0, 10)) <= 3 Then
            If TieneObjetosRobables(Victima.UserIndex) Then
            Call RobarObjeto(Ladron, Victima)
            Else
            EnviarPaquete Paquetes.mensajeinfo, Victima.Name & " no tiene objetos.", Ladron.UserIndex
            End If
        End If
    
    Case Is <= 99
       
        If Victima.Stats.GLD = 0 Then
            EnviarPaquete Paquetes.MensajeCompuesto, Chr$(33) & Victima.Name, Ladron.UserIndex
        Else
            N = RandomNumber(1150, 1220)
            If N > Victima.Stats.GLD Then N = Victima.Stats.GLD
            Victima.Stats.GLD = Victima.Stats.GLD - N
            Call AddtoVar(Ladron.Stats.GLD, N, MAXORO)
        End If
        
        If Int(RandomNumber(0, 10)) <= 4 Then
            If TieneObjetosRobables(Victima.UserIndex) Then
            Call RobarObjeto(Ladron, Victima)
            Else
            EnviarPaquete Paquetes.mensajeinfo, Victima.Name & " no tiene objetos.", Ladron.UserIndex
            End If
        End If
    
    Case 100
        If Victima.Stats.GLD = 0 Then
            EnviarPaquete Paquetes.MensajeCompuesto, Chr$(33) & Victima.Name, Ladron.UserIndex
        Else
            N = RandomNumber(1300, 1380)
            If N > Victima.Stats.GLD Then N = Victima.Stats.GLD
            Victima.Stats.GLD = Victima.Stats.GLD - N
            Call AddtoVar(Ladron.Stats.GLD, N, MAXORO)
        End If
        
        If Int(RandomNumber(0, 10)) <= 5 Then
            If TieneObjetosRobables(Victima.UserIndex) Then
            Call RobarObjeto(Ladron, Victima)
            Else
            EnviarPaquete Paquetes.mensajeinfo, Victima.Name & " no tiene objetos.", Ladron.UserIndex
            End If
        End If
    End Select
    'Exit Sub
    End If
Else 'No es clase ladron

    If RandomNumber(1, 200 - Ladron.Stats.UserSkills(eSkills.Robar)) < 20 Then
    N = RandomNumber(10, Ladron.Stats.UserSkills(eSkills.Robar) * 2)
    If Victima.Stats.GLD = 0 Then EnviarPaquete Paquetes.MensajeCompuesto, Chr$(33) & Victima.Name, Ladron.UserIndex: Exit Sub
    If N > Victima.Stats.GLD Then N = Victima.Stats.GLD
    Victima.Stats.GLD = Victima.Stats.GLD - N
    Call AddtoVar(Ladron.Stats.GLD, N, MAXORO)
    End If

End If 'Cerramos el no es ladron!


If N > 0 Then
    EnviarPaquete Paquetes.MensajeCompuesto, Chr$(29) & N & "," & Victima.Name, Ladron.UserIndex
    EnviarPaquete Paquetes.MensajeCompuesto, Chr$(30) & N, Victima.UserIndex
    EnviarPaquete Paquetes.WavSnd, Chr$(SOUND_SWING), Victima.UserIndex
    EnviarPaquete Paquetes.WavSnd, Chr$(SOUND_SWING), Ladron.UserIndex
    EnviarPaquete Paquetes.EnviarOro, Codify(Victima.Stats.GLD), Victima.UserIndex
    Call SendUserStatsBox(Ladron.UserIndex)
    Call SubirSkill(Ladron.UserIndex, eSkills.Robar)
    Call AddtoVar(Ladron.Reputacion.LadronesRep, 10, MAXREP)
End If

End Sub
