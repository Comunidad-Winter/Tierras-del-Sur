Attribute VB_Name = "ModFacciones"
Option Explicit

Private Type Armaduras
    BarDruiCazAseH As Integer
    BarDruiCazAseG As Integer
    ClerigoH As Integer
    ClerigoG As Integer
    PalGueH As Integer
    PalGueG As Integer
    MagDruiHM As Integer
    MagDruiHH As Integer
    MagDrioG As Integer
End Type


Private Type info
    Matados As Integer
    oro As Long
    Nivel As Byte
    Armadura As Armaduras
End Type

Private RequisitosReal(1 To 5) As info
Private RequisitosCaos(1 To 5) As info

Public Sub EnlistarPersonaje(personaje As User)
    Dim npcIndex As Integer
    
    npcIndex = personaje.flags.TargetNPC
    
    '¿Clikeo un npc?
    If npcIndex = 0 Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(131), personaje.UserIndex
        Exit Sub
    End If
            
    '¿Esta demasiado lejos?
    If distancia(personaje.pos, NpcList(npcIndex).pos) > 10 Then
        EnviarPaquete Paquetes.MensajeSimple, Chr$(4), personaje.UserIndex
        Exit Sub
    End If
    
    '¿Es un lider?
    If personaje.flags.TargetNpcTipo <> NPCTYPE_NOBLE Then
        EnviarPaquete Paquetes.DescNpc, Chr$(92) & ITS(NpcList(npcIndex).Char.charIndex), personaje.UserIndex
        Exit Sub
    End If
    
    ' ¿A donde se quiere enlistar?
    If NpcList(npcIndex).faccion = eAlineaciones.Real Then
        EnlistarArmadaReal personaje, NpcList(npcIndex)
    Else
        EnlistarCaos personaje, NpcList(npcIndex)
    End If
End Sub

            
Public Sub CargarRequisitos()
'ARMADA
RequisitosReal(1).oro = 0
RequisitosReal(1).Matados = 100
RequisitosReal(1).Nivel = 30

RequisitosReal(2).oro = 50000
RequisitosReal(2).Matados = 300
RequisitosReal(2).Nivel = 32

RequisitosReal(3).oro = 100000
RequisitosReal(3).Matados = 500
RequisitosReal(3).Nivel = 36

RequisitosReal(4).oro = 250000
RequisitosReal(4).Matados = 700
RequisitosReal(4).Nivel = 38

RequisitosReal(5).oro = 750000
RequisitosReal(5).Matados = 1000
RequisitosReal(5).Nivel = 40

'CAOS
RequisitosCaos(1).oro = 0
RequisitosCaos(1).Matados = 150
RequisitosCaos(1).Nivel = 25

RequisitosCaos(2).oro = 50000
RequisitosCaos(2).Matados = 250
RequisitosCaos(2).Nivel = 27

RequisitosCaos(3).oro = 100000
RequisitosCaos(3).Matados = 450
RequisitosCaos(3).Nivel = 30

RequisitosCaos(4).oro = 250000
RequisitosCaos(4).Matados = 650
RequisitosCaos(4).Nivel = 34

RequisitosCaos(5).oro = 750000
RequisitosCaos(5).Matados = 850
RequisitosCaos(5).Nivel = 37

'Armaduras caos
RequisitosCaos(2).Armadura.BarDruiCazAseH = 743
RequisitosCaos(2).Armadura.BarDruiCazAseG = 744
RequisitosCaos(2).Armadura.ClerigoH = 745
RequisitosCaos(2).Armadura.ClerigoG = 746
RequisitosCaos(2).Armadura.PalGueH = 747
RequisitosCaos(2).Armadura.PalGueG = 748
RequisitosCaos(2).Armadura.MagDruiHM = 749
RequisitosCaos(2).Armadura.MagDruiHH = 750
RequisitosCaos(2).Armadura.MagDrioG = 751

RequisitosCaos(3).Armadura.BarDruiCazAseH = 752
RequisitosCaos(3).Armadura.BarDruiCazAseG = 753
RequisitosCaos(3).Armadura.ClerigoH = 754
RequisitosCaos(3).Armadura.ClerigoG = 755
RequisitosCaos(3).Armadura.PalGueH = 756
RequisitosCaos(3).Armadura.PalGueG = 757
RequisitosCaos(3).Armadura.MagDruiHM = 758
RequisitosCaos(3).Armadura.MagDruiHH = 759
RequisitosCaos(3).Armadura.MagDrioG = 760

RequisitosCaos(4).Armadura.BarDruiCazAseH = 761
RequisitosCaos(4).Armadura.BarDruiCazAseG = 762
RequisitosCaos(4).Armadura.ClerigoH = 763
RequisitosCaos(4).Armadura.ClerigoG = 764
RequisitosCaos(4).Armadura.PalGueH = 765
RequisitosCaos(4).Armadura.PalGueG = 766
RequisitosCaos(4).Armadura.MagDruiHM = 767
RequisitosCaos(4).Armadura.MagDruiHH = 768
RequisitosCaos(4).Armadura.MagDrioG = 769

RequisitosCaos(5).Armadura.BarDruiCazAseH = 770
RequisitosCaos(5).Armadura.BarDruiCazAseG = 771
RequisitosCaos(5).Armadura.ClerigoH = 772
RequisitosCaos(5).Armadura.ClerigoG = 773
RequisitosCaos(5).Armadura.PalGueH = 774
RequisitosCaos(5).Armadura.PalGueG = 775
RequisitosCaos(5).Armadura.MagDruiHM = 776
RequisitosCaos(5).Armadura.MagDruiHH = 777
RequisitosCaos(5).Armadura.MagDrioG = 778

'Armaduras Real
RequisitosReal(2).Armadura.BarDruiCazAseH = 788
RequisitosReal(2).Armadura.BarDruiCazAseG = 789
RequisitosReal(2).Armadura.ClerigoH = 790
RequisitosReal(2).Armadura.ClerigoG = 791
RequisitosReal(2).Armadura.PalGueH = 792
RequisitosReal(2).Armadura.PalGueG = 793
RequisitosReal(2).Armadura.MagDruiHM = 794
RequisitosReal(2).Armadura.MagDruiHH = 795
RequisitosReal(2).Armadura.MagDrioG = 796

RequisitosReal(3).Armadura.BarDruiCazAseH = 797
RequisitosReal(3).Armadura.BarDruiCazAseG = 798
RequisitosReal(3).Armadura.ClerigoH = 799
RequisitosReal(3).Armadura.ClerigoG = 800
RequisitosReal(3).Armadura.PalGueH = 801
RequisitosReal(3).Armadura.PalGueG = 802
RequisitosReal(3).Armadura.MagDruiHM = 803
RequisitosReal(3).Armadura.MagDruiHH = 804
RequisitosReal(3).Armadura.MagDrioG = 805

RequisitosReal(4).Armadura.BarDruiCazAseH = 806
RequisitosReal(4).Armadura.BarDruiCazAseG = 807
RequisitosReal(4).Armadura.ClerigoH = 808
RequisitosReal(4).Armadura.ClerigoG = 809
RequisitosReal(4).Armadura.PalGueH = 810
RequisitosReal(4).Armadura.PalGueG = 811
RequisitosReal(4).Armadura.MagDruiHM = 812
RequisitosReal(4).Armadura.MagDruiHH = 813
RequisitosReal(4).Armadura.MagDrioG = 814
 
RequisitosReal(5).Armadura.BarDruiCazAseH = 815
RequisitosReal(5).Armadura.BarDruiCazAseG = 816
RequisitosReal(5).Armadura.ClerigoH = 817
RequisitosReal(5).Armadura.ClerigoG = 818
RequisitosReal(5).Armadura.PalGueH = 819
RequisitosReal(5).Armadura.PalGueG = 820
RequisitosReal(5).Armadura.MagDruiHM = 821
RequisitosReal(5).Armadura.MagDruiHH = 822
RequisitosReal(5).Armadura.MagDrioG = 823

End Sub
Public Sub EnlistarArmadaReal(ByRef personaje As User, ByRef Enlistador As npc)

Dim npcIndex As Integer
Dim MiObj As obj

'Un usuario NO puede ingresar a la armada si.
'* Si ya pertenece a la Armada real.
'* Si pertenece al caos.
'* Si es criminal.
'* Si mato a algun ciudadano.
'* Tiene los requisitos de nivel y oro.

If personaje.faccion.ArmadaReal = 1 Then
    EnviarPaquete Paquetes.DescNpc, Chr$(28) & ITS(Enlistador.Char.charIndex), personaje.UserIndex
    Exit Sub
End If

If personaje.faccion.FuerzasCaos = 1 Then
    EnviarPaquete Paquetes.DescNpc, Chr$(29) & ITS(Enlistador.Char.charIndex), personaje.UserIndex
    Exit Sub
End If

'If .faccion.RecibioExpInicialReal = 1 Then
'    EnviarPaquete Paquetes.DescNpc2, ITS(str(NpcList(.flags.TargetNPC).Char.charIndex)) & "¿Ahora quieres volver? ¡Vete de aquí!", UserIndex
'    Exit Sub
'End If

If personaje.IDCuenta = 0 Then
    EnviarPaquete Paquetes.DescNpc2, ITS(Enlistador.Char.charIndex) & "¿Quién eres?. Debes tener una Cuenta para poder enlistarte.", personaje.UserIndex
    Exit Sub
End If

If personaje.faccion.alineacion = eAlineaciones.caos Then
    EnviarPaquete Paquetes.DescNpc2, ITS(Enlistador.Char.charIndex) & "No se permiten integrantes del Ejército Escarlata en el Ejército Índigo!!!", personaje.UserIndex
    Exit Sub
End If

If personaje.faccion.CriminalesMatados < RequisitosReal(1).Matados Then
    EnviarPaquete Paquetes.DescNpc, Chr$(30) & ITS(Enlistador.Char.charIndex) & RequisitosReal(1).Matados & "," & personaje.faccion.CriminalesMatados, personaje.UserIndex
    Exit Sub
End If

If personaje.Stats.ELV < RequisitosReal(1).Nivel Then
    EnviarPaquete Paquetes.DescNpc, Chr$(79) & ITS(Enlistador.Char.charIndex) & RequisitosReal(1).Nivel & ",", personaje.UserIndex
    Exit Sub
End If

If personaje.faccion.CiudadanosMatados > 0 Then
    EnviarPaquete Paquetes.DescNpc, Chr$(32) & ITS(Enlistador.Char.charIndex), personaje.UserIndex
    Exit Sub
End If

MiObj.Amount = 1
MiObj.ObjIndex = 0

Select Case personaje.Raza

    Case eRazas.ElfoOscuro, eRazas.Elfo, eRazas.Humano
        
        If personaje.clase = eClases.Bardo Or personaje.clase = eClases.Druida Or personaje.clase = eClases.Cazador Or personaje.clase = eClases.asesino Then
            MiObj.ObjIndex = 779
        ElseIf personaje.clase = eClases.Clerigo Then
            MiObj.ObjIndex = 781
        ElseIf personaje.clase = eClases.Paladin Or personaje.clase = eClases.Guerrero Then
            MiObj.ObjIndex = 783
        ElseIf personaje.clase = eClases.Mago Or personaje.clase = eClases.Druida Then
            If personaje.Genero = eGeneros.Hombre Then
                MiObj.ObjIndex = 786
            Else
                MiObj.ObjIndex = 785
            End If
        End If
    
    Case eRazas.Gnomo, eRazas.Enano
        
        If personaje.clase = eClases.Bardo Or personaje.clase = eClases.Druida Or personaje.clase = eClases.Cazador Or personaje.clase = eClases.asesino Then
            MiObj.ObjIndex = 780
        ElseIf personaje.clase = eClases.Clerigo Then
            MiObj.ObjIndex = 782
        ElseIf personaje.clase = eClases.Paladin Or personaje.clase = eClases.Guerrero Then
            MiObj.ObjIndex = 784
        ElseIf personaje.clase = eClases.Mago Or personaje.clase = eClases.Druida Then
            If personaje.Genero = eGeneros.Hombre Then
                MiObj.ObjIndex = 787
            Else
                MiObj.ObjIndex = 787
            End If
        End If
        
End Select

'Tengo que darle una armadura?
If MiObj.ObjIndex > 0 Then
    'La mete si o si
    If Not MeterItemEnInventario(personaje.UserIndex, MiObj) Then
        EnviarPaquete Paquetes.mensajeinfo, "No tienes espacio en el inventario para recibir la armadura faccionaria. Debes hacer lugar para ella.", personaje.UserIndex, ToIndex
        Exit Sub
    End If
End If

personaje.faccion.RecibioArmaduraReal = 1
personaje.faccion.ArmadaReal = 1
personaje.faccion.RecompensasReal = 1

Call NuevoIntegranteReal(personaje.id)

'Si se hace de la armada cambia de alineacion
Call modPersonaje.CambiarAlineacion(personaje.UserIndex, eAlineaciones.Real)

' Informo
EnviarPaquete Paquetes.DescNpc, Chr$(74) & ITS(Enlistador.Char.charIndex), personaje.UserIndex
End Sub

Private Sub NuevoIntegranteReal(id As Long)
    conn.Execute "INSERT INTO " & DB_NAME_PRINCIPAL & ".ejercito_real(IDPJ) VALUES('" & id & "')", , adExecuteNoRecords
End Sub

Private Sub NuevoIntegranteCaos(id As Long)
    conn.Execute "INSERT INTO " & DB_NAME_PRINCIPAL & ".ejercito_caos(IDPJ) VALUES('" & id & "')", , adExecuteNoRecords
End Sub

Private Sub QuitarIntegranteCaos(id As Long)
    conn.Execute "DELETE FROM " & DB_NAME_PRINCIPAL & ".ejercito_caos WHERE IDPJ=" & id, , adExecuteNoRecords
End Sub
Private Sub QuitarIntegranteReal(id As Long)
    conn.Execute "DELETE FROM " & DB_NAME_PRINCIPAL & ".ejercito_real WHERE IDPJ=" & id, , adExecuteNoRecords
End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)
Dim Rec As Byte
 
With UserList(UserIndex)
Rec = .faccion.RecompensasReal + 1

    If .faccion.RecompensasReal = 0 Then
    EnviarPaquete Paquetes.DescNpc, Chr$(18) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
    Exit Sub
    ElseIf .faccion.RecompensasReal >= 5 Then
    EnviarPaquete Paquetes.DescNpc, Chr$(100) & ITS(str(NpcList(.flags.TargetNPC).Char.charIndex)), UserIndex
    Else
        If .faccion.CriminalesMatados < RequisitosReal(Rec).Matados Then
        EnviarPaquete Paquetes.DescNpc, Chr$(101) & ITS(str(NpcList(.flags.TargetNPC).Char.charIndex)) & RequisitosReal(Rec).Matados & " criminales," & .faccion.CriminalesMatados, UserIndex
        Exit Sub
        End If
        
        If .Stats.ELV < RequisitosReal(Rec).Nivel Then
        EnviarPaquete Paquetes.DescNpc, Chr$(102) & ITS(str(NpcList(.flags.TargetNPC).Char.charIndex)) & RequisitosReal(Rec).Nivel & ",", UserIndex
        Exit Sub
        End If
        
        If .Stats.GLD < RequisitosReal(Rec).oro Then
        EnviarPaquete Paquetes.DescNpc, Chr$(103) & ITS(str(NpcList(.flags.TargetNPC).Char.charIndex)) & RequisitosReal(Rec).oro & ",", UserIndex
        Exit Sub
        End If
        
        'Si esta todo ok..
        'Quitar armadura
        .faccion.RecompensasReal = Rec
        .Stats.GLD = .Stats.GLD - RequisitosReal(Rec).oro
       
        Call QuitarArmaduraReal(UserList(UserIndex))
        '********************
       '** ENTEGA DE LAS ARMADURA
        Dim MiObj As obj
        MiObj.Amount = 1
        
        Select Case .Raza
            
            Case eRazas.Humano, eRazas.Elfo, eRazas.ElfoOscuro
                
                If .clase = eClases.Bardo Or .clase = eClases.Druida Or .clase = eClases.Cazador Or .clase = eClases.asesino Then
                    MiObj.ObjIndex = RequisitosReal(Rec).Armadura.BarDruiCazAseH
                
                ElseIf .clase = eClases.Clerigo Then
                    MiObj.ObjIndex = RequisitosReal(Rec).Armadura.ClerigoH
                    
                ElseIf .clase = eClases.Paladin Or .clase = eClases.Guerrero Then
                    MiObj.ObjIndex = RequisitosReal(Rec).Armadura.PalGueH
                    
                ElseIf .clase = eClases.Mago Or .clase = eClases.Druida Then
                    If .Genero = eGeneros.Hombre Then
                        MiObj.ObjIndex = RequisitosReal(Rec).Armadura.MagDruiHH
                    Else
                        MiObj.ObjIndex = RequisitosReal(Rec).Armadura.MagDruiHM
                    End If
                
                End If
            
            Case eRazas.Gnomo, eRazas.Enano
                
                If .clase = eClases.Bardo Or .clase = eClases.Druida Or .clase = eClases.Cazador Or .clase = eClases.asesino Then
                    MiObj.ObjIndex = RequisitosReal(Rec).Armadura.BarDruiCazAseG
                ElseIf .clase = eClases.Clerigo Then
                    MiObj.ObjIndex = RequisitosReal(Rec).Armadura.ClerigoG
                ElseIf .clase = eClases.Paladin Or .clase = eClases.Guerrero Then
                    MiObj.ObjIndex = RequisitosReal(Rec).Armadura.PalGueG
                ElseIf .clase = eClases.Mago Or .clase = eClases.Druida Then
                    MiObj.ObjIndex = RequisitosReal(Rec).Armadura.MagDrioG
                End If
                
        End Select
        'Si o si tiene que meter la armadura en el inventario.
        Call MeterItemEnInventario(UserIndex, MiObj)
        'Dar armadura
        EnviarPaquete Paquetes.EnviarOro, Codify(UserList(UserIndex).Stats.GLD), UserIndex, ToIndex
        EnviarPaquete Paquetes.DescNpc2, ITS(str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex)) & "Aqui tienes tu recompensa noble guerrero!!!", UserIndex
    End If
End With
End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer)
'Si deja de ser caos pasa a ser neutro
Call CambiarAlineacion(UserIndex, eAlineaciones.Neutro)

Call QuitarArmaduraReal(UserList(UserIndex))

' Actualiza base de datos
Call QuitarIntegranteReal(UserList(UserIndex).id)

Call WarpUserChar(UserIndex, UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, False)

UserList(UserIndex).faccion.ArmadaReal = 0

EnviarPaquete Paquetes.MensajeSimple, Chr$(182), UserIndex
End Sub

Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer)
'Si deja de ser caos pasa a ser neutro
Call CambiarAlineacion(UserIndex, eAlineaciones.Neutro)

Call QuitarArmaduraCaos(UserList(UserIndex))

' Actualiza base de datos
Call QuitarIntegranteCaos(UserList(UserIndex).id)

Call WarpUserChar(UserIndex, UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, False)
   
UserList(UserIndex).faccion.FuerzasCaos = 0

EnviarPaquete Paquetes.MensajeSimple, Chr$(183), UserIndex
End Sub

Public Sub EnlistarCaos(ByRef personaje As User, ByRef Enlistador As npc)
    
Dim MiObj As obj

'Un usuario NO puede enlistarse al caos si
' * No es criminal.
' * Ya pertenece al caos.
' * Pertenece a la armada real
' + No tiene una cuenta.
' * Ya fue del caos
' * Ya fue de de la Armada (RecibioExpInicialReal, nunca se valida)
' * Tiene los requisitos de nivel y ciudadanos matados
If personaje.faccion.alineacion = eAlineaciones.Real Then
    EnviarPaquete Paquetes.DescNpc, Chr$(76) & ITS(Enlistador.Char.charIndex), personaje.UserIndex
    Exit Sub
End If

If personaje.faccion.FuerzasCaos = 1 Then
    EnviarPaquete Paquetes.DescNpc, Chr$(75) & ITS(Enlistador.Char.charIndex), personaje.UserIndex
    Exit Sub
End If

If personaje.faccion.ArmadaReal = 1 Then
    EnviarPaquete Paquetes.DescNpc, Chr$(77) & ITS(Enlistador.Char.charIndex), personaje.UserIndex
    Exit Sub
End If

If personaje.IDCuenta = 0 Then
    EnviarPaquete Paquetes.DescNpc2, ITS(Enlistador.Char.charIndex) & "¿Quién eres?. Debes tener una Cuenta para poder enlistarte.", personaje.UserIndex
    Exit Sub
End If

'Si era un miembro de la armada real no se pueda enlistar
If personaje.faccion.RecibioExpInicialReal = 1 Or personaje.faccion.RecibioArmaduraCaos = 1 Then 'Tomamos el valor de ahí: ¿Recibio la experiencia para entrar?
    EnviarPaquete Paquetes.DescNpc, Chr$(77) & ITS(Enlistador.Char.charIndex), personaje.UserIndex
    Exit Sub
End If

If personaje.faccion.CiudadanosMatados < RequisitosCaos(1).Matados Then
    EnviarPaquete Paquetes.DescNpc, Chr$(78) & ITS(Enlistador.Char.charIndex) & RequisitosCaos(1).Matados & "," & personaje.faccion.CiudadanosMatados & ",", personaje.UserIndex
    Exit Sub
End If

If personaje.Stats.ELV < RequisitosCaos(1).Nivel Then
    EnviarPaquete Paquetes.DescNpc, Chr$(79) & ITS(NpcList(personaje.flags.TargetNPC).Char.charIndex) & RequisitosCaos(1).Nivel & ",", personaje.UserIndex
    Exit Sub
End If


MiObj.Amount = 1
MiObj.ObjIndex = 0

'Obtengo la armadura que le corresponde
Select Case personaje.Raza

    Case eRazas.Humano, eRazas.Elfo, eRazas.ElfoOscuro
    
        If personaje.clase = eClases.Bardo Or personaje.clase = eClases.Druida Or personaje.clase = eClases.Cazador Or personaje.clase = eClases.asesino Then
            MiObj.ObjIndex = 734
        ElseIf personaje.clase = eClases.Clerigo Then
            MiObj.ObjIndex = 736
        ElseIf personaje.clase = eClases.Paladin Or personaje.clase = eClases.Guerrero Then
            MiObj.ObjIndex = 738
        ElseIf personaje.clase = eClases.Mago Or personaje.clase = eClases.Druida Then
            If personaje.Genero = eGeneros.Hombre Then
                MiObj.ObjIndex = 741
            Else
                MiObj.ObjIndex = 740
            End If
        End If
        
    Case eRazas.Gnomo, eRazas.Enano
    
        If personaje.clase = eClases.Bardo Or personaje.clase = eClases.Druida Or personaje.clase = eClases.Cazador Or personaje.clase = eClases.asesino Then
            MiObj.ObjIndex = 735
        ElseIf personaje.clase = eClases.Clerigo Then
            MiObj.ObjIndex = 737
        ElseIf personaje.clase = eClases.Paladin Or personaje.clase = eClases.Guerrero Then
            MiObj.ObjIndex = 739
        ElseIf personaje.clase = eClases.Mago Or personaje.clase = eClases.Druida Then
            If personaje.Genero = eGeneros.Hombre Then
                MiObj.ObjIndex = 742
            Else
                MiObj.ObjIndex = 742
            End If
        End If
        
End Select

'Le tengo que dar una armadura?
If MiObj.ObjIndex > 0 Then
    'Si o si tiene que meter la armadura en el inventario.
    If Not MeterItemEnInventario(personaje.UserIndex, MiObj) Then
        EnviarPaquete Paquetes.mensajeinfo, "No tienes espacio en el inventario para recibir la armadura faccionaria. Debes hacer lugar para ella.", personaje.UserIndex, ToIndex
        Exit Sub
    End If
End If

'Finalmente lo hacemos de la lengion
personaje.faccion.RecompensasCaos = 1
personaje.faccion.FuerzasCaos = 1
personaje.faccion.RecibioArmaduraCaos = 1

Call NuevoIntegranteCaos(personaje.id)
'Si se hace del caos cambia de alineacion
Call modPersonaje.CambiarAlineacion(personaje.UserIndex, eAlineaciones.caos)

'El NPC le dice que se enlisto
EnviarPaquete Paquetes.DescNpc, Chr$(88) & ITS(Enlistador.Char.charIndex), personaje.UserIndex

End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)
Dim Rec As Byte

With UserList(UserIndex)
Rec = .faccion.RecompensasCaos + 1

    If .faccion.RecompensasCaos = 0 Then
        EnviarPaquete Paquetes.DescNpc, Chr$(18) & ITS(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex), UserIndex
        Exit Sub
    ElseIf .faccion.RecompensasCaos >= 5 Then
        EnviarPaquete Paquetes.DescNpc, Chr$(100) & ITS(str(NpcList(.flags.TargetNPC).Char.charIndex)), UserIndex
        'Desea pasar a un rango que no existe PONER MENSAJE
    Else
        If .faccion.CiudadanosMatados < RequisitosCaos(Rec).Matados Then
            EnviarPaquete Paquetes.DescNpc, Chr$(101) & ITS(str(NpcList(.flags.TargetNPC).Char.charIndex)) & RequisitosCaos(Rec).Matados & " ciudadanos," & .faccion.CiudadanosMatados, UserIndex
            Exit Sub
        End If
        
        If .Stats.ELV < RequisitosCaos(Rec).Nivel Then
            EnviarPaquete Paquetes.DescNpc, Chr$(102) & ITS(str(NpcList(.flags.TargetNPC).Char.charIndex)) & RequisitosCaos(Rec).Nivel & ",", UserIndex
            Exit Sub
        End If
        
        If .Stats.GLD < RequisitosCaos(Rec).oro Then
            EnviarPaquete Paquetes.DescNpc, Chr$(103) & ITS(str(NpcList(.flags.TargetNPC).Char.charIndex)) & RequisitosCaos(Rec).oro & ",", UserIndex
            Exit Sub
        End If
        
        'Si esta todo ok..
        'Quitar armadura
        .faccion.RecompensasCaos = Rec
        .Stats.GLD = .Stats.GLD - RequisitosCaos(Rec).oro
        
        Dim slot As Byte
        
        For slot = 1 To .Stats.MaxItems
            If .Invent.Object(slot).ObjIndex > 0 Then
                If (ObjData(.Invent.Object(slot).ObjIndex).alineacion And eAlineaciones.caos) Then
                    QuitarUserInvItem UserIndex, slot, 1
                    UpdateUserInv False, UserIndex, slot
                    Exit For
                End If
            End If
        Next
        
        Call QuitarArmaduraCaos(UserList(UserIndex))
        
        '********************
       '** ENTEGA DE LAS ARMADURA
        Dim MiObj As obj
        MiObj.Amount = 1
        
        Select Case .Raza
        
            Case eRazas.Humano, eRazas.Elfo, eRazas.ElfoOscuro
            
                If .clase = eClases.Bardo Or .clase = eClases.Druida Or .clase = eClases.Cazador Or .clase = eClases.asesino Then
                    MiObj.ObjIndex = RequisitosCaos(Rec).Armadura.BarDruiCazAseH
                ElseIf .clase = eClases.Clerigo Then
                    MiObj.ObjIndex = RequisitosCaos(Rec).Armadura.ClerigoH
                ElseIf .clase = eClases.Paladin Or .clase = eClases.Guerrero Then
                    MiObj.ObjIndex = RequisitosCaos(Rec).Armadura.PalGueH
                ElseIf .clase = eClases.Mago Or .clase = eClases.Druida Then
                    If .Genero = eGeneros.Hombre Then
                        MiObj.ObjIndex = RequisitosCaos(Rec).Armadura.MagDruiHH
                    Else
                        MiObj.ObjIndex = RequisitosCaos(Rec).Armadura.MagDruiHM
                    End If
                End If
                
            Case eRazas.Gnomo, eRazas.Enano
            
                If .clase = eClases.Bardo Or .clase = eClases.Druida Or .clase = eClases.Cazador Or .clase = eClases.asesino Then
                    MiObj.ObjIndex = RequisitosCaos(Rec).Armadura.BarDruiCazAseG
                ElseIf .clase = eClases.Clerigo Then
                    MiObj.ObjIndex = RequisitosCaos(Rec).Armadura.ClerigoG
                ElseIf .clase = eClases.Paladin Or .clase = eClases.Guerrero Then
                    MiObj.ObjIndex = RequisitosCaos(Rec).Armadura.PalGueG
                ElseIf .clase = eClases.Mago Or .clase = eClases.Druida Then
                    If .Genero = eGeneros.Hombre Then
                        MiObj.ObjIndex = RequisitosCaos(Rec).Armadura.MagDrioG
                    End If
                End If
            
        End Select
        'Si o si tiene que meter la armadura en el inventario.
        Call MeterItemEnInventario(UserIndex, MiObj)
        'Dar armadura
        EnviarPaquete Paquetes.EnviarOro, Codify(UserList(UserIndex).Stats.GLD), UserIndex, ToIndex
        EnviarPaquete Paquetes.DescNpc2, ITS(str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.charIndex)) & "Aqui tienes tu recompensa noble guerrero!!!", UserIndex
    End If
End With
End Sub

Private Sub QuitarArmaduraReal(ByRef personaje As User)
Dim LaQuito As Boolean
Dim slot As Byte

LaQuito = False

For slot = 1 To personaje.Stats.MaxItems
    If personaje.Invent.Object(slot).ObjIndex > 0 Then
        If (ObjData(personaje.Invent.Object(slot).ObjIndex).alineacion And eAlineaciones.Real) Then
            QuitarUserInvItem personaje.UserIndex, slot, 1
            UpdateUserInv False, personaje.UserIndex, slot
            LaQuito = True
        End If
    End If
Next
               
'Sino esta en el inventario me fijo en la boveda
If LaQuito = True Then Exit Sub
    
For slot = 1 To MAX_BANCOINVENTORY_SLOTS
    If personaje.BancoInvent.Object(slot).ObjIndex > 0 Then
        If (ObjData(personaje.BancoInvent.Object(slot).ObjIndex).alineacion And eAlineaciones.Real) Then
            QuitarBancoInvItem personaje.UserIndex, slot, 1
            LaQuito = True
        End If
    End If
Next
    
If LaQuito = False Then
    Call LogError("No se pudo quitar la armadura real al usuario " & personaje.Name)
End If

End Sub

Private Sub QuitarArmaduraCaos(ByRef personaje As User)

Dim LaQuito As Boolean
Dim slot As Byte

For slot = 1 To personaje.Stats.MaxItems
    If personaje.Invent.Object(slot).ObjIndex > 0 Then
        If (ObjData(personaje.Invent.Object(slot).ObjIndex).alineacion And eAlineaciones.caos) Then
            QuitarUserInvItem personaje.UserIndex, slot, 1
            UpdateUserInv False, personaje.UserIndex, slot
            LaQuito = True
        End If
    End If
Next
           
'Sino esta en el inventario me fijo en la boveda
If LaQuito = True Then Exit Sub

For slot = 1 To MAX_BANCOINVENTORY_SLOTS
    If personaje.BancoInvent.Object(slot).ObjIndex > 0 Then
        If (ObjData(personaje.BancoInvent.Object(slot).ObjIndex).alineacion And eAlineaciones.caos) Then
            QuitarBancoInvItem personaje.UserIndex, slot, 1
            LaQuito = True
        End If
    End If
Next

If LaQuito = False Then
    Call LogError("No se pudo quitar la armadura CAOS al usuario " & personaje.Name)
End If

End Sub
