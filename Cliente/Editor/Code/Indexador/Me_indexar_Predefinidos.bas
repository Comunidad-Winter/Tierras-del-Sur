Attribute VB_Name = "Me_indexar_Predefinidos"
Option Explicit

Private Const Archivo = "Presets.ini"
Private Const HEAD_ELEMENTO = ""
Private Const CDM_IDENTIFICADOR = "PREDEFINIDO"

Public Function nuevo() As Integer
    
    #If Colaborativo = 0 Then
        'Busco alguno que este libre
        Dim elemento As Integer
        
        nuevo = -1
        
        For elemento = 1 To UBound(PresetsData)
            If Not existe(elemento) Then
                nuevo = elemento
                Exit For
            End If
        Next
        
        'No tengo slot libre. Creo uno
        If nuevo = -1 Then
            ReDim Preserve PresetsData(0 To UBound(PresetsData) + 1) As PresetData
            nuevo = UBound(PresetsData)
            NumPresets = nuevo
        End If
    #Else
        nuevo = CDM.cerebro.SolicitarRecurso(CDM_IDENTIFICADOR)
        
        If nuevo > UBound(PresetsData) Then
            ReDim Preserve PresetsData(0 To nuevo) As PresetData
            NumPresets = UBound(PresetsData)
        End If
        
        Call versionador.creado(CDM_IDENTIFICADOR, nuevo)
    #End If

End Function

Public Sub resetear(preset As PresetData)

    preset.alto = 0
    preset.ancho = 0
    preset.nombre = ""
        
End Sub
Public Function existe(ByVal id As Integer) As Boolean

    If id > UBound(PresetsData) Then
        existe = False
        Exit Function
    End If
    
    existe = (PresetsData(id).alto > 0 And PresetsData(id).ancho > 0)

End Function

Public Sub eliminar(id As Integer)
    Dim nombreBackup As String
    
    nombreBackup = PresetsData(id).nombre
    
    Call resetear(PresetsData(id))
    
    Call actualizarEnIni(PresetsData(id))
    
    If id = UBound(PresetsData) Then
        ReDim Preserve PresetsData(0 To UBound(PresetsData) - 1) As PresetData
        NumPresets = UBound(PresetsData)
    End If

    #If Colaborativo = 1 Then
        Call versionador.eliminado(CDM_IDENTIFICADOR, id, nombreBackup)
    #End If
End Sub

Public Sub actualizarEnIni(preset As PresetData)
    Dim X As Integer
    Dim Y As Integer
    Dim loopGrh As Byte
    Dim luz As tLuzPropiedades
    
    Dim identificadorPreset As String
    Dim archivoSalida As String
    Dim indiceMatriz As String
    Dim tempLong As Long
    Dim tempbyte1 As Byte
    Dim tempbyte2 As Byte
    
    archivoSalida = DBPath & Archivo
        
    If FileExist(archivoSalida, vbArchive) = False Then
        MsgBox "Falta el archivo '" & Archivo & "' en " & DBPath, vbCritical
        End
    End If
        
    identificadorPreset = preset.id
    
    'Guardo los datos generales
    Call WriteVar(archivoSalida, identificadorPreset, "NOMBRE", preset.nombre)
    Call WriteVar(archivoSalida, identificadorPreset, "ANCHO", preset.ancho)
    Call WriteVar(archivoSalida, identificadorPreset, "ALTO", preset.alto)
    
    For X = 1 To preset.ancho
        For Y = 1 To preset.alto
                    
                indiceMatriz = "(" & X & "," & Y & ")"
            
                With preset.infoPos(X, Y)
                        'Guardo los GRH
                        For loopGrh = 1 To CANTIDAD_CAPAS
                            Call WriteVar(archivoSalida, identificadorPreset, "GRH" & indiceMatriz & "(" & loopGrh & ")", .Graphic(loopGrh).GrhIndex)
                        Next
                        
                        'Guardo trigger
                        Call WriteVar(archivoSalida, identificadorPreset, "TRIGGER" & indiceMatriz, .Trigger)
                        
                        'Guardo bloqueo
                        If (.Trigger And eTriggers.TodosBordesBloqueados) = eTriggers.TodosBordesBloqueados Then
                            Call WriteVar(archivoSalida, identificadorPreset, "BLOQUEADO" & indiceMatriz, "1")
                        Else
                            Call WriteVar(archivoSalida, identificadorPreset, "BLOQUEADO" & indiceMatriz, "0")
                        End If
                        
                        'Particulas
                        For loopGrh = 0 To 2
                            If Not .Particles_groups(loopGrh) Is Nothing Then
                                Call WriteVar(archivoSalida, identificadorPreset, "PARTICULA" & indiceMatriz & "(" & loopGrh & ")", .Particles_groups(loopGrh).PGID)
                            Else
                                Call WriteVar(archivoSalida, identificadorPreset, "PARTICULA" & indiceMatriz & "(" & loopGrh & ")", 0)
                            End If
                        Next
                        
                        'Objeto
                        Call WriteVar(archivoSalida, identificadorPreset, "OBJETO" & indiceMatriz, CStr(.OBJInfo.objIndex & " " & .OBJInfo.Amount))
                        
                        'Npc
                        Call WriteVar(archivoSalida, identificadorPreset, "NPC" & indiceMatriz, .NpcIndex)
                        
                        'TileSet
                        Call WriteVar(archivoSalida, identificadorPreset, "TILESET" & indiceMatriz, .tile_texture & " " & .tile_number)
                        
                        'Luz
                        If .luz > 0 Then
                            'Obtengo las propiedades de la luz
                            DLL_Luces.Get_Light .luz, tempbyte1, tempbyte2, luz.LuzColor.r, luz.LuzColor.g, luz.LuzColor.b, luz.LuzRadio, luz.LuzBrillo, luz.LuzTipo, luz.luzInicio, luz.luzFin
                            'Guardo la info de la luz
                            Call WriteVar(archivoSalida, identificadorPreset, "LUZ" & indiceMatriz, luz.LuzRadio & " " & luz.LuzBrillo & " " & luz.LuzColor.r & " " & luz.LuzColor.g & " " & luz.LuzColor.b & " " & luz.LuzTipo & " " & luz.luzInicio & " " & luz.luzFin)
                            'Genero una copia en el sistema de luces de predefinidos
                            .luz = PresetsLucesBackup.crear(100, 100, luz.LuzColor.r, luz.LuzColor.g, luz.LuzColor.b, luz.LuzRadio, luz.LuzBrillo, luz.LuzTipo, luz.luzInicio, luz.luzFin)
                        Else
                            Call WriteVar(archivoSalida, identificadorPreset, "LUZ" & indiceMatriz, 0 & " " & 0 & " " & 0 & " " & 0 & " " & 0 & " " & 0 & " " & 0 & " " & 0)
                        End If
                End With
      
            
        Next Y
    Next X
    
    #If Colaborativo = 1 Then
        If existe(identificadorPreset) Then
            Call versionador.modificado(CDM_IDENTIFICADOR, identificadorPreset, PresetsData(identificadorPreset).nombre)
        End If
    #End If
    
    Debug.Print "Preset " & preset.nombre & " GUARDADO."
End Sub

Public Sub CargarPresets()
'On Error GoTo Fallo

    Dim i As Integer
    Dim Leer As New cIniManager
    Dim T() As String
    Dim tmp As String
    Dim alto As Integer
    Dim ancho As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim loopGrh As Byte
    Dim luz As tLuzPropiedades
    Dim indiceMatriz As String
        
    ' Chequeo si existe el archivo
    If FileExist(DBPath & Archivo, vbArchive) = False Then
        MsgBox "Falta el archivo '" & Archivo & "' en " & DBPath, vbCritical
        End
    End If
    
    ' Inicio
    Call Leer.Initialize(DBPath & Archivo)
    
    ' Obtengo la cantidad de elementos que hay
    NumPresets = CInt(val(Leer.getNameLastSection))
    
    ' Genero el sistema de luces en donde voy a guardar las luces de los predefinidos guardadas
    Set PresetsLucesBackup = New LucesManager
    Call PresetsLucesBackup.iniciar(500, ResultColorArray(1, 1), ANCHO_MAPA)
       
    ReDim PresetsData(0 To NumPresets)
    
    For i = 1 To NumPresets

        PresetsData(i).alto = CByte(val(Leer.getValue(i, "ALTO")))
        PresetsData(i).ancho = CByte(val(Leer.getValue(i, "ANCHO")))
        
            If PresetsData(i).alto > 0 And PresetsData(i).ancho > 0 Then
                
                PresetsData(i).nombre = Leer.getValue(i, "NOMBRE")
                PresetsData(i).id = i
                
                ReDim PresetsData(i).infoPos(1 To PresetsData(i).ancho, 1 To PresetsData(i).alto)
                'Cargo la información de cada tile
                For X = 1 To PresetsData(i).ancho
                
                    For Y = 1 To PresetsData(i).alto

                        indiceMatriz = "(" & X & "," & Y & ")"

                        With PresetsData(i).infoPos(X, Y)
                        
                            'Cargo los grh
                            For loopGrh = 1 To CANTIDAD_CAPAS
                                .Graphic(loopGrh).GrhIndex = val(Leer.getValue(i, "GRH" & indiceMatriz & "(" & loopGrh & ")"))
                            Next
                            
                            'Trigger
                            .Trigger = val(Leer.getValue(i, "TRIGGER" & indiceMatriz))
                            
                            '¿Esta bloqueado?
                            If (val(Leer.getValue(i, "BLOQUEADO" & indiceMatriz)) <> 0) Then
                                .Trigger = (.Trigger Or eTriggers.TodosBordesBloqueados)
                            End If

                            
                            'Particulas
                            For loopGrh = 0 To 2
                                Dim tmp_part As Integer
                                tmp_part = val(Leer.getValue(i, "PARTICULA" & indiceMatriz & "(" & loopGrh & ")"))
                                If tmp_part Then
                                    Set .Particles_groups(loopGrh) = New Engine_Particle_Group
                                    .Particles_groups(loopGrh).CargarPGID = tmp_part
                                End If
                            Next
                            
                            'Objeto
                            tmp = Leer.getValue(i, "OBJETO" & indiceMatriz)
                            If Len(tmp) Then
                                T = Split(tmp, " ")
                                If UBound(T) = 1 Then
                                    .OBJInfo.objIndex = val(T(0))
                                    .OBJInfo.Amount = val(T(1))
                                End If
                            End If
                            
                            'Npc
                            .NpcIndex = val(Leer.getValue(i, "NPC" & indiceMatriz))

                            'TileSet
                            tmp = Leer.getValue(i, "TILESET" & indiceMatriz)
                            If Len(tmp) Then
                                T = Split(tmp, " ")
                                If UBound(T) = 1 Then
                                    .tile_texture = val(T(0))
                                    .tile_number = val(T(1))
                                    '.tile_texture  = val(t(2))
                                End If
                            End If

                            'Traslado
                            'Hay que ver esto del guardado de las acciones
                            'tmp = Leer.GetValue(& i, "EXIT" & indiceMatriz)
                            'If Len(tmp) Then
                            '    t = Split(tmp, " ")
                            '    If UBound(t) = 2 Then
                            '        .exit.map = val(t(0))
                            '        .exit.x = val(t(1))
                            '        .exit.y = val(t(2))
                            '    End If
                            'End If
                            
                            'LEO LAS LUCES
                           tmp = Leer.getValue(i, "LUZ" & indiceMatriz)
                           
                            If Len(tmp) Then
                                T = Split(tmp, " ")
                                
                                If T(0) > 0 Then
                                
                                    luz.LuzColor.r = val(T(2))
                                    luz.LuzColor.g = val(T(3))
                                    luz.LuzColor.b = val(T(4))
                                    
                                    luz.LuzRadio = val(T(0))
                                    luz.LuzBrillo = val(T(1))
                                    luz.LuzTipo = val(T(5))
                                    luz.luzInicio = val(T(6))
                                    luz.luzFin = val(T(7))
    'luz.LuzRadio & " " & luz.LuzBrillo & " " & luz.LuzColor.r & " " & luz.LuzColor.g & " " & luz.LuzColor.b & " " & luz.LuzTipo & " " & luz.luzInicio & " " & luz.luzFin)
                                    .luz = PresetsLucesBackup.crear(100, 100, luz.LuzColor.r, luz.LuzColor.g, luz.LuzColor.b, luz.LuzRadio, luz.LuzBrillo, luz.LuzTipo, luz.luzInicio, luz.luzFin)
                                    Debug.Print "Generada la luz"; .luz, "Count"; PresetsLucesBackup.count
                                End If
                            End If
                            
                        End With
                        
                    Next Y
                
                Next X

            End If

    Next i

End Sub
