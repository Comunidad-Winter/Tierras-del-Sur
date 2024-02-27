Attribute VB_Name = "Me_Tools_Presets"
Option Explicit

Public Enum eHerramientasPresets
    ninguna = 0
    insertar = 1
End Enum

Public herramientaInternaPresets As eHerramientasPresets

Public idPresetSeleccionado As Integer


Public Sub click_insertarPreset()

    Dim tmp As Integer
    Dim tmp_tool As Long
        
    tmp = idPresetSeleccionado
        
    If tmp > 0 And tmp <= NumPresets Then
        
        herramientaInternaPresets = eHerramientasPresets.insertar
        
        With PresetsData(tmp)
        
            tmp_tool = (0 Or Tools.tool_triggers Or Tools.tool_obj Or Tools.tool_npc Or Tools.tool_tileset Or tool_grh Or tool_particles Or tool_luces)
   
            tmp_tool = Me_Tools_Seleccion.copiarDesdeMapBlock(.infoPos, ME_presets.PresetsLucesBackup, tmp_tool)
                        
            Call ME_Tools.selectToolMultiple(tmp_tool, "Insertar " & .nombre, frmMain.cmdInsertarPreset)
                        
            'Le aviso al sistema que el area no es de 1x1 sino que es de otro tamaño.
            Call ME_Tools.establecerAmpliacionDeArea(CInt(.ancho), CInt(.alto))
                
        End With
    
        Call actualizarListaUltimosUsados(tmp)
    Else
        MsgBox "Debe seleccionar un preset a insertar."
    End If

End Sub

Public Sub activarUltimaHerramientaPresets()
    
    Select Case herramientaInternaPresets
        Case eHerramientasPresets.insertar
            Call Me_Tools_Presets.click_insertarPreset
    End Select
    
End Sub

Public Function nuevoPreset(elementos As Tools) As Boolean
    Dim TempStr As String
    Dim creado As Boolean
    
    creado = False
    
    TempStr = InputBox("Ingrese el nombre para el elemento predefinido que está por crear a partir del area del mapa que seleccionó.", "Nuevo elemento predefinido")
                    
    If TempStr <> "" Then
    
        If ME_presets.obtenerIDPreset(TempStr) > 0 Then
            If MsgBox("El preset " & TempStr & " ya existe, ¿Deseas remplazarlo?", vbYesNo + vbExclamation) = vbNo Then
                Exit Function
            End If
        End If
        
        
        creado = Me_Tools_Seleccion.crearPresetDesdeMapa(areaSeleccionada, TempStr, elementos)
    End If
                    
    If creado Then
        Call ME_presets.cargarListaPresets
    End If
    
    nuevoPreset = creado
End Function

Public Sub elimimarDeListaUltimosUsados(idPredefinido As Integer)
    Dim loopElemento As Integer
        
    For loopElemento = 1 To frmMain.lstUltimosPredefinidosUtilizados.ListCount
        If val(frmMain.lstUltimosPredefinidosUtilizados.list(loopElemento - 1)) = idPredefinido Then
                Call frmMain.lstUltimosPredefinidosUtilizados.RemoveItem(loopElemento - 1)
            Exit For
        End If
    Next
    
End Sub


Public Sub actualizarPresetEnListaUltimosUsados(idPredefinido As Integer)
    Dim loopElemento As Integer
        
    For loopElemento = 1 To frmMain.lstUltimosPredefinidosUtilizados.ListCount
        If val(frmMain.lstUltimosPredefinidosUtilizados.list(loopElemento - 1)) = idPredefinido Then
                frmMain.lstUltimosPredefinidosUtilizados.list(loopElemento - 1) = idPredefinido & " - " & PresetsData(idPredefinido).nombre
            Exit For
        End If
    Next
    
End Sub

Public Sub actualizarListaUltimosUsados(idPredefinido As Integer)
    Dim loopElemento As Integer
    Dim existe As Boolean
    
    existe = False
    ' Busco si ya esta en la lista.
    For loopElemento = 1 To frmMain.lstUltimosPredefinidosUtilizados.ListCount
        If val(frmMain.lstUltimosPredefinidosUtilizados.list(loopElemento - 1)) = idPredefinido Then
        
             'Si esta, lo remuevo, excepto que este en la posicion numero 1
            'If loopElemento <> 1 Then
            '    Call frmMain.lstUltimosPredefinidosUtilizados.RemoveItem(loopElemento - 1)
           ' End If
            'Exit For
            
            ' No hago nada. Sino se esta moviendo todo el tempo el nombre en la lista y es imposible de encontrar
            existe = True
            Exit Sub
        End If
    Next
    
    'Si esta en la primer posicion, no necesito agregarlo y la lista mantiene la cantidad de elementos.
   ' If Not loopElemento = 1 Or frmMain.lstUltimosPredefinidosUtilizados.ListCount = 0 Then
   
    If Not existe Then
        ' Llegue a la maxima cantidad de Elementos ? Entonces borro el ultimo
        If frmMain.lstUltimosPredefinidosUtilizados.ListCount = 16 Then
            Call frmMain.lstUltimosPredefinidosUtilizados.RemoveItem(frmMain.lstUltimosPredefinidosUtilizados.ListCount - 1)
        End If
        
        'Agrego
        Call frmMain.lstUltimosPredefinidosUtilizados.AddItem(idPredefinido & " - " & PresetsData(idPredefinido).nombre, 0)
    End If

End Sub
