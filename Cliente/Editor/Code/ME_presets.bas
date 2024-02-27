Attribute VB_Name = "ME_presets"
Option Explicit

Public Type infoPresetPos
   
    Grh(1 To CANTIDAD_CAPAS)     As Integer
    part(0 To 2)    As Integer
    
    Luz_radio       As Integer
    Brillo_radio    As Integer
    Luz_color       As RGBCOLOR
    Luz_tipo        As Byte
    
    obj             As Integer
    obj_cant        As Integer
    
    exit            As WorldPos
    
    bloqueo         As Boolean
    
    Trigger         As Integer
    
    npc             As Integer
End Type

Public Type PresetData
    id As Integer
    nombre          As String
    ancho As Byte
    alto As Byte
    infoPos() As MapBlock
End Type

Public NumPresets As Long
Public PresetsData() As PresetData
Public PresetsLucesBackup As LucesManager


Public Function obtenerIDPreset(nombre As String)

    Dim i As Integer
    
    For i = 1 To NumPresets
        If UCase$(PresetsData(i).nombre) = UCase$(nombre) Then
            obtenerIDPreset = PresetsData(i).id
            Exit Function
        End If
    Next i
    
    obtenerIDPreset = 0
End Function

Public Sub cambiarNombrePreset(idpreset As Integer, nuevoNombre As String)
    PresetsData(idpreset).nombre = nuevoNombre
    Call Me_indexar_Predefinidos.actualizarEnIni(PresetsData(idpreset))
End Sub

Public Function agregarNuevoPreset(preset As PresetData) As Boolean
    
    Dim idpreset As Integer
    
    'Si ya existe uno con ese nombre, lo remplazo
    idpreset = obtenerIDPreset(preset.nombre)
    
    If idpreset = 0 Then
        'El preset no existe tengo que agregar uno totalmente nuevo
        preset.id = Me_indexar_Predefinidos.nuevo
    Else
        preset.id = idpreset
    End If
    
    If Not preset.id = -1 Then
        'Guardo
        PresetsData(preset.id) = preset
        
        Call Me_indexar_Predefinidos.actualizarEnIni(PresetsData(preset.id))
        
        agregarNuevoPreset = True
    Else
        agregarNuevoPreset = False
    End If
End Function


Public Sub cargarListaPresets()

Dim loopPreset As Integer

frmMain.ListaConBuscadorPresets.vaciar

For loopPreset = 1 To NumPresets
    With PresetsData(loopPreset)
        If .alto > 0 And .ancho > 0 Then
            Call frmMain.ListaConBuscadorPresets.addString(loopPreset, loopPreset & " - " & .nombre)
        End If
    End With
Next

End Sub
'
