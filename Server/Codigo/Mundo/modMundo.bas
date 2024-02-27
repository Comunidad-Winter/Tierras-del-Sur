Attribute VB_Name = "modMundo"
Option Explicit

Private Const TIEMPO_MAX_FOGATAS As Long = 900000

Public fraccionDelDia As Integer
Public ForzarDia As Byte


Public Sub LimpiarMundo()
    Call BorrarFogatasViejas
    Call BorrarItemsEnSuelo
End Sub


Private Sub BorrarItemsEnSuelo()
'Borra los items que esten tirados en el suelo
Dim iMap As Integer
Dim x, y As Integer

For iMap = 1 To NumMaps
    'En mapas seguro no limpi
    If MapInfo(iMap).Pk = True Then
        For x = X_MINIMO_USABLE To X_MAXIMO_USABLE
            For y = Y_MINIMO_USABLE To Y_MAXIMO_USABLE
                If (MapData(iMap, x, y).Trigger And eTriggers.TodosBordesBloqueados) = False Then
                    If MapData(iMap, x, y).OBJInfo.ObjIndex > 0 Then
                      If Not ObjData(MapData(iMap, x, y).OBJInfo.ObjIndex).ObjType = OBJTYPE_GUITA Then
                        #If TDSFacil Then
                            If ObjData(MapData(iMap, x, y).OBJInfo.ObjIndex).valor * MapData(iMap, x, y).OBJInfo.Amount < 500 Then
                        #Else
                            If ObjData(MapData(iMap, x, y).OBJInfo.ObjIndex).valor * MapData(iMap, x, y).OBJInfo.Amount < 100 Then
                        #End If
                            If ItemNoEsDeMapa(MapData(iMap, x, y).OBJInfo.ObjIndex) Then Call EraseObj(ToMap, 0, iMap, 10000, iMap, x, y)
                        End If
                      Else
                         'Elimino pilones de una moneda de oro.
                         If MapData(iMap, x, y).OBJInfo.Amount = 1 Then Call EraseObj(ToMap, 0, iMap, 1, iMap, x, y)
                      End If
                    End If
                End If
            Next y
        Next x
    End If
Next iMap

End Sub

Private Sub BorrarFogatasViejas()

    Dim iMap As Integer
    Dim fogataData As ItemMapaData
    Dim fechaActual As Long
    
    fechaActual = GetTickCount
    
    For iMap = 1 To NumMaps
        If MapInfo(iMap).Existe = False Then GoTo continue:
        
        If MapInfo(iMap).fogatas.Count > 0 Then
        
            Do While MapInfo(iMap).fogatas.Count > 0
            
                Set fogataData = MapInfo(iMap).fogatas.Item(1)
                
                If fogataData.fecha + TIEMPO_MAX_FOGATAS < fechaActual Then
                    Call MapInfo(iMap).fogatas.Remove(1)
                    Call EraseObj(ToMap, 0, iMap, 1, iMap, fogataData.x, fogataData.y)
                Else
                    ' Como las fogatas se ordenan de mas viejas a mas nuevas, si la primera no cumple las condiciones mucho menos las siguientes
                    GoTo continue:
                End If
            Loop
            
        End If
        
continue:
    Next

End Sub
