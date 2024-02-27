Attribute VB_Name = "ME_FIFO_Server"
Option Explicit

Private Type accionTileAccionServer
    idAccionEditor As Integer
    accionServer As iAccion
End Type


Private Sub obtenerListaAccionesServer(ByRef lista() As accionTileAccionServer)

    Dim listaAccionEditorUsadas As Collection
    
    'Solo me interesan las acciones que se usan en el mapa...
    Set listaAccionEditorUsadas = obtenerListaAccionesEditorUsadas()
    
    Dim accion As iAccionEditor
    
    If listaAccionEditorUsadas.count > 0 Then
        ReDim lista(1 To listaAccionEditorUsadas.count) As accionTileAccionServer
        
        Dim i As Integer
        
        i = 1
        For Each accion In listaAccionEditorUsadas
            
            lista(i).idAccionEditor = accion.getID
            Set lista(i).accionServer = accion.generarAccionReal
        
            i = i + 1
        Next
    Else
        ReDim lista(0)
    End If
    
End Sub

Private Function obtenerIDAccionServer(lista() As accionTileAccionServer, idAccionTileEditor As Integer) As Integer
    
        Dim i As Integer
        
        For i = 1 To UBound(lista)
        
            If lista(i).idAccionEditor = idAccionTileEditor Then
                obtenerIDAccionServer = i
            End If
        Next i
End Function


Private Sub persistirListaAcciones(archivoDestino As Integer, lista() As accionTileAccionServer)
    
    Dim loopC As Integer
    
    Put archivoDestino, , CInt(UBound(lista))
    
    For loopC = 1 To UBound(lista)
        Call lista(loopC).accionServer.persistir(archivoDestino)
    Next loopC

End Sub



Private Function obtenerListaAccionesEditorUsadas() As Collection

   
    Dim esUsada As Boolean
    
    Dim listaAccionesEditor As Collection
    
    Set obtenerListaAccionesEditorUsadas = New Collection
    Set listaAccionesEditor = ME_modAccionEditor.obtenerListaAccionesEditor()
    
    Dim accionEditor As iAccionEditor
    
    For Each accionEditor In listaAccionesEditor

        esUsada = ME_modAccionEditor.esAccionnUsada(accionEditor)
        
        If esUsada Then
            Call obtenerListaAccionesEditorUsadas.Add(accionEditor)
        End If
        
    Next accionEditor


End Function
Public Function Guardar_Mapa_SV(ByVal SaveAs As String) As Boolean

On Error GoTo ErrorSave

Dim handle As Integer
Dim tempInt As Integer
Dim y As Long
Dim x As Long
Dim ByFlags As Integer

Dim listaAccionesServer() As accionTileAccionServer

' Si el archivo existe lo borramos
If FileExist(SaveAs, vbNormal) = True Then
    Call Kill(SaveAs)
End If

'Abrimos el archivo
handle = FreeFile
Open SaveAs For Binary As handle
Seek handle, 1
    
    Put handle, , THIS_MAPA.nombre
    Put handle, , THIS_MAPA.numero
    
    'inf Header
    Put handle, , ANCHO_MAPA 'TODO Se supone que siempre va a ser cuadrado..
    Put handle, , tempInt
    Put handle, , tempInt
    Put handle, , tempInt
    Put handle, , tempInt
    
    '
    mapinfo.numero = THIS_MAPA.numero

    Call ME_Mundo.aplicar_traslados(mapinfo, mapdata)
    'Obtengo las acciones que se utilizaron y el ID con el que fueron utilizadas en el mapeditor
    Call obtenerListaAccionesServer(listaAccionesServer)
    'Persisto estas acciones
    Call persistirListaAcciones(handle, listaAccionesServer)
    
    'Write .map file
    For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
        For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
        
            'Seteo del flag
            ByFlags = 0
            
            '1) ¿Tiene un evento este tile?
            If Not mapdata(x, y).accion Is Nothing Then ByFlags = ByFlags Or 1
            '2)
            If mapdata(x, y).NpcIndex Then ByFlags = ByFlags Or 2
            '3)
            If mapdata(x, y).OBJInfo.objIndex Then ByFlags = ByFlags Or 4
            '4)
            If mapdata(x, y).Trigger Then ByFlags = ByFlags Or 8
            
            Put handle, , ByFlags

            '3b) Si este tile tiene una accion guarda el id relativo del numero de accion que se egravo previamente
            If Not mapdata(x, y).accion Is Nothing Then
                Put handle, , obtenerIDAccionServer(listaAccionesServer, mapdata(x, y).accion.getID)
            End If
            
            '4b) Aparece un npc?
            If mapdata(x, y).NpcIndex Then
                Put handle, , mapdata(x, y).NpcIndex
            End If
            
            '5b) Hay un objeto?
            If mapdata(x, y).OBJInfo.objIndex Then
                Put handle, , mapdata(x, y).OBJInfo.objIndex
                Put handle, , mapdata(x, y).OBJInfo.Amount
            End If
            
            '6b) Trigger
            If mapdata(x, y).Trigger Then
                Put handle, , mapdata(x, y).Trigger
            End If
            
        Next x
    Next y
    Close handle
    
    Guardar_Mapa_SV = True

Exit Function

ErrorSave:
If handle > 0 Then Close handle
    MsgBox "Error en GuardarV2, nro. " & Err.Number & " - " & Err.description
End Function

