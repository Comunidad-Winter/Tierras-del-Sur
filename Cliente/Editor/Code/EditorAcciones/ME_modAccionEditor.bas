Attribute VB_Name = "ME_modAccionEditor"
Option Explicit

Public listaAccionTileEditor As New Collection

Public listaAccionTileEditorUsando As Collection

Private idsLIbres As EstructurasLib.ColaConBloques

'/* Para el editor */
Public Const MAX_ACCIONES_DISTINTAS = 100

Public Sub agregarNuevaAccion(template As cAccionCompuestaEditor)
    Call template.setID(obtenerIDLIbre)
    Call listaAccionTileEditorUsando.Add(template)
End Sub

Public Function obtenerListaAccionesEditor() As Collection
    Set obtenerListaAccionesEditor = listaAccionTileEditorUsando
End Function

Public Sub eliminarAccion(idAccion As Integer)
    Dim accion As cAccionCompuestaEditor
    Dim loopc As Integer
    
    loopc = 1
    
    For Each accion In listaAccionTileEditorUsando
            If accion.getID = idAccion Then
                listaAccionTileEditorUsando.Remove loopc
                liberarID (idAccion)
                Exit For
            End If
            loopc = loopc + 1
    Next
End Sub

Public Function obtenerAccion(nombreAccion As String) As cAccionCompuestaEditor
    Dim accion As cAccionCompuestaEditor
    
    For Each accion In listaAccionTileEditorUsando
            If accion.iAccionEditor_getNombre = nombreAccion Then
                Set obtenerAccion = accion
                Exit For
            End If
    Next
End Function

Public Function obtenerAccionID(idAccion As Integer) As cAccionCompuestaEditor
    Dim accion As cAccionCompuestaEditor
    
    For Each accion In listaAccionTileEditorUsando
            If accion.getID = idAccion Then
                Set obtenerAccionID = accion
                Exit Function
            End If
    Next
    
    Set obtenerAccionID = Nothing
End Function

Public Sub iniciar()
    Call iniciarEstructuraIDsLibres
    Set listaAccionTileEditorUsando = New Collection
End Sub

Public Sub cargarListaAccionTile()
    Set listaAccionTileEditor = New Collection
    
    Call ME_modAccionEditor.listaAccionTileEditor.Add(ME_constructoresAccionEditor.construirAccionTileExitComun)
    'Call modAccionTileEditor.listaAccionTileEditor.Add(constructoresAccionEditor.construirAccionTileExitAutomaticoDerecha)
    'Call modAccionTileEditor.listaAccionTileEditor.Add(constructoresAccionEditor.construirAccionTileEjecutarEntidadWorldPos)
    'Call modAccionTileEditor.listaAccionTileEditor.Add(constructoresAccionEditor.construirAccionTileBloquearPase)
End Sub
Private Sub iniciarEstructuraIDsLibres()
    Dim idLibre As Integer
    Set idsLIbres = New EstructurasLib.ColaConBloques
    'Agrego de atras para adelante para que primero tomes los indexs más chicos
    'Los ultimos serán los primeros.
    For idLibre = MAX_ACCIONES_DISTINTAS To 1 Step -1
        Call idsLIbres.agregar(idLibre)
    Next
End Sub

Public Function obtenerIDLIbre() As Long
    obtenerIDLIbre = idsLIbres.sacar()
End Function

Public Sub eliminarID(id As Long)
    Call idsLIbres.eliminar(id)
End Sub

Public Sub liberarID(id As Long)
    Call idsLIbres.agregar(id)
End Sub

Public Sub refrescarListaUsando(Lista As ListBox)
    Dim aux As cAccionCompuestaEditor
    
    Lista.Clear
    For Each aux In ME_modAccionEditor.listaAccionTileEditorUsando
        If aux.esVisible() Then
            Lista.AddItem aux.getID & ") " & aux.iAccionEditor_getNombre
        End If
    Next
    
    Lista.Refresh

End Sub

Public Sub refrescarListaDisponibles(Lista As ListBox)
    
    Dim aux As iAccionEditor
    
    Lista.Clear
    For Each aux In listaAccionTileEditor
    
        Lista.AddItem aux.GetNombre
    Next
    Lista.Refresh
End Sub

Public Sub persistirListaAccionTileEditorUsando(archivoSalida As Integer)
    Dim accionTileEditor As iAccionEditor
    
    '1) Obtengo la cantidad que hay utilizandose en este mapa
    Put archivoSalida, , CInt(listaAccionTileEditorUsando.Count)
    
    '2) Persisto los elementos
    For Each accionTileEditor In listaAccionTileEditorUsando
        Call accionTileEditor.persistir(archivoSalida)
    Next
End Sub

Public Sub cargarListaAccionesEditorUsando(archivoFuente As Integer)
    Dim accionTileEditor As iAccionEditor
    Dim cantidadAccionTileEditor As Integer
    Dim contadorAccion As Integer
    Dim tipoAccion As Byte
        
    '1) Obtengo la cantidad que hay utilizandose en este mapa
    Get archivoFuente, , cantidadAccionTileEditor
    
    '2) Persisto los elementos
    For contadorAccion = 1 To cantidadAccionTileEditor
            
        Get archivoFuente, , tipoAccion
            
        If tipoAccion = 1 Then
            Set accionTileEditor = New cAccionCompuestaEditor
        Else
            Set accionTileEditor = New cAccionTileEditor
        End If
            
        Call accionTileEditor.Cargar(archivoFuente)
        
        Call ME_modAccionEditor.eliminarID(accionTileEditor.getID)
        
        Call listaAccionTileEditorUsando.Add(accionTileEditor)

    Next
End Sub

Public Sub eliminarAccionMapa(accionEditor As iAccionEditor)
    Dim X As Integer
    Dim Y As Integer
    
    For X = X_MINIMO_JUGABLE To X_MAXIMO_JUGABLE
        For Y = Y_MINIMO_JUGABLE To X_MAXIMO_JUGABLE
            If MapData(X, Y).accion Is accionEditor Then
                Set MapData(X, Y).accion = Nothing
            End If
        Next Y
    Next X
End Sub

Public Function esAccionnUsada(accionEditor As iAccionEditor) As Boolean
    Dim X As Integer
    Dim Y As Integer

    X = X_MINIMO_JUGABLE
    esAccionnUsada = False

    Do While (Not esAccionnUsada And X <= X_MAXIMO_JUGABLE)
    
        Y = Y_MINIMO_JUGABLE
        Do While (Not esAccionnUsada And Y <= Y_MAXIMO_JUGABLE)
        
            If MapData(X, Y).accion Is accionEditor Then
                esAccionnUsada = True
            End If
            
            Y = Y + 1
        Loop
        
        X = X + 1
    Loop

End Function
