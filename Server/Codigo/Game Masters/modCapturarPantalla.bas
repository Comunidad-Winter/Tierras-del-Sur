Attribute VB_Name = "modCapturarPantalla"
' Este modulo sirve para capturar la pantalla del usuario y guardarla en el servidor
Option Explicit

Private capturasEjecutandose As Collection

Private Const MAX_TIEMPO_TRANSFERENCIA As Long = 120000  ' 20 minutos

Public Sub iniciar()
    Set capturasEjecutandose = New Collection
End Sub
Public Sub capturarPantalla(GameMaster As User, nombrePersonaje As String)
    Dim capturaData As clsCapturaPantalla
    Dim indexCapturado As Integer
    
    indexCapturado = NameIndex(nombrePersonaje)
    
    ' ¿Esta Online?
    If indexCapturado = 0 Then
        EnviarPaquete Paquetes.mensajeinfo, "El personaje '" & nombrePersonaje & "' se encuentra OffLine.", GameMaster.UserIndex, ToIndex
        Exit Sub
    End If

    ' El identificador es la fecha
    Set capturaData = New clsCapturaPantalla
    
    capturaData.fecha = GetTickCount

    capturaData.nombreGM = GameMaster.Name

    capturaData.nombreUsuario = UCase$(nombrePersonaje)
    capturaData.idUsuario = UserList(indexCapturado).id
    
    capturaData.bytesTotales = 0
    capturaData.bytesTransferidos = 0
    capturaData.stream = 0
    
    ' Agregamos la info
    Call capturasEjecutandose.Add(capturaData)
    
    ' Le decimos al cliente que arranque
    EnviarPaquete Paquetes.transferenciaIniciar, LongToString(capturaData.fecha), indexCapturado, ToIndex

    ' Avisamos al GM que ya arranco
    EnviarPaquete Paquetes.mensajeinfo, "Captura de la pantalla de '" & nombrePersonaje & "' en marcha.", GameMaster.UserIndex, ToIndex
    
    Call LogDesarrollo("Captura de Pantalla -> Iniciado. " & GameMaster.Name & " sobre " & nombrePersonaje & ".")
End Sub

Private Sub eliminarCapturaData(capturaData As clsCapturaPantalla)
    Dim loopCaptura As Integer
    
    For loopCaptura = 1 To capturasEjecutandose.Count
        
        If capturasEjecutandose(loopCaptura) Is capturaData Then
            Call capturasEjecutandose.Remove(loopCaptura)
            Exit For
        End If
        
    Next
    
End Sub

Private Function obtenerCapturaData(idTranferencia As Long, idPersonaje As Long) As clsCapturaPantalla
    Dim capturaData As clsCapturaPantalla
        
    Set obtenerCapturaData = Nothing
    
    ' La buscamos
    For Each capturaData In capturasEjecutandose
    
        ' ¿Es?
        If capturaData.fecha = idTranferencia And capturaData.idUsuario = idPersonaje Then
            Set obtenerCapturaData = capturaData
            Exit For
        End If
        
    Next

End Function
Private Sub transferencia_Iniciar(personaje As User, data As String)
    Dim capturaData As clsCapturaPantalla
    Dim idTransferencia As Long
    
    Call LogDesarrollo("Captura de Pantalla -> Se Reciben datos de Inicio de " & personaje.Name & ".")
    
    idTransferencia = StringToLong(data, 2)
    
    ' Buscamos la transferencia
    Set capturaData = obtenerCapturaData(idTransferencia, personaje.id)

    If capturaData Is Nothing Then
        Call LogError("Captura de Pantalla -> " & personaje.Name & " el usuario envío información de inicio de una captura que no existe.")
        Exit Sub
    End If
    
    capturaData.bytesTotales = StringToLong(data, 6)
    capturaData.bytesTransferidos = 0

    ' Abrimos el stree
    capturaData.stream = FreeFile
    Open (App.Path & "/Capturas/" & capturaData.fecha & capturaData.nombreUsuario & ".bmp") For Binary Access Write As capturaData.stream
        
    ' Envio la transferencia ok
    EnviarPaquete Paquetes.transferenciaOK, "", personaje.UserIndex, ToIndex
    
    ' Log
    Call LogDesarrollo("Captura de Pantalla -> Sobre " & personaje.Name & " iniciada correctamente.")
End Sub
Private Sub transferencia_Finalizar(personaje As User, capturaData As clsCapturaPantalla)
    Dim indexGM As Integer
    
    ' Cerramos el Stream
    Close capturaData.stream
        
    ' Eliminar informacion de captura
    Call eliminarCapturaData(capturaData)
        
    ' Obtenemos al Game Master
    indexGM = NameIndex(capturaData.nombreGM)
    
    If indexGM > 0 Then
        EnviarPaquete Paquetes.mensajeinfo, "Finalizada la captura de " & capturaData.nombreUsuario & ".", indexGM, ToIndex
    End If
        
    ' Captura Guardada
    Call LogDesarrollo("Captura de Pantalla -> " & personaje.Name & " info guardada correctamente.")
End Sub

Private Sub transferencia_SumarDatos(personaje As User, data As String)
    Dim capturaData As clsCapturaPantalla
    Dim idTransferencia As Long
    Dim bytesRecibidos As Long
    Dim archivo As Integer
    
    Call LogDesarrollo("Captura de Pantalla -> Se Reciben datos de captura de " & personaje.Name & ".")
    
    idTransferencia = StringToLong(data, 2)
    
    ' Buscamos la transferencia
    Set capturaData = obtenerCapturaData(idTransferencia, personaje.id)

    If capturaData Is Nothing Then
        Call LogError("Captura de Pantalla -> " & personaje.Name & " el usuario envío información de una captura que no existe.")
        Exit Sub
    End If
    
    bytesRecibidos = Len(data) - 5
    
    ' Vamos a guardar esto
    Seek capturaData.stream, capturaData.bytesTransferidos + 1
    Put capturaData.stream, , mid$(data, 6)
    
    ' Contamos la cantidad de bytes que le sumamos
    capturaData.bytesTransferidos = capturaData.bytesTransferidos + bytesRecibidos
    
    If capturaData.bytesTransferidos = capturaData.bytesTotales Then
        ' Si estan todos los datos, la finalizamos
        Call transferencia_Finalizar(personaje, capturaData)
    Else
        ' Le pedimos la siguiente parte
        EnviarPaquete Paquetes.transferenciaOK, "", personaje.UserIndex, ToIndex
        
        ' Log
        Call LogDesarrollo("Captura de Pantalla -> " & personaje.Name & " info guardada correctamente.")
    End If
End Sub


Public Sub agregarDatos(personaje As User, data As String)

    Dim esInicio As Boolean
    
    Call LogDesarrollo("Captura de Pantalla -> Se reciben datos de " & personaje.Name & ".")
        
    esInicio = (StringToByte(data, 1) = 1)
    
    If esInicio Then
        Call transferencia_Iniciar(personaje, data)
    Else
        Call transferencia_SumarDatos(personaje, data)
    End If

End Sub

Private Sub cancelarCaptura(capturaData As clsCapturaPantalla)

    If capturaData.stream > 0 Then
        Close capturaData.stream
    End If
    
    Call LogDesarrollo("Captura de pantalla -> Se cancela " & capturaData.fecha & " sobre el usuario " & capturaData.nombreUsuario)
    
    Dim indexGM As Integer
    
    indexGM = NameIndex(capturaData.nombreGM)
    
    If indexGM > 0 Then
        EnviarPaquete Paquetes.mensajeinfo, "Se cancela la captura sobre " & capturaData.nombreUsuario, indexGM, ToIndex
    End If
    
End Sub
' Elimina todas las capturas que no se hayan completado en un plazo máximo de 4 minutos
Public Sub eliminarCorruptas()
    Dim ahora As Long
    Dim capturaData As clsCapturaPantalla
    Dim loopCaptura As Integer
    
    ' ¿Hay alguna pendiente?
    If capturasEjecutandose.Count = 0 Then Exit Sub
    
    ahora = GetTickCount
    
    loopCaptura = 1
    Do While capturasEjecutandose.Count > 0 And loopCaptura <= capturasEjecutandose.Count
    
        If capturasEjecutandose(loopCaptura).fecha < ahora - MAX_TIEMPO_TRANSFERENCIA Then
            Call cancelarCaptura(capturasEjecutandose(loopCaptura))
            Call capturasEjecutandose.Remove(loopCaptura)
        Else
            loopCaptura = loopCaptura + 1
        End If
    
    Loop

End Sub


