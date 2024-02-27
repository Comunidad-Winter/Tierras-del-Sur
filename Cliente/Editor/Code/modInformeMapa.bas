Attribute VB_Name = "modInformeMapa"
Option Explicit


Private Function obtenerInformeObjetos() As String
    Dim X, Y As Integer
    Dim objIndex As Integer
    Dim nombre As String
    
    obtenerInformeObjetos = ""
    For X = X_MINIMO_USABLE To X_MAXIMO_USABLE
        For Y = X_MINIMO_USABLE To X_MAXIMO_USABLE
            objIndex = mapdata(X, Y).OBJInfo.objIndex
            If objIndex > 0 Then
                nombre = ObjData(objIndex).Name
                obtenerInformeObjetos = obtenerInformeObjetos & format(X, "0##") & vbTab & format(Y, "0##") & vbTab & "-> " & nombre & " (" & mapdata(X, Y).OBJInfo.Amount & ")" & vbCrLf
            End If
        Next Y
    Next X

End Function

Private Function obtenerInformeCriaturas() As String
    Dim X, Y As Integer

    obtenerInformeCriaturas = ""
    For X = X_MINIMO_USABLE To X_MAXIMO_USABLE
        For Y = X_MINIMO_USABLE To X_MAXIMO_USABLE
            If mapdata(X, Y).NpcIndex > 0 Then
                obtenerInformeCriaturas = obtenerInformeCriaturas & format(X, "0##") & vbTab & format(Y, "0##") & vbTab & "-> " & NpcData(mapdata(X, Y).NpcIndex).Name & vbCrLf
            End If
        Next Y
    Next X

End Function

Private Function obtenerInformeEntidades() As String
    Dim X, Y As Integer
    Dim idEntidad As Integer
    Dim tempStr As String
    
    obtenerInformeEntidades = ""
    For X = X_MINIMO_USABLE To X_MAXIMO_USABLE
        For Y = X_MINIMO_USABLE To X_MAXIMO_USABLE
        
            If EntidadesMap(X, Y) > 0 Then
            
                idEntidad = EntidadesMap(X, Y)
                tempStr = format(X, "0##") & vbTab & format(Y, "0##") & vbTab & "-> "
                
                Do While idEntidad > 0
                    tempStr = tempStr & EntidadesIndexadas(Entidades(idEntidad).numeroIndexadoEntidad).nombre & " | "
                    idEntidad = Entidades(idEntidad).Next
                Loop
               
                obtenerInformeEntidades = obtenerInformeEntidades & tempStr & vbCrLf
            End If
        Next Y
    Next X

End Function

Private Function obtenerInformeLuces() As String
    Dim X, Y As Integer
    Dim cantidadLuces As Integer
    
    obtenerInformeLuces = ""
    cantidadLuces = 0
    
    For X = X_MINIMO_USABLE To X_MAXIMO_USABLE
        For Y = X_MINIMO_USABLE To X_MAXIMO_USABLE
        
            If mapdata(X, Y).luz > 0 Then
                cantidadLuces = cantidadLuces + 1
            End If
        Next Y
    Next X
    
    If cantidadLuces = 0 Then
        obtenerInformeLuces = "No hay luces en el mapa."
    Else
        obtenerInformeLuces = "Cantidad: " & cantidadLuces
    End If
End Function

Private Function obtenerInformeGraficos() As String
    Dim X, Y, loopCapa As Integer
    Dim cantidadLuces As Integer
    Dim capaCantidad(1 To CANTIDAD_CAPAS) As Integer
    
    obtenerInformeGraficos = ""
    
    For X = X_MINIMO_USABLE To X_MAXIMO_USABLE
        For Y = X_MINIMO_USABLE To X_MAXIMO_USABLE
            For loopCapa = 1 To CANTIDAD_CAPAS
                If mapdata(X, Y).Graphic(loopCapa).GrhIndex > 0 Then
                    capaCantidad(loopCapa) = capaCantidad(loopCapa) + 1
                End If
            Next loopCapa
        Next Y
    Next X
    
    For loopCapa = 1 To CANTIDAD_CAPAS
        obtenerInformeGraficos = obtenerInformeGraficos & "Capa " & loopCapa & ": " & capaCantidad(loopCapa) & vbCrLf
    Next loopCapa
End Function

Public Function obtenerInforme() As String
    Dim informe As String
    
    informe = "------- Objetos --------" & vbCrLf
    informe = informe & " X " & vbTab & " Y " & vbTab & "    Objeto (cantidad)" & vbCrLf
    informe = informe & obtenerInformeObjetos & vbCrLf

    informe = informe & "------- Criaturas --------" & vbCrLf
    informe = informe & " X " & vbTab & " Y " & vbTab & "    Criatura" & vbCrLf
    informe = informe & obtenerInformeCriaturas & vbCrLf
    
    informe = informe & "------- Entidades --------" & vbCrLf
    informe = informe & " X " & vbTab & " Y " & vbTab & "    Entidad" & vbCrLf
    informe = informe & obtenerInformeEntidades & vbCrLf
    
    informe = informe & "-------   Luces   --------" & vbCrLf
    informe = informe & obtenerInformeLuces & vbCrLf & vbCrLf
    
    informe = informe & "-------   Graficos -------" & vbCrLf
    informe = informe & obtenerInformeGraficos & vbCrLf
    
    
    obtenerInforme = informe

End Function
