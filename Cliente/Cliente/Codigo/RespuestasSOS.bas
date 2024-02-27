Attribute VB_Name = "CLI_RespuestasSOS"
Option Explicit
'Aca guardamos los valores de inicio y fin de los mensajes
'''''''''''''''''''''''''''''''''
Private Type infoRespuesta
    empieza As Integer
    longitud As Integer
    Titulo  As String
End Type

Public respuestasSOS() As infoRespuesta

Public Function cargarRespuestasSOS() As Boolean
    Dim linea As String
    Dim i, k, canal As Integer
    Dim posAnt As Long
    
    If FileExist(app.Path & "\Init\rsos.tds", vbArchive) Then
        canal = FreeFile
    
        Open app.Path & "\Init\rsos.tds" For Input As #canal
   
        posAnt = 0
        k = 0
        
        Do While Not EOF(canal)
     
            Line Input #canal, linea
       
            ReDim Preserve respuestasSOS(k)
        
            i = InStr(1, linea, "<=>")
           
            posAnt = posAnt + Len(mid(linea, 1, i - 1)) + 3
        
            respuestasSOS(k).Titulo = mid(linea, 1, i - 1)
            respuestasSOS(k).empieza = posAnt + 1
            respuestasSOS(k).longitud = Len(linea) - Len(respuestasSOS(k).Titulo) - 3
            posAnt = posAnt + respuestasSOS(k).longitud + 2
         
            k = k + 1
        Loop
     
        Close canal
        
        cargarRespuestasSOS = True
     Else
        cargarRespuestasSOS = False
     End If

End Function

Public Function obtenerRespuestaSOS(numero As Integer)
        Dim canal As Integer
        Dim aux2 As String
        
        aux2 = ""
        
        canal = FreeFile
            
        Open app.Path & "\Init\rsos.tds" For Input As #canal
        
        Seek canal, respuestasSOS(numero).empieza
        aux2 = Input(respuestasSOS(numero).longitud, canal)

        Close #canal
        
        obtenerRespuestaSOS = aux2
End Function

