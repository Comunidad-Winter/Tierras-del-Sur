Attribute VB_Name = "Me_indexar_Efectos"
Option Explicit

'Lo necesito apra genera rel archivo compilado. La estructura tindicefx del editor tiene
' el nombre, y en el cliente no.
Public Type tIndiceFxAuxiliar
    Animacion As Integer
    offsetX As Single
    offsetY As Single
    particula As Integer
    wav As Integer
End Type

Private Const archivo = "Efectos.ini"
Private Const archivo_compilado = "Efectos.ind"
Private Const HEAD_ELEMENTO = ""
Private Const CDM_IDENTIFICADOR = "EFECTO"

Public Function nuevo() As Integer
    
    #If Colaborativo = 0 Then
        'Busco alguno que este libre
        Dim elemento As Integer
        
        nuevo = -1
        
        For elemento = 1 To UBound(FxData)
            If Not existe(elemento) Then
                nuevo = elemento
                Exit For
            End If
        Next
        
        'No tengo slot libre. Creo uno
        If nuevo = -1 Then
            ReDim Preserve FxData(0 To UBound(FxData) + 1) As tIndiceFx
            nuevo = UBound(FxData)
        End If
    #Else
        
        nuevo = CDM.cerebro.SolicitarRecurso(CDM_IDENTIFICADOR)

        ' ¿Me entra en memoria?
        If nuevo > UBound(FxData) Then
            ReDim Preserve FxData(0 To nuevo) As tIndiceFx
        End If
    
        Call versionador.creado(CDM_IDENTIFICADOR, nuevo)
    #End If
    
End Function

Public Function eliminar(id As Integer)
    Dim nombreBackup As String
    
    nombreBackup = FxData(id).nombre
    
    Call resetear(FxData(id))
    
    Call actualizarEnIni(id)
    
    If id = UBound(FxData) Then
        ReDim Preserve FxData(0 To UBound(FxData) - 1) As tIndiceFx
    End If
    
    #If Colaborativo = 1 Then
        Call versionador.eliminado(CDM_IDENTIFICADOR, id, nombreBackup)
    #End If
    
End Function

Private Sub resetear(ByRef efecto As tIndiceFx)
    efecto.nombre = ""
    efecto.particula = 0
    efecto.Animacion = 0
    efecto.wav = 0
    
    efecto.offsetX = 0
    efecto.offsetY = 0
End Sub


Public Function existe(ByVal id As Integer) As Boolean
    
    Dim direccion As Byte
    
    existe = True
    
    If id > UBound(FxData) Then
        existe = False
        Exit Function
    End If
    
    If FxData(id).Animacion = 0 And FxData(id).particula = 0 And FxData(id).wav = 0 Then
        existe = False
    End If

End Function

'*****************************************************************************
'******************** COMPILACION *******************************************
Public Function compilar() As Boolean
    Dim archivo As Integer
    Dim efecto As tIndiceFxAuxiliar
    Dim direccion As Integer
    Dim i As Integer
    
    archivo = FreeFile

    Open Clientpath & "Init\" & archivo_compilado For Binary Access Write As #archivo
            
        'Guardamos la cantidad de cabezas
        Put #archivo, , CInt(UBound(FxData))
        
        For i = 1 To UBound(FxData)
        
            efecto.Animacion = FxData(i).Animacion
            efecto.particula = FxData(i).particula
            efecto.wav = FxData(i).wav
            
            efecto.offsetX = FxData(i).offsetX
            efecto.offsetY = FxData(i).offsetY
         
            Put #archivo, , efecto
                
        Next i
    
    Close #archivo
    
    compilar = True
End Function
'*****************************************************************************
'******************** PERSISTENCIA *******************************************

Public Function cargarDesdeIni() As Boolean
    Dim Soport  As New cIniManager
    Dim cantidad As Integer
    Dim loopElemento As Long
        
    If LenB(Dir(DBPath & archivo, vbArchive)) = 0 Then
        MsgBox "No existe " & archivo & " en la carpeta " & DBPath
        Exit Function
    End If

    Soport.Initialize DBPath & archivo
    
    cantidad = CInt(val(Soport.getNameLastSection))

    ReDim FxData(0 To cantidad) As tIndiceFx
    
    For loopElemento = 1 To cantidad
    
        With FxData(loopElemento)
        
            .nombre = Soport.getValue(HEAD_ELEMENTO & loopElemento, "NOMBRE")
            .Animacion = CInt(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "ANIMACION")))
            .wav = CInt(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "SONIDO")))
            .particula = CInt(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "PARTICULA")))
             
            .offsetX = CInt(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "OFFSETX")))
            .offsetY = CInt(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "OFFSETY")))
        End With
    Next
    
   
    cargarDesdeIni = True
    
End Function

Public Sub actualizarEnIni(ByVal Numero As Long)

    With FxData(Numero)

        WriteVar DBPath & archivo, HEAD_ELEMENTO & CStr(Numero), "NOMBRE", .nombre
                
        WriteVar DBPath & archivo, HEAD_ELEMENTO & CStr(Numero), "ANIMACION", .Animacion
        WriteVar DBPath & archivo, HEAD_ELEMENTO & CStr(Numero), "SONIDO", .wav
        WriteVar DBPath & archivo, HEAD_ELEMENTO & CStr(Numero), "PARTICULA", .particula
        
        WriteVar DBPath & archivo, HEAD_ELEMENTO & CStr(Numero), "OFFSETX", .offsetX
        WriteVar DBPath & archivo, HEAD_ELEMENTO & CStr(Numero), "OFFSETY", .offsetY
    End With
    
    #If Colaborativo = 1 Then
        If existe(Numero) Then
            Call versionador.modificado(CDM_IDENTIFICADOR, Numero, FxData(Numero).nombre)
        End If
    #End If
End Sub
