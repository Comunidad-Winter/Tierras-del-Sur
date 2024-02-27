Attribute VB_Name = "Me_indexar_Sonidos"
Option Explicit

Public Type tSonido
    nombre As String
    tipo As Byte 'Efecto = 0. Sonido = 1
End Type

Public Sonidos() As tSonido

Private Const Archivo = "Sonidos.ini"
Private Const archivo_compilado = "Sonidos.ind"
Private Const HEAD_ELEMENTO = ""
Private Const CDM_IDENTIFICADOR = "SONIDO"

Public Function nuevo() As Integer
    'Busco alguno que este libre
    Dim elemento As Integer
    
    #If Colaborativo = 0 Then
        nuevo = -1
        
        For elemento = 1 To UBound(Sonidos)
            If Not existe(elemento) Then
                nuevo = elemento
                Exit For
            End If
        Next
        
        'No tengo slot libre. Creo uno
        If nuevo = -1 Then
            ReDim Preserve Sonidos(0 To UBound(Sonidos) + 1) As tSonido
            nuevo = UBound(Sonidos)
        End If
    #Else
        
        nuevo = CDM.cerebro.SolicitarRecurso(CDM_IDENTIFICADOR)
        
        If nuevo > UBound(Sonidos) Then
            ReDim Preserve Sonidos(0 To nuevo) As tSonido
        End If
        
        Call versionador.creado(CDM_IDENTIFICADOR, nuevo)
    #End If
End Function

Public Function eliminar(id As Integer)
    Dim nombreBackup As String
    
    nombreBackup = Sonidos(id).nombre
    
    Call resetear(Sonidos(id))
    
    Call actualizarEnIni(id)
    
    If id = UBound(Sonidos) Then
        ReDim Preserve Sonidos(0 To UBound(Sonidos) - 1) As tSonido
    End If
    
    #If Colaborativo = 1 Then
        Call versionador.eliminado(CDM_IDENTIFICADOR, id, nombreBackup)
    #End If
End Function

Private Sub resetear(ByRef sonido As tSonido)
    sonido.nombre = ""
    sonido.tipo = 0
End Sub


Public Function existe(ByVal id As Integer) As Boolean
    Dim direccion As Byte
    
    existe = True
        
    If id > UBound(Sonidos) Then
        existe = False
        Exit Function
    End If
    
    If Sonidos(id).nombre = "" Then
        existe = False
    End If

End Function

Public Function compilar() As Boolean
'No se compila para el cliente
End Function

'*****************************************************************************
'******************** PERSISTENCIA *******************************************
Public Function cargarDesdeIni() As Boolean
    Dim Soport  As New cIniManager
    Dim cantidad As Integer
    Dim loopElemento As Long
        
    If LenB(Dir(DBPath & Archivo, vbArchive)) = 0 Then
        MsgBox "No existe " & Archivo & " en la carpeta " & DBPath
        Exit Function
    End If

    Soport.Initialize DBPath & Archivo
    
    cantidad = CInt(val(Soport.getNameLastSection))

    ReDim Sonidos(0 To cantidad) As tSonido

    For loopElemento = 1 To cantidad
    
        With Sonidos(loopElemento)
            .nombre = Soport.getValue(HEAD_ELEMENTO & loopElemento, "NOMBRE")
            .tipo = CByte(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "TIPO")))
        End With
    Next
    
    cargarDesdeIni = True
    
End Function

Public Sub actualizarEnIni(ByVal Numero As Long)
       
    With Sonidos(Numero)
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "NOMBRE", .nombre
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "TIPO", .tipo
    End With
    
    #If Colaborativo = 1 Then
        If existe(Numero) Then
            Call versionador.modificado(CDM_IDENTIFICADOR, Numero, Sonidos(Numero).nombre)
        End If
    #End If
End Sub
