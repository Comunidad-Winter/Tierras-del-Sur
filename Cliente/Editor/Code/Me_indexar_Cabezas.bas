Attribute VB_Name = "Me_indexar_Cabezas"
Option Explicit

Private Const Archivo = "Cabezas.ini"
Private Const archivo_compilado = "Cabezas.ind"
Private Const HEAD_ELEMENTO = ""
Private Const CDM_IDENTIFICADOR = "CABEZA"

Public Function nuevo() As Integer

    #If Colaborativo = 0 Then
        'Busco alguno que este libre
        Dim elemento As Integer
        nuevo = -1
    
        For elemento = 1 To UBound(HeadData)
            If Not existe(elemento) Then
                nuevo = elemento
                Exit For
            End If
        Next
        
        'No tengo slot libre. Creo uno
        If nuevo = -1 Then
            ReDim Preserve HeadData(0 To UBound(HeadData) + 1) As HeadData
            nuevo = UBound(HeadData)
        End If
    #Else
        nuevo = CDM.cerebro.SolicitarRecurso(CDM_IDENTIFICADOR)
        
        If nuevo > UBound(HeadData) Then
            ReDim Preserve HeadData(0 To nuevo) As HeadData
        End If
        
        Call versionador.creado(CDM_IDENTIFICADOR, nuevo)
    #End If

End Function

Public Function eliminar(id As Integer)
    Dim nombreBackup As String
    
    nombreBackup = HeadData(id).nombre
    
    Call resetear(HeadData(id))
    
    Call actualizarEnIni(id)
    
    If id = UBound(HeadData) Then
        ReDim Preserve HeadData(0 To UBound(HeadData) - 1) As HeadData
    End If

    #If Colaborativo = 1 Then
        Call versionador.eliminado(CDM_IDENTIFICADOR, id, nombreBackup)
    #End If
End Function

Private Sub resetear(cabeza As HeadData)

    Dim direccion As Integer
    
    For direccion = E_Heading.NORTH To E_Heading.WEST
        cabeza.Head(direccion).GrhIndex = 0
    Next
    
    cabeza.nombre = ""

End Sub

Public Function existe(ByVal id As Integer) As Boolean
    Dim direccion As Byte
    
    If id > UBound(HeadData) Then
        existe = False
        Exit Function
    End If
    
    For direccion = E_Heading.NORTH To E_Heading.WEST
        existe = existe Or (HeadData(id).Head(direccion).GrhIndex > 0)
    Next

End Function

Public Function compilar() As Boolean
    Dim Archivo As Integer
    Dim casco As tIndiceCabeza
    Dim direccion As Integer
    Dim i As Integer
    
    Archivo = FreeFile

    Open Clientpath & "Init\" & archivo_compilado For Binary Access Write As #Archivo
    
        'Guardamos la cantidad de cabezas
        Put #Archivo, , CInt(UBound(HeadData))
        
        For i = 1 To UBound(HeadData)
        
            For direccion = E_Heading.NORTH To E_Heading.WEST
                casco.Head(direccion) = HeadData(i).Head(direccion).GrhIndex
            Next
            
            Put #Archivo, , casco
                
        Next i
    
    Close #Archivo
    
    compilar = True
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

    ReDim HeadData(0 To cantidad) As HeadData
    
    For loopElemento = 1 To cantidad
    
        With HeadData(loopElemento)
        
            .nombre = Soport.getValue(HEAD_ELEMENTO & loopElemento, "NOMBRE")

            InitGrh .Head(E_Heading.NORTH), CInt(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "NORTE"))), 0
            InitGrh .Head(E_Heading.EAST), CInt(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "ESTE"))), 0
            InitGrh .Head(E_Heading.SOUTH), CInt(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "SUR"))), 0
            InitGrh .Head(E_Heading.WEST), CInt(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "OESTE"))), 0
        
        End With
    Next
    
    cargarDesdeIni = True
    
End Function


Public Sub actualizarEnIni(ByVal Numero As Long)
    
    With HeadData(Numero)
        
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "NOMBRE", .nombre
                
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "NORTE", .Head(E_Heading.NORTH).GrhIndex
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "SUR", .Head(E_Heading.SOUTH).GrhIndex
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "ESTE", .Head(E_Heading.EAST).GrhIndex
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "OESTE", .Head(E_Heading.WEST).GrhIndex
        
    End With
    
    #If Colaborativo = 1 Then
        If existe(Numero) Then
            Call versionador.modificado(CDM_IDENTIFICADOR, Numero, HeadData(Numero).nombre)
        End If
    #End If
End Sub
