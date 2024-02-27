Attribute VB_Name = "Me_indexar_Cascos"
Option Explicit

Private Const Archivo = "Cascos.ini"
Private Const archivo_compilado = "Cascos.ind"
Private Const HEAD_ELEMENTO = ""
Private Const CDM_IDENTIFICADOR = "CASCO"

Public Function nuevo() As Integer
    
    #If Colaborativo = 0 Then
        'Busco alguno que este libre
        Dim elemento As Integer
        
        nuevo = -1
        
        For elemento = 1 To UBound(CascoAnimData)
            If Not existe(elemento) Then
                nuevo = elemento
                Exit For
            End If
        Next
        
        'No tengo slot libre. Creo uno
        If nuevo = -1 Then
            ReDim Preserve CascoAnimData(0 To UBound(CascoAnimData) + 1) As HeadData
            nuevo = UBound(CascoAnimData)
        End If
    #Else
        
        nuevo = CDM.cerebro.SolicitarRecurso(CDM_IDENTIFICADOR)
        
        If nuevo > UBound(CascoAnimData) Then
            ReDim Preserve CascoAnimData(0 To nuevo) As HeadData
        End If
        
        Call versionador.creado(CDM_IDENTIFICADOR, nuevo)
    #End If
    
End Function

Public Function eliminar(id As Integer)
    Dim nombreBackup As String
    
    nombreBackup = CascoAnimData(id).nombre
    
    Call resetear(CascoAnimData(id))
    
    Call actualizarEnIni(id)
    
    If id = UBound(CascoAnimData) Then
        ReDim Preserve CascoAnimData(0 To UBound(CascoAnimData) - 1) As HeadData
    End If
    
    #If Colaborativo = 1 Then
        Call versionador.eliminado(CDM_IDENTIFICADOR, id, nombreBackup)
    #End If
    
End Function

Private Sub resetear(casco As HeadData)

    Dim direccion As Integer
    
    For direccion = E_Heading.NORTH To E_Heading.WEST
        casco.Head(direccion).GrhIndex = 0
    Next
    
    casco.nombre = ""

End Sub

Public Function existe(ByVal id As Integer) As Boolean
    
    Dim direccion As Byte
    
    If id > UBound(CascoAnimData) Then
        existe = False
        Exit Function
    End If
    
    For direccion = E_Heading.NORTH To E_Heading.WEST
        existe = existe Or (CascoAnimData(id).Head(direccion).GrhIndex > 0)
    Next

End Function


Public Function compilar() As Boolean
    Dim Archivo As Integer
    Dim cabeza As tIndiceCabeza
    Dim direccion As Integer
    Dim i As Integer
    
    Archivo = FreeFile

    Open Clientpath & "Init\" & archivo_compilado For Binary Access Write As #Archivo
    
        'Guardamos la cantidad de cabezas
        Put #Archivo, , CInt(UBound(CascoAnimData))
        
        For i = 1 To UBound(CascoAnimData)
        
            For direccion = E_Heading.NORTH To E_Heading.WEST
                cabeza.Head(direccion) = CascoAnimData(i).Head(direccion).GrhIndex
            Next
            
            Put #Archivo, , cabeza
                
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

    ReDim CascoAnimData(0 To cantidad) As HeadData
    
    For loopElemento = 1 To cantidad
    
        With CascoAnimData(loopElemento)
        
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
    With CascoAnimData(Numero)
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "NOMBRE", .nombre
                
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "NORTE", .Head(E_Heading.NORTH).GrhIndex
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "SUR", .Head(E_Heading.SOUTH).GrhIndex
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "ESTE", .Head(E_Heading.EAST).GrhIndex
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "OESTE", .Head(E_Heading.WEST).GrhIndex
    End With
    
    #If Colaborativo = 1 Then
        If existe(Numero) Then
            Call versionador.modificado(CDM_IDENTIFICADOR, Numero, CascoAnimData(Numero).nombre)
        End If
    #End If
End Sub
