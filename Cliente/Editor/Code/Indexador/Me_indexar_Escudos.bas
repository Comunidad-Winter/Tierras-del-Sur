Attribute VB_Name = "Me_indexar_Escudos"
Option Explicit

Private Const Archivo = "Escudos.ini"
Private Const archivo_compilado = "Escudos.ind"
Private Const HEAD_ELEMENTO = ""
Private Const CDM_IDENTIFICADOR = "ESCUDO"

Public Function nuevo() As Integer
    
    #If Colaborativo = 0 Then
        'Busco alguno que este libre
        Dim elemento As Integer
        
        nuevo = -1
        
        For elemento = 1 To UBound(ShieldAnimData)
            If Not existe(elemento) Then
                nuevo = elemento
                Exit For
            End If
        Next
        
        'No tengo slot libre. Creo uno
        If nuevo = -1 Then
            ReDim Preserve ShieldAnimData(0 To UBound(ShieldAnimData) + 1) As ShieldAnimData
            nuevo = UBound(ShieldAnimData)
        End If
    #Else
        nuevo = CDM.cerebro.SolicitarRecurso(CDM_IDENTIFICADOR)
        
        If nuevo > UBound(ShieldAnimData) Then
            ReDim Preserve ShieldAnimData(0 To nuevo) As ShieldAnimData
        End If
        
        Call versionador.creado(CDM_IDENTIFICADOR, nuevo)
    #End If

End Function

Public Function eliminar(id As Integer)
    Dim nombreBackup As String
    
    nombreBackup = ShieldAnimData(id).nombre
    
    Call resetear(ShieldAnimData(id))
    
    Call actualizarEnIni(id)
    
    If id = UBound(ShieldAnimData) Then
        ReDim Preserve ShieldAnimData(0 To UBound(ShieldAnimData) - 1) As ShieldAnimData
    End If
    
    #If Colaborativo = 1 Then
        Call versionador.eliminado(CDM_IDENTIFICADOR, id, nombreBackup)
    #End If
    
End Function

Private Sub resetear(escudo As ShieldAnimData)

    Dim direccion As Integer
    
    For direccion = E_Heading.NORTH To E_Heading.WEST
        escudo.ShieldWalk(direccion).GrhIndex = 0
    Next
    
    escudo.nombre = ""

End Sub


Public Function existe(ByVal id As Integer) As Boolean
    Dim direccion As Byte
    
    If id > UBound(ShieldAnimData) Then
        existe = False
        Exit Function
    End If
    
    For direccion = E_Heading.NORTH To E_Heading.WEST
        existe = existe Or (ShieldAnimData(id).ShieldWalk(direccion).GrhIndex > 0)
    Next

End Function
'*****************************************************************************
'******************** COMPILACION *******************************************

Public Function compilar() As Boolean
    Dim Archivo As Integer
    Dim escudo As tIndiceEscudo
    Dim direccion As Integer
    Dim i As Integer
    
    Archivo = FreeFile

    Open Clientpath & "Init\" & archivo_compilado For Binary Access Write As #Archivo
            
        'Guardamos la cantidad de cabezas
        Put #Archivo, , CInt(UBound(ShieldAnimData))
        
        For i = 1 To UBound(ShieldAnimData)
        
            For direccion = E_Heading.NORTH To E_Heading.WEST
                escudo.Walk(direccion) = ShieldAnimData(i).ShieldWalk(direccion).GrhIndex
            Next
            
            Put #Archivo, , escudo
                
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

    ReDim ShieldAnimData(0 To cantidad) As ShieldAnimData
    
    For loopElemento = 1 To cantidad
    
        With ShieldAnimData(loopElemento)
        
            .nombre = Soport.getValue(HEAD_ELEMENTO & loopElemento, "NOMBRE")

            InitGrh .ShieldWalk(E_Heading.NORTH), CInt(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "NORTE"))), 0
            InitGrh .ShieldWalk(E_Heading.EAST), CInt(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "ESTE"))), 0
            InitGrh .ShieldWalk(E_Heading.SOUTH), CInt(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "SUR"))), 0
            InitGrh .ShieldWalk(E_Heading.WEST), CInt(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "OESTE"))), 0
        
        End With
    Next
    
    cargarDesdeIni = True
    
End Function

Public Sub actualizarEnIni(ByVal Numero As Long)
    With ShieldAnimData(Numero)
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "NOMBRE", .nombre
                
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "NORTE", .ShieldWalk(E_Heading.NORTH).GrhIndex
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "SUR", .ShieldWalk(E_Heading.SOUTH).GrhIndex
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "ESTE", .ShieldWalk(E_Heading.EAST).GrhIndex
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "OESTE", .ShieldWalk(E_Heading.WEST).GrhIndex
    End With
    
    #If Colaborativo = 1 Then
        If existe(Numero) Then
            Call versionador.modificado(CDM_IDENTIFICADOR, Numero, ShieldAnimData(Numero).nombre)
        End If
    #End If
End Sub
