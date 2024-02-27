Attribute VB_Name = "Me_indexar_Armas"
Option Explicit

Private Const Archivo = "Armas.ini"
Private Const archivo_compilado = "Armas.ind"
Private Const HEAD_ELEMENTO = ""
Private Const CDM_IDENTIFICADOR = "ARMA"

Public Function nuevo() As Integer
  
    #If Colaborativo = 0 Then
    
        'Busco alguno que este libre
        Dim elemento As Integer
        nuevo = -1
        
        For elemento = 1 To UBound(WeaponAnimData)
            If Not existe(elemento) Then
                nuevo = elemento
                Exit For
            End If
        Next
        
        'No tengo slot libre. Creo uno
        If nuevo = -1 Then
            ReDim Preserve WeaponAnimData(0 To UBound(WeaponAnimData) + 1) As WeaponAnimData
            nuevo = UBound(WeaponAnimData)
        End If
    #Else
        
        nuevo = CDM.cerebro.SolicitarRecurso(CDM_IDENTIFICADOR)
        
        If nuevo > UBound(WeaponAnimData) Then
            ReDim Preserve WeaponAnimData(0 To nuevo) As WeaponAnimData
        End If
        
        Call versionador.creado(CDM_IDENTIFICADOR, nuevo)
    #End If
End Function

Public Function eliminar(id As Integer)
    Dim nombreBackup As String
    
    nombreBackup = WeaponAnimData(id).nombre
    
    Call resetear(WeaponAnimData(id))
    
    Call actualizarEnIni(id)
    
    If id = UBound(WeaponAnimData) Then
        ReDim Preserve WeaponAnimData(0 To UBound(WeaponAnimData) - 1) As WeaponAnimData
    End If
    
    #If Colaborativo = 1 Then
        Call versionador.eliminado(CDM_IDENTIFICADOR, id, nombreBackup)
    #End If
    
End Function

Private Sub resetear(arma As WeaponAnimData)

    Dim direccion As Integer
    
    For direccion = E_Heading.NORTH To E_Heading.WEST
        arma.WeaponWalk(direccion).GrhIndex = 0
    Next
    
    arma.nombre = ""

End Sub

Public Function existe(ByVal id As Integer) As Boolean
    
    Dim direccion As Byte
    
    If id = 0 Then existe = True: Exit Function
    
    If id > UBound(WeaponAnimData) Then
        existe = False
        Exit Function
    End If
    
    For direccion = E_Heading.NORTH To E_Heading.WEST
        existe = existe Or (WeaponAnimData(id).WeaponWalk(direccion).GrhIndex > 0)
    Next

End Function


Public Function compilar() As Boolean
    Dim Archivo As Integer
    Dim arma As tIndiceArma
    Dim direccion As Integer
    Dim i As Integer
    
    Archivo = FreeFile

    Open Clientpath & "Init\" & archivo_compilado For Binary Access Write As #Archivo
            
        'Guardamos la cantidad de Armas
        Put #Archivo, , CInt(UBound(WeaponAnimData))
        
        For i = 1 To UBound(WeaponAnimData)
        
            For direccion = E_Heading.NORTH To E_Heading.WEST
                arma.Walk(direccion) = WeaponAnimData(i).WeaponWalk(direccion).GrhIndex
            Next
            
            Put #Archivo, , arma
                
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

    ReDim WeaponAnimData(0 To cantidad - 1) As WeaponAnimData
    
    For loopElemento = 0 To cantidad - 1
    
        With WeaponAnimData(loopElemento)
        
            .nombre = Soport.getValue(HEAD_ELEMENTO & loopElemento, "NOMBRE")

            InitGrh .WeaponWalk(E_Heading.NORTH), CInt(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "NORTE"))), 0
            InitGrh .WeaponWalk(E_Heading.EAST), CInt(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "ESTE"))), 0
            InitGrh .WeaponWalk(E_Heading.SOUTH), CInt(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "SUR"))), 0
            InitGrh .WeaponWalk(E_Heading.WEST), CInt(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "OESTE"))), 0
        
        End With
    Next
    
    cargarDesdeIni = True
    
End Function

Public Sub actualizarEnIni(ByVal Numero As Long)
    With WeaponAnimData(Numero)
        
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "NOMBRE", .nombre
                
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "NORTE", .WeaponWalk(E_Heading.NORTH).GrhIndex
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "SUR", .WeaponWalk(E_Heading.SOUTH).GrhIndex
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "ESTE", .WeaponWalk(E_Heading.EAST).GrhIndex
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "OESTE", .WeaponWalk(E_Heading.WEST).GrhIndex
    End With


    #If Colaborativo = 1 Then
        If existe(Numero) Then
            Call versionador.modificado(CDM_IDENTIFICADOR, Numero, WeaponAnimData(Numero).nombre)
        End If
    #End If
End Sub
