Attribute VB_Name = "Me_indexar_Cuerpos"
Option Explicit

Private Const Archivo = "Cuerpos.ini"
Private Const archivo_compilado = "Cuerpos.ind"
Private Const HEAD_ELEMENTO = ""
Private Const CDM_IDENTIFICADOR = "CUERPO"


Public Function existe(ByVal id As Integer) As Boolean
    
    Dim direccion As Byte
    
    If id > UBound(BodyData) Then
        existe = False
        Exit Function
    End If
    
    For direccion = E_Heading.NORTH To E_Heading.WEST
        existe = existe Or (BodyData(id).Walk(direccion).GrhIndex > 0)
    Next

End Function

Public Function nuevo() As Integer

    #If Colaborativo = 0 Then
        'Busco alguno que este libre
        Dim elemento As Integer
        
        nuevo = -1
        
        For elemento = 1 To UBound(BodyData)
            If Not existe(elemento) Then
                nuevo = elemento
                Exit For
            End If
        Next
        
        'No tengo slot libre. Creo uno
        If nuevo = -1 Then
            ReDim Preserve BodyData(0 To UBound(BodyData) + 1) As BodyData
            nuevo = UBound(BodyData)
        End If
    #Else
        nuevo = CDM.cerebro.SolicitarRecurso(CDM_IDENTIFICADOR)
        
        If nuevo > UBound(BodyData) Then
            ReDim Preserve BodyData(0 To nuevo) As BodyData
        End If
        
        Call versionador.creado(CDM_IDENTIFICADOR, nuevo)
    #End If
End Function

Public Function eliminar(id As Integer)
    Dim nombreBackup As String
    
    nombreBackup = BodyData(id).nombre
    
    Call resetear(BodyData(id))
    
    Call actualizarEnIni(id)
    
    If id = UBound(BodyData) Then
        ReDim Preserve BodyData(0 To UBound(BodyData) - 1) As BodyData
    End If
    
    #If Colaborativo = 1 Then
        Call versionador.eliminado(CDM_IDENTIFICADOR, id, nombreBackup)
    #End If
End Function

Private Sub resetear(body As BodyData)

    Dim direccion As Integer
    
    For direccion = E_Heading.NORTH To E_Heading.WEST
        body.Walk(direccion).GrhIndex = 0
    Next
    
    body.HeadOffset.x = 0
    body.HeadOffset.y = 0
    
    body.nombre = ""

End Sub

Public Function compilar() As Boolean
Dim Archivo As Integer
Dim cuerpo As tIndiceCuerpo
Dim direccion As Integer
Dim i As Integer

Archivo = FreeFile


Open Clientpath & "Init\" & archivo_compilado For Binary Access Write As #Archivo

    'Guardamos la cantidad de cabezas
    Put #Archivo, , CInt(UBound(BodyData))
    
    For i = 1 To UBound(BodyData)
    
        For direccion = E_Heading.NORTH To E_Heading.WEST
            cuerpo.body(direccion) = BodyData(i).Walk(direccion).GrhIndex
        Next
        
        cuerpo.HeadOffsetX = BodyData(i).HeadOffset.x
        cuerpo.HeadOffsetY = BodyData(i).HeadOffset.y
            
        Put #Archivo, , cuerpo
            
    Next i

Close #Archivo

compilar = True

End Function

'*****************************************************************************
'******************** PERSISTENCIA *******************************************
Public Function cargarCuerpoEnIni() As Boolean
    Dim Soport  As New cIniManager
    Dim cantidadCuerpos As Integer
    Dim loopCuerpo     As Long
        
    If LenB(Dir(DBPath & "Cuerpos.ini", vbArchive)) = 0 Then
        MsgBox "No existe Cuerpos.ini en la carpeta " & DBPath
        Exit Function
    End If

    Soport.Initialize DBPath & Archivo
        
    cantidadCuerpos = CInt(val(Soport.getNameLastSection))

    ReDim BodyData(0 To cantidadCuerpos) As BodyData

    For loopCuerpo = 1 To cantidadCuerpos
    
        With BodyData(loopCuerpo)
        
            BodyData(loopCuerpo).nombre = Soport.getValue(HEAD_ELEMENTO & loopCuerpo, "NOMBRE")
            
            BodyData(loopCuerpo).HeadOffset.x = CInt(val(Soport.getValue(HEAD_ELEMENTO & loopCuerpo, "OFFSETX")))
            BodyData(loopCuerpo).HeadOffset.y = CInt(val(Soport.getValue(HEAD_ELEMENTO & loopCuerpo, "OFFSETY")))
            
            InitGrh BodyData(loopCuerpo).Walk(E_Heading.NORTH), CInt(val(Soport.getValue(HEAD_ELEMENTO & loopCuerpo, "NORTE"))), 0
            InitGrh BodyData(loopCuerpo).Walk(E_Heading.EAST), CInt(val(Soport.getValue(HEAD_ELEMENTO & loopCuerpo, "ESTE"))), 0
            InitGrh BodyData(loopCuerpo).Walk(E_Heading.SOUTH), CInt(val(Soport.getValue(HEAD_ELEMENTO & loopCuerpo, "SUR"))), 0
            InitGrh BodyData(loopCuerpo).Walk(E_Heading.WEST), CInt(val(Soport.getValue(HEAD_ELEMENTO & loopCuerpo, "OESTE"))), 0
        
        End With
    Next
    
    cargarCuerpoEnIni = True
    
End Function

Public Sub actualizarEnIni(ByVal Numero As Long)

    With BodyData(Numero)
        
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "NOMBRE", .nombre
                
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "NORTE", .Walk(E_Heading.NORTH).GrhIndex
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "SUR", .Walk(E_Heading.SOUTH).GrhIndex
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "ESTE", .Walk(E_Heading.EAST).GrhIndex
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "OESTE", .Walk(E_Heading.WEST).GrhIndex
        
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "OFFSETX", .HeadOffset.x
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "OFFSETY", .HeadOffset.y
    End With
    
    #If Colaborativo = 1 Then
        If existe(Numero) Then
            Call versionador.modificado(CDM_IDENTIFICADOR, Numero, BodyData(Numero).nombre)
        End If
    #End If
End Sub
