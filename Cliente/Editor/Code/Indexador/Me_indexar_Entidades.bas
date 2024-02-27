Attribute VB_Name = "Me_indexar_Entidades"
Option Explicit


Private Const Archivo = "Entidades.ini"
Private Const archivo_compilado = "Entidades.ind"
Private Const HEAD_ELEMENTO = ""
Private Const CDM_IDENTIFICADOR = "ENTIDAD"

Public Function nuevo() As Integer
    
    nuevo = -1
    
    #If Colaborativo = 0 Then
        Dim elemento As Integer
    
        'Busco alguno que este libre
        For elemento = 1 To UBound(EntidadesIndexadas)
            If Not existe(elemento) Then
                nuevo = elemento
                Exit For
            End If
        Next
    
    '    'No tengo slot libre. Creo uno
        If nuevo = -1 Then
            ReDim Preserve EntidadesIndexadas(0 To UBound(EntidadesIndexadas) + 1) As tIndiceEntidad
            nuevo = UBound(EntidadesIndexadas)
        End If
    #Else
        
        nuevo = CDM.cerebro.SolicitarRecurso(CDM_IDENTIFICADOR)
        
        If nuevo > UBound(EntidadesIndexadas) Then
            ReDim Preserve EntidadesIndexadas(0 To nuevo) As tIndiceEntidad
        End If
        
        Call versionador.creado(CDM_IDENTIFICADOR, nuevo)
    #End If

    If Not nuevo = -1 Then
        Call resetear(EntidadesIndexadas(nuevo))
    End If
End Function

Public Function eliminar(id As Integer)
    Dim nombreBackup As String
    
    nombreBackup = EntidadesIndexadas(id).nombre
    
    Call resetear(EntidadesIndexadas(id))
    
    Call actualizarEnIni(id)
    
    If id = UBound(EntidadesIndexadas) Then
        ReDim Preserve EntidadesIndexadas(0 To UBound(EntidadesIndexadas) - 1) As tIndiceEntidad
    End If
    
    #If Colaborativo = 1 Then
        Call versionador.eliminado(CDM_IDENTIFICADOR, id, nombreBackup)
    #End If
End Function

Private Sub resetear(entidad As tIndiceEntidad)

    entidad.tipo = eTipoEntidadVida.Nulo
    entidad.nombre = ""
    entidad.CrearAlMorir = 0
    entidad.Vida = 0
    entidad.Proyectil = 0
    
    'Sonidos
    ReDim entidad.Sonidos(0)
    
    'Sonidos al pegar
    ReDim entidad.SonidosAlPegar(0)
    
    'Particulas
    ReDim entidad.Particulas(0)
    
    'Graficos
    ReDim entidad.Graficos(0)
    
    'Luz
    entidad.luz.LuzTipo = 0
    entidad.luz.LuzRadio = 0
    
    entidad.luz.LuzColor.r = 0
    entidad.luz.LuzColor.g = 0
    entidad.luz.LuzColor.b = 0

End Sub

Public Function existe(ByVal id As Integer) As Boolean
 
    If id > UBound(EntidadesIndexadas()) Then existe = False:    Exit Function
    
    If EntidadesIndexadas(id).tipo = eTipoEntidadVida.Nulo Then existe = False: Exit Function
    
    existe = True

End Function

Public Function compilar() As Boolean
    Dim Archivo As Integer
    Dim cabeza As tIndiceCabeza
    Dim direccion As Integer
    Dim i As Integer
    
    Archivo = FreeFile

    Open Clientpath & "Init\" & archivo_compilado For Binary Access Write As #Archivo
    
        'Escribimos la version del archivo
        Put #Archivo, , CLng(0)
        
        'Guardamos la cantidad de cabezas
        Put #Archivo, , CLng(UBound(CascoAnimData))
        
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
    Dim tmp As String
    Dim vector() As String
    Dim loopParte As Integer
    
    If LenB(Dir(DBPath & Archivo, vbArchive)) = 0 Then
        MsgBox "No existe " & Archivo & " en la carpeta " & DBPath
        Exit Function
    End If

    Soport.Initialize DBPath & Archivo
        
    cantidad = CInt(val(Soport.getNameLastSection))

    ReDim EntidadesIndexadas(0 To cantidad) As tIndiceEntidad
    
    For loopElemento = 1 To cantidad
    
        With EntidadesIndexadas(loopElemento)
            'Generales
            .nombre = Soport.getValue(HEAD_ELEMENTO & loopElemento, "NOMBRE")
            
            .tipo = val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "TIPO"))
            .Vida = val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "VIDA"))
            .Proyectil = val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "PROYECTIL"))
            
            If Not .tipo = eTipoEntidadVida.Nulo Then
                'Cargamos los gráficos
                tmp = Soport.getValue(HEAD_ELEMENTO & loopElemento, "GRAFICOS")
                
                vector = Split(tmp, "-")
                
                ReDim .Graficos(UBound(vector))
                For loopParte = 0 To UBound(vector)
                    .Graficos(loopParte) = val(vector(loopParte))
                Next loopParte
        
                'Cargamos los sonidos
                tmp = Soport.getValue(HEAD_ELEMENTO & loopElemento, "SONIDOS")
                vector = Split(tmp, " ")
    
                ReDim .Sonidos(UBound(vector))
                For loopParte = 0 To UBound(vector)
                    .Sonidos(loopParte) = val(vector(loopParte))
                    'Loop infinito?
                    If right(vector(loopParte), 1) = "L" Then .Sonidos(loopParte) = .Sonidos(loopParte) * -1
                Next loopParte
                
                'Cargamos la informacion de que pasa cuando la entidad pierde vida
                tmp = Soport.getValue(HEAD_ELEMENTO & loopElemento, "ALPERDERVIDA")
                vector = Split(tmp, "-")
                
                ReDim .SonidosAlPegar(UBound(vector))
                For loopParte = 0 To UBound(vector)
                    .SonidosAlPegar(loopParte) = val(vector(loopParte))
                Next loopParte
                
                'Cargamos las particulas
                tmp = Soport.getValue(HEAD_ELEMENTO & loopElemento, "PARTICULAS")
                vector = Split(tmp, "-")
    
                ReDim .Particulas(UBound(vector))
                For loopParte = 0 To UBound(vector)
                    .Particulas(loopParte) = val(vector(loopParte))
                Next loopParte
    
                'Luz
                tmp = Soport.getValue(HEAD_ELEMENTO & loopElemento, "LUZ")
                vector = Split(tmp, "-")
                
                .luz.LuzRadio = val(vector(0))
                .luz.LuzTipo = val(vector(1))
                .luz.LuzBrillo = val(vector(2))
                
                .luz.LuzColor.r = val(vector(3))
                .luz.LuzColor.g = val(vector(4))
                .luz.LuzColor.b = val(vector(5))
                
                .luz.luzInicio = val(vector(6))
                .luz.luzFin = val(vector(7))
            End If
        End With
    Next
    
    cargarDesdeIni = True
    
End Function
Private Function JoinS(vector() As Integer, delimitador As String) As String
    Dim loopParte As Integer
    Dim tmp As String
    
    For loopParte = LBound(vector) To UBound(vector)
        tmp = tmp & delimitador & vector(loopParte)
    Next
    
    JoinS = mid$(tmp, 2)
End Function
Public Sub actualizarEnIni(ByVal Numero As Long)
    
    With EntidadesIndexadas(Numero)
            
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "NOMBRE", .nombre
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "TIPO", .tipo
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "VIDA", .Vida
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "PROYECTIL", .Proyectil
        
        'Graficos
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "GRAFICOS", JoinS(.Graficos, "-")
        
        'Sonidos
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "SONIDOS", JoinS(.Sonidos, " ")
        
        'Al perder vida
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "ALPERDERVIDA", JoinS(.SonidosAlPegar, "-")
        
        'Particulas
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "PARTICULAS", JoinS(.Particulas, "-")
        
        'Luz
        WriteVar DBPath & Archivo, HEAD_ELEMENTO & CStr(Numero), "LUZ", .luz.LuzRadio & "-" & .luz.LuzTipo & "-" & .luz.LuzBrillo & "-" & .luz.LuzColor.r & "-" & .luz.LuzColor.g & "-" & .luz.LuzColor.b & "-" & .luz.luzInicio & "-" & .luz.luzFin
    End With
    
    #If Colaborativo = 1 Then
        If existe(Numero) Then
            Call versionador.modificado(CDM_IDENTIFICADOR, Numero, EntidadesIndexadas(Numero).nombre)
        End If
    #End If
End Sub
