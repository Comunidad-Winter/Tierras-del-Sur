Attribute VB_Name = "Me_indexar_Pisos"
Option Explicit

Private Const archivo = "Pisos.ini"
Private Const archivo_compilado = "Pisos.ind"
Private Const HEAD_ELEMENTO = ""
Private Const CDM_IDENTIFICADOR = "PISO"

Private Const CANTIDAD_FORMATOS = 3 ' Caminos Chicos,  Caminos Grandes, Costas

Private Type MatrizRepresentacion
    datos() As String
    formato As eFormatoTileSet
End Type

' Por cada Formato, sus matrices
Private MatricesTranformacion(1 To CANTIDAD_FORMATOS) As MatrizRepresentacion

Public TmpTilesetsNum     As Integer


Public Function nuevo() As Integer
    
    
    #If Colaborativo = 0 Then
    
        Dim tileset As Integer
        
        For tileset = 1 To Tilesets_count
            If Tilesets(tileset).stage_count = 0 Then
                obtenerIDTileSetLibre = tileset
                Exit Function
            End If
        Next
        
        'Tengo que agregar un nuevo slot
        Tilesets_count = Tilesets_count + 1

        ReDim Preserve Tilesets(1 To Tilesets_count)

        nuevo = Tilesets_count
    #Else
    
        nuevo = CDM.cerebro.SolicitarRecurso(CDM_IDENTIFICADOR)
        
        'Tengo que agregar un nuevo slot
        If nuevo > UBound(Tilesets) Then
            Tilesets_count = nuevo
            ReDim Preserve Tilesets(0 To Tilesets_count) As TilesetStruct
        End If
        
        Call versionador.creado(CDM_IDENTIFICADOR, nuevo)
    #End If
    
End Function

Public Function existe(ByVal id As Integer) As Boolean

    If id > UBound(Tilesets) Then
        existe = False
        Exit Function
    End If
    
    existe = Not (Tilesets(id).filenum = 0)
End Function


Public Sub eliminar(ByVal id As Integer)
    Dim nombreBackup As String
    
    nombreBackup = Tilesets(id).nombre
    
    Call resetear(Tilesets(id))
    
    Call actualizarEnIni(id)
    
    If id = UBound(Tilesets) Then
        ReDim Preserve Tilesets(0 To UBound(Tilesets) - 1) As TilesetStruct
        Tilesets_count = UBound(Tilesets)
    End If
  
    #If Colaborativo = 1 Then
        Call versionador.eliminado(CDM_IDENTIFICADOR, id, nombreBackup)
    #End If
End Sub
Public Sub resetear(ByRef tileset As TilesetStruct)
    tileset.stage_actual = 0
    tileset.stage_count = 0
    tileset.stage_mu = 0
    
    ReDim tileset.stages(0)
    
    tileset.filenum = 0
    tileset.Olitas = 0
        
    tileset.nombre = ""
    tileset.anim = 0
End Sub

Private Function cargarRepresentacionMatricesTransformacion()
    Dim Soport  As New cIniManager
    Dim loopFormato As Byte
    Dim loopMatriz As Byte
    Dim cantidadMatrices As Integer
    Dim formato As eFormatoTileSet
    
    ' Iniciamos
    Soport.Initialize DBPath & "matrices.ini"
    
    ' Cargamos las representaciones de las matrices de transformacion
    For loopFormato = 1 To CANTIDAD_FORMATOS
    
        cantidadMatrices = CInt(val(Soport.getValue(loopFormato, "CANTIDAD")))
        formato = CInt(val(Soport.getValue(loopFormato, "FORMATO")))
        
        ' Redimensionamos
        ReDim MatricesTranformacion(loopFormato).datos(1 To cantidadMatrices)
        
        MatricesTranformacion(loopFormato).formato = formato
        
        ' Cargamos la informacion de cada matriz para este formato
        For loopMatriz = 1 To cantidadMatrices
            MatricesTranformacion(loopFormato).datos(loopMatriz) = Soport.getValue(loopFormato, loopMatriz)
        Next
        
    Next
       
    Set Soport = Nothing
End Function
Public Function CargarTilesetsIni(ByVal streamfile As String) As Boolean
    '*****************************************************************
    'Menduz
    '*****************************************************************
    
    Dim Soport  As New cIniManager
    Dim cantidad As Integer
    Dim loopElemento As Integer
    Dim loopStage As Integer
    Dim loopPaso As Integer
    Dim pasos As String
    
    If LenB(Dir(DBPath & archivo, vbArchive)) = 0 Then
        MsgBox "No existe " & archivo & " en la carpeta " & DBPath
        Exit Function
    End If
    
    ' Cargamos las representaciones de las matrices de transformacion
    Call cargarRepresentacionMatricesTransformacion

    Soport.Initialize DBPath & archivo
        
    Tilesets_count = CInt(val(Soport.getNameLastSection))
    
    ReDim Tilesets(0 To Tilesets_count)
    
    For loopElemento = 1 To Tilesets_count
        With Tilesets(loopElemento)
            .stage_count = CByte(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "Graficos")))
            .stage_mu = timeGetTime
            .anim = CInt(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "Animacion")))
            .nombre = Soport.getValue(HEAD_ELEMENTO & loopElemento, "Nombre")
            .Olitas = CInt(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "Olitas")))
            
            If .stage_count > 0 Then
                .stage_actual = 1
                
                ReDim .stages(1 To .stage_count)
                For loopStage = 1 To .stage_count
                    .stages(loopStage) = CInt(val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "Grh" & loopStage)))
                Next loopStage
                .filenum = .stages(1)
            End If
            
            ' Nuevo formato de pisos
            .formato = val(Soport.getValue(HEAD_ELEMENTO & loopElemento, "formato"))
            
            ' Referencia a la textura
            .referencia.textura = val(ReadField(1, Soport.getValue(HEAD_ELEMENTO & loopElemento, "referencia"), Asc(" ")))
            .referencia.numero = val(ReadField(2, Soport.getValue(HEAD_ELEMENTO & loopElemento, "referencia"), Asc(" ")))
            
            ' Sistema de efectos de sonido al pisar
            pasos = Soport.getValue(HEAD_ELEMENTO & loopElemento, "pasos")
            
            If Len(pasos) > 0 Then
                For loopPaso = 0 To UBound(.EfectoPisada)
                    .EfectoPisada(loopPaso) = STI(pasos, loopPaso * 2)
                Next loopPaso
            End If
         
        End With
        
        If Tilesets(loopElemento).formato = eFormatoTileSet.camino_chico Or Tilesets(loopElemento).formato = eFormatoTileSet.camino_grande_parte2 Or Tilesets(loopElemento).formato = eFormatoTileSet.costa_tipo_1_parte2 Then
            ' Genero la matriz dinamica
            Call setearMatrizTranformacion(loopElemento)
        End If
    Next loopElemento
    
      
            
    CargarTilesetsIni = True
End Function

Private Sub generarMatriz(matrizOUT() As TileSetInfoTile, dimensionOUT As Byte, ByVal representacionMatriz As String, ts_textura As Integer, ts_textura_comienzo As Integer, ts_camino As Integer, Optional ByVal ts_caminoP2 As Integer = -1)

Dim x As Byte
Dim y As Byte
Dim loopElemento As Integer
Dim seccionS As String

' Redimensiono la matriz
Dim infoMatriz() As String

' Remplazo la matriz las texturas intervinientes

' - Seteo la textura original
representacionMatriz = Replace$(representacionMatriz, "A", ts_textura)

' - Seteo la textura de caminos 1
representacionMatriz = Replace$(representacionMatriz, "B", ts_camino)

' - Seteo la textura de caminos 2
If ts_caminoP2 > -1 Then
    representacionMatriz = Replace$(representacionMatriz, "C", ts_caminoP2)
End If

' Voy a setear la matriz
infoMatriz = Split(representacionMatriz, " ")
loopElemento = 0

For y = 0 To 15
    For x = 0 To 15
    
        ' Seteo el tileset
        matrizOUT(dimensionOUT, x, y).textura = val(ReadField(1, CStr(infoMatriz(loopElemento)), Asc(",")))
         
        ' Para el número de tile me voy a fijar si hace una referencia dinamica
        seccionS = ReadField(2, CStr(infoMatriz(loopElemento)), Asc(","))
        
        '¿Hace referencia a la seccion?
        If left$(seccionS, 1) = "S" Then
            matrizOUT(dimensionOUT, x, y).numero = val(mid$(seccionS, 2)) + ts_textura_comienzo
        Else
            matrizOUT(dimensionOUT, x, y).numero = val(seccionS)
        End If
              
        
        loopElemento = loopElemento + 1
    Next x
Next y

End Sub

Private Function obtenerDatosFormato(formato As eFormatoTileSet) As Byte
Dim loopFormato As Byte

For loopFormato = 1 To CANTIDAD_FORMATOS
    
    If MatricesTranformacion(loopFormato).formato = formato Then
        obtenerDatosFormato = loopFormato
        Exit Function
    End If

Next

End Function

Public Function obtenerCantidadVirtuales(formato As eFormatoTileSet) As Byte

Dim loopFormato As Byte

For loopFormato = 1 To CANTIDAD_FORMATOS
    If MatricesTranformacion(loopFormato).formato = formato Then
        obtenerCantidadVirtuales = UBound(MatricesTranformacion(loopFormato).datos)
        Exit Function
    End If
Next

obtenerCantidadVirtuales = 0

End Function

Public Sub setearMatrizTranformacion(idtileset As Integer)

Dim numeroGrupoMatrices As Byte
Dim loopMatriz As Byte
Dim ts_textura As Integer
Dim ts_textura_comienzo As Integer ' Seccion donde comienza
Dim ts_camino As Integer
Dim ts_camino_p2 As Integer
Dim cantidadMatrices As Byte

Dim slotMatrizDatos As Byte

With Tilesets(idtileset)
    
    If .formato = camino_chico Then
    
        ts_textura = .referencia.textura ' La textura del piso
        ts_textura_comienzo = .referencia.numero ' El tile X,Y superior izquierdo
        ts_camino = idtileset ' El camino chico es el actual tileset
        ts_camino_p2 = -1 ' No tiene una segunda parte
        
    ElseIf .formato = camino_grande_parte2 Then
        
        ts_camino_p2 = idtileset
        ts_camino = .referencia.textura ' La primer parte de los caminos grandes
        ts_textura = Tilesets(ts_camino).referencia.textura
        ts_textura_comienzo = Tilesets(ts_camino).referencia.numero
        
    ElseIf .formato = costa_tipo_1_parte2 Then
    
        ts_camino_p2 = idtileset ' La segunda parte soy yo
        ts_camino = .referencia.textura ' La primera parte
        ts_textura = Tilesets(.referencia.textura).referencia.textura ' La textura del agua
        ts_textura_comienzo = 0
        
    Else
        ' No necesita
        Exit Sub
    End If
    
    slotMatrizDatos = obtenerDatosFormato(.formato)
    
    cantidadMatrices = UBound(MatricesTranformacion(slotMatrizDatos).datos)
    
End With
    
ReDim Tilesets(idtileset).matriz_transformacion(1 To cantidadMatrices, 0 To 15, 0 To 15)
    
' Generamos las matrices seteandolas done corresponde
For loopMatriz = 1 To cantidadMatrices
    Call generarMatriz(Tilesets(idtileset).matriz_transformacion, loopMatriz, MatricesTranformacion(slotMatrizDatos).datos(loopMatriz), ts_textura, ts_textura_comienzo, ts_camino, ts_camino_p2)
Next loopMatriz




End Sub

Public Sub actualizarEnIni(idtileset As Integer)
    Dim num As Byte
    Dim loopPaso As Integer
    Dim pasos As String
    
    With Tilesets(idtileset)
    
        WriteVar DBPath & archivo, HEAD_ELEMENTO & idtileset, "Graficos", .stage_count
        WriteVar DBPath & archivo, HEAD_ELEMENTO & idtileset, "Animacion", .anim
        WriteVar DBPath & archivo, HEAD_ELEMENTO & idtileset, "Nombre", .nombre
        WriteVar DBPath & archivo, HEAD_ELEMENTO & idtileset, "Olitas", .Olitas
    
        For num = 1 To .stage_count
            WriteVar DBPath & archivo, HEAD_ELEMENTO & idtileset, "Grh" & num, .stages(num)
        Next
    
        ' Nuevo formato de pisos
        WriteVar DBPath & archivo, HEAD_ELEMENTO & idtileset, "formato", .formato
        WriteVar DBPath & archivo, HEAD_ELEMENTO & idtileset, "referencia", .referencia.textura & " " & .referencia.numero

        ' Pisadas
        pasos = ""
        For loopPaso = 0 To 50
           pasos = pasos & ITS(.EfectoPisada(loopPaso))
        Next
        
        WriteVar DBPath & archivo, HEAD_ELEMENTO & idtileset, "pasos", pasos
    End With
   
    #If Colaborativo = 1 Then
        If existe(idtileset) Then
            Call versionador.modificado(CDM_IDENTIFICADOR, idtileset, Tilesets(idtileset).nombre)
        End If
    #End If
End Sub
Public Function compilar() As Boolean
    Dim handle  As Integer
    Dim tileset As Integer
    Dim num     As Long
        
    handle = FreeFile()
    Open IniPath & archivo_compilado For Binary Access Write As handle
        
    Put handle, , CInt(Tilesets_count)
    
    For tileset = 1 To Tilesets_count
        With Tilesets(tileset)
            If .stage_count > 0 Then
                Put handle, , tileset
                
                Put handle, , .stage_count
                Put handle, , .anim
                                
                Put handle, , .Olitas
                
                For num = 1 To .stage_count
                    Put handle, , .stages(num)
                Next
            End If
        End With
    Next tileset
    
    Close handle

    compilar = True
    
End Function



Public Function obtenerIDTileSet(nombreTileSet As String) As Integer
    Dim tileset As Integer
    
    For tileset = 1 To Tilesets_count
        If Tilesets(tileset).nombre = nombreTileSet Then
           obtenerIDTileSet = tileset
           Exit Function
        End If
    Next tileset

    obtenerIDTileSet = 0
End Function
