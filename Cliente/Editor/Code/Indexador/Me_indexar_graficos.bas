Attribute VB_Name = "Me_indexar_Graficos"
Option Explicit

Public Const SEPARADOR_PROPIEDADES = "-"

Private Const archivo = "Graficos.ini"
Private Const archivo_compilado = "Graficos.ind"
Private Const HEAD_ELEMENTO = ""
Private Const CDM_IDENTIFICADOR = "GRAFICO"


Public Function existe(ByVal ID As Integer) As Boolean
    
    Dim direccion As Byte
    
    If ID > UBound(GrhData) Then
        existe = False
        Exit Function
    End If
    
    If GrhData(ID).NumFrames = 0 Then
        existe = False
    Else
        existe = True
    End If

End Function

Public Sub establecerConfigBasica(idGrafico As Integer, nombre As String, ancho As Integer, alto As Integer, idImagen As Integer, offsetX As Integer, offsetY As Integer, Optional ByVal IDUnico As String = "")

Dim offsetParaAjuste As Position

' Seteamos
With GrhData(idGrafico)

    .ID = IDUnico
    .nombreGrafico = nombre
    .filenum = idImagen

    .NumFrames = 1
    ReDim .frames(1)
    .frames(1) = idGrafico
    .Speed = 0

    .sx = offsetX
    .sy = offsetY

    .pixelHeight = alto
    .pixelWidth = ancho
 
    'Nuevas propiedades
    .EfectoPisada = 0
    .perteneceAunaAnimacion = False
    .esInsertableEnMapa = True
    
    Dim i As Integer
    For i = 1 To CANTIDAD_CAPAS
        .Capa(i) = True
    Next i
    
End With


Call Me_indexar_Graficos.obtenerOffsetAjustadoTile(GrhData(idGrafico), offsetParaAjuste.x, offsetParaAjuste.y)

GrhData(idGrafico).offsetX = offsetParaAjuste.x
GrhData(idGrafico).offsetY = offsetParaAjuste.y

End Sub
' Busca un grafico por su nombre
Public Function existeNombre(ByVal nombre_ As String) As Boolean
    
    existeNombre = (obtenerIDNombre(nombre_) > 0)

End Function

' Busca un grafico por su nombre
Public Function existeIDUnico(ByVal id_ As String) As Boolean
    
    existeIDUnico = (obtenerIDPorIDUnico(id_) > 0)

End Function

' Busca un grafico por su nombre
Public Function obtenerIDPorIDUnico(ByVal id_ As String) As Integer
    
Dim loopGrh As Integer


For loopGrh = 1 To grhCount

    If existe(loopGrh) Then
        If GrhData(loopGrh).ID = id_ Then
            obtenerIDPorIDUnico = loopGrh
            Exit Function
        End If
    End If

Next

obtenerIDPorIDUnico = 0

End Function

' Busca un grafico por su nombre
Public Function obtenerIDNombre(ByVal nombre_ As String) As Integer
    
Dim loopGrh As Integer
Dim nombre As String

nombre = UCase$(nombre_)

For loopGrh = 1 To grhCount

    If existe(loopGrh) Then
    
        If UCase$(GrhData(loopGrh).nombreGrafico) = nombre Then
            obtenerIDNombre = loopGrh
            Exit Function
        End If
    End If

Next

obtenerIDNombre = 0

End Function

Private Function numeroConSignoAString(numero As Integer) As String
     If numero <= 0 Then
        numeroConSignoAString = Abs(numero)
    Else
        numeroConSignoAString = "+" & numero
    End If
End Function

Private Function StringANumeroConSigno(numero As String) As Integer

    If left$(numero, 1) = "+" Then
        StringANumeroConSigno = val(numero)
    Else
        StringANumeroConSigno = val(numero) * -1
    End If
End Function


Public Sub actualizarEnIni(ByVal numeroDeGrh As Integer)
    Dim datos As String
    Dim n As Byte
    Dim capasAplica As Byte
    Dim loopCapa As Byte
    Dim offsetNeto As Position
    
    With GrhData(numeroDeGrh)

        If .NumFrames > 0 Then
            If .NumFrames = 1 Then
                datos = "1-" & CStr(.filenum) & SEPARADOR_PROPIEDADES & CStr(.sx) & SEPARADOR_PROPIEDADES & CStr(.sy) & SEPARADOR_PROPIEDADES & CStr(.pixelWidth) & SEPARADOR_PROPIEDADES & CStr(.pixelHeight)
            ElseIf .NumFrames > 1 Then
                datos = CStr(.NumFrames)
                
                For n = 1 To .NumFrames
                  datos$ = datos & SEPARADOR_PROPIEDADES & CStr(.frames(n))
                Next
                datos = datos & SEPARADOR_PROPIEDADES & CStr(.Speed)
            End If
        
            datos = datos & SEPARADOR_PROPIEDADES & .nombreGrafico
                    
            If .esInsertableEnMapa Then
                datos = datos & "-1"
            Else
                datos = datos & "-0"
            End If
        
            If .perteneceAunaAnimacion Then
                datos = datos & "-1"
            Else
                datos = datos & "-0"
            End If
            
            If .esInsertableEnMapa Then
                capasAplica = 0
                
                For loopCapa = 1 To CANTIDAD_CAPAS
                    If .Capa(loopCapa) Then Call BS_Byte_On(capasAplica, loopCapa)
                Next loopCapa
                
                datos = datos & "-" & capasAplica
            Else
                datos = datos & "-" & capasAplica
            End If
                        
            ' Dummy, ex centrado en 32
            datos = datos & "-0"
  
            
            ' Efecto al escuchar la pisada sobre el piso
            datos = datos & "-" & .EfectoPisada
            
            Call calcularOffsetNeto(GrhData(numeroDeGrh), offsetNeto.x, offsetNeto.y)
            
            ' Ubicación del gráfico dentro de la Grilla del Juego
            datos = datos & "-" & numeroConSignoAString(offsetNeto.x) & "-" & numeroConSignoAString(offsetNeto.y)
            
            ' Sombra. Tamaño y offset
            datos = datos & "-" & .SombrasSize & "-" & numeroConSignoAString(.SombraOffsetX) & "-" & numeroConSignoAString(.SombraOffsetY)

            ' Identificador unico
            datos = datos & "-" & .ID
        Else
            datos = "0-0-0-0-0-0--0-0-0"
        End If
        
        'Escribo
        WriteVar DBPath & archivo, HEAD_ELEMENTO & CStr(numeroDeGrh), "Datos", datos
    End With
    
    #If Colaborativo = 1 Then
        If existe(numeroDeGrh) Then
            Call versionador.modificado(CDM_IDENTIFICADOR, numeroDeGrh, GrhData(numeroDeGrh).nombreGrafico)
        End If
    #End If
End Sub

Public Sub CargarListaGraficosComunes()

    Dim i As Integer
    
    frmMain.ListaConBuscadorGraficos.vaciar
    For i = 1 To grhCount
        If GrhData(i).NumFrames <> 0 Then
            If GrhData(i).esInsertableEnMapa Then
                Call frmMain.ListaConBuscadorGraficos.addString(CInt(i), CStr(i & " - " & GrhData(i).nombreGrafico))
            End If
        End If
    Next i

End Sub


Public Function CargarGraficosIni() As Boolean

    Dim Soport  As New cIniManager
    Dim Grh     As Long
    Dim datos As String
    Dim loopCapa As Byte
    
    If LenB(Dir(DBPath & "Graficos.ini", vbArchive)) = 0 Then
        MsgBox "No existe Graficos.ini en la carpeta " & DBPath
        Exit Function
    End If

    Soport.Initialize DBPath & archivo
    
    grhCount = CInt(val(Soport.getNameLastSection))
    
    ReDim Preserve GrhData(0 To grhCount)
    
    For Grh = 1 To grhCount
    
        datos = Soport.getValue(HEAD_ELEMENTO & Grh, "Datos")
                
        If Len(datos$) Then
            indexar_from_string Grh, datos
        Else
            GrhData(Grh).filenum = 0
            GrhData(Grh).NumFrames = 0
            GrhData(Grh).sx = 0
            GrhData(Grh).sy = 0
            GrhData(Grh).pixelWidth = 0
            GrhData(Grh).pixelHeight = 0
            
            GrhData(Grh).perteneceAunaAnimacion = False
            GrhData(Grh).esInsertableEnMapa = False
            GrhData(Grh).nombreGrafico = ""
            
            ' Centrado
            GrhData(Grh).offsetX = 0
            GrhData(Grh).offsetY = 0
            
            ' Sombra
            GrhData(Grh).SombrasSize = 0
            GrhData(Grh).SombraOffsetX = 0
            GrhData(Grh).SombraOffsetY = 0
            
            For loopCapa = 1 To CANTIDAD_CAPAS
                GrhData(Grh).Capa(loopCapa) = False
            Next loopCapa
            
        End If
    Next
    
    CargarGraficosIni = True
    
    Exit Function
    
ErrorHandler:
MsgBox "Error cargando Graficos.ini en grafico numero: " & Grh
End Function


Public Sub resetear(Grh As GrhData)
    Grh.filenum = 0
    Grh.NumFrames = 0
    Grh.sx = 0
    Grh.sy = 0
    Grh.pixelWidth = 0
    Grh.pixelHeight = 0
        
    Grh.esInsertableEnMapa = False
    Grh.nombreGrafico = ""
    Grh.perteneceAunaAnimacion = False
End Sub
Private Sub indexar_from_string(ByVal Grh As Long, ByRef datos As String)
    Dim frames  As Long
    Dim loopCapa As Byte
    Dim CapaAplica As Byte
    Dim baseCampos As Byte
    Dim cantidadCampos As Byte
    Dim offsetNeto As Position
    
    With GrhData(Grh)
    
    .NumFrames = ReadField(1, datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES))
    
    If .NumFrames > 0 Then
    
        ' Cantidad de campos de informacion
        cantidadCampos = FieldCount(datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES))
        
        ReDim .frames(1 To .NumFrames)
        
        If .NumFrames = 1 Then
        
            .filenum = val(ReadField(2, datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES)))
            .sx = val(ReadField(3, datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES)))
            .sy = val(ReadField(4, datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES)))
            .pixelWidth = val(ReadField(5, datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES)))
            .pixelHeight = val(ReadField(6, datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES)))
                                
            .frames(1) = Grh
            
            ' A partir de donde comienza a leerlos campos generales
            baseCampos = 7
        ElseIf .NumFrames > 1 Then
        
            For frames = 1 To .NumFrames
                .frames(frames) = val(ReadField(frames + 1, datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES)))
                If .frames(frames) <= 0 Or .frames(frames) > grhCount Then
                    GoTo ErrorHandler
                End If
            Next
            
            .Speed = CCVal(ReadField(.NumFrames + 2, datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES)))
            If .Speed <= 0 Then GoTo ErrorHandler
            
            baseCampos = .NumFrames + 3
            
            'Compute width and height
            .pixelHeight = GrhData(.frames(1)).pixelHeight
            If .pixelHeight <= 0 Then
                GoTo ErrorHandler
            End If

            .pixelWidth = GrhData(.frames(1)).pixelWidth
            If .pixelWidth <= 0 Then
                GoTo ErrorHandler
            End If

        End If
        
        
        ' GENERALES PARA ANIMACIONES Y GRAFICOS
        .nombreGrafico = ReadField(baseCampos, datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES))
        .esInsertableEnMapa = (val(ReadField(baseCampos + 1, datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES))) = "1")
        .perteneceAunaAnimacion = (val(ReadField(baseCampos + 2, datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES))) = "1")
            
        CapaAplica = val(ReadField(baseCampos + 3, datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES)))
        
        ' Capas donde se puede utilizar
        For loopCapa = 1 To CANTIDAD_CAPAS
            .Capa(loopCapa) = HelperBitWise.BS_Byte_Get(CapaAplica, loopCapa)
        Next
            
        ' Dummy 4
        '.centrarEn32 = ReadField(baseCampos + 4, datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES)) = "1"
        

        ' Efecto de la pisada
        If cantidadCampos >= baseCampos + 5 Then
            .EfectoPisada = val(ReadField(baseCampos + 5, datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES)))
        Else
            .EfectoPisada = 0
        End If
            
        ' Ubicación del gráfico dentro de la Grilla del Juego
        If cantidadCampos >= baseCampos + 7 Then
            offsetNeto.x = StringANumeroConSigno(ReadField(baseCampos + 6, datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES)))
            offsetNeto.y = StringANumeroConSigno(ReadField(baseCampos + 7, datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES)))
        Else
            offsetNeto.x = 0
            offsetNeto.y = 0
        End If
        
            
        ' Sombras
        If cantidadCampos >= baseCampos + 10 Then
            .SombrasSize = maxi(mini(CCVal(ReadField(baseCampos + 8, datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES))), 255), 0)
            .SombraOffsetX = StringANumeroConSigno(ReadField(baseCampos + 9, datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES)))
            .SombraOffsetY = StringANumeroConSigno(ReadField(baseCampos + 10, datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES)))
        Else
            .SombrasSize = 0
            .SombraOffsetX = 0
            .SombraOffsetY = 0
        End If
        
        ' ID
        If cantidadCampos >= baseCampos + 11 Then
           .ID = ReadField(baseCampos + 11, datos, Asc(Me_indexar_Graficos.SEPARADOR_PROPIEDADES))
        Else
            .ID = ""
        End If
        
            
    End If
    

    End With

    Call establecerOffsetBruto(GrhData(Grh), offsetNeto.x, offsetNeto.y)
        
Exit Sub
ErrorHandler:
MsgBox "Error indexar_from_string en grafico numero: " & Grh
End Sub

' Por defecto el juego calcula el offset para que el gráfico quede en el centro de un tile.
Public Sub obtenerOffsetNatural(Grh As GrhData, ByRef offsetX As Integer, ByRef offsetY As Integer)
    
    offsetX = 16 - Grh.pixelWidth / 2 ' Lo centro
    offsetY = 32 - Grh.pixelHeight

End Sub

' Obtiene el offset que debe tener un gráfico para que este se ajuste al borde de un tile de manera horizontal
Public Sub obtenerOffsetAjustadoTile(Grh As GrhData, ByRef offsetX As Integer, ByRef offsetY As Integer)
    
    If Grh.pixelWidth = 32 Then
        offsetX = 0
    Else
        '¿Es impar?. El objetivo es que siempre quede ajustado a la grilla
        If Grh.pixelWidth \ 32 Mod 2 > 0 Then
            offsetX = -((Grh.pixelWidth - 32) \ 2)
        Else
            offsetX = -(Grh.pixelWidth \ 2)
        End If
    End If
    
    offsetY = 32 - Grh.pixelHeight

End Sub


' Devuelve en offsetX e offset y los valores netos del offset del gráfico que en su estructura contiene el valor bruto.
Public Sub calcularOffsetNeto(ByRef Grh As GrhData, ByRef offsetX As Integer, ByRef offsetY As Integer)
    Dim offsetNatural As Position
    
    Call Me_indexar_Graficos.obtenerOffsetNatural(Grh, offsetNatural.x, offsetNatural.y)
    
    offsetX = Grh.offsetX - offsetNatural.x
    offsetY = Grh.offsetY - offsetNatural.y
End Sub

' Devuelve en offsetX e offset y los valores brutos del offset del gráfico a partir del offset neto
Public Sub establecerOffsetBruto(Grh As GrhData, ByVal offsetXNeto As Integer, ByVal offsetYNeto As Integer)
    Dim offsetNatural As Position
    
    Call Me_indexar_Graficos.obtenerOffsetNatural(Grh, offsetNatural.x, offsetNatural.y)
    
    Grh.offsetX = offsetXNeto + offsetNatural.x
    Grh.offsetY = offsetYNeto + offsetNatural.y
End Sub


Public Function compilar() As Boolean
    
    Dim handle  As Integer
    Dim Grh     As Long
    Dim Frame   As Long
    Dim mayorSlot As Integer
   
    handle = FreeFile()
    
    ' Obtengo el mayor slot. En vez de guardar la cantidad total de SLot, solo guardo hasta el numero que use
    mayorSlot = 0
    
    For Grh = grhCount To 1 Step -1
    
        If existe(Grh) Then
            mayorSlot = Grh
            Exit For
        End If
    
    Next Grh
    
    Open IniPath & archivo_compilado For Binary Access Write As handle
    
    ' Guardamos la cantidad
    Put #handle, , mayorSlot
    
    For Grh = 1 To mayorSlot
            
        With GrhData(Grh)
            
            If existe(Grh) Then
                
                ' Ponemos la cantidad de numeros de frames
                Put handle, , .NumFrames
        
                If .NumFrames = 1 Then
                    Put handle, , .filenum
                    Put handle, , .sx
                    Put handle, , .sy
                    Put handle, , .pixelWidth
                    Put handle, , .pixelHeight
                    Put handle, , .offsetX
                    Put handle, , .offsetY
                    Put handle, , .SombrasSize
                ElseIf GrhData(Grh).NumFrames > 1 Then
                    For Frame = 1 To GrhData(Grh).NumFrames
                        Put handle, , GrhData(Grh).frames(Frame)
                    Next
                    Put handle, , GrhData(Grh).Speed
                End If

            Else
                Put handle, , CInt(0)
            End If
            
        End With
    Next Grh

    'Cerramos el archivo
    Close #handle
    
    compilar = True
        
End Function

Public Function eliminar(ByVal ID As Integer)
    Dim nombreBackup As String
    
    nombreBackup = GrhData(ID).nombreGrafico
    
    Call resetear(GrhData(ID))
    
    Call actualizarEnIni(ID)
    
    If ID = UBound(GrhData) Then
        ReDim Preserve GrhData(0 To UBound(GrhData) - 1) As GrhData
        grhCount = UBound(GrhData)
    End If
    
    #If Colaborativo = 1 Then
        Call versionador.eliminado(CDM_IDENTIFICADOR, ID, nombreBackup)
    #End If
    
End Function

Public Function nuevo() As Long

    #If Colaborativo = 0 Then
    
    #Else
        nuevo = CDM.cerebro.SolicitarRecurso(CDM_IDENTIFICADOR)
        
        ' Me entra en la memoria?
        If nuevo > UBound(GrhData) Then
            ReDim Preserve GrhData(0 To nuevo) As GrhData
            grhCount = nuevo
        End If
        
        Call versionador.creado(CDM_IDENTIFICADOR, nuevo)
    #End If
    
End Function
