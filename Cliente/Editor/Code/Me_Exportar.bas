Attribute VB_Name = "Me_Exportar"
Option Explicit

' Capturar pantalla

' Variables de configuracion
Private tilesAncho As Byte ' Cantidad de tiles de ancho que muestra al render
Private tilesAlto As Byte ' Cantidad de tiles de alto que muestra el render
Private escala As Single ' Escala en al cual se va a guardar la imagen
Private guardarComoPNG As Boolean ' ¿La imagen final va como PNG?
Private capturarPantalla_Prefijo As String '¿Con que nombre identificamos a los archivos?

' Variables generadas a partir de las de configuracion
Private fraccionesCantidad As Position ' Cantidad de fracciones

' Backup
Private capturarPantallaPosicionInicial As Position '¿En que X e Y estaba la camara antes de arrancar?

' Variables del proceso
Private bCapturarPantalla As Boolean '¿Estamos capturando la pantalla?
Private fraccionesActual As Position '¿En que fracción estoy actualmente?
Private ultimaImagenGenerada_ As String

Public Function ultimaImagenGenerada() As String
    ultimaImagenGenerada = ultimaImagenGenerada_
End Function

Public Function progreso() As Single
    progreso = ((fraccionesActual.X * fraccionesActual.Y) / (fraccionesCantidad.X * fraccionesCantidad.Y)) * 100
End Function

Public Function capturandoPantalla() As Boolean
    capturandoPantalla = bCapturarPantalla
End Function

Public Sub capturarPantalla(tilesAncho_ As Integer, tilesAlto_ As Integer, escala_ As Single, prefijoNombre As String, png As Boolean)
    ' Flag que indica que estamos trabajando
    bCapturarPantalla = True
    guardarComoPNG = png
    escala = escala_
    tilesAncho = tilesAncho_
    tilesAlto = tilesAlto_
    
    ' Cuando se termina de capturar la pantalla, el personaje vuelve a su posicion inicial
    capturarPantallaPosicionInicial.X = UserPos.X
    capturarPantallaPosicionInicial.Y = UserPos.Y
    
    ' Posición inicial donde comienza la camara (el personaje)
    UserPos.X = (tilesAncho \ 2) + 1
    UserPos.Y = (tilesAlto \ 2) + 1
    
    ' ¿Cuantas fracciones en X e Y voy a tener?
    fraccionesCantidad.X = Round(SV_Constantes.ANCHO_MAPA / tilesAncho + 0.5, 0)
    fraccionesCantidad.Y = Round(SV_Constantes.ALTO_MAPA / tilesAlto + 0.5, 0)
    
    ' Fraccion en X e Y inicial
    fraccionesActual.X = 1
    fraccionesActual.Y = 1
    
    ' Como vamos a guardar los distintos archivos que se van generando.
    capturarPantalla_Prefijo = prefijoNombre
End Sub


' Esta funcion recibe la pantalla, le saca una foto y la guarda
Public Sub generarFraccionPantalla(D3DDevice As Direct3DDevice8, D3DX As D3DX8)

    ' Tomamos la pantalla
    Call ME_Render.capturarPantalla(D3DDevice, D3DX, 1, 1, capturarPantalla_Prefijo & "_" & fraccionesActual.X & "_" & fraccionesActual.Y & ".temp")
        
    ' Avanzo a la siguiente fraccion en X e Y
    If fraccionesActual.X = fraccionesCantidad.X Then
        fraccionesActual.X = 1
        fraccionesActual.Y = fraccionesActual.Y + 1
    Else
        fraccionesActual.X = fraccionesActual.X + 1
    End If
        
    If fraccionesActual.Y > fraccionesCantidad.Y Then
        bCapturarPantalla = False
    End If
    
    ' ¿Termine de capturar?
    If Not bCapturarPantalla Then
        ' Vuelvo a la posición inicial del personaje
        UserPos.X = capturarPantallaPosicionInicial.X
        UserPos.Y = capturarPantallaPosicionInicial.Y
        
        ' Generamos la imagen final
        Call generarImagenFinalMapa
    Else
        ' Avanzo la camara al centro de la siguiente fraccion
        ' El mínimo entre la siguiente posición y el borde del mapa
        UserPos.X = mini(((tilesAncho \ 2) + 1) + (fraccionesActual.X - 1) * tilesAncho, SV_Constantes.ANCHO_MAPA - ((tilesAncho \ 2) + 1))
        UserPos.Y = mini(((tilesAlto \ 2) + 1) + (fraccionesActual.Y - 1) * tilesAlto, SV_Constantes.ALTO_MAPA - ((tilesAlto \ 2) + 1))
    End If


    ' TODO Marce No sé si esto es correcto, pero sino el suelo no se ve al moverme
    Call rm2a
    Cachear_Tiles = True
End Sub

' Esta funcion se ejecuta cuando se tiene todas las partes que conforman un mapa
' Retorna la ruta de la imagen generada
Private Function generarImagenFinalMapa() As String
    Dim X As Integer
    Dim Y As Integer
    Dim ancho As Long
    Dim alto As Long
    Dim archivoImagen As String
    Dim freeimage1 As Long
    Dim dir_name  As String
    
    Dim puntoX As Single
    Dim puntoY As Single
    
    Dim altoFraccion As Single
    Dim anchoFraccion As Single
    
    Dim escalaMod As Integer
    
    ' Carpeta destino donde voy a guardar la imagen final
    dir_name = OPath & "Imagenes\"
    
    escalaMod = 32 * escala
    
    ' Ancho de la imagen
    ancho = SV_Constantes.ANCHO_MAPA * escalaMod
    alto = SV_Constantes.ALTO_MAPA * escalaMod
    
    ' Ancho de cada fraccion que genere
    altoFraccion = tilesAlto * escalaMod
    anchoFraccion = tilesAncho * escalaMod
    
    ' Redimensiono el picture que me va a permitir unir las distintas partes
    frmExportarAux.picResized.Cls
    frmExportarAux.picResized.width = ancho
    frmExportarAux.picResized.height = alto
    
    ' Pego en el Picture cada parte
    For Y = 1 To fraccionesCantidad.Y
        For X = 1 To fraccionesCantidad.X
            ' Ruta donde esta la parte
            archivoImagen = dir_name & CStr(THIS_MAPA.numero) & "_" & X & "_" & Y & ".temp"
            
            'Cargo la imagen
            frmExportarAux.picOriginal.Cls
            frmExportarAux.picOriginal.Picture = LoadPicture(archivoImagen)

            puntoX = mins(CInt((X - 1) * anchoFraccion), ancho - anchoFraccion)
            puntoY = mins(CInt((Y - 1) * altoFraccion), alto - altoFraccion)
            
            ' Pego la imagen
            frmExportarAux.picResized.PaintPicture frmExportarAux.picOriginal.Picture, puntoX, puntoY, anchoFraccion, altoFraccion, 0, 0, frmExportarAux.picOriginal.width, frmExportarAux.picOriginal.height
            
            ' La elimino
            Call Kill(archivoImagen)
        Next X
    Next Y

    ' Guardamos la imagen en BMP
    Call SavePicture(frmExportarAux.picResized.Image, dir_name & CStr(THIS_MAPA.numero) & ".bmp")
    
    ' ¿Debo pasarla a PNG?
    If guardarComoPNG Then
        ' Convertimos el BMP en un PNG asi pesa menos
        freeimage1 = FreeImage_Load(FIF_BMP, dir_name & CStr(THIS_MAPA.numero) & ".bmp", 0)
        Call FreeImage_Save(FIF_png, freeimage1, dir_name & CStr(THIS_MAPA.numero) & ".png", 0)
        
        ' Eliminamos el temporal
        Call Kill(dir_name & CStr(THIS_MAPA.numero) & ".bmp")
        
        ' Retornamos la ruta
        ultimaImagenGenerada_ = dir_name & CStr(THIS_MAPA.numero) & ".png"
    Else
        ultimaImagenGenerada_ = dir_name & CStr(THIS_MAPA.numero) & ".bmp"
    End If
    
End Function


Public Sub fusionarImagenesDeMapas(escala As Single)
    Dim loopX As Integer
    Dim loopY As Integer
    
    Dim posx As Integer
    Dim posy As Integer
    
    Dim archivoImagen As String
    
    Dim ancho_mapa_pixels As Integer 'Cuanto mide cada mapa a dibujar
    Dim alto_mapa_pixeles As Integer
    
    Dim ancho_borde_pixels As Integer
    Dim alto_borde_pixels As Integer
    
    Dim ancho_imagen_pixeles As Integer
    Dim alto_imagen_pixeles As Integer
    
    Dim escalaMod As Integer
    
    ' Genero una escala (pixeles por tile) entero
    escalaMod = (32 * escala)
    
    ' Tamaño del Mapa en pixeles
    ancho_mapa_pixels = (X_MAXIMO_USABLE - X_MINIMO_USABLE + 1) * escalaMod
    alto_mapa_pixeles = (Y_MAXIMO_USABLE - Y_MINIMO_USABLE + 1) * escalaMod
    
    ' Borde del mapa que no se utiliza
    ancho_borde_pixels = BORDE_TILES_INUTILIZABLE * escalaMod
    alto_borde_pixels = BORDE_TILES_INUTILIZABLE * escalaMod
    
    ' Inicializo el picture que uso como herrameinta
    frmExportarAux.picResized.Cls
    
    ' Establezco el tamaño total de la imagen, el tamaño de cada mapa por la cantidad más los bordes finales
    ancho_imagen_pixeles = ancho_mapa_pixels * UBound(ME_Mundo.MapasArray, 1) + ancho_borde_pixels * 2
    alto_imagen_pixeles = alto_mapa_pixeles * UBound(ME_Mundo.MapasArray, 2) + alto_borde_pixels * 2
    
    frmExportarAux.picResized.width = ancho_imagen_pixeles
    frmExportarAux.picResized.height = alto_imagen_pixeles
    
    ' Configuraciones generales. Dibujo los números
    frmExportarAux.picResized.BackColor = vbBlack
    frmExportarAux.picResized.ForeColor = vbCyan
    frmExportarAux.picResized.font.Size = 36
    frmExportarAux.picResized.font.bold = True
    
    ' Recorro mapa por mapa
    posy = 0
    For loopY = 1 To UBound(ME_Mundo.MapasArray, 2)
        posx = 0
        For loopX = 1 To UBound(ME_Mundo.MapasArray, 1)
        
            If ME_Mundo.MapasArray(loopX, loopY).numero > 0 Then
            
                ' Obtengo la imagen
                archivoImagen = OPath & "Imagenes\" & ME_Mundo.MapasArray(loopX, loopY).numero & ".bmp"
                
                If FileExist(archivoImagen) Then 'Si el archivo existe, lo dibujamos, sino no, el fondo es negro ya.
                    'Cargo la imagen original
                    frmExportarAux.picOriginal.Picture = LoadPicture(archivoImagen)
                    ' Pego la imagen
                    frmExportarAux.picResized.PaintPicture frmExportarAux.picOriginal.Picture, posx, posy, frmExportarAux.picOriginal.width, frmExportarAux.picOriginal.height, 0, 0, frmExportarAux.picOriginal.width, frmExportarAux.picOriginal.height
                    ' Liberamos memoria
                    frmExportarAux.picOriginal.Cls
                End If
                
                ' Dibujamos el numero del mapa
                frmExportarAux.picResized.CurrentX = posx + ancho_borde_pixels
                frmExportarAux.picResized.CurrentY = posy + alto_borde_pixels
                frmExportarAux.picResized.Print ME_Mundo.MapasArray(loopX, loopY).numero
            End If
            
            posx = posx + ancho_mapa_pixels
            
        Next loopX
        
        posy = posy + alto_mapa_pixeles

    Next loopY
        
    ' Dibujamos las lineas divisoras del mapa
    frmExportarAux.picResized.DrawWidth = 1
    
    ' Lineas horizontales
    posy = alto_borde_pixels
    For loopY = 1 To UBound(ME_Mundo.MapasArray, 2) - 1
        posy = posy + alto_mapa_pixeles
        frmExportarAux.picResized.Line (0, posy)-(frmExportarAux.picResized.width - 1, posy), vbWhite
    Next loopY
    
    ' Lineas verticales
    posx = ancho_borde_pixels
    For loopX = 1 To UBound(ME_Mundo.MapasArray, 1) - 1
        posx = posx + ancho_mapa_pixels
        frmExportarAux.picResized.Line (posx, 0)-(posx, frmExportarAux.picResized.height - 1), vbWhite
    Next
    
    ' Guardamos en BMP
    Call SavePicture(frmExportarAux.picResized.Image, OPath & "Imagenes\" & ME_Mundo.obtenerNombreZonaActual & ".temp")
    
    ' Convertimos el BMP en un PNG asi pesa menos
    Dim freeimage1 As Long
    freeimage1 = FreeImage_Load(FIF_BMP, OPath & "Imagenes\" & ME_Mundo.obtenerNombreZonaActual & ".temp", 0)
    Call FreeImage_Save(FIF_png, freeimage1, OPath & "Imagenes\" & ME_Mundo.obtenerNombreZonaActual & ".png", 0)
    
    ' Eliminamos el temporal
    Call Kill(OPath & "Imagenes\" & ME_Mundo.obtenerNombreZonaActual & ".temp")
        
    ' Liberamos memoria
    frmExportarAux.picResized.Cls
End Sub
