VERSION 5.00
Begin VB.Form frmExportarAux 
   Caption         =   "Formulario Auxiliar"
   ClientHeight    =   1425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   95
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   189
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOriginal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   960
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picResized 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2400
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer tmrGenerarZonaImagen 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   840
   End
   Begin VB.Timer tmrCheckGenerarMiniMapa 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   840
   End
End
Attribute VB_Name = "frmExportarAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MiniMapa
Private MiniMapaImagen_Aviso As Boolean
Private MiniMapaImagen_MapaEnZona As Position
Private MiniMapaImagen_Generando As Boolean
Private MiniMapaImagen_Escala As Single

Public Sub exportarMapa(escala As Single, png As Boolean, aviso As Boolean)

    ' Guardamos la configuración
    MiniMapaImagen_Aviso = aviso
    MiniMapaImagen_Generando = True
    
    ' Ponemos el cartel de trabajando.
    Call frmMain.ocultarRender("Exportando mapa a imágen.")
    
    ' Primero copiamos los bordes
    Call CopiarBordesMapaActual
    
    ' Cargamos el formulario auxiliar
    load frmExportarAux
    
    ' Ahora le decimos que debe capturar todas las fracciones del mapa.
    ' El prefijo con el que se guarda es el número del mapa
    Call Me_Exportar.capturarPantalla(modPantalla.TilesPantalla.X, modPantalla.TilesPantalla.Y, escala, CStr(THIS_MAPA.numero), png)
    
    ' Detectamos que termino con este timer
    frmExportarAux.tmrCheckGenerarMiniMapa.Enabled = True
End Sub

Public Sub exportarZona(escala As Single)

    ' Guardamos la escala y donde comenzamos
    MiniMapaImagen_Escala = escala
    MiniMapaImagen_MapaEnZona.X = 0
    MiniMapaImagen_MapaEnZona.Y = 1

    'Vamos a iniciar el proceso en el próximo ciclo
    frmExportarAux.tmrGenerarZonaImagen.Enabled = True
        
    ' Ocultamos el render y mostramos el mensaje que queremos
    Call frmMain.ocultarRender("Exportando Zona a imágen." & vbCrLf & "0%")
End Sub

Private Sub tmrGenerarZonaImagen_Timer()
    Dim listo As Boolean
    Dim total As Integer
    Dim hecho As Integer

    ' Voy a procesar la proxima parte del minimapa cuando no este haciendo una parte
    If MiniMapaImagen_Generando Then Exit Sub
    
    Do While Not listo
        
        ' Busco el próximo mapa al cual le debo generar la imagen
        If MiniMapaImagen_MapaEnZona.X = UBound(ME_Mundo.MapasArray, 1) Then
            MiniMapaImagen_MapaEnZona.Y = MiniMapaImagen_MapaEnZona.Y + 1
            MiniMapaImagen_MapaEnZona.X = 1
        Else
            MiniMapaImagen_MapaEnZona.X = MiniMapaImagen_MapaEnZona.X + 1
        End If

        If MiniMapaImagen_MapaEnZona.Y > UBound(ME_Mundo.MapasArray, 2) Then
            listo = True
        ElseIf ME_Mundo.MapasArray(MiniMapaImagen_MapaEnZona.X, MiniMapaImagen_MapaEnZona.Y).numero Then
            listo = True
        End If
    Loop
    

    If MiniMapaImagen_MapaEnZona.Y > UBound(ME_Mundo.MapasArray, 2) Then
        ' Termine!.
        Me.tmrGenerarZonaImagen.Enabled = False
        
        ' Mensaje
        frmMain.lblTapaSolapas = "Exportando Zona a imágen." & vbCrLf & "Finalizando..."
        frmMain.Refresh
        
        ' Juntamos las imagenes
        Call fusionarImagenesDeMapas(MiniMapaImagen_Escala)
        
        ' Mostramos nuevamente la pantalla
        Call frmMain.mostrarRender
        
        ' Abrimos la carpeta donde esta el archivo
        Call Shell("explorer /select," & OPath & "Imagenes\" & ME_Mundo.obtenerNombreZonaActual & ".png", vbMaximizedFocus)
    Else
        ' Abro el mapa
        If frmMain.ABRIR_Mapa(ME_Mundo.MapasArray(MiniMapaImagen_MapaEnZona.X, MiniMapaImagen_MapaEnZona.Y).numero, False) Then
            ' Lo exporto
            Call exportarMapa(MiniMapaImagen_Escala, False, False)
        End If

        ' Actualizamos
        total = UBound(ME_Mundo.MapasArray, 1) * UBound(ME_Mundo.MapasArray, 2) 'alto * ancho
        hecho = MiniMapaImagen_MapaEnZona.X + (UBound(ME_Mundo.MapasArray, 1) * (MiniMapaImagen_MapaEnZona.Y - 1))
        frmMain.lblTapaSolapas = "Exportando Zona a imágen." & vbCrLf & Round(hecho / total * 100, 2) & "%"
    End If
    
End Sub

Private Sub tmrCheckGenerarMiniMapa_Timer()
    
    Dim archivoImagen As String ' Archivo de imagen generado
    
    If Not Me_Exportar.capturandoPantalla Then
    
        Me.tmrCheckGenerarMiniMapa.Enabled = False
        
        archivoImagen = Me_Exportar.ultimaImagenGenerada()
        
        ' ¿Tengo que avisar?
        If MiniMapaImagen_Aviso Then
            'Abrimos la carpeta donde esta el archivo
            Call Shell("explorer /select," & archivoImagen, vbMaximizedFocus)
        
            'Mostramos el Render habitual
            Call frmMain.mostrarRender
        End If
        
        ' Terminamos
        MiniMapaImagen_Generando = False
    
    Else
        ' Mostramos el progreso
        Call frmMain.ocultarRender("Exportando mapa a imágen." & vbCrLf & FormatNumber(Me_Exportar.progreso, 2))
    End If
    
End Sub
