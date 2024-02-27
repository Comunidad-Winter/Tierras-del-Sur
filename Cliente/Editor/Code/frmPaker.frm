VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfigurarRecursos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar/Cambiar archivos de Recursos"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   Icon            =   "frmPaker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   695
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBorrarMultiple 
      Caption         =   "Borrar Multiple"
      Height          =   360
      Left            =   6840
      TabIndex        =   26
      Top             =   7560
      Width           =   1530
   End
   Begin EditorTDS.UpDownText udBorrarMaximo 
      Height          =   375
      Left            =   5160
      TabIndex        =   25
      Top             =   7560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      MaxValue        =   25000
      MinValue        =   0
      Enabled         =   -1  'True
   End
   Begin EditorTDS.UpDownText udBorrarMinimo 
      Height          =   375
      Left            =   3720
      TabIndex        =   24
      Top             =   7560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      MaxValue        =   25000
      MinValue        =   0
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generar Deletes"
      Height          =   360
      Left            =   8520
      TabIndex        =   22
      Top             =   6960
      Width           =   1650
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar"
      Height          =   360
      Left            =   4920
      TabIndex        =   21
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdExtraerTodos 
      Caption         =   "Extraer todos"
      Height          =   360
      Left            =   6720
      TabIndex        =   20
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Height          =   360
      Left            =   3750
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPaker.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Actualizar lista de archivos"
      Top             =   675
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   4200
      Top             =   6150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enpaquetado"
      Height          =   6255
      Left            =   4920
      TabIndex        =   0
      Top             =   480
      Width           =   5415
      Begin VB.CheckBox chkConNumero 
         Appearance      =   0  'Flat
         Caption         =   "Con número en el nombre"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1440
         TabIndex        =   23
         Top             =   5760
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdVersiones 
         Caption         =   "Versiones"
         Height          =   360
         Left            =   120
         TabIndex        =   17
         Top             =   5760
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Mapas(juego)"
         Height          =   315
         Index           =   4
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Sonidos"
         Height          =   315
         Index           =   3
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdExtraerRecursos 
         Caption         =   "Extraer Seleccionado"
         Height          =   375
         Left            =   3120
         TabIndex        =   13
         Top             =   5760
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Mapas(ME)"
         Height          =   315
         Index           =   0
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Interface"
         Height          =   315
         Index           =   1
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   5100
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   5175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Imagenes"
         Height          =   315
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Archivos en disco"
      Height          =   6255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   4695
      Begin VB.CommandButton cmdAyuda 
         Caption         =   "?"
         Height          =   360
         Left            =   3360
         TabIndex        =   19
         Top             =   1000
         Width           =   270
      End
      Begin VB.CommandButton cmdCrearEmpaquetado 
         Caption         =   "Crear nuevo pak con los archivos seleccionados"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   5760
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.CommandButton cmdParchear 
         Caption         =   "->"
         Enabled         =   0   'False
         Height          =   4695
         Left            =   4080
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton examinar_in 
         Height          =   555
         Left            =   3990
         Picture         =   "frmPaker.frx":200C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Seleccionar carpeta"
         Top             =   240
         Width           =   615
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         Height          =   4710
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   9
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txtCarpetaEntrada 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "C:\..\...\"
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox txtFiltroExtensiones 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Text            =   "bmp; png"
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de archivo / filtro:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(nuevo en paquetado)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   10215
   End
End
Attribute VB_Name = "frmConfigurarRecursos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tTipoImagen
    numeroComplemento As Byte '0: origina. (1, 2, 3, 4) : complemento N (Blend One, Color Add, Normal Map, Agus Secret)
    nombreOriginal As String ' El nombre del archivo sin la referencia a los complementos
    nombreCompleto As String ' El nombre completo del archivo
    idImagen As Integer ' El id con el que se agrego este grafico
    archivo As String
End Type

' Empaquetado actual
Dim nombreVersionado As String
Dim Pak As clsEnpaquetado

Private Type tComplemento
    abreviatura As String
    nombre As String
    numero As Byte
End Type

Private Complementos() As tComplemento


Private Sub cmdAyuda_Click()
    Dim mensajeAyuda As String
    
    ' Ayuda para agregar uno nuevo
    mensajeAyuda = "Al agregar una nueva imagen, el nombre debe seguir el siguiente estandar: " & vbCrLf & _
                    "La imagen común: nombre_descriptivo.extension" & vbCrLf & _
                    "Complemento 1 (Blend One): nombre_descriptivo_bo.extension" & vbCrLf & _
                    "Complemento 2 (Color Add): nombre_decriptivo_ca.extension" & vbCrLf & _
                    "Complemento 3 (Normal Map): nombre_descriptivo_nm.extension" & vbCrLf & _
                    "El sistema automaticamente le asignará a cada archivo agregado un identificador númerico unico."

    ' Para remaplzar
    mensajeAyuda = mensajeAyuda & vbCrLf & vbCrLf
    
    mensajeAyuda = mensajeAyuda & "Para remplazar un recurso existente, el archivo nuevo debe tener el mismo número o nombre que el que se va a remplazar. "

    MsgBox mensajeAyuda, vbInformation
End Sub

Private Sub cmdBorrar_Click()
    Dim idElemento As Integer
    Dim INFOHEADER As INFOHEADER
    Dim Data(0) As Byte
    Dim resultado As VbMsgBoxResult
    
    ' TODO DESHACORDEAR QUE SEA SOLO PARA GRAFICOS
    
    ' ¿Seleccionó algo?
    If List1.listIndex = -1 Then
        Call MsgBox("Tenés que seleccionar el recurso que queres eliminar.", vbExclamation, Me.caption)
        Exit Sub
    End If
    
    ' Obtengo el identificador
    idElemento = val(Replace(Split(List1.list(List1.listIndex), " - ")(0), "*", ""))
    
    If idElemento = 0 Then
        Call MsgBox("Tenés que seleccionar el recurso que queres eliminar.", vbExclamation, Me.caption)
        Exit Sub
    End If

    ' Obtengo el header
    Call pakGraficos.IH_Get(idElemento, INFOHEADER)
                
    resultado = MsgBox("¿Estás seguro de que queres eliminar el recurso '" & pakGraficos.Cabezal_GetFilenameName(idElemento) & "'", vbQuestion + vbYesNo, Me.caption)
    
    If resultado = vbYes Then
    
        ' Lo elimino del CDM.
        Call versionador.eliminado("RECURSO_IMAGEN", idElemento, pakGraficos.Cabezal_GetFilenameName(idElemento))
        
        ' Le nuleo el nombre
        INFOHEADER.originalname = Xor_String("--------------------------------", INFOHEADER.cript)
        
        ' Parcheo con array 0 para generar este version eliminada
        Call pakGraficos.ParchearByteArray(Data, idElemento, INFOHEADER)
        
        ' Actualizo
        List1.list(List1.listIndex) = "--------------------------------"
        
        ' Mensaje de alerta
        Call MsgBox("Recurso eliminado correctamente.", vbInformation, Me.caption)
    End If
End Sub

Private Sub cmdBorrarMultiple_Click()
Dim loopRecurso As Integer
Dim INFOHEADER As INFOHEADER
Dim Data(0) As Byte
    
For loopRecurso = Me.udBorrarMinimo.value To Me.udBorrarMaximo.value

    ' Lo elimino del CDM.
    Call versionador.eliminado("RECURSO_IMAGEN", loopRecurso, pakGraficos.Cabezal_GetFilenameName(loopRecurso))
    
    ' Obtengo el header
    Call pakGraficos.IH_Get(loopRecurso, INFOHEADER)
    
    ' Le nuleo el nombre
    INFOHEADER.originalname = Xor_String("--------------------------------", INFOHEADER.cript)
        
    ' Parcheo con array 0 para generar este version eliminada
    Call pakGraficos.ParchearByteArray(Data, loopRecurso, INFOHEADER)
        
    ' Actualizo
    List1.list(List1.listIndex) = "--------------------------------"
Next

' Mensaje de alerta
Call MsgBox("Recursos eliminados correctamente.", vbInformation, Me.caption)
End Sub

Private Sub cmdExtraerTodos_Click()

Dim loopElemento As Integer
Dim idElemento As Integer
Dim destino As String
Dim infoArchivo() As String

For loopElemento = 0 To Me.List1.ListCount - 1

    If InStr(1, List1.list(loopElemento), "--------------------------------") = 0 Then
    
        idElemento = val(Replace(Split(List1.list(loopElemento), " - ")(0), "*", ""))
        infoArchivo = Split(Pak.Cabezal_GetFilenameName(idElemento), ".", 2)
        idElemento = val(infoArchivo(0))
        
        destino = Me.txtCarpetaEntrada & Pak.Cabezal_GetFilenameName(idElemento)
        
        Call Pak.Extraer(idElemento, destino)
    End If
Next

Call MsgBox("Se extrayeron todos los recursos de este tipo.", vbInformation, Me.caption)

End Sub

Private Sub cmdVersiones_Click()

' ¿Selecciono algo?
If List1.listIndex = -1 Then
    Call MsgBox("Tenés que seleccionar el recurso que queres extraer.", vbExclamation, Me.caption)
    Exit Sub
End If

' Cargamos el formulario
load frmPakRollback

Set frmPakRollback.Pak = Pak
frmPakRollback.NumeroArchivoSeleccionado = val(Replace(Split(List1.list(List1.listIndex), " - ")(0), "*", ""))
Pak.Add_To_Listbox_Versiones frmPakRollback.lstVersiones, frmPakRollback.NumeroArchivoSeleccionado

frmPakRollback.Show
End Sub

Private Sub Command1_Click()
    Me.File1.Refresh
End Sub

Private Sub cmdCrearEmpaquetado_Click()

Dim archivoDestino As String
Dim archivos() As String
Dim tmp_pak As clsEnpaquetado
Dim loopArchivo As Integer
Dim cantidadSeleccionados As Integer
Dim tempInt As Integer

archivoDestino = SaveAs

If Len(archivoDestino) Then
    
   ' Contamos la cantidad de archivos seleccionados. El file no tiene SelCount!
    cantidadSeleccionados = 0
    For loopArchivo = 0 To Me.File1.ListCount - 1
        If Me.File1.Selected(loopArchivo) Then cantidadSeleccionados = cantidadSeleccionados + 1
    Next

    ' ¿Seleccionó los archivos?
    If cantidadSeleccionados = 0 Then
        Call MsgBox("Tenes que seleccionar los archivos que queres que formen parte del empaquetado.", vbExclamation, Me.caption)
        Exit Sub
    End If
    
    ' Creamos el empaquetado
    Set tmp_pak = New clsEnpaquetado
    
    'Redimensionamos donde lo vamos a guardar
    ReDim archivos(cantidadSeleccionados - 1)

   ' Guardamos los archivos
    tempInt = 0
    For loopArchivo = 0 To Me.File1.ListCount - 1
        If Me.File1.Selected(loopArchivo) Then
            archivos(tempInt) = Me.txtCarpetaEntrada & "\" & Me.File1.list(loopArchivo)
            archivos(tempInt) = Replace$(archivos(tempInt), "\\", "\")
            tempInt = tempInt + 1
        End If
    Next

   ' Generamos el paquete
    If tmp_pak.CrearDesdeCarpeta(archivoDestino, archivos) Then
        SetPak tmp_pak
        MsgBox "Nuevo empaquetado creado con exito.", vbInformation, Me.caption
    Else
        MsgBox "Se ha producido un error y no fue posible crear el paquete.", vbExclamation, Me.caption
    End If
End If

End Sub



Private Sub cmdExtraerRecursos_Click()
    Dim i As Integer
    Dim destino As String
    Dim carpetaDestino As String
        
    ' ¿Selecciono algo?
    If List1.listIndex = -1 Then
        MsgBox "Selecciona en la lista superior algún archivo para extraer.", vbExclamation, Me.caption
        Exit Sub
    End If
    
    i = val(Replace(Split(List1.list(List1.listIndex), " - ")(0), "*", ""))
    
    carpetaDestino = Me.txtCarpetaEntrada
    
    ' Armo el destino
    destino = carpetaDestino & IIf(Me.chkConNumero.value = 1, i & ".", "") & Pak.Cabezal_GetFileNameSinComplementos(i)
    
    ' Extraigo
    If Pak.Puedo_Extraer(i) Then
        If Pak.Extraer(i, destino) Then
            MsgBox "El recurso se ha extraido correctamente. Puede encontrarlo en '" & destino & "'.", vbInformation, Me.caption
        Else
            MsgBox "Lo sentimos. Se ha producido un error al intentar extraer el recurso.", vbExclamation, Me.caption
        End If
    Else
         MsgBox "No tenes permiso para extraer este archivo.", vbExclamation, Me.caption
    End If
    
    ' Refrescamos la lista
    Me.File1.Refresh
End Sub


Private Function obtenerNombreSinComplementoTexto(nombreArchivo As String) As String

    Dim Pos As Integer
    
    Pos = InStrRev(nombreArchivo, "_")
    
    If Pos > -1 Then
        obtenerNombreSinComplementoTexto = mid$(nombreArchivo, 1, Pos - 1)
    Else
        obtenerNombreSinComplementoTexto = 0
    End If
    
End Function

Private Function obtenerImagenDeComplemento(idImagen As Integer, numeroComplemento As Byte) As Integer

    Dim tmpInfoHeader As INFOHEADER
    
    '¿Ya lo tiene como complemento?
    Call Pak.IH_Get(idImagen, tmpInfoHeader)
            
    Select Case numeroComplemento
        Case 1
            obtenerImagenDeComplemento = tmpInfoHeader.complemento_1
        Case 2
            obtenerImagenDeComplemento = tmpInfoHeader.complemento_2
        Case 3
            obtenerImagenDeComplemento = tmpInfoHeader.complemento_3
        Case 4
            obtenerImagenDeComplemento = tmpInfoHeader.complemento_4
    End Select

End Function
Private Function obtenerNumeroComplementoNombreArchivo(nombreArchivo As String) As Byte

    Dim Pos As Integer
    Dim abreviaturaComplemento As String
    Dim loopComplemento As Byte
    
    Pos = InStrRev(nombreArchivo, "_")
    
    If Pos = -1 Then obtenerNumeroComplementoNombreArchivo = 0: Exit Function
    
    abreviaturaComplemento = mid$(nombreArchivo, Pos + 1)
    
    obtenerNumeroComplementoNombreArchivo = 0
    
    For loopComplemento = LBound(Complementos) To UBound(Complementos)
    
        If Complementos(loopComplemento).abreviatura = abreviaturaComplemento Then
            obtenerNumeroComplementoNombreArchivo = Complementos(loopComplemento).numero
            Exit For
        End If
        
    Next loopComplemento
    
    
End Function

Private Function esUnSprite(idImagen As Integer) As Integer

    Dim loopGrafico As Integer
    Dim cantidad As Byte
    
    cantidad = 0
    
    For loopGrafico = 1 To UBound(GrhData)
    
        If Me_indexar_Graficos.existe(loopGrafico) Then
        
            If GrhData(loopGrafico).filenum = idImagen Then cantidad = cantidad + 1
        
            If cantidad > 1 Then esUnSprite = True: Exit Function
        End If
        
    Next
    
    esUnSprite = False

End Function


Private Function crearGraficoBase(nombre As String, idImagen As Integer, rutaImagen As String, Optional ByVal ID As Integer = -1) As Integer

Dim imgDatos As ImgDimType
Dim imgExt As String
Dim idGrafico As Integer
Dim tmpGrhData As GrhData
Dim actualizando As Boolean
Dim loopCapa As Integer
Dim OffsetAjustado As Position

' Obtenemos info del archivo de  imagen
Call HelperImage.getImgDim(rutaImagen, imgDatos, imgExt)

' Solicitamos un ID para la nueva
If ID = -1 Then
    actualizando = False
    idGrafico = Me_indexar_Graficos.nuevo
Else
    actualizando = True
    idGrafico = ID
End If

If idGrafico = -1 Then
    crearGraficoBase = -1
    Exit Function
End If

' Si estoy actualizando y es un sprite (esta en varios grh) no hago ningun cambio en los grhdata
If actualizando Then
    If esUnSprite(idImagen) Then
        crearGraficoBase = ID
        Exit Function
    End If
End If

' Seteamos
With tmpGrhData

    .nombreGrafico = nombre
    .filenum = idImagen

    .NumFrames = 1
    ReDim .frames(1)
    .frames(1) = idGrafico
    .Speed = 0

    .sx = 0
    .sy = 0

    .pixelHeight = imgDatos.height
    .pixelWidth = imgDatos.width
 
    .EfectoPisada = 0
    
    .perteneceAunaAnimacion = False
    .esInsertableEnMapa = True

    ' Insertable en todas las capas
    For loopCapa = 1 To CANTIDAD_CAPAS
        .Capa(loopCapa) = True
    Next loopCapa
        
End With

Call Me_indexar_Graficos.obtenerOffsetAjustadoTile(tmpGrhData, OffsetAjustado.x, OffsetAjustado.y)

tmpGrhData.offsetX = OffsetAjustado.x
tmpGrhData.offsetY = OffsetAjustado.y

' Calculamos las propiedades variables
'Call Me_indexar_Graficos.calcularPropiedadesVariables(GrhData(idGrafico))

' Guardamos en memoria
GrhData(idGrafico) = tmpGrhData

' Guardamos en disco
Call Me_indexar_Graficos.actualizarEnIni(idGrafico)

crearGraficoBase = idGrafico
End Function

Private Function validarArchivos() As String
    Dim i As Long
    Dim nombreArchivo As String
    Dim archivo As String
    Dim extension As String
    Dim error As String
    Dim infoArchivo() As String
    
    error = vbNullString
    
    For i = 0 To File1.ListCount - 1
        If File1.Selected(i) Then
        
            archivo = File1.Path & "\" & File1.list(i)
            infoArchivo = Split(File1.list(i), ".", 5) ' Nombre.Complemento1.Complemento2.Complemento3.extension
            nombreArchivo = infoArchivo(0)
            extension = infoArchivo(UBound(infoArchivo))
        
            
            If Not FileExist(archivo) Then
                ' ¿El archivo existe?
                error = error & "El archivo '" & archivo & " ' no existe." & vbNewLine
            ElseIf Len(nombreArchivo) + Len(extension) + 1 > 32 Then
                ' La máxima longitud del nombre no puede ser de más de 32
                error = error & "El nombre del archivo '" & archivo & " ' supera el máximo de 32 carácteres para nombre de imagen.'"
            End If
            
        End If
    Next
    
    validarArchivos = error
    
End Function

Private Sub AutoIndexarImagen(imagenesNuevas() As tTipoImagen, mensajeExito As String, mensajeError As String, indexeAutomaticamente As Boolean)
Dim i As Integer
Dim respuesta As VbMsgBoxResult
Dim TmpInt As Integer
Dim nombreGrafico As String
Dim loopGrafico As Integer
Dim idGrafico As Integer

indexeAutomaticamente = False
' Archivo nuevos que NO son complementos si se respeto el estandar de nombres
For i = 0 To UBound(imagenesNuevas)
    '¿No es un complemento
    If imagenesNuevas(i).numeroComplemento = 0 Then
    
        ' ¿Ya esta indexada?
        idGrafico = 0
        For loopGrafico = 1 To UBound(GrhData)
        
            If Me_indexar_Graficos.existe(loopGrafico) Then
                If GrhData(loopGrafico).filenum = imagenesNuevas(i).idImagen Then
                    idGrafico = loopGrafico
                    Exit For
                End If
            End If
        Next
        
                
        If idGrafico > 0 Then
            ' Actualizo el grafico (ejemplo cambio de tamaño)
            Call crearGraficoBase(GrhData(idGrafico).nombreGrafico, imagenesNuevas(i).idImagen, imagenesNuevas(i).archivo, idGrafico)
            
            ' Mensaje
            mensajeExito = mensajeExito & "Se actualizó el gráfico '" & idGrafico & " - " & GrhData(idGrafico).nombreGrafico & "' relacionado a la imágen '" & imagenesNuevas(i).nombreOriginal & "'." & vbNewLine
        Else
    
            respuesta = MsgBox("¿Queres indexar automaticamente la imagen '" & imagenesNuevas(i).nombreOriginal & "'?. Se configurará un unico gráfico que abarcará toda la imagen.", vbInformation + vbYesNo)
            
            If respuesta = vbYes Then
            
                ' Le pregungo el nombre
                nombreGrafico = InputBox("¿Qué nombre le queres poner al gráfico que se va a crear?", "Crear Gráfico automaticamente", mid$(imagenesNuevas(i).nombreCompleto, 1, InStrRev(imagenesNuevas(i).nombreCompleto, ".") - 1))
                
                ' Ultima opción para darle al cancelar
                If nombreGrafico = "" Then
                    mensajeError = mensajeError & "No se indexó automáticamente la imágen '" & imagenesNuevas(i).nombreOriginal & "' porque no se eligió un nombre.'" & vbNewLine
                    GoTo continue
                End If
                
                ' Creamos el grafico
                TmpInt = crearGraficoBase(nombreGrafico, imagenesNuevas(i).idImagen, imagenesNuevas(i).archivo)
                    
                If Not TmpInt = -1 Then
                    mensajeExito = mensajeExito & "Se indexó automáticamente la imágen '" & imagenesNuevas(i).nombreOriginal & "' creandose el gráfico '" & TmpInt & " - " & nombreGrafico & "'." & vbNewLine
                Else
                    mensajeError = mensajeError & "No se pudo indexar automáticamente la imágen '" & imagenesNuevas(i).nombreOriginal & "'. No fue posible obtener un slot para guardar el gráfico.'" & vbNewLine
                End If
                    
                indexeAutomaticamente = True
            End If
            
        End If
        
    End If
continue:
Next i

End Sub
Private Sub AutoRelacionarComplementos(imagenesNuevas() As tTipoImagen, mensajeExito As String, mensajeError As String)

Dim loopComplemento As Byte
Dim complementoActual As Integer
Dim supuestoComplemento As String
Dim loopImagen As Integer
Dim TmpInt As Integer
Dim respuesta As VbMsgBoxResult
Dim elMismo As Boolean

' Archivo nuevos que tienden a NO ser complementos
For loopImagen = 0 To UBound(imagenesNuevas)
    
    '¿No es un complemento
    If imagenesNuevas(loopImagen).numeroComplemento = 0 Then
                        
        '¿Tienen algun complemento?
        ' Recorro de 1 a 4 complementos...
        For loopComplemento = LBound(Complementos) To UBound(Complementos)
            
            ' Genero el nombre que deberia tener el complemento
            supuestoComplemento = mid$(imagenesNuevas(loopImagen).nombreCompleto, 1, InStrRev(imagenesNuevas(loopImagen).nombreCompleto, ".") - 1) & "_" & Complementos(loopComplemento).abreviatura & "." & mid$(imagenesNuevas(loopImagen).nombreCompleto, InStrRev(imagenesNuevas(loopImagen).nombreCompleto, ".") + 1)
        
            ' Lo trato de obtener
            TmpInt = Pak.obtenerIndiceArchivo(supuestoComplemento)
        
            ' Parece que tiene el complemento
            If TmpInt > 0 Then
                
                '¿Ya lo tenia?. Obtengo el complemeto de la imagen que ya esta en el infoheader
                elMismo = (obtenerImagenDeComplemento(imagenesNuevas(loopImagen).idImagen, loopComplemento) = TmpInt)
                
                respuesta = vbNo
                If Not elMismo Then respuesta = MsgBox("¿El archivo '" & TmpInt & " - " & supuestoComplemento & "' parece ser el complemento " & Complementos(loopComplemento).numero & " (" & Complementos(loopComplemento).nombre & ")" & " de la imagen '" & imagenesNuevas(loopImagen).idImagen & " - " & imagenesNuevas(loopImagen).nombreOriginal & "'. ¿Los relacionamos?.", vbInformation + vbYesNo)

                If respuesta = vbYes Then
                    ' Relacionar ID Imagen Original, Numero Complemento, ID Imagen Complemento
                    Call relacionarImagenComunConComplemento(imagenesNuevas(loopImagen).idImagen, Complementos(loopComplemento).numero, TmpInt)
                
                    ' Marcamos que fue modificado
                    Call versionador.modificado(nombreVersionado, imagenesNuevas(loopImagen).idImagen, imagenesNuevas(loopImagen).nombreCompleto)
                    
                    ' Si no es el mismo no hay novedad!
                    mensajeExito = mensajeExito & "Se relacionó '" & TmpInt & " - " & supuestoComplemento & "' como complemento " & Complementos(loopComplemento).numero & " (" & Complementos(loopComplemento).nombre & ")" & " de la imagen '" & imagenesNuevas(loopImagen).idImagen & " - " & imagenesNuevas(loopImagen).nombreOriginal & "'." & vbNewLine
                End If
            End If
        Next
        
    End If
    
Next loopImagen

' Archivo nuevos que tienden a SER complementos
For loopImagen = 0 To UBound(imagenesNuevas)
    
    '¿Es un complemento
    If imagenesNuevas(loopImagen).numeroComplemento > 0 And Not imagenesNuevas(loopImagen).numeroComplemento = 255 Then
    
        ' Si sigue el formato, el nombre original va a ser del gráfico original del que este es el complemento
        TmpInt = Pak.obtenerIndiceArchivo(imagenesNuevas(loopImagen).nombreOriginal)
    
        ' ¿Existe el archivo original?
        If TmpInt > 0 Then
        
            ' Ya esta relacionado este complemento a la imagen original?. Tal vez se hizo en paso anterior
            complementoActual = obtenerImagenDeComplemento(TmpInt, imagenesNuevas(loopImagen).numeroComplemento)
                 
            If complementoActual = imagenesNuevas(loopImagen).idImagen Then GoTo continue
            
            ' Preguntamos
            respuesta = vbNo
            respuesta = MsgBox("¿El archivo '" & imagenesNuevas(loopImagen).nombreCompleto & "' parece ser el complemento " & imagenesNuevas(loopImagen).numeroComplemento & " de la imagen '" & TmpInt & " - " & imagenesNuevas(loopImagen).nombreOriginal & "'. ¿Los relacionamos?.", vbInformation + vbYesNo)
        
            If respuesta = vbYes Then
                Call relacionarImagenComunConComplemento(TmpInt, imagenesNuevas(loopImagen).numeroComplemento, imagenesNuevas(loopImagen).idImagen)
                
                ' Marcamos que fueron modificados
                Call versionador.modificado(nombreVersionado, imagenesNuevas(loopImagen).idImagen, imagenesNuevas(loopImagen).nombreCompleto)
                Call versionador.modificado(nombreVersionado, TmpInt, Pak.Cabezal_GetFileNameSinComplementos(TmpInt))
                    
                mensajeExito = mensajeExito & "Se relacionó '" & imagenesNuevas(loopImagen).nombreCompleto & "' como complemento " & imagenesNuevas(loopImagen).numeroComplemento & " de la imagen '" & TmpInt & " - " & imagenesNuevas(loopImagen).nombreOriginal & "'." & vbNewLine
            End If
        End If
        
    End If

continue:
Next loopImagen


End Sub
Private Sub cmdParchear_Click()

Dim archivo As String ' Archivo con ruta incluida
Dim nombreArchivo As String ' Solo el nombre del archivo sin los complementos
Dim extension As String ' La extension de un archivo

Dim infoArchivo() As String ' Estructura auxiliar para parseo

Dim loopArchivo As Long
Dim respuesta As VbMsgBoxResult
Dim creado As Boolean
Dim omitir As Boolean

Dim error As String
Dim mensaje As String
Dim TmpInt As Integer

Dim indexeAutomaticamente As Boolean
Dim imagenesNuevas() As tTipoImagen

indexeAutomaticamente = False

' Donde guardamos la info de las cosas que vamos parcheando
ReDim Preserve imagenesNuevas(0)
imagenesNuevas(0).numeroComplemento = 255

error = validarArchivos()

' Primero los valido todos, para que no se generen errores chotos en el medio
If Not error = vbNullString Then
    MsgBox error & mensaje, vbExclamation, Me.caption
    Exit Sub
End If

' Analizamos todos los archivos
For loopArchivo = 0 To File1.ListCount - 1
    If File1.Selected(loopArchivo) Then
        
        creado = False
        omitir = False
        
        archivo = File1.Path & "\" & File1.list(loopArchivo)
        infoArchivo = Split(File1.list(loopArchivo), ".", 5) ' Nombre.Complemento1.Complemento2.Complemento3.extension
        nombreArchivo = infoArchivo(0) '
        extension = infoArchivo(UBound(infoArchivo))

        ' PASO 1. Obtenemos el indice donde se va a guardar la imagen
        If Not IsNumeric(nombreArchivo) Then 'Si no es un número es que se desea agregar un nuevo recurso o tal vez remplazar uno que tiene ese mismo nombre
            
            'Si no establecio un número, lo tiene que establecer de alguna manera
            #If Colaborativo = 0 Then
                Do While TmpInt = 0 And Not omitir
                        TmpInt = val(InputBox("Ingresa el número de recurso al cual quiere remplazar o crear con el archivo " & File1.list(loopArchivo) & ".", "Agregar o remplazar archivo de recurso"))
                        If TmpInt = 0 Then 'No ingreso ningun numero para el grafico
                            omitir = (MsgBox("Si no ingresa un número que lo relacione no es posible agregar / modificar el gráfico " & File1.list(loopArchivo) & " zSeguro que deseas no agregarlo?", vbYesNo, File1.list(loopArchivo)) = vbYes)
                        End If
                Loop
            #Else
                ' Buscamos en el empaquetado si ya existe un archivo con ese nombre
                TmpInt = Pak.obtenerIndiceArchivo(nombreArchivo & "." & extension)
                
                '¿Existe? ¿Quiere remplazarlo?
                If TmpInt > 0 Then
                    respuesta = MsgBox("El recurso " & TmpInt & " tiene un archivo con el nombre " & nombreArchivo & ". ¿Desea remplazarlo?", vbQuestion + vbYesNo, "Remplazar recurso")
                Else
                    respuesta = vbNo
                End If
                    
                'Si no quiere remplazarlo
                If respuesta = vbNo Then
                    'Le avisamos que debe tener uno
                    respuesta = MsgBox("Se creará un nuevo recurso con el archivo '" & File1.list(loopArchivo) & "'. ¿Estás seguro?", vbQuestion + vbYesNo, "Agregar nuevo recurso")
                            
                    If respuesta = vbYes Then
                        TmpInt = CInt(CDM.cerebro.SolicitarRecurso(nombreVersionado))
                                
                        If TmpInt = -1 Then
                            error = error & "No se ha podido agregar un nuevo recurso con el archivo '" & File1.list(loopArchivo) & "'. El Cerebro de Mono no responde. Error: " & CDM.cerebro.ultimoError & "." & vbNewLine
                            omitir = True
                        Else
                            creado = True
                        End If
                    Else
                        omitir = True
                        error = error & "No se hizo nada con '" & File1.list(loopArchivo) & "'." & vbNewLine
                    End If
                End If
            #End If
            
        Else
            ' Si es un número se desea remplazar o agregar en un indice especifico de recurso
            If val(nombreArchivo) > 32678 Then
                error = error & "El número de recurso " & nombreArchivo & " (" & File1.list(loopArchivo) & ") es incorrecto. El máximo valor posible es 32678." & vbNewLine
                omitir = True
            Else
                TmpInt = val(nombreArchivo)
            End If
        End If
            
        ' PASO 2. Parcheamos
        
        ' ¿Finalmente confirmo agregar el grafico?
        If Not omitir Then
        
            ' Guardo datos de la nueva imagen que agregue
            With imagenesNuevas(UBound(imagenesNuevas))
            
                ' ¿Esta imagen puede ser un complemento?
                ' ACLARACION: Si usa el formato viejo de nombres, no va a detectarlo como complemento.
                ' se va a relacionar el complemento cuando parchea
                .numeroComplemento = obtenerNumeroComplementoNombreArchivo(nombreArchivo)
                            
                If .numeroComplemento Then
                    ' Obtenemos el nombre sin la marca de complemento
                    .nombreOriginal = obtenerNombreSinComplementoTexto(nombreArchivo) & "." & extension
                Else
                    ' El nombre no tiene marca de complemento
                    .nombreOriginal = nombreArchivo & "." & extension
                End If
                
                ' El nombre tal cual esta siendo parcheando
                .nombreCompleto = nombreArchivo & "." & extension
                ' El id del slot donde se guarda
                .idImagen = TmpInt
                ' La ruta completa donde se encuentra el archivo
                .archivo = archivo
            End With
                            
            ' Nuevo Slot
            ReDim Preserve imagenesNuevas(UBound(imagenesNuevas) + 1)
            imagenesNuevas(UBound(imagenesNuevas)).numeroComplemento = 255 ' 255 = "Vacio"
        
            ' Confirmo que tenga privilegios para modificar esto
            If Pak.Puedo_Editar(TmpInt) Then
                
                ' Parcheo
                Pak.Parchear TmpInt, archivo
                
                ' Recargo la textura
                BorrarTexturaDeMemoria TmpInt
   
                ' Guardamos mensaje y le avisamos al versionador que hay un elemento modificado
                If creado Then
                    mensaje = mensaje & "Se creó el recurso " & TmpInt & " con el archivo '" & File1.list(loopArchivo) & "'." & vbNewLine
                    Call versionador.creado(nombreVersionado, TmpInt)
                Else
                    mensaje = mensaje & "Se actualizó el recurso " & TmpInt & " con el archivo '" & File1.list(loopArchivo) & "'." & vbNewLine
                    Call versionador.modificado(nombreVersionado, TmpInt, nombreArchivo)
                End If
            Else
                error = error & "No tenés permiso para editar el slot numero " & TmpInt & vbNewLine
            End If
        End If
    End If
Next loopArchivo

' Para mejorar la calidad de vida

' PASO 3. Opcional. Relacionamos automaticamente los complementos
Call AutoRelacionarComplementos(imagenesNuevas, mensaje, error)

' Paso 4. Preguntamos si quiere indexar automaticamente las imagenes que agrego en graficos individuales
Call AutoIndexarImagen(imagenesNuevas, mensaje, error, indexeAutomaticamente)

' Paso 5. Mensaje final que resume todo lo hecho
If error <> "" Then
    MsgBox error & mensaje, vbExclamation, Me.caption
ElseIf mensaje <> "" Then
    Call MsgBox("La operación se ha completado exitosamente." & vbNewLine & mensaje, vbOKOnly + vbInformation, Me.caption)
Else
    MsgBox "Debes seleccionar algun archivo para parchear.", vbExclamation
End If

' Actualizamos la lista que se muestra a las personas
SetPak Pak

'Actualizo la lista
If indexeAutomaticamente Then CargarListaGraficosComunes
End Sub

Private Function relacionarImagenComunConComplemento(idImagen As Integer, numeroComplemento As Byte, idImagenComplemento As Integer) As Boolean
    Dim tmpInfoHeader As INFOHEADER
    
    ' Obtenemos info del Header
    Call Pak.IH_Get(idImagen, tmpInfoHeader)

    Select Case numeroComplemento
        Case 1
            tmpInfoHeader.complemento_1 = idImagenComplemento
        Case 2
            tmpInfoHeader.complemento_2 = idImagenComplemento
        Case 3
            tmpInfoHeader.complemento_3 = idImagenComplemento
        Case 4
            tmpInfoHeader.complemento_4 = idImagenComplemento
    End Select
                 
    ' Guardamos
    Call Pak.IH_Mod(idImagen, tmpInfoHeader)
          
    relacionarImagenComunConComplemento = False
    
End Function
Private Sub Command2_Click()
Dim loopElemento As Integer
Dim strDelete As String
Dim idElemento As Integer
Dim archivo As Integer
Dim destino As String

destino = app.Path & "/" & "delete_recursos.sql"

' Genero
strDelete = "UPDATE recursos SET ESTADO='CONFIRMADO' WHERE tipo='RECURSO_IMAGEN' AND ID IN("

For loopElemento = 0 To Me.List1.ListCount - 1
    idElemento = val(Replace(Split(List1.list(loopElemento), " - ")(0), "*", ""))
        
    strDelete = strDelete & idElemento & ","
Next

strDelete = mid$(strDelete, 1, Len(strDelete) - 1) & ")"
        
' Guardo en el archivo
archivo = FreeFile

Open destino For Output As #archivo
        Print #archivo, strDelete
Close #archivo

' Aviso
Call MsgBox("Se generó el archivo de delete de recursos en '" & destino & "'.")
End Sub

Private Sub examinar_in_Click()

    Dim TempStr As String
        
    TempStr = modFolderBrowse.Seleccionar_Carpeta("Selecciona la carpeta que tiene los archivos a parserar.", Me.txtCarpetaEntrada)
    
    If Not TempStr = "" Then
        Call cambiarCarpetaDeEntrada(TempStr)
    End If
End Sub

Private Sub File1_DblClick()
    cmdParchear_Click
End Sub

Private Function obtenerCarpetaDefault() As String
    Dim Carpeta As String
    
    Carpeta = ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("ArchivosEntrada")
    
    If Not FolderExist(Carpeta) Then
        Carpeta = app.Path
        Call ME_Configuracion_Usuario.actualizarVariableConfiguracion("ArchivosEntrada", Carpeta)
    End If
    
    obtenerCarpetaDefault = Carpeta
End Function

Private Sub mostrarEdicion()

Me.cmdExtraerRecursos.Visible = True
Me.chkConNumero.Visible = True
Me.cmdVersiones.Visible = True

cmdParchear.Visible = True
cmdParchear.Enabled = True

cmdCrearEmpaquetado.Visible = True
cmdCrearEmpaquetado.Enabled = True

End Sub
Private Sub Form_Load()

' Permisos: Esto es grave! Deberia denunciarlo
If Not cerebro.Usuario.tienePermisos("RECURSOS", ePermisosCDM.lectura) Then End

Me.txtCarpetaEntrada = obtenerCarpetaDefault

File1.Path = Me.txtCarpetaEntrada

Call Option2_Click(2)

If pakMapas Is Nothing Then
    Option2(4).Enabled = False
Else
    Option2(4).Enabled = True
End If

Call filtrar

' Cargo los complementos
ReDim Complementos(1 To 4) As tComplemento

Complementos(1).numero = 1
Complementos(1).nombre = "Blend One"
Complementos(1).abreviatura = "bo"

Complementos(2).numero = 2
Complementos(2).nombre = "Color Add"
Complementos(2).abreviatura = "ca"

Complementos(3).nombre = "Normal Map"
Complementos(3).numero = 3
Complementos(3).abreviatura = "nm"

Complementos(4).nombre = "Agus Scret"
Complementos(4).numero = 4
Complementos(4).abreviatura = "as"

' Permisos
If cerebro.Usuario.tienePermisos("RECURSOS", ePermisosCDM.escritura) Then
    Call mostrarEdicion
End If
    
End Sub

Private Sub SetPak(obj As clsEnpaquetado)
If Not obj Is Nothing Then
    Set Pak = obj
    Pak.Add_To_Listbox_Permisos List1, -1, 0
    Label1.caption = "Editando: " & Pak.Path_res
Else
    MsgBox "EL objeto no existe."
End If
End Sub


Private Sub Option2_Click(Index As Integer)
    Select Case Index
    Case 0
        ' Mapas
        SetPak pakMapasME
        nombreVersionado = "RECURSO_MAPA"
    Case 1
        ' GUI
        SetPak pakGUI
        nombreVersionado = "RECURSO_INTERFACE"
    Case 2
        ' GRH
        SetPak pakGraficos
        nombreVersionado = "RECURSO_IMAGEN"
    Case 3
        ' SONIDOS
        SetPak pakSonidos
        nombreVersionado = "RECURSO_SONIDO"
    Case 4
        SetPak pakMapas
    End Select
End Sub

Private Sub txtCarpetaEntrada_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then Call cambiarCarpetaDeEntrada(txtCarpetaEntrada.text)

End Sub


Private Sub cambiarCarpetaDeEntrada(ByVal nuevaRuta As String)

    If nuevaRuta = vbNullString Then Exit Sub
    
    ' Aseguramos la barra
    If right$(nuevaRuta, 1) <> "\" Then nuevaRuta = nuevaRuta & "\"
        
    ' ¿Existe?
    If Not FolderExist(nuevaRuta) Then
        Call MsgBox("La carpeta '" & nuevaRuta & "'no existe.", vbExclamation, Me.caption)
        nuevaRuta = obtenerCarpetaDefault
    End If
    
    ' Asignamos
    File1.Path = nuevaRuta
    File1.Refresh
    Me.txtCarpetaEntrada = nuevaRuta
    
    ' Guardamos
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("ArchivosEntrada", nuevaRuta)
   
End Sub

Private Sub filtrar()
    Dim extensiones() As String
    Dim loopExtension As Byte
    Dim corregido As String
    
    If Len(Trim$(txtFiltroExtensiones.text)) = 0 Then
        corregido = "*"
    Else
        extensiones = Split(txtFiltroExtensiones.text, ";")
        corregido = ""
        For loopExtension = LBound(extensiones) To UBound(extensiones)
            If Len(Trim$(extensiones(loopExtension))) > 0 Then
                If Len(corregido) > 0 Then corregido = corregido & ";"
                
                corregido = corregido & "*." & Trim$(extensiones(loopExtension))
            End If
        Next
    End If
    File1.pattern = corregido
End Sub
Private Sub txtFiltroExtensiones_Change()
  Call filtrar
End Sub

Function SaveAs() As String
    Dim tmp_path As String
    
    cdl.filter = "Empaquetado TDS (*.TDS)|*.TDS"
    cdl.flags = cdlOFNHideReadOnly
    cdl.InitDir = Clientpath & "Graficos\"
    cdl.FileName = "Parche.TDS"
    cdl.DefaultExt = "TDS"
    cdl.DialogTitle = "Guardar como..."
    cdl.ShowSave
    SaveAs = cdl.FileName
End Function
