VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportarConfiguracionGraficos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Configuración de Gráficos"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImportarConfiguracionGraficos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   499
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   StartUpPosition =   1  'CenterOwner
   Begin EditorTDS.TextConListaConBuscador Listado 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   503
      CantidadLineasAMostrar=   10
   End
   Begin VB.CommandButton cmdImportarIndex 
      Caption         =   "Cargar desde Archivo"
      Height          =   375
      Left            =   3240
      Picture         =   "frmImportarConfiguracionGraficos.frx":1CCA
      TabIndex        =   9
      ToolTipText     =   "Importar configuración de gráficos"
      Top             =   6600
      Width           =   2895
   End
   Begin MSComDlg.CommonDialog oFile 
      Left            =   1320
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPortapapeles 
      Caption         =   "Cargar desde Portapapeles"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6600
      Width           =   2895
   End
   Begin VB.CommandButton cmdImportar 
      Caption         =   "Importar"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   7080
      Width           =   2895
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   7080
      Width           =   2895
   End
   Begin VB.TextBox txtConfiguracionGraficos 
      Appearance      =   0  'Flat
      Height          =   4335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2160
      Width           =   6015
   End
   Begin VB.CheckBox chkCapas 
      Appearance      =   0  'Flat
      Caption         =   "Capa 4"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   12
      Top             =   1080
      Width           =   855
   End
   Begin VB.CheckBox chkCapas 
      Appearance      =   0  'Flat
      Caption         =   "Capa 5"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   5160
      TabIndex        =   13
      Top             =   1080
      Width           =   855
   End
   Begin VB.CheckBox chkCapas 
      Appearance      =   0  'Flat
      Caption         =   "Capa 3"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   11
      Top             =   1080
      Width           =   855
   End
   Begin VB.CheckBox chkCapas 
      Appearance      =   0  'Flat
      Caption         =   "Capa 2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   10
      Top             =   1080
      Width           =   855
   End
   Begin VB.CheckBox chkCapas 
      Appearance      =   0  'Flat
      Caption         =   "Capa 1"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblForCapas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Capas donde se puede insertar estos gráficos:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   3885
   End
   Begin VB.Label lblImagen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imágen donde está el o los gráficos:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3045
   End
   Begin VB.Label lblInstrucciones 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmImportarConfiguracionGraficos.frx":200C
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   6060
   End
End
Attribute VB_Name = "frmImportarConfiguracionGraficos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Si agregue algún gráfico esta variable es verdadera
Private huboCambios As Boolean

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdImportar_Click()

    '¿Selecciono la imagen?
    If Me.Listado.obtenerIDValor = 0 Then
        Call MsgBox("Tenes que seleccionar el recurso de imagen en donde está el o los gráficos cuya configuración queres cargar.", vbExclamation, Me.caption)
        Exit Sub
    End If

    '¿Puso algo como configuración?
    If Len(Trim$(Me.txtConfiguracionGraficos.text)) = 0 Then
        Call MsgBox("Tenés que poner la configuración de los gráficos que necesitas configurar. Esta configuración te la da el Generador de Imágenes.", vbExclamation, Me.caption)
        Exit Sub
    End If
    
    ' Importamos
    Call importarConfiguracion
End Sub

Private Sub setEstadoBotones(estado As Boolean)
    Me.cmdCerrar.Enabled = estado
    Me.cmdImportar.Enabled = estado
    Me.cmdImportarIndex.Enabled = estado
    Me.cmdPortapapeles.Enabled = estado
End Sub

Private Sub cmdImportarIndex_Click()

Dim archivo As String
Dim config As String
Dim handle As Integer

On Error GoTo BotonCancelar

oFile.CancelError = True 'Si la persona toca cancelar, se genera un error
oFile.filter = "Configuración de Gráficos (*.gconf)|*.gconf"
oFile.flags = cdlOFNHideReadOnly
oFile.DefaultExt = "gconf"
oFile.DialogTitle = "Seleccioné el archivo con la configuración de Gráficos a Importar"
oFile.ShowOpen

archivo = oFile.FileName

' ¿Existe?
If Not FileExist(archivo, vbArchive) Then
    Call MsgBox("El archivo seleccionado no existe", vbExclamation, Me.caption)
    Exit Sub
End If

' ¿No es demasiado grande no?
If HelperFiles.getFileSize(archivo) > 10240 Then '10Kb.
    Call MsgBox("El archivo es demasiado grande. Es extraño que un archivo de configuración de Gráficos sea tan grande.", vbExclamation, Me.caption)
    Exit Sub
End If

' Obtenemos la informacion del archivo. Lo leo todo de una ya que no va a ser significativamente grande.
config = LeerArchivo(archivo)

' Cargamos la data
Me.txtConfiguracionGraficos.text = config

Exit Sub
BotonCancelar:
    Err.Clear
    Exit Sub
    
End Sub


Private Function crearGraficos(configs As Collection, errores As Collection) As Integer

    Dim config As cConfiguracionIndexGrafico
    Dim idGrafico As Integer
    Dim id As Integer
    Dim exitosos As Integer
    Dim loopCapa As Integer
    
    exitosos = 0
    
    ' Recorremos todas las configuraciones. Algunos serán para modificar y otros para agregar
    For Each config In configs
       
        '¿Ya existe?
        idGrafico = Me_indexar_Graficos.obtenerIDPorIDUnico(config.id)

        If idGrafico = 0 Then idGrafico = Me_indexar_Graficos.nuevo()
                    
        If idGrafico = -1 Then
            Call errores.Add("No se pudo obtener un identificador para el gráfico  " & config.nombre & ". Por favor, intente más tarde o informele a un administrador.")
            GoTo continue_for
        End If
            
        ' Ponemos los datos
        Call Me_indexar_Graficos.establecerConfigBasica(idGrafico, config.nombre, config.ancho, config.alto, config.imagen, config.x, config.y, config.id)
        
        ' Seteamos las capas adicionalmente
        If config.capas = 0 Then
            GrhData(idGrafico).esInsertableEnMapa = False
        Else
            GrhData(idGrafico).esInsertableEnMapa = True
        End If
        For loopCapa = 1 To CANTIDAD_CAPAS
            GrhData(idGrafico).Capa(loopCapa) = HelperBitWise.BS_Byte_Get(config.capas, loopCapa)
        Next
              
        ' Guardamos
        Call Me_indexar_Graficos.actualizarEnIni(idGrafico)
            
        ' Contamos
        exitosos = exitosos + 1

        huboCambios = True
continue_for:
    Next


crearGraficos = exitosos

End Function


Private Sub importarConfiguracion()
    Dim lineas() As String
    Dim loopLinea As Integer
    Dim infoGrafico() As String
    Dim idImagen As Integer
    Dim configs As Collection
    Dim configGrafico As cConfiguracionIndexGrafico
    Dim resultado As VbMsgBoxResult
    Dim todoSi As Boolean ' Para evitar tener que confirman varias veces el remplazo de graficos
    Dim huboError As Boolean
    Dim errores As Collection
    Dim capasAplica As Byte
    Dim loopCapa As Integer
    
    Set configs = New Collection
    
    ' Imagen donde están los gráficos
    idImagen = Me.Listado.obtenerIDValor

    ' Capas en donde puede ser aplicado
    capasAplica = 0
    For loopCapa = 1 To CANTIDAD_CAPAS
        If Me.chkCapas(loopCapa) Then Call BS_Byte_On(capasAplica, loopCapa)
    Next loopCapa
    
    ' Separamos por los saltos de linea
    huboError = False
    todoSi = False
    lineas = Split(Me.txtConfiguracionGraficos.text, vbCrLf)
    
    For loopLinea = 0 To UBound(lineas)
        
        ' Salteamos lineas en blanco
        If Len(Trim$(lineas(loopLinea))) = 0 Then GoTo continue_for
        
        ' Separo cada campo
        infoGrafico = Split(lineas(loopLinea), ";")
    
        ' ¿Estan los datos?
        If Not UBound(infoGrafico) = 5 Then
            Call MsgBox("Hay un problema en la linea " & (loopLinea + 1) & ". " & vbCrLf & "'" & lineas(loopLinea) & "'" & vbCrLf & vbCrLf & "La configuración no es correcta. El formato es:" & vbCrLf & "identificador unico;nombre del grafico;ancho;alto;posicion en x dentro de la imagen; posición en y dentro de la imagen", vbExclamation, Me.caption)
            huboError = True
            Exit For
        End If
        
        Set configGrafico = New cConfiguracionIndexGrafico
        
        ' Tomo los datos
        configGrafico.id = Trim$(infoGrafico(0))
        configGrafico.nombre = Trim$(infoGrafico(1))
        configGrafico.ancho = CInt(val(infoGrafico(2)))
        configGrafico.alto = CInt(val(infoGrafico(3)))
        configGrafico.imagen = idImagen
        configGrafico.x = CInt(val(infoGrafico(4)))
        configGrafico.y = CInt(val(infoGrafico(5)))
        configGrafico.capas = capasAplica
        
        ' ¿ya existe?
        'If Not todoSi Then
        '    If Me_indexar_Graficos.existeNombre(configGrafico.nombre) Then
        '        resultado = MsgBox("Ya existe un gráfico con el nombre '" & configGrafico.nombre & "'. ¿Querés remplazarlo?. Sino queres remplazarlo lo mejor es que le pongas otro nombre, así evitamos conflictos. " & vbCrLf & "Toca en 'No' para no importar este gráfico y 'Cancelar' para parar el proceso y restaurar los ultimos cambios.", vbQuestion + vbYesNoCancel, Me.caption)
               
        '       If resultado = vbCancel Then
        '            huboError = True
        '            Exit For
        '       ElseIf resultado = vbNo Then
        '            huboError = True
        '            GoTo continue_for
        '       Else
        '            todoSi = (MsgBox("¿Querés remplazar todos los próximos gráficos que ya existan?", vbQuestion + vbYesNo, Me.caption) = vbYes)
        '       End If
        '    End If
        'End If
        
        ' Agregamos a la lista
        Call configs.Add(configGrafico)
        
        '
continue_for:
    Next
    
    If huboError Then Exit Sub
    
    If configs.count = 0 Then
        Call MsgBox("No hay ninguna configuración de gráfico válida para ser importada.", vbExclamation, Me.caption)
        Exit Sub
    End If
   
    ' Pregungo por las dudas?
    resultado = MsgBox("Se van a importar " & configs.count & " gráficos. Esto puede tardar unos instantes. ¿Estás seguro?", vbQuestion + vbYesNo, Me.caption)
   
    If resultado = vbNo Then Exit Sub
   
    Call setEstadoBotones(False)
   
    Set errores = New Collection
    
    Call crearGraficos(configs, errores)
    
    If errores.count Then
        Call MsgBox("Se produjeron " & errores.count & " errores al importar los gráficos:" & vbCrLf & modColeccion.Coleccion_Join(errores), vbExclamation, Me.caption)
    Else
        Call MsgBox("Se importaron correctamente todos los gráficos.", vbInformation, Me.caption)
    End If
    
    Call setEstadoBotones(True)
End Sub
Private Sub cmdPortapapeles_Click()

    If Clipboard.GetFormat(vbCFText) Then
        Me.txtConfiguracionGraficos.text = Clipboard.GetText(vbCFText)
        Call MsgBox("Información copiada desde el portapapeles.", vbInformation, Me.caption)
    Else
        Call MsgBox("No hay información en el portapapeles.", vbExclamation, Me.caption)
    End If

End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim elementos() As modEnumerandosDinamicos.eEnumerado
    
    huboCambios = False
    ' Imagenes disponibles
    elementos = modEnumerandosDinamicos.obtenerEnumeradosDinamicos("IMAGENES")
    
    For i = LBound(elementos) To UBound(elementos)
        Call Me.Listado.addString(elementos(i).valor, elementos(i).valor & " - " & elementos(i).nombre)
    Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Para que la lista quede actualizada
    If huboCambios Then frmConfigurarGraficos.seActualizoGraficos
End Sub

