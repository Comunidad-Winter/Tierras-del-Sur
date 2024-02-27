VERSION 5.00
Begin VB.Form frmEditorGenerico 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor"
   ClientHeight    =   6870
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   8850
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "formulario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   458
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGenerarDelete 
      Caption         =   "Generar delete"
      Height          =   360
      Left            =   3480
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.CommandButton cmdExportar 
      Height          =   270
      Left            =   2640
      Picture         =   "formulario.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Exportar datos"
      Top             =   60
      Width           =   360
   End
   Begin VB.CommandButton cmdEliminarMasivo 
      Caption         =   "Eliminar masivo"
      Height          =   360
      Left            =   5160
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   1530
   End
   Begin EditorTDS.ListaConBuscador lstSecciones 
      Height          =   4695
      Left            =   0
      TabIndex        =   18
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   8281
   End
   Begin VB.TextBox txSearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6840
      TabIndex        =   12
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Guardar seccion"
      Height          =   465
      Left            =   0
      TabIndex        =   11
      Top             =   5280
      Width           =   3015
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      Height          =   345
      Left            =   1440
      TabIndex        =   10
      Top             =   5880
      Width           =   1575
   End
   Begin VB.VScrollBar sclPanel 
      Height          =   5895
      LargeChange     =   10
      Left            =   8520
      Max             =   100
      TabIndex        =   0
      Top             =   840
      Width           =   300
   End
   Begin VB.CommandButton cmdSuprimir 
      Caption         =   "Eliminar"
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Frame frmPropiedades 
      Caption         =   "Propiedades"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   3120
      TabIndex        =   3
      Top             =   720
      Width           =   5415
      Begin VB.PictureBox picCont 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   5655
         Left            =   120
         ScaleHeight     =   5655
         ScaleWidth      =   5175
         TabIndex        =   6
         Top             =   240
         Width           =   5175
         Begin VB.Frame frameMain 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   4575
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   5175
            Begin VB.Frame templateFrame 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   615
               Index           =   0
               Left            =   240
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   20
               Top             =   2760
               Visible         =   0   'False
               Width           =   4695
            End
            Begin EditorTDS.TextConListaConBuscador templateListaSimple 
               Height          =   285
               Index           =   0
               Left            =   240
               TabIndex        =   16
               Top             =   720
               Visible         =   0   'False
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   503
            End
            Begin VB.TextBox txbox 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   8
               Top             =   4560
               Visible         =   0   'False
               Width           =   855
            End
            Begin EditorTDS.UpDownText numerico 
               Height          =   315
               Index           =   0
               Left            =   240
               TabIndex        =   14
               Top             =   360
               Visible         =   0   'False
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   556
               MaxValue        =   0
               MinValue        =   0
            End
            Begin EditorTDS.TextBoxConValidador templateText 
               Height          =   285
               Index           =   0
               Left            =   240
               TabIndex        =   15
               Top             =   1080
               Visible         =   0   'False
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   503
            End
            Begin VB.ListBox lst 
               Appearance      =   0  'Flat
               Height          =   930
               Index           =   0
               Left            =   240
               Style           =   1  'Checkbox
               TabIndex        =   17
               Top             =   1440
               Visible         =   0   'False
               Width           =   4695
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "----------------------------"
               Height          =   195
               Index           =   0
               Left            =   1680
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   9
               Top             =   120
               Visible         =   0   'False
               Width           =   1680
            End
         End
      End
   End
   Begin VB.CommandButton cmdCancelarCambios 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptarCambios 
      Caption         =   "Aceptar"
      Height          =   345
      Left            =   1440
      TabIndex        =   1
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label lblEstadisticas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad: XXX"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   5040
      Width           =   1020
   End
   Begin VB.Label lblBuscarPropiedad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar propiedad"
      Height          =   195
      Left            =   5520
      TabIndex        =   19
      Top             =   405
      Width           =   1245
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   0
      Width           =   5655
   End
   Begin VB.Label lblSecciones 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Elementos"
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
      TabIndex        =   4
      Top             =   120
      Width           =   885
   End
End
Attribute VB_Name = "frmEditorGenerico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Se va a utilizar algun item para identificar a la seccion? Cual?
Public ITEM_SENALADOR As String
Public ITEM_TIPO As String
Public ITEM_TIPO_PLURAL As String
Public ITEM_VERSIONADO As String

Private senaladorSeleccionado As String

Dim currentfile As cFileINI
Dim currentSection As cSection

Dim myTDA As TDAList
    
Dim lngIncrement As Single

Private elementosModificados As New Collection
Private elementosCreados As New Collection
Private elementosEliminados As New Collection

Private Sub cerrar()

    Set elementosModificados = Nothing
    Set elementosCreados = Nothing
    Set elementosEliminados = Nothing

    Set currentfile = Nothing
    Set currentSection = Nothing
End Sub
'###################
' Interaccion PANEL
'###################

Private Sub cmdAceptarCambios_Click()

    'Al mismo tiempo que guardo el archivo, guardo en el versionador las alteraciones que hice

    If Not ITEM_VERSIONADO = "" Then
        Dim tElementoManipulado As cElementoModificado

        For Each tElementoManipulado In elementosCreados
            Call versionador.creado(ITEM_VERSIONADO, tElementoManipulado.id, tElementoManipulado.nombre)
        Next
        
        For Each tElementoManipulado In elementosEliminados
            Call versionador.eliminado(ITEM_VERSIONADO, tElementoManipulado.id, tElementoManipulado.nombre)
        Next
        
        For Each tElementoManipulado In elementosModificados
            Call versionador.modificado(ITEM_VERSIONADO, tElementoManipulado.id, tElementoManipulado.nombre)
        Next
        
        Set elementosCreados = New Collection
        Set elementosEliminados = New Collection
        Set elementosModificados = New Collection
    End If
    
    'Guardamos
    currentfile.save

    Unload Me
End Sub

Private Sub cmdCancelarCambios_Click()
    'Salimos sin guardar
    Unload Me
End Sub

Private Sub cmdExportar_Click()
    load frmExportarDataEditorGenerico
    
    Call frmExportarDataEditorGenerico.iniciar(currentfile)
    
    frmExportarDataEditorGenerico.Show vbModal, Me
End Sub

Private Sub cmdGenerarDelete_Click()
    Dim loopElemento As Integer
    Dim strDelete As String
    Dim seccion As cSection
    Dim archivo As Integer

    strDelete = "UPDATE recursos SET ESTADO='CONFIRMADO' WHERE tipo='" & ITEM_VERSIONADO & "' AND ID IN("

    For loopElemento = 1 To currentfile.getSectionCount
        Set seccion = currentfile.getSectionByListIndex(loopElemento)
        
        strDelete = strDelete & seccion.getName & ","
    Next

    strDelete = mid$(strDelete, 1, Len(strDelete) - 1) & ")"
    
    archivo = FreeFile
    
    Open "C:\generico_" & ITEM_VERSIONADO & ".txt" For Output As #archivo
        Print #archivo, strDelete
    Close #archivo
End Sub

Private Sub cmdNuevo_Click()
    Dim elemento As Integer
    Dim section As cSection
    Dim i As Long

    Me.cmdNuevo.Enabled = False
    
    If Not ITEM_VERSIONADO = "" Then
        elemento = CDM.cerebro.SolicitarRecurso(ITEM_VERSIONADO)
    Else
        elemento = CInt(val(InputBox("Ingrese el nombre de la nueva seccion.")))
    End If
    
    If elemento > 0 Then
        Set section = currentfile.newSection(CStr(elemento))
        
        'Agrego
        lstSecciones.addString elemento, elemento & " - "
        
        'Selecciono el ultimo que cree
        Call lstSecciones.seleccionarID(elemento)
        
        Dim elementoManipulado As New cElementoModificado
        elementoManipulado.id = elemento
        elementoManipulado.nombre = ""
        
        Call elementosCreados.Add(elementoManipulado)
        
        Call actualizarEstadisticas
    Else
       Call MsgBox("No se ha podido crear el " & ITEM_TIPO & ". Por favor, intente más tarde o contacte a un Administrador.", vbExclamation, Me.caption)
    End If
    
    Me.cmdNuevo.Enabled = True
End Sub

Private Sub cmdSave_Click()
    Dim item As cItem
    
    If SetSectionFromTDA(myTDA, currentSection) Then
        '¿Cambio el valor del identificador?
        If Len(ITEM_SENALADOR) > 0 Then
            Set item = currentSection.getItemByName(ITEM_SENALADOR)
    
            If Not senaladorSeleccionado = item.getValue Then
                Call lstSecciones.cambiarNombre(val(currentSection.getName), val(currentSection.getName) & " - " & item.getValue)
                senaladorSeleccionado = item.getValue
            End If
        End If
        
        'Guardamos en la lista de elementos modificados
        Dim elementoManipulado As New cElementoModificado
        elementoManipulado.id = currentSection.getName
        
        If Not ITEM_SENALADOR = "" Then
            elementoManipulado.nombre = currentSection.getItemByName(ITEM_SENALADOR).getValue
        Else
            elementoManipulado.nombre = ""
        End If
        
        Call elementosModificados.Add(elementoManipulado)
    Else
        MsgBox "Hay algun(os) parametro(s) incorrecto(s), porfavor asegurate de que todo este bien."
    End If
End Sub

Private Sub cmdSuprimir_Click()
    Dim sec As Long
    sec = lstSecciones.obtenerIDValor
    
    If MsgBox("¿Estas seguro de eliminar el " & ITEM_TIPO & " '" & lstSecciones.obtenerValor & "'?", vbExclamation Or vbYesNo) = vbYes Then
        Call eliminar(sec)
    End If
End Sub

Private Sub eliminar(ByVal sec As Integer)
    Dim elementoManipulado As New cElementoModificado
    elementoManipulado.id = sec
    
    If Not ITEM_SENALADOR = "" Then
        elementoManipulado.nombre = currentSection.getItemByName(ITEM_SENALADOR)
    Else
        elementoManipulado.nombre = ""
    End If
    
    'Eliminamos la seccion
    Call currentfile.delSection(CStr(sec))
        
    ' Eliminamos de la lista
    Call Me.lstSecciones.eliminar(sec)
        
    ' Guardamos en la lista de elementos eliminados
    Call elementosEliminados.Add(elementoManipulado)
        
    Call actualizarEstadisticas
End Sub

Public Sub iniciar(configuracion As cFileJSON)
    
    'Configuramos el TDA. Establecemos los elementos templates con los que va a trabajar el generador
    InitTDA myTDA, Me.picCont, Me.frameMain, Me.templateText, _
                Me.lst, Me.templateListaSimple, Me.lbl, Me.templateFrame, Me.lblStatus, Me.numerico
    
                
    'Generamos el formulario con la estructura en blanco
    LoadTDAFromItems myTDA, configuracion.getItems()
    
    'Refrescamos la barra (?)
    frmEditorGenerico.refreshScrollbar
    
    'Titulo
    frmEditorGenerico.lblSecciones = UCase$(left$(Me.ITEM_TIPO_PLURAL, 1)) & mid$(Me.ITEM_TIPO_PLURAL, 2)
    
    'Arranca todo deshabilitado
    Call modPosicionarFormulario.setEnabledHijos(False, Me.frmPropiedades, Me)
End Sub


Public Sub seleccionar(ByVal elemento As Long)
    Call Me.lstSecciones.seleccionarID(elemento)
End Sub

'###################
' Carga de PANEL
'###################

Private Sub cargarSecciones()
    Dim i As Long, sec As cSection
    Dim item As cItem
    Dim nombreDescriptivo As Integer
    
    lstSecciones.vaciar
    
    For i = 1 To currentfile.getSectionCount
        Set sec = currentfile.getSectionByListIndex(i)

        If Not ITEM_SENALADOR = "" Then
            Set item = sec.getItemByName(ITEM_SENALADOR)
            
            If Not item Is Nothing Then
                lstSecciones.addString CInt(val(sec.getName)), CInt(val(sec.getName)) & " - " & item.getValue
                GoTo continue
            End If
        End If
        
        lstSecciones.addString CInt(val(sec.getName)), CInt(val(sec.getName)) & " - " & sec.getName
        
        
continue:
    Next i
End Sub
Public Sub showFile(ByRef file As cFileINI)
    Set currentfile = file
    
    Call cargarSecciones
    
    Call actualizarEstadisticas
End Sub

Private Sub cmdEliminarMasivo_Click()
    Dim desde As Integer
    Dim hasta As Integer
    Dim loopE As Integer
    
    desde = val(InputBox("Desde"))
    hasta = val(InputBox("hasta"))
    
    For loopE = desde To hasta
        Call eliminar(loopE)
    Next
End Sub

Private Sub Form_Load()
    senaladorSeleccionado = ""
    ITEM_VERSIONADO = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim respuesta As VbMsgBoxResult
    
    If elementosModificados.count + elementosEliminados.count + elementosCreados.count > 0 Then
        respuesta = MsgBox("Hay elemento que modificaste y no se guardaron ¿Seguro que queres cerrar la pantalla?. Para que se guarden tenes que pulsar Aceptar.", vbInformation + vbYesNo, "Ignorar cambios")
    
        If respuesta = vbNo Then
            Cancel = 1
            Exit Sub
        End If
        Call cerrar
    End If
End Sub

Private Sub lstSecciones_Change(Valor As String, id As Integer)
    Dim key As String, vArr() As String, v As String
    Dim item As cItem
    
    If lstSecciones.obtenerIDValor > 0 Then
        Set currentSection = currentfile.getSectionByName(lstSecciones.obtenerIDValor)
        
        If Not currentSection Is Nothing Then
            'Blanqueamos
            SetNullValuesInTDA myTDA
            
            'Establecemos los nuevos valores
            SetSectionInTDA myTDA, currentSection
            
            'Habilito los controles
            Call modPosicionarFormulario.setEnabledHijos(True, Me.frmPropiedades, Me)
            
            'Guardo la identificación de la seccion actual
            If Len(ITEM_SENALADOR) > 0 Then
                Set item = currentSection.getItemByName(ITEM_SENALADOR)
                
                senaladorSeleccionado = item.getValue
            End If
        Else
            MsgBox "La sección no existe. Avise al programador.", vbCritical
        End If
    End If
End Sub

Private Sub actualizarEstadisticas()
    Me.lblEstadisticas = "Cantidad: " & currentfile.getSectionCount
End Sub
Private Sub sclPanel_Scroll()
    Call actualizarVisibilidad
End Sub

Private Function quitarTildes(texto As String) As String
    quitarTildes = Replace$(texto, "á", "a")
    quitarTildes = Replace$(quitarTildes, "é", "e")
    quitarTildes = Replace$(quitarTildes, "í", "i")
    quitarTildes = Replace$(quitarTildes, "ó", "o")
    quitarTildes = Replace$(quitarTildes, "ú", "u")
End Function

Private Sub txSearch_Change()
    Dim searchControl
    Dim parentFrame
    Dim found As Boolean
    Dim Y As Long
    Dim cadenaBuscada As String
    
    If txSearch.Text = "" Then
        txSearch.BackColor = vbWhite
        Exit Sub
    Else
        txSearch.BackColor = vbRed
    End If
    
    cadenaBuscada = UCase$(quitarTildes(txSearch.Text))
    
    found = False
    For Each searchControl In myTDA.labels
        If InStr(UCase$(quitarTildes(searchControl.caption)), cadenaBuscada) > 0 Then
            found = True
            Exit For
        End If
    Next
    
    ' Si no esta en los Labels busco en los frames
    If Not found Then
        For Each searchControl In myTDA.frames
            If InStr(UCase$(quitarTildes(searchControl.caption)), cadenaBuscada) > 0 Then
                found = True
                Exit For
            End If
        Next
    End If
    
    If found Then
        txSearch.BackColor = vbGreen

        If TypeName(searchControl) = "Label" Then
            Set parentFrame = searchControl.Container
            If parentFrame <> frameMain Then Y = parentFrame.top
        End If
        
        Y = Y + searchControl.top - 150
                 
        sclPanel.value = IIf((Y / (frameMain.Height - picCont.Height)) * sclPanel.max < 0, 0, (Y / (frameMain.Height - picCont.Height)) * sclPanel.max)
    End If
    
End Sub

Public Sub refreshScrollbar()
    lngIncrement = (frameMain.Height - picCont.Height) / sclPanel.max
    If lngIncrement < 0 Then sclPanel.Enabled = False Else sclPanel.Enabled = True
End Sub

Private Sub actualizarVisibilidad()
    frameMain.top = -(sclPanel.value * lngIncrement)
End Sub
Private Sub sclPanel_Change()
   Call actualizarVisibilidad
End Sub
