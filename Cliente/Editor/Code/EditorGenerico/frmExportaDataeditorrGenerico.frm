VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExportarDataEditorGenerico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tierras del Sur - Exportar información"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6900
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExportaDataeditorrGenerico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmExportarDatos 
      Caption         =   "Exportar"
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin MSComDlg.CommonDialog cdlGuardarComo 
         Left            =   1920
         Top             =   1920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin EditorTDS.TextConListaConBuscador txtFila 
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   360
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   2895
      End
      Begin VB.ListBox lstColumnas 
         Appearance      =   0  'Flat
         Height          =   2055
         Left            =   3960
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   360
         Left            =   3720
         TabIndex        =   3
         Top             =   2640
         Width           =   3015
      End
      Begin VB.OptionButton optExportarTipo 
         Appearance      =   0  'Flat
         Caption         =   "Excel"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   1440
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Label lblFormato 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Formato de salida:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblColumna 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Columnas:"
         Height          =   195
         Left            =   3120
         TabIndex        =   6
         Top             =   360
         Width           =   750
      End
      Begin VB.Label lblFilas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filas:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmExportarDataEditorGenerico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private currentfile As cFileINI

Private Function toString(item As cItem) As String
    Dim seleccionado As Integer
    Dim enumerado As Collection
    Dim loopItem As Integer
    Dim items() As cItem
    Dim seleccionados() As String
    Dim fuente() As eEnumerado
    Dim loopFuente As Integer
    
    Select Case item.getType
                
            Case ItemType.e_Cadena, ItemType.e_Numerico
                toString = item.getValue
                        
            Case ItemType.e_Enumerado
                
                If item.getCombinado Then
                
                  seleccionados = Split(item.getValue, ",")
                    
                    For Each enumerado In item.getValues
                        If existeEnArrayString(enumerado.item(1), seleccionados) Then
                            toString = toString & " - " & enumerado.item(2)
                        End If
                    Next
                Else
                    seleccionado = item.getValue
    
                    For Each enumerado In item.getValues
                        If enumerado.item(1) = seleccionado Then
                            toString = enumerado.item(2)
                            Exit For
                        End If
                    Next
                End If
        
            Case ItemType.e_EnumeradoDinamico
                
                seleccionados = Split(item.getValue, ",")
                
                fuente = modEnumerandosDinamicos.obtenerEnumeradosDinamicos(item.getFuente)
                
                For loopItem = LBound(seleccionados) To UBound(seleccionados)
                    For loopFuente = LBound(fuente) To UBound(fuente)
                        If fuente(loopFuente).valor = CLng(val(seleccionados(loopItem))) Then
                             toString = toString & " - " & fuente(loopFuente).nombre
                        End If
                    Next
                Next
                
            Case ItemType.e_MixedValue
            
                items = item.getItems
                
                For loopItem = LBound(items) To UBound(items)
                    toString = toString & "-" & toString(items(loopItem))
                Next
    End Select
                
    

End Function

Public Sub iniciar(currentfile_ As cFileINI)

    Dim items() As cItem
    Dim loopItem As Integer
    
    Set currentfile = currentfile_
    
    items = currentfile.getSectionByListIndex(1).getItems
    
    Call Me.lstColumnas.Clear
    Call Me.txtFila.limpiarLista
    
    For loopItem = 1 To UBound(items)
        
       Call Me.lstColumnas.AddItem(items(loopItem).getHumanoKey)
       Call Me.txtFila.addString(loopItem, items(loopItem).getHumanoKey)
        
    Next
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdExportar_Click()
    Dim i As Integer
    Dim sec As cSection
    Dim fila As String
    Dim loopColumna As Integer
    Dim salida As String
    Dim items() As cItem
    Dim item As cItem
    
    salida = ""

    'Obtengo los items
    items = currentfile.getSectionByListIndex(1).getItems
    
    salida = items(Me.txtFila.obtenerIDValor).getHumanoKey & ";"
    For loopColumna = 1 To UBound(items)
        
        '¿Esta seleccionado?
        If Me.lstColumnas.Selected(loopColumna - 1) Then
            Set item = items(loopColumna)
            salida = salida & item.getHumanoKey & ";"
        End If
    Next
    
    salida = salida & vbNewLine
    
    ' Recorro todas las secciones (que seran las filas)
    For i = 1 To currentfile.getSectionCount
    
        ' Obtengo la fila
        Set sec = currentfile.getSectionByListIndex(i)
        
        'Obtengo los items
        items = sec.getItems
        
        'El primero será
        salida = salida & toString(items(Me.txtFila.obtenerIDValor))
        
        For loopColumna = 1 To UBound(items)
        
            '¿Esta seleccionado?
            If Me.lstColumnas.Selected(loopColumna - 1) Then
            
                Set item = items(loopColumna)
                
                salida = salida & ";" & toString(item)
            End If
        Next
        
        'Salto de linea
        salida = salida & vbCrLf
               
    Next i
    
    ' Guardamos!
    If guardar(salida) Then
        MsgBox "Datos exportados correctamente.", vbInformation, Me.caption
        MsgBox "Entrá al Excel. " & vbNewLine & "1) Anda a la pestaña Datos, hace clic en 'Desde texto'." & vbNewLine & "2) Seleccioná el archivo que acabas de crear." & vbNewLine & "3) En 'Tipo de archivo' seleccioná 'Delimitados'." & vbNewLine & "4) Pulsá siguiente." & vbNewLine & "5) Tildá donde dice 'Punto y coma'." & vbNewLine & "6) Siguiente." & vbNewLine & "7) Finalizar. ", vbInformation, Me.caption
    End If
End Sub

Public Function GuardarComo() As String
    Dim tmp_path As String
    
    On Error GoTo hayerror:

    Me.cdlGuardarComo.DefaultExt = "csv"
    Me.cdlGuardarComo.DialogTitle = "Guardar como..."
    Me.cdlGuardarComo.ShowSave
    
    tmp_path = cdlGuardarComo.FileName
    If FileExist(tmp_path, vbNormal) = True Then
        If MsgBox("¿Desea sobrescribir " & tmp_path & "?", vbExclamation + vbYesNo) = vbNo Then
            GuardarComo = vbNullString
            Exit Function
        Else
            Kill tmp_path
        End If
    End If
    GuardarComo = tmp_path
    
    Exit Function
hayerror:
    GuardarComo = vbNullString

End Function

Private Function guardar(contenido As String) As Boolean
    
    Dim archivo As Integer
    Dim ruta As String
    
    ruta = GuardarComo
    
    guardar = False
    
    If ruta = vbNullString Then Exit Function
    
    archivo = FreeFile
    
    Open ruta For Output As archivo
        Print #archivo, contenido
    Close #archivo
    
    guardar = True
End Function

