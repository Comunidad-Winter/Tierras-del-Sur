VERSION 5.00
Begin VB.Form frmAdminEventos 
   Caption         =   "Administrador de eventos"
   ClientHeight    =   7005
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelarEvento 
      Caption         =   "Cancelar Evento"
      Height          =   360
      Left            =   8520
      TabIndex        =   18
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton cmdVerEventos 
      Caption         =   "Ver eventos"
      Height          =   375
      Left            =   6480
      TabIndex        =   13
      Top             =   4320
      Width           =   1815
   End
   Begin VB.ListBox lstEstadoEventos 
      Appearance      =   0  'Flat
      Height          =   2955
      Left            =   6480
      TabIndex        =   12
      Top             =   1200
      Width           =   3975
   End
   Begin VB.CommandButton cmdActualizarEventos 
      Caption         =   "Actualizar"
      Height          =   735
      Left            =   3960
      TabIndex        =   11
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdVerRings 
      Caption         =   "Ver Rings"
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   4320
      Width           =   3135
   End
   Begin VB.CommandButton cmdVerDescansos 
      Caption         =   "Ver Descansos"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   2895
   End
   Begin VB.ListBox lstEstadoRings 
      Appearance      =   0  'Flat
      Height          =   2955
      Left            =   3240
      TabIndex        =   8
      Top             =   1200
      Width           =   3135
   End
   Begin VB.ListBox lstEstadoDescansos 
      Appearance      =   0  'Flat
      Height          =   2955
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   2895
   End
   Begin VB.CommandButton cmdRecargarDescansos 
      Caption         =   "Recargar descansos"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Frame frmRetos 
      Caption         =   "Retos"
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Width           =   10335
      Begin VB.CommandButton cmdHabilitarResuRetos 
         Caption         =   "Habilitar resu en retos"
         Height          =   360
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton cmdHabilitarPlantado 
         Caption         =   "Habilitar Plantado"
         Height          =   360
         Left            =   4200
         TabIndex        =   20
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton cmdEstado3vs3 
         Caption         =   "Habilitar 3vs3"
         Height          =   360
         Left            =   2160
         TabIndex        =   19
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton cmdActualizarCantidadRetos 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   6240
         TabIndex        =   17
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtMaximaCantidadRetos 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4200
         TabIndex        =   16
         Text            =   "9"
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdHabilitarRetoOro 
         Caption         =   "Habilitar retos por oro"
         Height          =   375
         Left            =   2160
         TabIndex        =   15
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton cmdHabilitarRetoItems 
         Caption         =   "Habilitar reto por items"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton cmdActualizarRetosActivos 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblCantidadRetosActivos 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdRecargarRings 
      Caption         =   "Recargar Rings"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   4800
      Width           =   3135
   End
   Begin VB.Label lbltamanioEventos 
      Caption         =   "Tamaño de la lista de eventos:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   3645
   End
   Begin VB.Label lblEventosCantidad 
      Caption         =   "Cantidad de eventos desarrollandose:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmAdminEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cerrar_Click()
Unload Me
End Sub

Private Sub cmdActualizarCantidadRetos_Click()
modRetos.cantidadMaximaRetos = Me.txtMaximaCantidadRetos
MsgBox "Actualizado"

End Sub

Private Sub cmdActualizarEventos_Click()
Me.lblEventosCantidad = modEventos.getCantidadEnventos
Me.lbltamanioEventos = modEventos.getMayorIndex
End Sub

Private Sub cmdActualizarRetosActivos_Click()
    Me.lblCantidadRetosActivos = "Cantidad retos activos:" & modRetos.getCantidadRetosActivos
End Sub

Private Sub cmdCancelarEvento_Click()
    Dim nombreEvento As String
    Dim inicioNombre As Byte
    Dim finNombre As Byte
    
    If Me.lstEstadoEventos.ListIndex >= 0 Then
        
        nombreEvento = Me.lstEstadoEventos.List(Me.lstEstadoEventos.ListIndex)
        inicioNombre = InStr(1, nombreEvento, "-") + 1
        finNombre = InStr(inicioNombre, nombreEvento, "-")
        
        If inicioNombre > 0 And finNombre > inicioNombre Then
            nombreEvento = mid$(nombreEvento, inicioNombre, finNombre - inicioNombre)
        End If
        Call modEventos.cancelarEvento(nombreEvento)
    Else
        MsgBox "Debe seleccionar un evento para cancelar"
    End If
End Sub

Private Sub actualizarBotonesHabilitacionRetos()

If Not modRetos.permitir3vs3 Then
    Me.cmdEstado3vs3.Caption = "Activar 3vs3"
Else
    Me.cmdEstado3vs3.Caption = "Desactivar 3vs3"
End If

If Not modRetos.permitirPlantado Then
    Me.cmdHabilitarPlantado.Caption = "Activar Plantados"
Else
    Me.cmdHabilitarPlantado.Caption = "Desactivar Plantado"
End If

If Not modRetos.permitirItems Then
    Me.cmdHabilitarRetoItems.Caption = "Activar retos por items"
Else
    Me.cmdHabilitarRetoItems.Caption = "Desactivar retos por items"
End If

If Not modRetos.permitirOro Then
    Me.cmdHabilitarRetoOro.Caption = "Activar retos por oro"
Else
    Me.cmdHabilitarRetoOro.Caption = "Desactivar retos por oro"
End If

If Not modRetos.permitirResu Then
    Me.cmdHabilitarResuRetos.Caption = "Activar retos con resu"
Else
    Me.cmdHabilitarResuRetos.Caption = "Desactivar retos con resu"
End If


End Sub
Private Sub cmdEstado3vs3_Click()

modRetos.permitir3vs3 = Not modRetos.permitir3vs3

Call actualizarBotonesHabilitacionRetos

End Sub

Private Sub cmdHabilitarPlantado_Click()

modRetos.permitirPlantado = Not modRetos.permitirPlantado

Call actualizarBotonesHabilitacionRetos

End Sub

Private Sub cmdHabilitarResuRetos_Click()

modRetos.permitirResu = Not modRetos.permitirResu

Call actualizarBotonesHabilitacionRetos

End Sub

Private Sub cmdHabilitarRetoItems_Click()

modRetos.permitirItems = Not modRetos.permitirItems

Call actualizarBotonesHabilitacionRetos

End Sub

Private Sub cmdHabilitarRetoOro_Click()

modRetos.permitirOro = Not modRetos.permitirOro

Call actualizarBotonesHabilitacionRetos

End Sub

Private Sub cmdRecargarDescansos_Click()
    Call modDescansos.reCargarZonasDescanso
    Call MsgBox("Los descansos fueron re-cargados.", vbInformation, "Tierras del Sur")
End Sub

Private Sub cmdRecargarRings_Click()
    Call modRings.reCargarRings
    Call MsgBox("Los rings fueron re-cargados.", vbInformation, "Tierras del Sur")
End Sub

'Private Sub cmdReintentarSum_Click()
'    Dim resultado As VbMsgBoxResult
'    Dim cantidadOffline As Integer
'    Dim listaOffline As String
'
'    resultado = MsgBox("¿REINTENTAR?", vbOKCancel)
'
'    If resultado = vbOK Then
'        cantidadOffline = modSumoneoAutomatico.reintentarSumonear(listaOffline)
'        If cantidadOffline = 0 Then
'            MsgBox "Genial!. Todos sumoneados"
'        Else
'            MsgBox "Hay " & cantidadOffline & " offline: " & listaOffline
'        End If
'
'    End If
'End Sub

'Private Sub cmdResetearSumoneo_Click()
'    Dim resultado As VbMsgBoxResult
'
'    resultado = MsgBox("¿RESET?", vbOKCancel)
'
'    If resultado = vbOK Then
'        Call modSumoneoAutomatico.resetSumonedos
'    End If
'End Sub

'Private Sub cmdSumonearADescansos_Click()
'    Dim infoManual As String
'    Dim infoParseada() As tipoParserEquipo
'    Dim cantidadOffline As Integer
'    Dim resultado As VbMsgBoxResult
'    Dim listaOffline As String
'    resultado = MsgBox("¿SUMONEAR?", vbOKCancel)
'
'    If resultado = vbOK Then
'        infoManual = Me.txtEquiposManual
'        Call quitarEnters(infoManual)
'        Me.txtEquiposManual = infoManual
'
'        Debug.Print "a" & infoManual & "a"
'        Call modParserTextBox.parsear(infoManual, infoParseada)
'
'        Call modSumoneoAutomatico.resetSumonedos
'        cantidadOffline = modSumoneoAutomatico.sumonearParseados(infoParseada, listaOffline)
'
'        If cantidadOffline = 0 Then
'            MsgBox "Genial!. Todos sumoneados"
'        Else
'            MsgBox "Hay " & cantidadOffline & " offline: " & listaOffline
'        End If
'    End If
'End Sub
Private Sub quitarEnters(cadena As String)
    If InStrRev(cadena, vbCrLf) = Len(cadena) - 1 Then
        cadena = mid$(cadena, 1, Len(cadena) - 2)
        Call quitarEnters(cadena)
    End If
End Sub

Private Sub cmdSumonearADescansos_Click()

End Sub

Private Sub cmdVerDescansos_Click()
    Me.lstEstadoDescansos.Clear
    Call modDescansos.verEstado(Me.lstEstadoDescansos)
End Sub

Private Sub cmdVerEventos_Click()
    Me.lstEstadoEventos.Clear
    Call modEventos.verEstadoEventos(Me.lstEstadoEventos)
End Sub

Private Sub cmdVerRings_Click()
    Me.lstEstadoRings.Clear
    Call modRings.verEstado(Me.lstEstadoRings)
End Sub

Private Sub Form_Load()
Me.lblEventosCantidad = modEventos.getCantidadEnventos
Me.lbltamanioEventos = modEventos.getMayorIndex
Me.lblCantidadRetosActivos = "Cantidad retos activos:" & modRetos.getCantidadRetosActivos

Call actualizarBotonesHabilitacionRetos
End Sub

Private Sub lstEstadoEventos_Click()
Me.lstEstadoEventos.ToolTipText = Me.lstEstadoEventos.List(Me.lstEstadoEventos.ListIndex)
End Sub

Private Sub formatearTextNumero(Text As TextBox)
    Dim texto As String
    
    texto = Replace$(Text, ".", "")
    Text = FormatNumber(val(texto), 0, vbTrue, vbFalse, vbTrue)
    Text.SelStart = Len(Text.Text)
End Sub

Private Sub txtEquiposManual_Change()

End Sub
