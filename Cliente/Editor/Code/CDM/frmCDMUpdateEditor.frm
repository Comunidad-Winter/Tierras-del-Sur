VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCDMUpdateEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tierras del Sur - Editor del Mundo"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCDMUpdateEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer trmArrancar 
      Interval        =   1
      Left            =   1920
      Top             =   3840
   End
   Begin VB.Frame frmMensaje 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   1080
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   5655
      Begin MSComctlLib.ProgressBar pgrProgresoUpdate 
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   720
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblMensaje 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Actualizando"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   1845
      End
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6000
      TabIndex        =   2
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   4440
      TabIndex        =   1
      Top             =   3960
      Width           =   1350
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5741
      _Version        =   393216
      AllowBigSelection=   0   'False
      HighLight       =   0
      FillStyle       =   1
      ScrollBars      =   2
      PictureType     =   1
      Appearance      =   0
   End
   Begin VB.Label lblBuscando 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscando novedades..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   5955
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblUltimasModificaciones 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ultimas modificaciones desde la ultima actualización:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4485
   End
End
Attribute VB_Name = "frmCDMUpdateEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents repositorio As clsCDM
Attribute repositorio.VB_VarHelpID = -1

Private Sub cmdActualizar_Click()
    Me.cmdActualizar.Enabled = False
    Me.frmMensaje.Visible = True
    
    Call repositorio.Repositorio_Actualizar
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set repositorio = CDM.cerebro
End Sub

Private Sub mostrarNovedades(novedades As Collection)
    Dim fila As Integer
    Dim lTextWidth As Long
    Dim intMultiplier As Integer
    Dim Version As Dictionary
    Dim lColWidth As Long
    
    Dim hayNovedades As Boolean
    
    hayNovedades = False
    
    If Not novedades Is Nothing Then
        If novedades.count > 0 Then
            hayNovedades = True
        End If
    End If
    
    With Me.MSFlexGrid1
    
        If hayNovedades Then
            .rows = novedades.count + 1
        Else
            .rows = 1
        End If
        
        .Cols = 4
            
        .TextMatrix(0, 0) = "Version"
        .TextMatrix(0, 1) = "Fecha"
        .TextMatrix(0, 2) = "Editor"
        .TextMatrix(0, 3) = "Novedades"
            
        .ColWidth(0) = 800
        .ColWidth(1) = 2000
        .ColWidth(2) = 1200
        .ColWidth(3) = .width - .ColWidth(0) - .ColWidth(1) - .ColWidth(2) - 250
            
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
    
        fila = 1
        
        If hayNovedades Then
          For Each Version In novedades
              'Cargamos los datos
              .TextMatrix(fila, 0) = Version.item("numero")
              .TextMatrix(fila, 1) = Version.item("fecha")
              .TextMatrix(fila, 2) = Version.item("usuario")
              .TextMatrix(fila, 3) = Version.item("comentario")
                  
              'Nos aseguramos que entre el texto de novedades
              lColWidth = .ColWidth(3)
              lTextWidth = TextWidth(.TextMatrix(fila, 3))
        
               If lTextWidth Mod lColWidth = 0 Then
                  intMultiplier = lTextWidth / lColWidth
              Else
                  intMultiplier = lTextWidth / lColWidth
                  intMultiplier = intMultiplier + 1
              End If
                  
              .RowHeight(fila) = intMultiplier * .RowHeight(fila)
              .WordWrap = True
        
              fila = fila + 1
          Next
        End If
     
        .Visible = True
    End With
    
    If hayNovedades Then
        Me.cmdActualizar.Enabled = True
    Else
        Me.lblUltimasModificaciones.caption = "No hay novedades para descargar."
        Me.lblUltimasModificaciones.ForeColor = &H8000&
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set repositorio = Nothing
    Unload Me
End Sub

Private Sub repositorio_actualizado(Version As Long)
    Dim respuesta As VbMsgBoxResult
    Dim comando As String
    
    Me.frmMensaje.Visible = False
    
    If Version > 0 Then
        ' Avisamos
        MsgBox "El editor del mundo ha sido actualizado a la versión " & Version & ".", vbInformation
        
        ' Alertamos de reiniciar
        respuesta = MsgBox("Es necesario reiniciar el editor para que las actualizaciones se apliquen. ¿Desea reiniciar el editor?", vbExclamation + vbYesNo)
        
        '¿Reiniciamos?
        If respuesta = vbYes Then
            ' Parametros:
            '1) Ejecutable del Editor.
            '2) Tiempo de espera
            '3) El ejecutable que tiene que iniciar cuando finaliza el tiempo de espera
            comando = Chr$(34) & app.Path & "\Updater.exe" & Chr$(34) & " 10 " & Chr$(34) & app.Path & "\EditorTDS.exe" & Chr$(34)
            Call Shell(comando, vbNormalFocus)
            Call frmMain.salirDelEditor(True)
        Else
            Unload Me
        End If

    Else
        Me.cmdActualizar.Enabled = True
        MsgBox "No se ha podido actualizar el editor. Error " & repositorio.ultimoError & ".", vbExclamation, "Error al actualizar"
    End If
End Sub

Private Sub repositorio_Progreso(actual As Single, Maximo As Single)
    Me.pgrProgresoUpdate.max = Maximo
    Me.pgrProgresoUpdate.value = actual
End Sub

Private Sub trmArrancar_Timer()
    Me.trmArrancar.Enabled = False ' Solo se ejecuta una vez
    '******************************
    Dim novedades As Collection
        
    Set novedades = repositorio.Repositorio_ObtenerNovedades
    
    If Not repositorio.ultimoError = "" Then
        Me.lblBuscando.caption = "No fue posible obtener novedades."
        Me.cmdCancelar.Enabled = True
        MsgBox "Ha ocurrido un error y no es posible obtener información del Cerebro de Mono. " & repositorio.ultimoError & ".", vbExclamation
    Else
        Call mostrarNovedades(novedades)
    End If
End Sub
