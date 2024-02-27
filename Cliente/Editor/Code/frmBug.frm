VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmBug 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tierras del Sur - Editor del Mundo"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5700
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBug.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmReportar_Manual 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CheckBox chkAdjuntarInfoPC_Manual 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Adjuntar info técnica de mi sistema (importante cuando las cosas no se ven como deberian)"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   5280
         Width           =   5175
      End
      Begin VB.CommandButton cmdCancelar_Manual 
         Caption         =   "Cancelar"
         Height          =   360
         Left            =   1920
         TabIndex        =   13
         Top             =   5880
         Width           =   1575
      End
      Begin VB.CommandButton cmdEnviarReporte_Manual 
         Caption         =   "Enviar reporte"
         Height          =   360
         Left            =   3720
         TabIndex        =   12
         Top             =   5880
         Width           =   1590
      End
      Begin VB.TextBox txtExplicacionDetallada 
         Appearance      =   0  'Flat
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   3480
         Width           =   5295
      End
      Begin VB.TextBox txtQuePaso 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   5295
      End
      Begin VB.TextBox txtQueActividad 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmBug.frx":1CCA
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Width           =   5280
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "¿Qué pasó?. (Ej: No se hizo la acción querida, se tildó el programa)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   5295
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTituloReportar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reportar error"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   2970
      End
      Begin VB.Label lblQueActividad 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "¿Qué actividad estabas realizando?. (Ej: agregando un objeto, usando el menú Configurar Mapa)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   5295
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame frmReportar_Copia 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6225
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton cmdCerrarEditor_Copia 
         Caption         =   "Cerrar Editor"
         Height          =   360
         Left            =   0
         TabIndex        =   25
         Top             =   5760
         Width           =   1575
      End
      Begin RichTextLib.RichTextBox rtbDetallesError 
         Height          =   5055
         Left            =   0
         TabIndex        =   1
         Top             =   480
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   8916
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"frmBug.frx":1D8D
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   360
         Left            =   3480
         TabIndex        =   2
         Top             =   5760
         Width           =   1935
      End
      Begin VB.Label lblOrden 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copia el texto que aquí debajo se encuentra y publicalo en el foro privado. Gracias."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   5415
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame frmReportar_Automatico 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton cmdCerrarEditor_Automatico 
         Caption         =   "Cerrar Editor"
         Height          =   360
         Left            =   120
         TabIndex        =   26
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CheckBox chkAdjuntarInfoPC_Automatico 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Adjuntar info técnica de mi sistema (importante cuando las cosas no se ven como deberian)"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   19
         Top             =   5280
         Width           =   5175
      End
      Begin VB.TextBox txtQueHiciste 
         Appearance      =   0  'Flat
         Height          =   2175
         Left            =   0
         TabIndex        =   18
         Top             =   2160
         Width           =   5415
      End
      Begin VB.CommandButton cmdCancelar_Automatico 
         Caption         =   "Cancelar"
         Height          =   360
         Left            =   1920
         TabIndex        =   17
         Top             =   5880
         Width           =   1710
      End
      Begin VB.CommandButton cmdEnviarReporte_Automatico 
         Caption         =   "Enviar reporte"
         Height          =   360
         Left            =   3720
         TabIndex        =   16
         Top             =   5880
         Width           =   1710
      End
      Begin VB.Label lblUps 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "¡Ups! Hemos fallado."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   4185
      End
      Begin VB.Label lblDescripcion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Se ha producido un error inesperado. Sí podés, te recomendamos que guardes el mapa y re-inicies el editor."
         Height          =   390
         Left            =   0
         TabIndex        =   23
         Top             =   600
         Width           =   5610
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPedido 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmBug.frx":1E08
         Height          =   435
         Left            =   0
         TabIndex        =   22
         Top             =   1080
         Width           =   5610
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblErrorContinuacion2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "¿Qué fue lo último que hiciste antes de que el Editor falle?."
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
         Left            =   0
         TabIndex        =   21
         Top             =   1800
         Width           =   4890
      End
      Begin VB.Label lblErrorContinuacion3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Te pedimos que intentes hacer lo mismo para ver si surge nuevamente el error. Si es así, encontraste la causa que produce el bug."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   0
         TabIndex        =   20
         Top             =   4440
         Width           =   5445
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmBug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private enviando As Boolean

'Informacion del error
Private metodo As String
Private error As String

Private infopc As String
Private infoPCCargada As Boolean
Private WithEvents consolaDeWindows As clsConsolaWindows
Attribute consolaDeWindows.VB_VarHelpID = -1


Private Sub cmdCancelar_Automatico_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Manual_Click()
    Unload Me
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Public Sub crearEnBlanco()
    metodo = ""
    error = ""
    Me.frmReportar_Manual.visible = True
    Me.frmReportar_Copia.visible = False
End Sub
Public Sub crear(metodoError As String, errorDescripcion As String)
    Me.frmReportar_Copia.visible = False
    Me.frmReportar_Manual.visible = False
    Me.rtbDetallesError.Text = ""
    
    metodo = metodoError
    error = errorDescripcion
End Sub

Private Sub cmdCerrarEditor_Automatico_Click()
    End
End Sub

Private Sub cmdCerrarEditor_Copia_Click()
    End
End Sub

Private Sub cmdEnviarReporte_Automatico_Click()
    Dim infopc As String
        
    'Deshabilitamos el formulario
    Call modPosicionarFormulario.setEnabledHijos(False, Me.frmReportar_Automatico, Me)
    
    'Chequeamos si hay que obtener la info de la computadora
    If Me.chkAdjuntarInfoPC_Automatico.value = 1 Then
        Me.cmdEnviarReporte_Automatico.caption = "Obteniendo datos..."
        infopc = ObtenerDatosPC
    Else
        infopc = ""
    End If
    
    Me.cmdEnviarReporte_Automatico.caption = "Enviando..."

    If CDM.cerebro.RepotarBug(error, metodo, Trim$(Me.txtQueHiciste), infopc) = False Then
        MsgBox "Uh, parece que realmente hay problemas. No se ha podido enviar el reporte. Copiá la información que te mostraremos ahora y avisanos a través del foro.", vbExclamation
        
        'Ponemos en rl tich el error y tambien el error del CDM
        Me.rtbDetallesError.Text = metodo & vbCrLf & vbCrLf & error & vbCrLf & vbCrLf & CDM.cerebro.ultimoError
        Me.frmReportar_Copia.visible = True
    Else
        'Se pudo enviar correctamente. Confirmamos y salimos
        MsgBox "El reporte ha sido enviado. Gracias!", vbInformation, Me.caption
        'Cerramos
        Unload Me
    End If
    
End Sub

Private Sub cmdEnviarReporte_Manual_Click()
    Dim infopc As String
    
    'Chequeamos que haya puesto todo
    If Len(Trim$(Me.txtQuePaso)) = 0 Or _
        Len(Trim$(Me.txtQueActividad)) = 0 Or _
        Len(Trim$(Me.txtExplicacionDetallada)) = 0 Then
         MsgBox "Completa todosl los campos por favor.", vbExclamation, Me.caption
        Exit Sub
    End If
    
    'Deshabilitamos el formulario para que no toque minetras seprocesa
    Call modPosicionarFormulario.setEnabledHijos(False, Me.frmReportar_Manual, Me)
    
    '¿Necesitamos tomar los specs?
    If Me.chkAdjuntarInfoPC_Manual.value = 1 Then
        Me.cmdEnviarReporte_Manual.caption = "Obteniendo datos..."
        infopc = ObtenerDatosPC
    Else
        infopc = ""
    End If

    Me.cmdEnviarReporte_Manual.caption = "Enviando..."
    
    'Enviamos la data
    If CDM.cerebro.RepotarBug(Trim$(Me.txtQuePaso), Trim$(Me.txtQueActividad), Trim$(Me.txtExplicacionDetallada), infopc) = False Then
        MsgBox "No se ha podido enviar el reporte. Copiá la información y avisanos a través del foro.", vbExclamation, Me.caption
        
        'Restablecemos el formulario
        Me.cmdEnviarReporte_Manual.caption = "Enviar Reporte"
        Call modPosicionarFormulario.setEnabledHijos(True, Me.frmReportar_Manual, Me)
    Else
        'Se pudo enviar correctamente. Confirmamos y salimos
        MsgBox "El reporte ha sido enviado. Gracias!", vbInformation, Me.caption
        Unload Me
    End If
End Sub

Private Function ObtenerDatosPC() As String
    Dim archivo As String
    Dim break As Boolean
    
    Set consolaDeWindows = New clsConsolaWindows
        
    archivo = app.Path & "\spects.txt"
    
    'DXDiag funciona de manera asincronica. La consola retorna automaticamente
    'ni bien se ejecuta el comando
    If Not FileExist(archivo, vbArchive) Then
        Call consolaDeWindows.RunCommand(frmMain.hwnd, "dxdiag /t " & archivo)
                    
        break = False
        
        Do While Not FileExist(archivo, vbArchive) Or Not break
            If FileExist(archivo, vbArchive) Then
                If FileLen(archivo) > 0 Then
                    break = True
                End If
            End If
            DoEvents
        Loop
        
        Call consolaDeWindows.parar
    End If
    
    ObtenerDatosPC = FileText(archivo)

End Function


Private Function FileText(FileName As String) As String
    Dim handle As Integer
    handle = FreeFile
    Open FileName$ For Input Access Read Lock Write As #handle
    FileText = Input$(LOF(handle), handle)
    Close #handle
End Function

