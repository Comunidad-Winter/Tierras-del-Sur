VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmCDM 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Cerebro de mono"
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin EditorTDS.asyncDownload asyncDownload1 
      Height          =   0
      Left            =   2160
      TabIndex        =   11
      Top             =   240
      Width           =   0
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin VB.TextBox DataArrival 
      Height          =   285
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   4200
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5880
      Top             =   0
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5400
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   4920
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Esconder"
      Height          =   255
      Left            =   6600
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command 
      Height          =   255
      Left            =   6360
      TabIndex        =   7
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picLista 
      BackColor       =   &H00000040&
      Height          =   4935
      Left            =   120
      ScaleHeight     =   4875
      ScaleWidth      =   7635
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CommandButton cmdEnviarCDM 
         Caption         =   "Enviar"
         Height          =   375
         Left            =   5520
         TabIndex        =   10
         Top             =   4320
         Width           =   1935
      End
      Begin VB.ListBox list_cdm 
         Appearance      =   0  'Flat
         Height          =   4080
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   120
         Width           =   7395
      End
   End
   Begin VB.PictureBox Picture 
      BackColor       =   &H00000000&
      Height          =   4935
      Left            =   120
      ScaleHeight     =   4875
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   480
      Width           =   7695
      Begin VB.TextBox txtOutput 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0072899A&
         Height          =   4095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   600
         Width           =   7395
      End
      Begin MSComctlLib.ProgressBar pb 
         Height          =   135
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar pbp 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Cerebro de mono"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmCDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public wsconnected As Boolean

Public TamañoTotal As Long
Public TamañoActual As Long
Public FilenameActual As String

Private Sub cmdEnviarCDM_Click()
CDM_EnviarTodo_real
Me.picLista.visible = False
End Sub

Private Sub cmdHide_Click()
Me.visible = False
Me.picLista.visible = False
If prgRun = False Or frmMain.visible = False Then
DoEvents
End
End If
End Sub

Private Sub Command_Click()
CDM_Update
End Sub

Private Sub Form_Load()
   ' If webb Is Nothing Then
       ' Set webb = New clsWEBA
     '   DoEvents
      '  If webb.Initialize(Me.Winsock) = False Then
          '  MsgBox "Error, no se pudo crear Winsock."
     '       CDMOutput "WINSOCK FAIL!"
      '  Else
         '   CDMOutput "Winsock OK!"
         '   Timer.Enabled = True
            'Timer4.Enabled = True
     '   End If
 '   End If
End Sub

Private Sub Inet_StateChanged(ByVal State As Integer)
Select Case State
    Case icError
        CDMDownloading = False
    Case icResponseCompleted
        Dim vtData As Variant
        Dim tempArray() As Byte
        Dim FileSize As Long
        Dim tma As Long
        'pro.Max = 101
        
        FileSize = val(Inet.GetHeader("Content-length"))
        
        If FileSize Then
            pb.max = FileSize
        Else
            If CDM_Current_File.size Then
                FileSize = CDM_Current_File.size
                pb.max = FileSize
            End If
        End If
        Debug.Print "Descargando"; CDM_Current_File.FileName
        
        Dim handle As Integer
        handle = FreeFile()
        
        tma = 0
        DoEvents
        
        FilenameActual = CDM_Current_File.FileName
        TamañoTotal = pb.max
        
        'If vWindowCDM Is Nothing Then
        '    Set vWindowCDM = New vWCDM
        '    GUI_Load vWindowCDM
        'End If
        'vWindowCDM.Show
            GUI_SetFocus vWindowCDM

        Open CDM_TMP_PATH & CDM_Current_File.FileName For Binary Access Write As handle
            vtData = 1
            Do While Not Len(vtData) = 0
                vtData = Inet.GetChunk(1024, icByteArray)
                DoEvents
                tempArray = vtData
                Put handle, , tempArray
                
                tma = tma + UBound(tempArray)
                If FileSize Then pb.Value = tma
                
                TamañoActual = tma
                
                'pro.value = tma / FileSize * 100
                'Label1.Caption = Str$(Round(tma / 1024)) & "KB descargados"
                DoEvents
            Loop
        Close handle
        
        handle = FreeFile()
        With CDM_Current_File
        
        Open CDM_TMP_PATH & .FileName For Binary Access Read As handle
            .data_lenght = EOF(handle)
            ReDim .Data(.data_lenght)
            Get handle, , .Data
        Close handle
        
        End With
        
        'If Not vWindowCDM Is Nothing Then
        '    vWindowCDM.Hide
        'End If
        
        
        CDMDownloading = False
End Select
End Sub

Private Sub list_cdm_Click()
CDM_LeerListbox list_cdm
End Sub

Private Sub picLista_Click()

End Sub

Private Sub Timer_Timer()
    webb.TryRequest
End Sub

Private Sub Timer4_Timer()
    CDM_Update
    Timer4.Enabled = False
End Sub

Private Sub webb_RecibeDatosWeb(datos As String, raw As Boolean)
'Marce On error resume next
Dim splite() As String
If Len(datos) Then
    If Strings.Left(datos, 1) = "$" Then
        Script_Downloaded datos
    ElseIf Strings.Left(datos, 1) = "*" Then
        Debug.Print val(mid(datos, 2))
        CDM_DataSent val(mid(datos, 2))
    ElseIf Strings.Left(datos, 1) = "#" Then
        CDM_Enviar val(Right(datos, Len(datos) - 1))
    ElseIf Strings.Left(datos, 2) = "][" Then
        MsgBox "Contraseña incorrecta!!!!", vbCritical, "Cerebro de mono"
        CDM_UserPrivs = 0
        CDM_Password = ""
        CDM_PasswordMD5 = ""
        CDM_UserID = 0
    ElseIf Strings.Left(datos, 2) = "[[" Then
        MsgBox "Usuario incorrecto!!!!", vbCritical, "Cerebro de mono"
        CDM_Username = ""
        CDM_UserPrivs = 0
        CDM_Password = ""
        CDM_PasswordMD5 = ""
        CDM_UserID = 0
    ElseIf Strings.Left(datos, 1) = "{" Then
        splite = Split(Right(datos, Len(datos) - 1), "|")
        CDM_Username = splite(0)
        CDM_UserID = val(splite(1))
        CDM_UserPrivs = val(splite(2))
        CDM_UserSession = val(splite(3))
        'MsgBox "¡Bienvenido al Cerebro de Mono " & CDM_Username & "!"
        ActualizarPuedoMapas
        frmLogin.Hide
    ElseIf Strings.Left(datos, 2) = "_[" Then
        GUI_Alert "La sesion del Cerebro de mono se perdió, esto puede suceder POR COMPARTIR EL USUARIO o abrir varios editores al mismo tiempo. Ahora vamos a intentar recuperarla..."
        CDM_Login CDM_Username, CDM_Password
    Else
        CDMOutput datos
        Debug.Print datos
    End If
    CharList(UserCharIndex).nombre = CDM_Username
    
    frmLogin.Caption = "Cerebro de mono"
End If
End Sub

Private Sub Winsock_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    pb.Value = pb.min
End Sub

Private Sub Winsock_SendComplete()
    pb.Value = pb.max
    If Not vWindowCDM Is Nothing Then
        vWindowCDM.Hide
    '    Set vWindowCDM = Nothing
    End If
End Sub

Private Sub Winsock_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    pb.max = bytesSent + bytesRemaining + 1
    pb.Value = bytesSent

    TamañoTotal = bytesSent + bytesRemaining
    TamañoActual = bytesSent
End Sub
