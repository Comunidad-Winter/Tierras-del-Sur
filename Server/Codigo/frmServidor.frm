VERSION 5.00
Begin VB.Form frmServidor 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Servidor"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   409
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   323
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command26 
      Caption         =   "Reset Listen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   20
      Top             =   5760
      Width           =   1455
   End
   Begin VB.PictureBox picFuera 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   120
      ScaleHeight     =   271
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   6
      Top             =   120
      Width           =   4590
      Begin VB.VScrollBar VS1 
         Height          =   4095
         LargeChange     =   50
         Left            =   4320
         SmallChange     =   17
         TabIndex        =   19
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picCont 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   4095
         Left            =   0
         ScaleHeight     =   273
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   289
         TabIndex        =   7
         Top             =   0
         Width           =   4334
         Begin VB.CommandButton Command7 
            Caption         =   "Recargar consultas anticheat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   25
            Top             =   3480
            Width           =   4050
         End
         Begin VB.CommandButton adminEventos 
            Caption         =   "Admin Eventos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   2520
            Width           =   4095
         End
         Begin VB.CommandButton cmdDesactivarDenunciar 
            Caption         =   "Desactivar denunciar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   23
            Top             =   3240
            Width           =   4095
         End
         Begin VB.CommandButton Command27 
            Caption         =   "Debug UserList"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   3720
            Width           =   4095
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Pausar el servidor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   3000
            Width           =   4095
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Actualizar npcs.dat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   2760
            Width           =   4095
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Reload Server.ini"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   2280
            Width           =   4095
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Update MOTD"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   2040
            Width           =   4095
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Debug listening socket"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1800
            Width           =   4095
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Debug Npcs"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1560
            Width           =   4095
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Stats de los slots"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1320
            Width           =   4095
         End
         Begin VB.CommandButton cmdRecargarAPIManager 
            Caption         =   "Recargar API Manager"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1080
            Width           =   4095
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Actualizar hechizos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   4095
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Reiniciar"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   4095
         End
         Begin VB.CommandButton Command6 
            Caption         =   "ReSpawn Guardias en posiciones originales"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   4095
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Actualizar objetos.dat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   4095
         End
      End
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Boton Magico para apagar server"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   5280
      Width           =   4095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cargar BackUp del mundo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   4920
      Width           =   4095
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Guardar todos los personajes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4680
      Width           =   4095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Hacer un Backup del mundo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   4440
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   3480
      TabIndex        =   0
      Top             =   5760
      Width           =   945
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Reset sockets"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      Height          =   1335
      Left            =   120
      Top             =   4320
      Width           =   4335
   End
End
Attribute VB_Name = "frmServidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub adminEventos_Click()
frmAdminEventos.Show
End Sub

Private Sub cmdDesactivarDenunciar_Click()
    denunciarActivado = Not denunciarActivado
End Sub

Private Sub cmdRecargarAPIManager_Click()
    'Inicio el Manager
    If Not API_Manager.iniciarManager Then
        LogError "No se pudo conectar con el Manager"
    End If
End Sub

Private Sub Command1_Click()
Call LoadOBJData
End Sub


Private Sub Command11_Click()
frmConID.Show
End Sub

Private Sub Command12_Click()
frmDebugNpc.Show
End Sub

Private Sub Command13_Click()
frmDebugSocket.Visible = True
End Sub

Private Sub Command14_Click()
Call LoadMotd
End Sub

Private Sub Command16_Click()
Call LoadSini
End Sub

Private Sub Command17_Click()
Call CargaNpcsDat
End Sub

Private Sub Command18_Click()
Me.MousePointer = 11
Call GuardarUsuarios(1)
Me.MousePointer = 0
MsgBox "Grabado de personajes OK!"
End Sub


Private Sub Command2_Click()
frmServidor.Visible = False
End Sub

Private Sub Command20_Click()
If MsgBox("Esta seguro que desea reiniciar los sockets ? Se cerrarán todas las conexiones activas.", vbYesNo, "Reiniciar Sockets") = vbYes Then
    Call WSApiReiniciarSockets
End If
End Sub

'Barrin 29/9/03
Private Sub Command21_Click()
If EnPausa = False Then
    EnPausa = True
    EnviarPaquete Paquetes.Pausa, "", 0, ToAll
    Command21.Caption = "Reanudar el servidor"
Else
    EnPausa = False
    EnviarPaquete Paquetes.Pausa, "", 0, ToAll
    Command21.Caption = "Pausar el servidor"
End If
End Sub

Private Sub Command23_Click()
If MsgBox("Esta seguro que desea hacer WorldSave, guardar pjs y cerrar ?", vbYesNo, "Apagar Magicamente") = vbYes Then
    Me.MousePointer = 11
    FrmStat.Show

    Call cerrarServidorGracefull
    
    End
End If
End Sub

Public Sub cerrarServidorGracefull()

    'WorldSave
    Call DoBackUp
    'commit experiencia
    Call mdParty.ActualizaExperiencias
    'Guardar Pjs
    Call GuardarUsuarios
    ' Actuaizo el numero de usuaros
    Call Admin.actualizarOnlinesDB(True)
    'Chauuu
    Unload frmMain

End Sub

Private Sub Command26_Click()
'Cierra el socket de escucha
If SockListen >= 0 Then Call apiclosesocket(SockListen)
'Inicia el socket de escucha
SockListen = ListenForConnect(Puerto, hWndMsg, "")
End Sub

Private Sub Command27_Click()
frmUserList.Show
End Sub
Private Sub Command3_Click()
If MsgBox("¡¡Atencion!! Si reinicia el servidor puede provocar la perdida de datos de los usarios. ¿Desea reiniciar el servidor de todas maneras?", vbYesNo) = vbYes Then
    Me.Visible = False
    Call Restart
End If
End Sub

Private Sub Command4_Click()

    Me.MousePointer = 11
    FrmStat.Show
    
    Call DoBackUp
    
    Me.MousePointer = 0
    MsgBox "WORLDSAVE OK!!"
End Sub

Private Sub Command5_Click()
'Se asegura de que los sockets estan cerrados e ignora cualquier err
If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

FrmStat.Show

If FileExist(App.Path & "\logs\errores.log", vbNormal) Then Kill App.Path & "\logs\errores.log"
If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\Connect.log"
If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
If FileExist(App.Path & "\logs\Resurrecciones.log", vbNormal) Then Kill App.Path & "\logs\Resurrecciones.log"
If FileExist(App.Path & "\logs\Teleports.Log", vbNormal) Then Kill App.Path & "\logs\Teleports.Log"

Call apiclosesocket(SockListen)

Dim loopC As Integer

For loopC = 1 To MaxUsers
    Call CloseSocket(loopC)
Next

LastUser = 0
NumUsers = 0
NumUsersPremium = 0

ReDim NpcList(1 To MAXNPCS) As npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer

Call LoadSini
Call LoadOBJData

SockListen = ListenForConnect(Puerto, hWndMsg, "")

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

End Sub

Private Sub Command6_Click()
Call ReSpawnOrigPosNpcs
End Sub

Private Sub Command7_Click()

If Anticheat_MemCheck.cargarChequeosData Then
    Call MsgBox("Chequeos recargados exitosamente.", vbInformation)
Else
    Call MsgBox("Error al cargar los chequeos.", vbExclamation)
End If

End Sub

Private Sub Command8_Click()
Call CargarHechizos
End Sub

Private Sub Form_Deactivate()
frmServidor.Visible = False
End Sub

Private Sub Form_Load()

Command20.Visible = True
Command26.Visible = True

VS1.min = 0
If picCont.Height > picFuera.ScaleHeight Then
    VS1.max = picCont.Height - picFuera.ScaleHeight
Else
    VS1.max = 0
End If
picCont.Top = -VS1.value
End Sub

Private Sub VS1_Change()
picCont.Top = -VS1.value
End Sub

Private Sub VS1_Scroll()
picCont.Top = -VS1.value
End Sub
