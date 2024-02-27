VERSION 5.00
Begin VB.Form frmCharInfo 
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   Caption         =   "Información del personaje"
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5370
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   441
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   358
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image backup 
      Height          =   135
      Left            =   5280
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Boton 
      Height          =   420
      Index           =   4
      Left            =   4080
      Top             =   6120
      Width           =   1050
   End
   Begin VB.Image Boton 
      Height          =   390
      Index           =   3
      Left            =   3000
      Top             =   6135
      Width           =   1035
   End
   Begin VB.Image Boton 
      Height          =   405
      Index           =   2
      Left            =   1920
      Top             =   6120
      Width           =   1035
   End
   Begin VB.Image Boton 
      Height          =   315
      Index           =   1
      Left            =   1230
      Top             =   6165
      Width           =   975
   End
   Begin VB.Image Boton 
      Height          =   435
      Index           =   0
      Left            =   240
      Top             =   6090
      Width           =   930
   End
   Begin VB.Label Ciudadanos 
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudadanos asesinados:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   5280
      Width           =   4695
   End
   Begin VB.Label criminales 
      BackStyle       =   0  'Transparent
      Caption         =   "Criminales asesinados:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   5520
      Width           =   4695
   End
   Begin VB.Label reputacion 
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   5760
      Width           =   4695
   End
   Begin VB.Label Solicitudes 
      BackStyle       =   0  'Transparent
      Caption         =   "Solicitudes para ingresar a clanes:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   3360
      Width           =   4695
   End
   Begin VB.Label solicitudesRechazadas 
      BackStyle       =   0  'Transparent
      Caption         =   "Solicitudes rechazadas:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3600
      Width           =   4695
   End
   Begin VB.Label fundo 
      BackStyle       =   0  'Transparent
      Caption         =   "Fundo el clan:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3840
      Width           =   4695
   End
   Begin VB.Label lider 
      BackStyle       =   0  'Transparent
      Caption         =   "Veces fue lider de clan:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   4080
      Width           =   4695
   End
   Begin VB.Label integro 
      BackStyle       =   0  'Transparent
      Caption         =   "Clanes que integro:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4320
      Width           =   4695
   End
   Begin VB.Label faccion 
      BackStyle       =   0  'Transparent
      Caption         =   "Faccion:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   4560
      Width           =   4695
   End
   Begin VB.Label Nombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   4695
   End
   Begin VB.Label Nivel 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label Clase 
      BackStyle       =   0  'Transparent
      Caption         =   "Clase:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label Raza 
      BackStyle       =   0  'Transparent
      Caption         =   "Raza:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label Genero 
      BackStyle       =   0  'Transparent
      Caption         =   "Genero:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   4695
   End
   Begin VB.Label Oro 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   4695
   End
   Begin VB.Label Banco 
      BackStyle       =   0  'Transparent
      Caption         =   "Banco:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   4695
   End
   Begin VB.Label status 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Width           =   4695
   End
End
Attribute VB_Name = "frmCharInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public frmmiembros As Boolean
Public frmsolicitudes As Boolean
Public Selecionado As Byte

Public Sub parseCharInfo(ByVal Rdata As String)
If frmmiembros Then
    Boton(3).Visible = False
    Boton(4).Visible = False
    Boton(1).Visible = True
    Boton(2).Visible = False
Else
    Boton(3).Visible = True
    Boton(4).Visible = True
    Boton(1).Visible = False
    Boton(2).Visible = True
End If

Dim Raza As Byte
Dim Clase As Byte
Dim Genero As Byte
Dim Nivel As Byte
Dim oro As Long
Dim Banco As Long
Dim fundoClan As Byte
Dim echadas As Integer
Dim Solicitudes As Integer
Dim solicitudesRechazadas As Integer
Dim VecesLider As Byte
Dim armada As Byte
Dim caos As Byte
Dim ciudadanosMatados As Integer
Dim criminalesMatados As Integer
Dim clanFundado As String
Dim criminal As Byte
Dim promedio As Long

Raza = Asc(mid(Rdata, 1, 1))
Clase = Asc(mid(Rdata, 2, 1))
Genero = Asc(mid(Rdata, 3, 1))
Nivel = Asc(mid(Rdata, 4, 1))
oro = StringToLong(Rdata, 5)
Banco = StringToLong(Rdata, 9)
promedio = StringToLong(Rdata, 13)
criminal = val(mid(Rdata, 17, 1))
fundoClan = val(mid(Rdata, 18, 1))
echadas = STI(Rdata, 19)
Solicitudes = STI(Rdata, 21)
solicitudesRechazadas = STI(Rdata, 23)
VecesLider = StringToByte(Rdata, 25)
armada = mid(Rdata, 26, 1)
caos = mid(Rdata, 27, 1)

ciudadanosMatados = STI(Rdata, 28)
criminalesMatados = STI(Rdata, 30)

clanFundado = mid(Rdata, 32, InStr(32, Rdata, "¬") - 32)

Nombre.Caption = "Nombre: " & mid(Rdata, InStr(32, Rdata, "¬") + 1)

Me.Raza.Caption = "Raza: " & ListaRazas(Raza)
If Clase <= UBound(ListaClases) Then
    Me.Clase.Caption = "Clase: " & ListaClases(Clase)
Else
    Me.Clase.Caption = "Clase: Staff"
End If
Me.Genero.Caption = "Genero: " & ListaGeneros(Genero)
Me.Nivel.Caption = "Nivel: " & Nivel


Me.oro.Caption = "Oro: " & FormatNumber(oro, 0, vbTrue)
Me.Banco.Caption = "Banco: " & FormatNumber(Banco, 0, vbTrue)

If criminal = 0 Then
    status.Caption = "Ejército Índigo"
Else
    status.Caption = "Ejército Escarla"
End If

Me.Solicitudes.Caption = "Solicitudes para ingresar a clanes: " & Solicitudes
Me.solicitudesRechazadas.Caption = "Solicitudes rechazadas: " & solicitudesRechazadas

If fundoClan = 1 Then
    fundo.Caption = "Fundo el clan: " & clanFundado
Else
    fundo.Caption = "Fundo el clan: Ninguno"
End If

lider.Caption = "Veces fue lider de clan: " & VecesLider
integro.Caption = "Clanes que integro: " & Solicitudes - solicitudesRechazadas

If armada = 1 Then
    faccion.Caption = "Faccion: Ejercito Real"
ElseIf caos = 1 Then
    faccion.Caption = "Faccion: Fuerzas del caos"
Else
    faccion.Caption = "Faccion: Ninguna"
End If

Ciudadanos.Caption = "Índigos asesinados: " & ciudadanosMatados
criminales.Caption = "Escarlatas asesinados: " & criminalesMatados
If criminal = 1 Then
   promedio = promedio * -1
End If
 reputacion.Caption = "Reputacion: " & promedio
Me.Show vbModal, frmMain
End Sub

Private Sub Boton_Click(Index As Integer)
Select Case Index
Case 0 'cerrar
Unload Me
Case 1 'Echar
            EnviarPaquete Paquetes.EcharGuild, right(Nombre, Len(Nombre) - 8)
            frmmiembros = False
            frmsolicitudes = False
            Unload frmGuildLeader
            Unload Me
Case 2
            EnviarPaquete Paquetes.EnviarGuildComen, right(Nombre, Len(Nombre) - 8)
Case 3
            EnviarPaquete Paquetes.RechazarGuild, right(Nombre, Len(Nombre) - 8)
            frmmiembros = False
            frmsolicitudes = False
            Unload frmGuildLeader
            EnviarPaquete Paquetes.GuildInfo
            Unload Me
Case 4
            frmmiembros = False
            frmsolicitudes = False
            EnviarPaquete Paquetes.AceptarGuild, right(Nombre, Len(Nombre) - 8)
            Unload frmGuildLeader
            EnviarPaquete Paquetes.GuildInfo
            Unload Me
End Select
End Sub

Private Sub Boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Selecionado <> Index Then
        Boton(Selecionado).tag = "0"
        Boton(Selecionado).Picture = backup.Picture
    End If
    
    If Boton(Index).tag <> "1" Then
        Boton(Index).tag = "1"
        Selecionado = Index
        backup.Picture = Boton(Selecionado).Picture
        Call DameImagen(Boton(Index), Index + 57)
    End If
End Sub

Private Sub Form_Load()
Call CambiarCursor(frmCharInfo)
Call DameImagen(Boton(0), 52)
Call DameImagen(Boton(1), 53)
Call DameImagen(Boton(2), 54)
Call DameImagen(Boton(3), 55)
Call DameImagen(Boton(4), 56)
backup.Picture = Boton(Selecionado).Picture
DameImagenForm Me, 103
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Boton(Selecionado).tag = "1" Then
    Boton(Selecionado).tag = "0"
    Boton(Selecionado).Picture = backup.Picture
    End If
End Sub

