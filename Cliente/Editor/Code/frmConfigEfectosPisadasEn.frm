VERSION 5.00
Begin VB.Form frmConfigEfectosPisadasEn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Efectos de Pisadas Para"
   ClientHeight    =   5790
   ClientLeft      =   2985
   ClientTop       =   1980
   ClientWidth     =   15810
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfigEfectosPisadasEn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   386
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1054
   StartUpPosition =   1  'CenterOwner
   Begin EditorTDS.TextConListaConBuscador txtEfectoGlobal 
      Height          =   285
      Left            =   3360
      TabIndex        =   11
      Top             =   240
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   503
      CantidadLineasAMostrar=   0
   End
   Begin VB.CommandButton cmdAplicarATodos 
      Caption         =   "Aplicar a todos"
      Height          =   360
      Left            =   6120
      TabIndex        =   10
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame frmSector 
      Caption         =   "Sector"
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
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2415
      Begin VB.ComboBox cmbSectorY 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   330
         Width           =   855
      End
      Begin VB.ComboBox cmbSectorX 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   330
         Width           =   735
      End
      Begin VB.Label lblSectorY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         TabIndex        =   8
         Top             =   330
         Width           =   135
      End
      Begin VB.Label lblSectorX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   350
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   8160
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   8160
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.Frame frmTile 
      Height          =   4935
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   15735
      Begin VB.ComboBox txtEfectoSonido 
         Height          =   315
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblNumeroTile 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Label lblEfectoGeneral 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Efecto de Pisada"
      Height          =   195
      Left            =   0
      TabIndex        =   12
      Top             =   110
      Width           =   1200
   End
End
Attribute VB_Name = "frmConfigEfectosPisadasEn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CANTIDAD_ALTO As Byte
Private CANTIDAD_ANCHO As Byte
Private CANTIDAD_POR_TILE As Byte

Private Const CANTIDAD_MOSTRAR_ALTO = 8
Private Const CANTIDAD_MOSTRAR_ANCHO = 8

Private datos() As Integer


Public Sub iniciar(cantidadAlto As Byte, cantidadAncho As Byte, informacion() As Integer, padre As Form)
    Dim x As Byte
    Dim y As Byte
    Dim numero As Integer
    Dim numeroControl As Integer
    Dim loopControl As Integer
    Dim i As Integer
    Dim elementos() As modEnumerandosDinamicos.eEnumerado
    
    CANTIDAD_ALTO = cantidadAlto
    CANTIDAD_ANCHO = cantidadAncho
    
    ' Guardo los datos iniciales
    datos = informacion
    
    ' Cargos los sectores
    For x = 1 To CInt(CANTIDAD_ANCHO / CANTIDAD_MOSTRAR_ANCHO)
        Me.cmbSectorX.AddItem x
    Next
    
    For y = 1 To CInt(CANTIDAD_ALTO / CANTIDAD_MOSTRAR_ALTO)
        Me.cmbSectorY.AddItem y
    Next
    
    Me.txtEfectoGlobal.CantidadLineasAMostrar = 5

    ' Creamos los combos
    numeroControl = 0
    For y = 0 To CANTIDAD_MOSTRAR_ALTO - 1
        For x = 0 To CANTIDAD_MOSTRAR_ANCHO - 1
        
            
            numero = y * CANTIDAD_ALTO + x
            
            If numeroControl > 0 Then
                load Me.lblNumeroTile(numeroControl)
            
                With Me.lblNumeroTile(numeroControl)
                    .left = .left + Me.lblNumeroTile(0).width * x
                    .top = .top + (Me.lblNumeroTile(0).height + Me.txtEfectoSonido(0).height) * y
                    .Visible = True
                    .caption = ""
                End With
            
                load Me.txtEfectoSonido(numeroControl)
            
                With Me.txtEfectoSonido(numeroControl)
                    .left = .left + Me.txtEfectoSonido(0).width * x
                    .top = .top + (Me.lblNumeroTile(0).height + Me.txtEfectoSonido(0).height) * y
                    .Visible = True
                End With
        
            End If
            
            numeroControl = numeroControl + 1
        Next x
    Next y
    
   
    ' Imagenes que se muestra como ola cuando el personaje se mueve por un pozo de la textura
    elementos = modEnumerandosDinamicos.obtenerEnumeradosDinamicos("EFECTOS_PISADAS")
        
    Call Me.txtEfectoGlobal.limpiarLista
    
    For i = LBound(elementos) To UBound(elementos)
    
        For loopControl = 0 To Me.txtEfectoSonido.count - 1
            Me.txtEfectoSonido(loopControl).AddItem (elementos(i).valor & " - " & elementos(i).nombre)
            Me.txtEfectoSonido(loopControl).itemData(Me.txtEfectoSonido(loopControl).NewIndex) = elementos(i).valor
            
        Next
        
        ' Opcion de seteo vacico
        Call Me.txtEfectoGlobal.addString(elementos(i).valor, elementos(i).valor & " - " & elementos(i).nombre)
            
    Next
    
    'Me.frmTile.width = Me.txtEfectoSonido(numeroControl).left + Me.txtEfectoSonido(numeroControl).width
   ' Me.frmTile.height = Me.txtEfectoSonido(numeroControl).top + Me.txtEfectoSonido(numeroControl).height
    'Me.width = Me.frmTile.width + 200
   ' Me.height = Me.frmTile.height + 200
    Me.Refresh
    
    Me.cmbSectorX.listIndex = 0
    Me.cmbSectorY.listIndex = 0

    Call seleccionarSector(1, 1)
End Sub

Private Sub seleccionarSector(sectorX As Byte, sectorY As Byte)
    Dim loopX As Integer
    Dim loopY As Integer
    
    Dim comienzo As Integer
    Dim numeroCampo As Integer
    
    comienzo = (sectorX - 1) * CANTIDAD_MOSTRAR_ANCHO + ((sectorY - 1) * CANTIDAD_MOSTRAR_ALTO * CANTIDAD_ANCHO)
    
    ' Actualizo
    numeroCampo = 0
    
    For loopY = 0 To CANTIDAD_MOSTRAR_ALTO - 1
    
        For loopX = 0 To CANTIDAD_MOSTRAR_ANCHO - 1
            Me.lblNumeroTile(numeroCampo).caption = comienzo
        
        
            comienzo = comienzo + 1
            numeroCampo = numeroCampo + 1
        Next loopX
        
        comienzo = comienzo + (CANTIDAD_ANCHO - CANTIDAD_MOSTRAR_ANCHO)
    Next loopY
    
End Sub

Private Sub cmbSectorX_Click()
     Call seleccionarSector(Me.cmbSectorX.listIndex + 1, Me.cmbSectorY.listIndex + 1)
End Sub


Private Sub cmbSectorY_Click()
     Call seleccionarSector(Me.cmbSectorX.listIndex + 1, Me.cmbSectorY.listIndex + 1)
End Sub

Private Sub cmdAplicarATodos_Click()
    Dim x As Integer
    Dim y As Integer
    Dim numeroControl As Integer
    Dim idSeleccionado As Integer
    Dim idEnCombo As Integer
    Dim listIndex As Integer
    Dim loopElemento As Integer
    
    idSeleccionado = Me.txtEfectoGlobal.obtenerIDValor
    
    ' Busco el listindex en los combos
    For loopElemento = 0 To Me.txtEfectoSonido(0).ListCount - 1
        If Me.txtEfectoSonido(0).itemData(loopElemento) = idSeleccionado Then
            listIndex = loopElemento
            Exit For
        End If
    Next
    
    ' Selecciono
    For y = 0 To CANTIDAD_MOSTRAR_ALTO - 1
        For x = 0 To CANTIDAD_MOSTRAR_ANCHO - 1
 
            Me.txtEfectoSonido(numeroControl).listIndex = listIndex
            
            numeroControl = numeroControl + 1
        Next x
    Next y
    
End Sub
