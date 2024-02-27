VERSION 5.00
Begin VB.Form frmConfigurarTeclas 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5010
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
   ScaleHeight     =   482
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   334
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtValorActual 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "A"
      Top             =   480
      Width           =   975
   End
   Begin VB.Image botons 
      Height          =   315
      Index           =   2
      Left            =   3330
      Top             =   6525
      Width           =   1215
   End
   Begin VB.Image botons 
      Height          =   315
      Index           =   1
      Left            =   1920
      Top             =   6525
      Width           =   1215
   End
   Begin VB.Image botons 
      Height          =   315
      Index           =   0
      Left            =   540
      Top             =   6525
      Width           =   1215
   End
   Begin VB.Label lblError 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "La tecla ya se encuentra utilizada en la acción dasdsadasdasdad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   4875
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPresione 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Presione la tecla a la cual le quiere asignar la acción"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   5280
      Visible         =   0   'False
      Width           =   4365
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion de la tecla"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1890
   End
End
Attribute VB_Name = "frmConfigurarTeclas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type ConfigTecla
    Nombre As String
    tecla As Integer
    default As Integer
End Type

Private Type TeclaNombre
    tecla As Integer
    Nombre As String
End Type

Private configTeclas() As ConfigTecla
Private teclasEspeciales() As TeclaNombre

Private Selecionado As Byte


Private Sub botons_Click(Index As Integer)
    Select Case Index
    
        Case 0
             Unload Me
        Case 1
        
            Dim loopTecla As Integer
        
            For loopTecla = 0 To UBound(configTeclas)
                txtValorActual(loopTecla).text = getKeyName(configTeclas(loopTecla).default)
                configTeclas(loopTecla).tecla = configTeclas(loopTecla).default
            Next loopTecla
        
        Case 2
        
            Call guardarConfiguracion
            Unload Me
    
    End Select
End Sub


Private Sub botons_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Selecionado <> Index Then
        botons(Selecionado).tag = "0"
        botons(Selecionado).Picture = Nothing
    End If
    
    If botons(Index).tag <> "1" Then
        botons(Index).tag = "1"
        Selecionado = Index
        Call DameImagen(botons(Index), Index + 167)
    End If

End Sub

Private Sub Form_Load()

    Call DameImagenForm(Me, 166)

    Call cargarConfiguracion
    
    Call visualizarConfiguracion
End Sub

Private Sub cargarConfiguracion()
    ReDim configTeclas(0 To 16) As ConfigTecla
    
    configTeclas(0).Nombre = "Caminar hacia arriba"
    configTeclas(0).tecla = vbKeyNorte
    configTeclas(0).default = vbKeyUp
    
    configTeclas(1).Nombre = "Caminar hacia abajo"
    configTeclas(1).tecla = vbKeySur
    configTeclas(1).default = vbKeyDown
    
    configTeclas(2).Nombre = "Caminar hacia la izquierda"
    configTeclas(2).tecla = vbKeyOeste
    configTeclas(2).default = vbKeyLeft
    
    configTeclas(3).Nombre = "Caminar hacia la derecha"
    configTeclas(3).tecla = vbKeyEste
    configTeclas(3).default = vbKeyRight
    
    configTeclas(4).Nombre = "Pegarle a tus enemigos"
    configTeclas(4).tecla = vbKeyPegar
    configTeclas(4).default = vbKeyControl
    
    configTeclas(5).Nombre = "Activar/desactivar la musica"
    configTeclas(5).tecla = vbKeyMusica
    configTeclas(5).default = vbKeyM
    
    configTeclas(6).Nombre = "Agarrar objetos del suelo"
    configTeclas(6).tecla = vbKeyAgarrarItem
    configTeclas(6).default = vbKeyA
    
    configTeclas(7).Nombre = "Tirar objetos al suelo"
    configTeclas(7).tecla = vbKeyTirarItem
    configTeclas(7).default = vbKeyT
    
    configTeclas(8).Nombre = "Modo combate"
    configTeclas(8).tecla = vbKeyModoCombate
    configTeclas(8).default = vbKeyC
    
    configTeclas(9).Nombre = "Equipar objeto del inventario"
    configTeclas(9).tecla = vbKeyEquiparItem
    configTeclas(9).default = vbKeyE
    
    configTeclas(10).Nombre = "Mostrar/ocultar nombres"
    configTeclas(10).tecla = vbKeyMostrarNombre
    configTeclas(10).default = vbKeyN
    
    configTeclas(11).Nombre = "Domar criaturas"
    configTeclas(11).tecla = vbKeyDomar
    configTeclas(11).default = vbKeyD
    
    configTeclas(12).Nombre = "Ocultar mi personaje"
    configTeclas(12).tecla = vbKeyOcultar
    configTeclas(12).default = vbKeyO
    
    configTeclas(13).Nombre = "Usar objeto del inventario"
    configTeclas(13).tecla = vbKeyUsar
    configTeclas(13).default = vbKeyU
    
    configTeclas(14).Nombre = "DesLaguear personaje"
    configTeclas(14).tecla = vbKeyLag
    configTeclas(14).default = vbKeyL
    
    configTeclas(15).Nombre = "Consola de clanes"
    configTeclas(15).tecla = vbKeyConsolaClanes
    configTeclas(15).default = vbKeyZ
    
    configTeclas(16).Nombre = "Meditar"
    configTeclas(16).tecla = vbKeyMeditar
    configTeclas(16).default = vbKeyF6
    
    ReDim teclasEspeciales(0 To 24) As TeclaNombre
    
    teclasEspeciales(0).Nombre = "Control"
    teclasEspeciales(0).tecla = vbKeyControl
    
    teclasEspeciales(1).Nombre = "Arriba"
    teclasEspeciales(1).tecla = vbKeyUp
    
    teclasEspeciales(2).Nombre = "Abajo"
    teclasEspeciales(2).tecla = vbKeyDown
    
    teclasEspeciales(3).Nombre = "Izquierda"
    teclasEspeciales(3).tecla = vbKeyLeft
    
    teclasEspeciales(4).Nombre = "Derecha"
    teclasEspeciales(4).tecla = vbKeyRight
    
    teclasEspeciales(5).Nombre = "F6"
    teclasEspeciales(5).tecla = vbKeyF6
    
    teclasEspeciales(6).Nombre = "Espacio"
    teclasEspeciales(6).tecla = vbKeySpace
    
    teclasEspeciales(7).Nombre = "Shift"
    teclasEspeciales(7).tecla = vbKeyShift
    
    teclasEspeciales(8).Nombre = "Insert"
    teclasEspeciales(8).tecla = vbKeyInsert
    
    teclasEspeciales(9).Nombre = "Delete"
    teclasEspeciales(9).tecla = vbKeyDelete
    
    teclasEspeciales(10).Nombre = "Menú"
    teclasEspeciales(10).tecla = vbKeyMenu
    
    teclasEspeciales(11).Nombre = "Page Up"
    teclasEspeciales(11).tecla = vbKeyPageUp
    
    teclasEspeciales(12).Nombre = "Page Down"
    teclasEspeciales(12).tecla = vbKeyPageDown
    
    teclasEspeciales(13).Nombre = "Fin"
    teclasEspeciales(13).tecla = vbKeyEnd
    
    teclasEspeciales(14).Nombre = "Inicio"
    teclasEspeciales(14).tecla = vbKeyHome
    
    Dim loopNumPad As Byte
    
    For loopNumPad = 0 To 9
        teclasEspeciales(15 + loopNumPad).Nombre = "NumPad " & loopNumPad
        teclasEspeciales(15 + loopNumPad).tecla = vbKeyNumpad0 + loopNumPad
    Next loopNumPad
    
    
End Sub

Private Function getKeyName(KeyCode As Integer) As String
    Dim loopTecla As Byte
    
    If KeyCode > 1000 Then
        getKeyName = "Mouse " & KeyCode - 1000
        Exit Function
    End If

    For loopTecla = 0 To UBound(teclasEspeciales)
        If teclasEspeciales(loopTecla).tecla = KeyCode Then
            getKeyName = teclasEspeciales(loopTecla).Nombre
            Exit Function
        End If
    Next loopTecla
    
    getKeyName = Chr$(KeyCode)
End Function


Private Sub guardarConfiguracion()
    vbKeyNorte = configTeclas(0).tecla
    vbKeySur = configTeclas(1).tecla
    vbKeyOeste = configTeclas(2).tecla
    vbKeyEste = configTeclas(3).tecla
    vbKeyPegar = configTeclas(4).tecla
    vbKeyMusica = configTeclas(5).tecla
    vbKeyAgarrarItem = configTeclas(6).tecla
    vbKeyTirarItem = configTeclas(7).tecla
    vbKeyModoCombate = configTeclas(8).tecla
    vbKeyEquiparItem = configTeclas(9).tecla
    vbKeyMostrarNombre = configTeclas(10).tecla
    vbKeyDomar = configTeclas(11).tecla
    vbKeyOcultar = configTeclas(12).tecla
    vbKeyUsar = configTeclas(13).tecla
    vbKeyLag = configTeclas(14).tecla
    vbKeyConsolaClanes = configTeclas(15).tecla
    vbKeyMeditar = configTeclas(16).tecla
End Sub


Private Sub visualizarConfiguracion()
    Dim loopTecla As Integer
    
    For loopTecla = 0 To UBound(configTeclas)
        If Not loopTecla = 0 Then
            Load Me.lblNombre(loopTecla)
            Load Me.txtValorActual(loopTecla)
            
            Me.lblNombre(loopTecla).top = Me.lblNombre(loopTecla - 1).top + 22
            Me.txtValorActual(loopTecla).top = Me.txtValorActual(loopTecla - 1).top + 22
            
            Me.lblNombre(loopTecla).Visible = True
            Me.txtValorActual(loopTecla).Visible = True
        End If
        
        Me.lblNombre(loopTecla).Caption = configTeclas(loopTecla).Nombre
        Me.txtValorActual(loopTecla).text = getKeyName(configTeclas(loopTecla).tecla)
        
    Next loopTecla
    
    
    Me.lblPresione.top = Me.lblNombre(loopTecla - 1).top + 25
    Me.lblError.top = Me.lblPresione.top
End Sub

Private Function isTeclaValida(KeyCode As Integer)

    If KeyCode >= vbKeyF1 And KeyCode <= vbKeyF12 Then
        isTeclaValida = False
        Exit Function
    End If
    
    isTeclaValida = True
   
End Function

Private Sub Form_Unload(Cancel As Integer)
    frmOpciones.Enabled = True
End Sub

Private Sub lblNombre_Click(Index As Integer)
    Me.txtValorActual(Index).SetFocus
End Sub

Private Sub txtValorActual_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    If Not isTeclaValida(KeyCode) Then
        Exit Sub
    End If
    
    Call setKeyCode(Index, KeyCode)
End Sub

Private Sub setKeyCode(Index As Integer, KeyCode As Integer)
    Dim config As ConfigTecla
    
    config = GetConfigForKey(KeyCode)
    
    Me.lblPresione.Visible = False
    
    If config.tecla > 0 Then
        Me.lblError.Visible = True
        Me.lblError.Caption = "La tecla '" & getKeyName(KeyCode) & "' se utilizá en " & config.Nombre & "."
        Exit Sub
    End If
    
    Me.lblError.Visible = False
    
    txtValorActual(Index).text = getKeyName(KeyCode)
    configTeclas(Index).tecla = KeyCode
End Sub
Private Sub txtValorActual_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <= 2 Then Exit Sub
    
    Call setKeyCode(Index, Button + 1000)
End Sub

Private Sub txtValorActual_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblPresione.Visible = True
    Me.lblError.Visible = False
    
    Debug.Print Button
End Sub

Private Function GetConfigForKey(KeyCode As Integer) As ConfigTecla
    Dim loopTecla As Integer
    
    For loopTecla = 0 To UBound(configTeclas)
        If configTeclas(loopTecla).tecla = KeyCode Then
            GetConfigForKey = configTeclas(loopTecla)
            Exit Function
        End If
    Next loopTecla
End Function

