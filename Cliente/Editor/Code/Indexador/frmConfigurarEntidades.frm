VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfigurarEntidades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Entidades"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9285
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfigurarEntidades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   9285
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer trmActualizarEstadoEntidad 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   5400
   End
   Begin VB.Frame frmVidaEntidad 
      Height          =   600
      Left            =   3600
      TabIndex        =   36
      Top             =   4680
      Visible         =   0   'False
      Width           =   2535
      Begin VB.HScrollBar scrollVidaEntidad 
         Height          =   255
         Left            =   240
         Max             =   100
         TabIndex        =   37
         Top             =   300
         Value           =   100
         Width           =   2055
      End
      Begin VB.Label lblVida 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vida: X"
         Height          =   195
         Left            =   960
         TabIndex        =   38
         Top             =   100
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdRestablecer 
      Caption         =   "Restablecer"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6480
      TabIndex        =   15
      Top             =   3900
      Width           =   2655
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "Aplicar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   3600
      TabIndex        =   14
      Top             =   3900
      Width           =   2535
   End
   Begin VB.CommandButton cmdProbar 
      Caption         =   "Probar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   3600
      TabIndex        =   13
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton cmdSimular 
      Caption         =   "Simular"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6480
      TabIndex        =   12
      Top             =   4320
      Width           =   2655
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   6480
      TabIndex        =   11
      Top             =   4920
      Width           =   2655
   End
   Begin VB.CommandButton cmdEliminar_Entidades 
      Caption         =   "Elminar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   1800
      TabIndex        =   3
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdNuevo_Entidades 
      Caption         =   "Nueva"
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Frame frmPropiedades 
      Caption         =   "Propiedades"
      Height          =   3735
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   5775
      Begin TabDlg.SSTab SSTab1 
         Height          =   3375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   5953
         _Version        =   393216
         Tabs            =   6
         Tab             =   3
         TabsPerRow      =   6
         TabHeight       =   520
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "frmConfigurarEntidades.frx":1CCA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lblNumero"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblNumeroResultado"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblNombre"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblEntidadCreadaAlMorir"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtNombre"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "optVida(0)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "optVida(1)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "chkProyectil"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtEntidadAlMorir"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtVida"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "Gráficos"
         TabPicture(1)   =   "frmConfigurarEntidades.frx":1CE6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "GridGraficos"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Particulas"
         TabPicture(2)   =   "frmConfigurarEntidades.frx":1D02
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "gridParticulas"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Sonidos"
         TabPicture(3)   =   "frmConfigurarEntidades.frx":1D1E
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "GridSonidos"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "chkLoopSonido"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "Luz"
         TabPicture(4)   =   "frmConfigurarEntidades.frx":1D3A
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "frmLuz"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).Control(1)=   "chkConLuz"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).ControlCount=   2
         TabCaption(5)   =   "Al dañarla"
         TabPicture(5)   =   "frmConfigurarEntidades.frx":1D56
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "lblAclaracionAlDanarla"
         Tab(5).Control(1)=   "gridAlPegar"
         Tab(5).ControlCount=   2
         Begin EditorTDS.UpDownText txtVida 
            Height          =   315
            Left            =   -71400
            TabIndex        =   47
            Top             =   1080
            Width           =   1335
            _extentx        =   2355
            _extenty        =   556
            maxvalue        =   1000000
            minvalue        =   0
         End
         Begin EditorTDS.GridTextConAutoCompletar gridAlPegar 
            Height          =   2450
            Left            =   -74880
            TabIndex        =   45
            Top             =   820
            Width           =   5175
            _extentx        =   9128
            _extenty        =   4313
         End
         Begin EditorTDS.TextConListaConBuscador txtEntidadAlMorir 
            Height          =   285
            Left            =   -72840
            TabIndex        =   17
            Top             =   1440
            Width           =   2775
            _extentx        =   4895
            _extenty        =   503
         End
         Begin VB.CheckBox chkConLuz 
            Appearance      =   0  'Flat
            Caption         =   "Con luz"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   -74640
            TabIndex        =   35
            Top             =   360
            Width           =   855
         End
         Begin VB.Frame frmLuz 
            Height          =   2895
            Left            =   -74760
            TabIndex        =   23
            Top             =   360
            Width           =   5055
            Begin VB.CheckBox chkPrendeEn 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "Prende en horarios"
               ForeColor       =   &H80000008&
               Height          =   220
               Left            =   360
               TabIndex        =   44
               Top             =   1920
               Width           =   1695
            End
            Begin VB.Frame FraHorario 
               Height          =   855
               Left            =   240
               TabIndex        =   39
               Top             =   1920
               Width           =   4575
               Begin VB.HScrollBar horaInicioLuz 
                  Enabled         =   0   'False
                  Height          =   255
                  LargeChange     =   4
                  Left            =   1080
                  Max             =   96
                  Min             =   1
                  TabIndex        =   41
                  Top             =   240
                  Value           =   1
                  Width           =   2055
               End
               Begin VB.HScrollBar horaFinLuz 
                  Enabled         =   0   'False
                  Height          =   255
                  LargeChange     =   4
                  Left            =   1080
                  Max             =   96
                  Min             =   1
                  TabIndex        =   40
                  Top             =   480
                  Value           =   1
                  Width           =   2055
               End
               Begin VB.Label lblInicio00 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Inicio: 00:00"
                  Height          =   255
                  Left            =   0
                  TabIndex        =   43
                  Top             =   240
                  Width           =   975
               End
               Begin VB.Label lblFin00 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Fin: 00:00"
                  Height          =   255
                  Left            =   0
                  TabIndex        =   42
                  Top             =   480
                  Width           =   975
               End
            End
            Begin VB.CheckBox chkUtilizarBrillo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "Utilizar brillo"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   360
               TabIndex        =   34
               Top             =   1080
               Width           =   1215
            End
            Begin VB.CheckBox chkAnimacionFuego 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "Animacion fuego"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   3000
               TabIndex        =   30
               Top             =   360
               Width           =   1575
            End
            Begin VB.CheckBox chkLuzCuadrada 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "Luz cuadrada"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   3000
               TabIndex        =   29
               Top             =   720
               Width           =   1335
            End
            Begin VB.HScrollBar scrollLuzRadio 
               Height          =   255
               LargeChange     =   5
               Left            =   1080
               Max             =   15
               Min             =   1
               TabIndex        =   28
               Top             =   360
               Value           =   3
               Width           =   1695
            End
            Begin VB.Frame FraBrillo 
               Caption         =   "Brillo"
               Height          =   720
               Left            =   240
               TabIndex        =   24
               Top             =   1080
               Width           =   4575
               Begin VB.HScrollBar luz_luminosidad 
                  Height          =   255
                  LargeChange     =   10
                  Left            =   120
                  Max             =   254
                  TabIndex        =   26
                  Top             =   360
                  Value           =   50
                  Width           =   3975
               End
               Begin VB.CommandButton cmd 
                  Caption         =   "?"
                  Height          =   255
                  Left            =   5160
                  TabIndex        =   25
                  Top             =   2280
                  Width           =   255
               End
               Begin VB.Label luz_luminosidad_lbl 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Brillo: 50%"
                  Height          =   255
                  Left            =   600
                  TabIndex        =   27
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Label luces_color 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000003&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1080
               TabIndex        =   33
               Top             =   720
               Width           =   1695
            End
            Begin VB.Label lblRadioLuz 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Radio: 3"
               Height          =   255
               Left            =   240
               TabIndex        =   32
               Top             =   360
               Width           =   735
            End
            Begin VB.Label lblColorLuz 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Color"
               Height          =   195
               Left            =   240
               TabIndex        =   31
               Top             =   720
               Width           =   375
            End
         End
         Begin VB.CheckBox chkLoopSonido 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   22
            ToolTipText     =   "Tildar para que el sonido se repita indefinidamente"
            Top             =   2880
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CheckBox chkProyectil 
            Appearance      =   0  'Flat
            Caption         =   "Es un proyectil"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   -74760
            TabIndex        =   21
            Top             =   1800
            Width           =   1455
         End
         Begin VB.OptionButton optVida 
            Appearance      =   0  'Flat
            Caption         =   "Milisegundos de vida"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   -73320
            TabIndex        =   8
            Top             =   1080
            Width           =   1935
         End
         Begin VB.OptionButton optVida 
            Appearance      =   0  'Flat
            Caption         =   "Puntos de vida"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   -74760
            TabIndex        =   7
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtNombre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -73920
            TabIndex        =   5
            Top             =   720
            Width           =   2175
         End
         Begin EditorTDS.GridTextConAutoCompletar gridParticulas 
            Height          =   2895
            Left            =   -74880
            TabIndex        =   16
            Top             =   360
            Width           =   5175
            _extentx        =   9128
            _extenty        =   5106
         End
         Begin EditorTDS.GridTextConAutoCompletar GridGraficos 
            Height          =   2895
            Left            =   -74880
            TabIndex        =   19
            Top             =   360
            Width           =   5175
            _extentx        =   9128
            _extenty        =   5106
         End
         Begin EditorTDS.GridTextConAutoCompletar GridSonidos 
            Height          =   2895
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   5175
            _extentx        =   9128
            _extenty        =   5106
         End
         Begin VB.Label lblAclaracionAlDanarla 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cada vez que la entidad reste puntos de vida, se escuchara el sonido correspondiente"
            Height          =   555
            Left            =   -74880
            TabIndex        =   46
            Top             =   465
            Width           =   5250
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblEntidadCreadaAlMorir 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Al morir crear la entidad:"
            Height          =   255
            Left            =   -74760
            TabIndex        =   18
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label lblNombre 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            Height          =   195
            Left            =   -74760
            TabIndex        =   10
            Top             =   720
            Width           =   555
         End
         Begin VB.Label lblNumeroResultado 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            Height          =   195
            Left            =   -73920
            TabIndex        =   9
            Top             =   480
            Width           =   540
         End
         Begin VB.Label lblNumero 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Número"
            Height          =   195
            Left            =   -74760
            TabIndex        =   6
            Top             =   480
            Width           =   555
         End
      End
   End
   Begin EditorTDS.ListaConBuscador lstEntidades 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _extentx        =   5741
      _extenty        =   8281
   End
End
Attribute VB_Name = "frmConfigurarEntidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private idEntidadActualPrueba As Integer
Private idUltimaEntidad As Integer

Private Sub setEstadoEditor(estado As Boolean)
  Call modPosicionarFormulario.setEnabledHijos(estado, Me.SSTab1, Me)
End Sub

Private Sub chkConLuz_Click()
    Call setEnabledHijos((Me.chkConLuz.value = 1), Me.frmLuz, Me)
End Sub

Private Sub chkPrendeEn_Click()
    Call setEnabledHijos((Me.chkPrendeEn.value = 1), Me.FraHorario, Me)
End Sub

Private Sub chkUtilizarBrillo_Click()
    Call setEnabledHijos((Me.chkUtilizarBrillo.value = 1), Me.FraBrillo, Me)
End Sub


Private Sub guardarEn(entidad As tIndiceEntidad)
    Dim loopParte As Integer
    
    'Generales
    With entidad
        .nombre = Trim$(Me.txtNombre)
        .Vida = CInt(Me.txtVida.value)
        .Proyectil = Me.chkProyectil.value
    
        'Tipo de vida
        For loopParte = Me.optVida.LBound To Me.optVida.UBound
            If Me.optVida(loopParte).value Then .tipo = loopParte + 1: Exit For
        Next

        'Graficos
        ReDim entidad.Graficos(0 To Me.GridGraficos.obtenerCantidadCampos - 1)
    
        For loopParte = 0 To Me.GridGraficos.obtenerCantidadCampos - 1
            .Graficos(loopParte) = Me.GridGraficos.obtenerID(loopParte)
        Next
    
        'Sonidos cuando pierde vida la criatura
        ReDim entidad.SonidosAlPegar(0 To Me.gridAlPegar.obtenerCantidadCampos - 1)
    
        For loopParte = 0 To Me.gridAlPegar.obtenerCantidadCampos - 1
            .SonidosAlPegar(loopParte) = Me.gridAlPegar.obtenerID(loopParte)
        Next
    
        'Sonido
        ReDim entidad.Sonidos(0 To Me.GridSonidos.obtenerCantidadCampos - 1)
    
        For loopParte = 0 To Me.GridSonidos.obtenerCantidadCampos - 1
            .Sonidos(loopParte) = Me.GridSonidos.obtenerID(loopParte) * IIf(Me.GridSonidos.getValorDinamico("chkRepetirSonido", CByte(loopParte)) = "1", -1, 1)
        Next
        
        'Particulas
        ReDim entidad.Particulas(0 To Me.gridParticulas.obtenerCantidadCampos - 1)
    
        For loopParte = 0 To Me.gridParticulas.obtenerCantidadCampos - 1
            .Particulas(loopParte) = Me.gridParticulas.obtenerID(loopParte)
        Next
        
        'Luz
        If Me.chkConLuz.value = 1 Then
            .luz.LuzRadio = val(Me.scrollLuzRadio.value)
            .luz.LuzBrillo = Me.luz_luminosidad.value
        
            If Me.chkPrendeEn.value = 1 Then
                .luz.luzInicio = Me.horaInicioLuz.value
                .luz.luzFin = Me.horaFinLuz.value
            Else
                .luz.luzInicio = 0
                .luz.luzFin = 0
            End If
        
            .luz.LuzTipo = 0
             
            .luz.LuzTipo = .luz.LuzTipo Or IIf(Me.chkAnimacionFuego.value, TipoLuces.Luz_Fuego, 0)
            .luz.LuzTipo = .luz.LuzTipo Or IIf(Me.chkLuzCuadrada.value, TipoLuces.Luz_Cuadrada, 0)
            .luz.LuzTipo = .luz.LuzTipo Or IIf(Me.chkUtilizarBrillo.value, TipoLuces.Luz_Normal, 0)

            'Ponemos el color en el label
            VBC2RGBC Me.luces_color.BackColor, entidad.luz.LuzColor
        Else
            .luz.LuzRadio = 0
            .luz.LuzTipo = 0
            .luz.LuzBrillo = 0
            .luz.LuzColor.r = 0
            .luz.LuzColor.g = 0
            .luz.LuzColor.b = 0
            .luz.luzInicio = 0
            .luz.luzFin = 0
        End If
    End With

End Sub

Private Sub cerrar()
    'Si hay una entidad viva la matamos
    If idEntidadActualPrueba > 0 Then
        Engine_Entidades.Entidades_SetVidaActual idEntidadActualPrueba, 0
    End If
End Sub
Private Sub cmdAceptar_Click()
    Call cerrar
    Unload Me
End Sub

Private Sub cmdAplicar_Click()
    Dim prueba As tIndiceEntidad
    Dim id As Integer
    
    id = val(Me.lblNumeroResultado.caption)
    
    Call guardarEn(prueba)
    
    EntidadesIndexadas(id) = prueba

    Call Me_indexar_Entidades.actualizarEnIni(id)
    
    Call Me.lstEntidades.cambiarNombre(id, id & " - " & EntidadesIndexadas(id).nombre)
End Sub

Private Sub cmdEliminar_Entidades_Click()
    Dim confirma As VbMsgBoxResult
    Dim idElemento As Integer
    
    If Not Me.lstEntidades.obtenerValor = "" Then
        
        idElemento = Me.lstEntidades.obtenerIDValor
        
        confirma = MsgBox("¿Está seguro de que desea eliminar la entidad '" & Me.lstEntidades.obtenerValor & "'?", vbYesNo + vbExclamation, Me.caption)
        
        If confirma = vbYes Then
            Call Me_indexar_Entidades.eliminar(idElemento)
            'Lo borramos de la lista
            Call Me.lstEntidades.eliminar(CLng(idElemento))
        End If
    End If
End Sub

Private Sub cmdNuevo_Entidades_Click()
    Dim nuevo As Integer
    Dim error As Boolean
    
    error = False
    Me.cmdNuevo_Entidades.Enabled = False
    
    'Obtengo el nuevo id
    nuevo = Me_indexar_Entidades.nuevo
    
    If nuevo = -1 Then
        error = True
        MsgBox "No se ha podido obtener espacio para agregar una nueva entidad. Por favor, intenta más tarde o contactate con un administrador del sistema.", vbExclamation
    End If
    
    If Not error Then
        'Lo selecciono
        If Me.lstEntidades.seleccionarID(CLng(nuevo)) = False Then
            Call Me.lstEntidades.addString(nuevo, nuevo & " - ")
            Call Me.lstEntidades.seleccionarID(CLng(nuevo))
        End If
    End If
    
    Me.cmdNuevo_Entidades.Enabled = True
    'Cuando se haga clic en "Aplicar" se guarda
End Sub

Private Sub cmdProbar_Click()
    Dim id As Integer
    Dim tempId As Integer
    Dim X As Integer
    Dim Y As Integer
    
    id = val(Me.lblNumeroResultado.caption)
    
    If cmdProbar.caption = "Probar" Then
    
        idUltimaEntidad = idUltimaEntidad + 1
        
        'Simulamos que el server nos da el ID de la entidad
        tempId = SV_Simulador.ObtenerIDEntidad()
        
        ' Convertimos de pixeles a x, y
        Call ConvertCPtoTP(frmMain.clicX, frmMain.clicY, modPantalla.PixelesPorTile.X, modPantalla.PixelesPorTile.Y, X, Y)
        
        tempId = Engine_Entidades.Entidades_Crear_Indexada(X, Y, tempId, EntidadesIndexadas(id))
        
        scrollVidaEntidad.min = 1
        scrollVidaEntidad.max = EntidadesIndexadas(id).Vida
        scrollVidaEntidad.value = EntidadesIndexadas(id).Vida
        
        If EntidadesIndexadas(id).Proyectil = 1 Then
            Engine_Entidades.Entidades_SetCharDestino tempId, UserCharIndex
        End If
        
        idEntidadActualPrueba = tempId
        
        Me.frmVidaEntidad.visible = True
        
        Call scrollVidaEntidad_Change
        
        cmdProbar.caption = "Parar"
        
        If EntidadesIndexadas(id).tipo = eTipoEntidadVida.Tiempo Then
            Me.trmActualizarEstadoEntidad.Enabled = True
            Me.trmActualizarEstadoEntidad.Interval = 100
            Me.cmdProbar.Enabled = False
        End If
        
    Else
        cmdProbar.caption = "Probar"
        
        'La matamos
        Call Engine_Entidades.Entidades_SetVidaActual(idEntidadActualPrueba, 0)
        'Simulamos que tambien lo hizo el server
        Call SV_Simulador.EliminarIDEntidad(idEntidadActualPrueba)
        
        idEntidadActualPrueba = 0
        
        Me.frmVidaEntidad.visible = False
    End If
End Sub

Private Sub cmdRestablecer_Click()
    Dim id As Integer
    
    id = val(Me.lblNumeroResultado.caption)
    
    Call cargarEnEditor(EntidadesIndexadas(id), id)
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    idUltimaEntidad = 0
    
    For i = 1 To UBound(EntidadesIndexadas)
        If Me_indexar_Entidades.existe(i) Then
            Call Me.lstEntidades.addString(i, i & " - " & EntidadesIndexadas(i).nombre)
        End If
    Next
      
    'Entidades al morir
    Call actualizarEntidadesDisponibles
    Me.txtEntidadAlMorir.CantidadLineasAMostrar = 5
    
    'Particulas
    Call Me.gridParticulas.setNombreCampos("Particula")
    Call Me.gridParticulas.iniciar
    
    Call Me.gridParticulas.addString(CInt(0), 0 & " - Sin Particulas")
    For i = 1 To UBound(GlobalParticleGroup)
       Call Me.gridParticulas.addString(i, i & " - " & GlobalParticleGroup(i).GetNombre())
    Next
        
    Call Me.gridParticulas.seleccionarID(0, 0)
    
    'Graficos
    Call Me.GridGraficos.setNombreCampos("Grafico")
    
    Call Me.GridGraficos.iniciar
    
    Call Me.GridGraficos.addString(CInt(0), 0 & " - Sin Grafico")
    For i = 1 To UBound(GrhData)
        If GrhData(i).NumFrames > 0 Then Call Me.GridGraficos.addString(i, i & " - " & GrhData(i).nombreGrafico)
    Next
    
    Call Me.GridGraficos.seleccionarID(0, 0)
    
    'Cargamos los sonidos
    Call Me.GridSonidos.setNombreCampos("Sonido")
    
    'Le agrego un control más que es un CheckBox.
    Call Me.GridSonidos.agregarControlDinamico(Me.chkLoopSonido, "chkRepetirSonido", "Repetir")
    Call Me.GridSonidos.iniciar
    
    Call Me.GridSonidos.addString(CInt(0), 0 & " - Sin Sonido")
    For i = 1 To UBound(Me_indexar_Sonidos.Sonidos)
        If Me_indexar_Sonidos.existe(i) Then
            Call Me.GridSonidos.addString(i, i & " - " & Sonidos(i).nombre)
        End If
    Next
    Call Me.GridSonidos.seleccionarID(0, 0)
    
    'Cargamos los sonidos
    Call Me.gridAlPegar.setNombreCampos("Sonido")
    
    'Le agrego un control más que es un CheckBox.
    Call Me.gridAlPegar.iniciar
    
    Call Me.gridAlPegar.addString(CInt(0), 0 & " - Sin Sonido")
    For i = 1 To UBound(Me_indexar_Sonidos.Sonidos)
        If Me_indexar_Sonidos.existe(i) Then
            Call Me.gridAlPegar.addString(i, i & " - " & Sonidos(i).nombre)
        End If
    Next
    Call Me.gridAlPegar.seleccionarID(0, 0)
        
    'Como no tengo ninguno seleccionado, deshabilito el editor
    Call setEstadoEditor(False)
End Sub

Private Sub actualizarDescripciones(grid As GridTextConAutoCompletar)
    Dim cantidad As Byte
    Dim estado As Integer
    Dim puntos As Integer
    Dim puntosTecho As Integer
    Dim puntosVida As Integer
    Dim descripcion As String
    
    Dim estados() As String
    ReDim estados(0 To grid.obtenerCantidadCampos - 1)
    
    If Me.optVida(1).value = True Then
        descripcion = " milisegundos."
    Else
        descripcion = " puntos de vida."
    End If
    
    cantidad = grid.obtenerCantidadCampos
    puntosVida = CInt(Me.txtVida.value)
    puntosTecho = puntosVida
    
    If cantidad = 1 Then
        Call grid.setDescripcion(0, "Durante toda la vida")
    Else
        puntosTecho = puntosVida
        
        For estado = 0 To cantidad - 1
            'puntos > X
             puntos = ((-(estado + 0.5) / (cantidad - 1)) + 1) * puntosVida
            If puntos < 0 Then puntos = 0
                    
            estados(estado) = "Entre " & puntos + 1 & " y " & puntosTecho & descripcion
        
            puntosTecho = puntos
        Next
        
        
        If Me.optVida(1).value = True Then
            For estado = cantidad To 1 Step -1
                Call grid.setDescripcion(cantidad - estado, estados(estado - 1))
            Next
        Else
            For estado = 0 To cantidad - 1
                Call grid.setDescripcion(CByte(estado), estados(estado))
            Next
        End If
        
    End If
    
      
       
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call cerrar
    Unload Me
End Sub


Private Sub gridAlPegar_CantidadElementoChange()
    Call actualizarDescripciones(gridAlPegar)
End Sub

Private Sub GridGraficos_CantidadElementoChange()
    'Cambio la cantidad de elementos
    'Recalculo la descripcion
    Call actualizarDescripciones(GridGraficos)
End Sub

Private Sub actualizarEntidadesDisponibles()
    Dim i As Integer

    Call Me.txtEntidadAlMorir.limpiarLista
    
    Call Me.txtEntidadAlMorir.addString(0, 0 & " - Ninguna")
            
    For i = 1 To UBound(EntidadesIndexadas)
        If Me_indexar_Entidades.existe(i) Then
            Call Me.txtEntidadAlMorir.addString(i, i & " - " & EntidadesIndexadas(i).nombre)
        End If
    Next
End Sub

Private Sub gridParticulas_CantidadElementoChange()
    Call actualizarDescripciones(gridParticulas)
End Sub

Private Sub GridSonidos_CantidadElementoChange()
    Call actualizarDescripciones(GridSonidos)
End Sub

Private Sub cargarEnEditor(entidad As tIndiceEntidad, id As Integer)
    Dim loopParte As Byte
    Me.txtNombre = entidad.nombre
    Me.txtVida.value = entidad.Vida
    
    If entidad.tipo = eTipoEntidadVida.puntos Then
        Me.optVida(0).value = True
    Else
        Me.optVida(1).value = True
    End If
    
    Me.chkProyectil.value = entidad.Proyectil
    
    Me.lblNumeroResultado = id
    
    Call Me.txtEntidadAlMorir.seleccionarID(entidad.CrearAlMorir)

    Me.GridGraficos.limpiar
    For loopParte = 0 To UBound(entidad.Graficos)
        Call Me.GridGraficos.seleccionarID(loopParte, entidad.Graficos(loopParte))
    Next
    
    'Particulas
    Me.gridParticulas.limpiar
    For loopParte = 0 To UBound(entidad.Particulas)
        Call Me.gridParticulas.seleccionarID(loopParte, entidad.Particulas(loopParte))
    Next
    
    'Sonidos
    Me.GridSonidos.limpiar
    For loopParte = 0 To UBound(entidad.Sonidos)
        Call Me.GridSonidos.seleccionarID(loopParte, IIf(entidad.Sonidos(loopParte) < 0, entidad.Sonidos(loopParte) * -1, entidad.Sonidos(loopParte)))
        Call Me.GridSonidos.setValorDinamico("chkRepetirSonido", loopParte, IIf(entidad.Sonidos(loopParte) < 0, 1, 0))
    Next
    
    'Sonidos al pegar
    Me.gridAlPegar.limpiar
    For loopParte = 0 To UBound(entidad.SonidosAlPegar)
        Call Me.gridAlPegar.seleccionarID(loopParte, entidad.SonidosAlPegar(loopParte))
    Next
    
    'Luz
    If entidad.luz.LuzRadio > 0 Then
        Me.chkConLuz.value = 1
               
        Me.luces_color.BackColor = RGB(entidad.luz.LuzColor.r, entidad.luz.LuzColor.g, entidad.luz.LuzColor.b)

        Me.scrollLuzRadio.value = entidad.luz.LuzRadio
            
        Me.chkLuzCuadrada.value = IIf((entidad.luz.LuzTipo And TipoLuces.Luz_Cuadrada), 1, 0)
        Me.chkAnimacionFuego.value = IIf((entidad.luz.LuzTipo And TipoLuces.Luz_Fuego), 1, 0)
        
        Me.chkUtilizarBrillo.value = IIf(entidad.luz.LuzBrillo > 0, 1, 0)
        Me.luz_luminosidad.value = entidad.luz.LuzBrillo
        
        If (entidad.luz.luzInicio > 0 And entidad.luz.luzFin > 0) Then
            Me.chkPrendeEn.value = 1
            
            Me.horaInicioLuz.value = entidad.luz.luzInicio
            Me.horaFinLuz.value = entidad.luz.luzFin
        Else
            Me.chkPrendeEn.value = 0
        End If
        
        Call actualizarRadioLuz(entidad.luz.LuzRadio)
     Else
        Me.chkConLuz.value = 0
        Me.luces_color.BackColor = &H80000003
    End If
    
    Call setEstadoEditor(True)
    
    Me.cmdAceptar.Enabled = True
    Me.cmdAplicar.Enabled = True
    Me.cmdRestablecer.Enabled = True
    Me.cmdProbar.Enabled = True
    Me.cmdEliminar_Entidades.Enabled = True
End Sub

Private Sub horaFinLuz_Change()
    Call actualizarBarraHorariaLuz
End Sub

Private Sub horaInicioLuz_Change()
    Call actualizarBarraHorariaLuz
End Sub

Private Sub actualizarBarraHorariaLuz()
    lblInicio00.caption = "Inicio: " & obtener_hora_fraccion(horaInicioLuz.value)
    lblFin00.caption = "Fin: " & obtener_hora_fraccion(horaFinLuz.value)
End Sub
Private Sub lstEntidades_Change(valor As String, id As Integer)
    Call cargarEnEditor(EntidadesIndexadas(id), id)
End Sub

Private Sub luces_color_Click()

    frmMain.ColorDialog.flags = cdlCCRGBInit
    frmMain.ColorDialog.Color = luces_color.BackColor
    frmMain.ColorDialog.ShowColor
    Me.luces_color.BackColor = frmMain.ColorDialog.Color

End Sub

Private Sub luz_luminosidad_Change()
    Call actualizarBrilloLuz(Me.luz_luminosidad.value)
End Sub

Private Sub luz_luminosidad_Scroll()
    Call actualizarBrilloLuz(Me.luz_luminosidad.value)
End Sub

Public Sub actualizarRadioLuz(radio As Byte)
    Me.lblRadioLuz = "Radio: " & radio
End Sub

Public Sub actualizarBrilloLuz(brillo As Byte)
    Me.luz_luminosidad_lbl = "Brillo: " & Round((255 - brillo) / 2.55, 1) & "%"
End Sub
Private Sub scrollLuzRadio_Change()
    Call actualizarRadioLuz(Me.scrollLuzRadio)
End Sub

Private Sub scrollLuzRadio_Scroll()
    Call actualizarRadioLuz(Me.scrollLuzRadio)
End Sub

Private Sub scrollVidaEntidad_Change()
    If idEntidadActualPrueba > 0 Then
        Engine_Entidades.Entidades_SetVidaActual idEntidadActualPrueba, scrollVidaEntidad.value
        
        Me.lblVida.caption = "Vida: " & scrollVidaEntidad.value
    End If
End Sub

Private Sub trmActualizarEstadoEntidad_Timer()
    
    Dim posicion As Integer
    Dim ahora As Long
    
    posicion = Entidades_Buscar(idUltimaEntidad)
    
    If posicion >= 0 Then
        ahora = GetTimer
                
        Me.scrollVidaEntidad.value = IIf(Entidades(posicion).MuereEnTick - ahora > 0, Entidades(posicion).MuereEnTick - ahora, 1)
    Else
        'Simulamos que la mato el server
        Call SV_Simulador.EliminarIDEntidad(idEntidadActualPrueba)
        
        cmdProbar.caption = "Probar"
              
        idEntidadActualPrueba = 0
        
        Me.cmdProbar.Enabled = True
        Me.frmVidaEntidad.visible = False
        Me.trmActualizarEstadoEntidad.Enabled = False
    End If
End Sub
