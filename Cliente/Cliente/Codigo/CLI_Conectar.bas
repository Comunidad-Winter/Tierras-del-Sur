Attribute VB_Name = "modDibujarInterface"
Option Explicit

Private antorcha1 As Engine_Particle_Group
Private antorcha2 As Engine_Particle_Group
Private cargando As Engine_Particle_Group
Private yaDibuje As Boolean

Private tiempoComienzo As Long
Private posicioninicial As position

Private tiempoLogo As Long
Private tiempoFondo As Long
Private tiempoAntorchaIzquierda As Long
Private tiempoAntorchaDerecha As Long
Private tiempoGUI As Long

Private vwLogin As vwLogin
Private renderizar As Boolean

Private progresoBackGround As clsProgreso

Private RECT As RECT

Private fondoOffsetX As Integer
Private posicionFogata1 As position
Private posicionFogata2 As position

Public Sub DibujarInterface()
    On Error GoTo errh
    
    If Not Device_Test_Cooperative_Level Then Exit Sub
 
    D3DDevice.Clear 1, RECT, D3DCLEAR_TARGET, 0, 0, 0
    D3DDevice.BeginScene

    DibujarConectar

    D3DDevice.EndScene
    D3DDevice.Present RECT, RECT, frmConnect.picInv.hWnd, ByVal 0
    Exit Sub
errh:
    LogError "DrawInv: " & D3DX.GetErrorString(Err.Number) & " Desc: " & Err.Description & " #: " & Err.Number
End Sub

Public Sub Show()
    renderizar = True
End Sub

Public Sub Hide()
    renderizar = False
    Call GUI_Quitar(vwLogin)
    Set antorcha1 = Nothing
    Set antorcha2 = Nothing
    Set vwLogin = Nothing
End Sub

Public Sub comenzarConectar()
    RECT.bottom = Engine_Resolution.pixelesAlto
    RECT.right = Engine_Resolution.pixelesAncho
    
    posicionFogata1.X = 410
    posicionFogata1.Y = 360
        
    posicionFogata2.X = 885
    posicionFogata2.Y = 360
        
    If Engine_Resolution.resolucionActual = RESOLUCION_43 Then
        fondoOffsetX = 128
        posicionFogata1.X = posicionFogata1.X - fondoOffsetX
        posicionFogata2.X = posicionFogata2.X - fondoOffsetX
    End If
    
    renderizar = True
    tiempoComienzo = GetTimer
    
    tiempoLogo = tiempoComienzo + 1250
    tiempoFondo = tiempoComienzo + 1250
    tiempoAntorchaIzquierda = tiempoComienzo
    tiempoAntorchaDerecha = tiempoComienzo + 250
    tiempoGUI = tiempoFondo
    
    posicioninicial.X = Engine_Resolution.pixelesAncho / 2 - 256 / 2
    posicioninicial.Y = Engine_Resolution.pixelesAlto / 2 - 256 / 2 - Engine_Resolution.pixelesAlto / 10
End Sub
Private Sub DibujarConectar()

    If Not renderizar Then Exit Sub
    
    Dim tiempoActual As Long
    
        tiempoActual = GetTimer

    If tiempoActual > tiempoFondo Then
       Call Engine_GrhDraw.Grh_Render_size(GRH_FONDO, 0, 0, 0, -1, 0, Engine_Resolution.pixelesAncho, Engine_Resolution.pixelesAlto, True, fondoOffsetX, 0)
       Call Engine_GrhDraw.Grh_Render_size(GRH_ANTORCHA_DERECHA, posicionFogata2.X - 35, posicionFogata2.Y - 125)
       Call Engine_GrhDraw.Grh_Render_size(GRH_ANTORCHA_IZQUIERDA, posicionFogata1.X - 35, posicionFogata1.Y - 125)
    End If
            
    If yaDibuje And frmPres.Visible Then
        Call frmConnect.SetFocus
        frmPres.Visible = False
    End If

    Dim Progreso As Single

    If tiempoActual > tiempoLogo Then
        Progreso = 1
        
        If vwLogin Is Nothing Then
            Dim Ventana As vWindow
            
            Set vwLogin = New vwLogin
            Set Ventana = vwLogin
            
            Call GUI_Load(Ventana)
            Call Ventana.Show
        End If
        
    Else
        Progreso = CosInterp(0, 1, (tiempoActual - tiempoComienzo) / (tiempoLogo - tiempoComienzo))
    End If
     
    If tiempoActual > tiempoFondo Then
        If progresoBackGround Is Nothing Then
            Set progresoBackGround = New clsProgreso
            progresoBackGround.SetTicks GetTimer + 1000
            progresoBackGround.SetRango 255, 1
        End If
    
        Engine.Draw_FilledBox 0, 0, D3DWindow.BackBufferWidth, D3DWindow.BackBufferHeight, Alphas(progresoBackGround.Calcular), 0, 0
        
        Call text_render_graphic("Dungeon Tenebris", Engine_Resolution.pixelesAncho - 120, Engine_Resolution.pixelesAlto - 30, mzCTalkMuertos)
    End If
    
    ' Antorchas
    If tiempoActual > tiempoAntorchaIzquierda Then
        If antorcha1 Is Nothing Then
            Set antorcha1 = New Engine_Particle_Group
            antorcha1 = 2
            Call antorcha1.SetPosAbsolute(posicionFogata1.X, posicionFogata1.Y)
        End If
        antorcha1.Render
    End If
    
    If tiempoActual > tiempoAntorchaDerecha Then
        If antorcha2 Is Nothing Then
            Set antorcha2 = New Engine_Particle_Group
            antorcha2 = 2
    
            Call antorcha2.SetPosAbsolute(posicionFogata2.X, posicionFogata2.Y)
        End If
        
        Call antorcha2.Render
        
       
    End If
    
    ' Interacciones
    If tiempoActual > tiempoFondo Then
        Call GUI_Render
        
    End If
    
    ' Logo de TDS
    Call Grh_Render_Simple_box(TEXT_LOGO_TDS, posicioninicial.X + 68 * Progreso, posicioninicial.Y - 105 * Progreso, -1, 256 - 128 * Progreso)
    
    yaDibuje = True
End Sub

Public Sub mostrarCuenta()
    Call vwLogin.mostrarCuenta
End Sub
Public Sub mostrarError(error As Byte, errordesc As String)
    If Not vwLogin Is Nothing Then
        Call vwLogin.mostrarError(error, errordesc)
    End If
End Sub
