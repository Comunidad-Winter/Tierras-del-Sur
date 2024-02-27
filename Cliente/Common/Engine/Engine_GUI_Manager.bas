Attribute VB_Name = "Engine_GUI_Manager"
Option Explicit

Public Enum eGuiColorFill
    dSolid = 0
    dHorizontal = 1
    dVertical = 2
End Enum

Public Enum eGuiTiposTexto
    ALFANUMERICO = 0
    numerico = 1
    Contraseña = 2
End Enum


Public vWindowLast As vWindow
Public vWindowRoot As vWindow 'Se hace render de atrás para adelante.
Public vWindowCurr As vWindow

Public vTextoAlerta As String
Public vAlerteTitle As String

Public pakGUI                   As New clsEnpaquetado

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
                ByVal lpPrevWndFunc As Long, _
                ByVal hWnd As Long, _
                ByVal Msg As Long, _
                ByVal wParam As Long, _
                ByVal lParam As Long) As Long

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
                
Private Declare Function GetParent Lib "user32" ( _
                ByVal hWnd As Long) As Long

Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" ( _
                ByVal hWnd As Long, _
                ByVal lpString As String) As Long

Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A
Private Const CB_GETDROPPEDSTATE = &H157

Private ClaseMouse As New clsMouseWheel

Private MouseX As Integer
Private MouseY As Integer

Private ObjetoIterador As vWindow

Private Sub GUI_ReiniciarIterador()
Set ObjetoIterador = vWindowLast
End Sub

Private Function GUI_Iterar(v As vWindow) As Boolean
Set v = ObjetoIterador
If Not ObjetoIterador Is Nothing Then
    ObjetoIterador.GetPrev ObjetoIterador
End If
GUI_Iterar = Not (v Is Nothing)
End Function

Public Sub GUI_Render()

'    If Not vWindowLast Is Nothing Then
'    GUI_RenderRInverse vWindowLast
'        If vWindowLast.Render = False Then
'            Set vWindowLast = Nothing
'        End If
'    End If

    If Not vWindowRoot Is Nothing Then
        GUI_RenderRInverse vWindowRoot
    End If
End Sub


Private Sub GUI_RenderRInverse(v As vWindow)
Dim vN As vWindow
Dim rN As Boolean
rN = v.GetNext(vN)
If v.Render = False Then GUI_Quitar v
If rN Then GUI_RenderRInverse vN
End Sub


Private Sub GUI_RenderR(v As vWindow)
Dim vN As vWindow
Dim rN As Boolean
rN = v.GetPrev(vN)
If v.Render = False Then GUI_Quitar v
If rN Then GUI_RenderR vN
End Sub

Public Sub GUI_Move(X%, Y%)
    If Not vWindowLast Is Nothing Then
        vWindowLast.SetPos X, Y
    End If
End Sub

Public Sub GUI_Load(NuevaVentana As vWindow)
    If NuevaVentana Is vWindowLast Then Exit Sub
    
    If vWindowRoot Is Nothing Then 'La lista está vacía
        NuevaVentana.SetNext Nothing
        NuevaVentana.SetPrev Nothing
        Set vWindowRoot = NuevaVentana
        Set vWindowLast = NuevaVentana
    Else
        vWindowLast.SetNext NuevaVentana
        NuevaVentana.SetPrev vWindowLast
        NuevaVentana.SetNext Nothing
        Set vWindowLast = NuevaVentana
    End If
End Sub

Public Function GUI_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long) As Boolean
    Dim v As vWindow
    
    If Not vWindowRoot Is Nothing Then
        
        GUI_ReiniciarIterador
        
        While GUI_Iterar(v) And GUI_MouseMove = False
            Dim Controles As vControles
            If v.IsVisible Then
                Set Controles = v.GetControl
                If Not (Controles Is Nothing) Then
                    GUI_MouseMove = Controles.MouseMove(Button, Shift, X, Y, 0)
                End If
            End If
        Wend
    End If
    
    MouseX = X
    MouseY = Y
End Function

Public Function GUI_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long) As Boolean
    Dim v As vWindow

    
    If Not vWindowRoot Is Nothing Then
        GUI_ReiniciarIterador
        
        While GUI_Iterar(v) And GUI_MouseUp = False
            Dim Controles As vControles
            
            If v.IsVisible Then
                Set Controles = v.GetControl
                If Not (Controles Is Nothing) Then
                    GUI_MouseUp = Controles.MouseUp(Button, Shift, X, Y)
                End If
            End If
        Wend
        
    End If
    
    MouseX = X
    MouseY = Y
End Function

Public Function GUI_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long) As Boolean
    Dim v As vWindow
    
    If Not vWindowRoot Is Nothing Then
        
        GUI_ReiniciarIterador
        
        While GUI_Iterar(v) And GUI_MouseDown = False
            Dim Controles As vControles
            If v.IsVisible Then
                Set Controles = v.GetControl
                If Not (Controles Is Nothing) Then
                    GUI_MouseDown = Controles.MouseDown(Button, Shift, X, Y)
                End If
            End If
        Wend
    End If
    
    MouseX = X
    MouseY = Y
End Function

Public Function GUI_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer) As Boolean
    
    If Not vWindowLast Is Nothing Then
        GUI_ReiniciarIterador
        Dim v As vWindow
        While GUI_Iterar(v) And GUI_KeyDown = False
            Dim Controles As vControles
            
            If v.IsVisible Then
                Set Controles = v.GetControl
                If Not (Controles Is Nothing) Then
                    GUI_KeyDown = Controles.KeyDown(KeyCode, Shift)
                End If
            End If
            
        Wend
    End If
End Function

Public Function GUI_AdvanceFoucs() As Boolean
    If vWindowLast Is Nothing Then
        Exit Function
    End If

    Dim Controles As vControles
    
    Set Controles = vWindowLast.GetControl
    
    If Not (Controles Is Nothing) Then
        GUI_AdvanceFoucs = Controles.AdvanceFocus
    End If
    
End Function

Public Function GUI_Keypress(ByVal KeyAscii As Integer) As Boolean
    If Not vWindowLast Is Nothing Then
        GUI_ReiniciarIterador
        Dim v As vWindow
        While GUI_Iterar(v) And GUI_Keypress = False
            Dim Controles As vControles
            
            If v.IsVisible Then
                Set Controles = v.GetControl
                If Not (Controles Is Nothing) Then
                    GUI_Keypress = Controles.KeyPress(KeyAscii)
                End If
            End If
        Wend
    End If
End Function

Public Function GUI_KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer) As Boolean
    If Not vWindowLast Is Nothing Then
        GUI_ReiniciarIterador
        Dim v As vWindow
        While GUI_Iterar(v) And GUI_KeyUp = False
            Dim Controles As vControles
            
            If v.IsVisible Then
                Set Controles = v.GetControl
                If Not (Controles Is Nothing) Then
                    GUI_KeyUp = Controles.KeyUp(KeyCode, Shift)
                End If
            End If
        Wend
    End If
End Function

Public Function GUI_MouseWheel(MouseKeys As Long, Rotation As Long, X As Long, Y As Long) As Boolean
    Dim Delta As Integer
    
    If Rotation > 0 Then
        Delta = 1
    Else
        Delta = -1
    End If
    
    If Not vWindowLast Is Nothing Then
        GUI_ReiniciarIterador
        Dim v As vWindow
        While GUI_Iterar(v) And GUI_MouseWheel = False
            Dim Controles As vControles
            
            If v.IsVisible Then
                Set Controles = v.GetControl
                If Not (Controles Is Nothing) Then
                    GUI_MouseWheel = Controles.MouseMove(MouseKeys And &HFFFF, 0, MouseX, MouseY, Delta)
                End If
            End If
            
        Wend
    End If
    
    If GUI_MouseWheel = False Then
        #If esMe = 1 Then
            If Delta = 1 Then
                Call ME_Tools.rotarHerramientaInterna(False)
            Else
                Call ME_Tools.rotarHerramientaInterna(True)
            End If
        #End If
    End If
End Function

Public Sub GUI_Quitar(Ventana As vWindow)
' Render   >   >   >   >   >    VENTANA ACTIVA
'         _1_ _2_ _3_ _4_ _5_
' Root   |0_2|1_3|2_4|3_5|4_0|  Last = 5
'
' Quitar: 3
'
'         _1_ _2_ _4_ _5_
' Root   |0_2|1_4|2_5|4_0|      Last = 5
If Not (Ventana Is Nothing) Then
    Dim vPrev As vWindow
    Dim vNext As vWindow
    Dim pPrev As Boolean
    Dim pNext As Boolean

    Ventana.Hide
    
    pPrev = Ventana.GetPrev(vPrev)
    pNext = Ventana.GetNext(vNext)
    
    If pPrev Then
        vPrev.SetNext vNext
    Else
        If vWindowRoot Is Ventana Then
            Set vWindowRoot = vNext 'Es el primer item de la lista
        End If
    End If
    
    If pNext Then
        vNext.SetPrev vPrev
    Else
        If vWindowLast Is Ventana Then
            Set vWindowLast = vPrev 'Es el ultimo item de la lista
        End If
    End If
End If
End Sub

Public Sub GUI_SetFocus(Ventana As vWindow) 'Mueve Ventana al root de la lista
    If Not Ventana Is Nothing Then
        If Not Ventana Is vWindowLast Then
            GUI_Quitar Ventana  'La quita del medio de la lista
            GUI_Load Ventana    'La agrego denuevo al final de la lista
        End If
    End If
End Sub

Public Function GUI_Alert(texto As String, Optional ByVal Titulo As String = "Alerta")
    vTextoAlerta = texto
    'vAlerteTitle = Titulo
   ' Dim v As New vWAlert
   ' GUI_Load v
   ' v.setTitulo Titulo
    
End Function

Public Sub GUI_RenderDialog(ByVal X As Integer, ByVal Y As Integer, ByVal w As Integer, ByVal h As Integer, Optional ByRef title As String = vbNullString, Optional ByRef Sender As vWindow = Nothing, Optional ByVal progress As Single = 1)
Static body As Box_Vertex
Static Titulo As Box_Vertex
Static Border As Box_Vertex

If body.rhw0 = 0 Then
    With body
    .rhw0 = 1
    .rhw1 = 1
    .rhw2 = 1
    .rhw3 = 1

    End With
    Titulo = body
    Border = body
End If
Dim BorderOffset As Integer

BorderOffset = 1 * progress

Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse Nothing
Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
If BorderOffset Then
    
    With Border
        .x0 = X - BorderOffset
        .x1 = .x0
        .x2 = X + w + BorderOffset
        .x3 = .x2
        .y0 = Y + h + BorderOffset
        .y1 = Y - BorderOffset
        .y2 = .y0
        .y3 = .y1
        .color0 = &HE6E7EA Or Alphas(&HFF * progress)
        .Color2 = .color0
        .Color1 = .color0 '&H333333 Or Alphas(&HFA * progress)
        .color3 = .Color1
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, .x0, TL_size
    End With
End If


If title <> vbNullString Then
    If Sender Is vWindowLast Then
        With Titulo
            .color0 = &H979797 Or Alphas(&HFF * progress)
            .Color2 = .color0
            .Color1 = &HBEBEBE Or Alphas(&HFF * progress)
            .color3 = .Color1
        End With
    Else
        With Titulo
            .color0 = Alphas(&HB0 * progress)
            .Color2 = .color0
            .Color1 = .color0 '&H00000 Or Alphas(&HFA * progress)
            .color3 = .Color1
        End With
    End If
    With Titulo
        .x0 = X
        .x1 = X
        .x2 = X + w
        .x3 = .x2
        .y0 = Y + 24 * progress
        .y1 = Y
        .y2 = .y0
        .y3 = .y1
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, .x0, TL_size
    End With
    If progress = 1 Then Engine.text_render_graphic title, X + 6, Y + 4
    With body
        .x0 = X
        .x1 = X
        .x2 = X + w
        .x3 = .x2
        .y0 = Y + h
        .y1 = Y + 24 * progress
        .y2 = .y0
        .y3 = .y1
        .color0 = &H111111 Or Alphas(&HFE * progress)
        .Color2 = .color0
        .Color1 = .color0
        .color3 = .color0
        Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse Nothing
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, .x0, TL_size
    End With
Else
    With body
        .color0 = &H888888 Or Alphas(&HFE * progress)
        .Color2 = .color0
        .Color1 = .color0
        .color3 = .color0
        .x0 = X
        .x1 = X
        .x2 = X + w
        .x3 = .x2
        .y0 = Y + h
        .y1 = Y
        .y2 = .y0
        .y3 = .y1
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, .x0, TL_size
    End With
End If




End Sub


Public Function mouseWindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim MouseKeys As Long
  Dim Rotation As Long
  Dim Xpos As Long
  Dim Ypos As Long
  Dim fFrm As Form

  Select Case Lmsg
  
    Case WM_MOUSEWHEEL
    
      MouseKeys = wParam And 65535
      Rotation = wParam / 65536
      Xpos = lParam And 65535
      Ypos = lParam / 65536

      GUI_MouseWheel MouseKeys, Rotation, Xpos, Ypos

  End Select
  
  mouseWindowProc = CallWindowProc(GetProp(Lwnd, "PrevWndProc"), Lwnd, Lmsg, wParam, lParam)
End Function

Public Sub IniciarRuedaMouse(ByVal hWnd As Long)
    ClaseMouse.WheelHook hWnd
End Sub


Public Function CastearvWindow(obj As vWindow) As vWindow
    Set CastearvWindow = obj
End Function
