Attribute VB_Name = "CLI_DibujarInventario"

Option Explicit

Private InvRect As RECT
Private grhFondoInventario As Integer
Private cantidadColumnas As Integer

Public MousePress As Byte
Public MouseX As Integer 'Las muevo aca para q sea mas rapido llamarlas y no tenga q poner frmmain. porq me da paja
Public MouseY As Integer
Public ItemDragued As Byte
Public MousePressX As Integer
Public MousePressY As Integer
Public MousePressPosX As Integer
Public MousePressPosY As Integer

Public itemElegido As Integer

Public dibujar_tooltip_inv As Integer
Public inv_tooltip_counter As Long
Public lineaUnoDescripcionObjeto As String
Public lineaDosDescripcionObjeto As String

Sub ItemClick(X As Integer, Y As Integer, picInv As PictureBox)
    Dim lPreItem As Long
    Dim MX As Integer
    Dim MY As Integer

    If X > 0 And Y > 0 And X < picInv.ScaleWidth And Y < picInv.ScaleHeight Then
        MX = X \ 32 + 1
        MY = Y \ 32 + 1
        lPreItem = (MX + (MY - 1) * cantidadColumnas)
        If lPreItem <= MAX_INVENTORY_SLOTS Then
            'If UserInventory(lPreItem).GrhIndex > 0 Then
                itemElegido = lPreItem
            'End If
        End If
    End If
End Sub

Public Function MouseMove(X As Single, Y As Single)
    Dim MX As Integer
    Dim MY As Integer
    Dim aux As Integer
    
    MX = X \ 32 + 1
    MY = Y \ 32 + 1
    aux = (MX + (MY - 1) * cantidadColumnas)
    
    Dim isAlgoSeleccionado As Boolean
    
    isAlgoSeleccionado = False
    
    If aux > 0 And aux < MAX_INVENTORY_SLOTS Then
        If frmMain.picInv.MousePointer = vbDefault Then
            If Not dibujar_tooltip_inv = aux Then
                If UserInventory(aux).OBJIndex > 0 Then
                    dibujar_tooltip_inv = aux
                    inv_tooltip_counter = GetTickCount
                
                    Dim comienzoParentesis As Integer
                          
                    comienzoParentesis = InStr(1, UserInventory(aux).Name, "(")
                    
                    If comienzoParentesis = 0 And Len(UserInventory(aux).Name) > 15 Then
                        comienzoParentesis = InStr(1, UserInventory(aux).Name, "+")
                    End If
                    
                    If comienzoParentesis > 0 Then
                        lineaUnoDescripcionObjeto = Trim$(left$(UserInventory(aux).Name, comienzoParentesis - 1))
                        lineaDosDescripcionObjeto = mid$(UserInventory(aux).Name, comienzoParentesis)
                    Else
                        lineaUnoDescripcionObjeto = UserInventory(aux).Name
                        lineaDosDescripcionObjeto = vbNullString
                    End If
                    isAlgoSeleccionado = True
                End If
                
            Else
                isAlgoSeleccionado = True
            End If
        End If
    End If
    
    If isAlgoSeleccionado = False Then
        dibujar_tooltip_inv = 0
        inv_tooltip_counter = 0
    End If
    
    MouseMove = isAlgoSeleccionado
End Function

    
Public Sub DrawInv()
    On Error GoTo errh
    
    DoEvents
    
    If frmMain.Visible = False Then Exit Sub
    
    If Not Device_Test_Cooperative_Level Then Exit Sub
    
    D3DDevice.Clear 1, InvRect, D3DCLEAR_TARGET, 0, 0, 0
    D3DDevice.BeginScene

    DibujarInv

    D3DDevice.EndScene
    D3DDevice.Present InvRect, ByVal 0, frmMain.picInv.hWnd, ByVal 0
    Exit Sub
errh:
    LogError "DrawInv: " & D3DX.GetErrorString(Err.Number) & " Desc: " & Err.Description & " #: " & Err.Number
End Sub


' Dibuja un inventario sobre un picture box. Toma el tamaño del mismo para calcular cuantos items entran.
Public Sub DrawInvSimple(picInv As PictureBox, inventario() As Inventory, slotSeleccionado As Byte)
    On Error GoTo errh
    
    Dim RECT As RECT
    
    RECT.left = 0
    RECT.top = 0
    RECT.bottom = picInv.Height
    RECT.right = picInv.width
    
    If Not Device_Test_Cooperative_Level Then Exit Sub
    
    D3DDevice.Clear 1, RECT, D3DCLEAR_TARGET, 0, 0, 0
    D3DDevice.BeginScene

    Call DibujarInvAuxiliar(RECT.right \ TilePixelWidth, RECT.left \ TilePixelHeight, slotSeleccionado, inventario)

    D3DDevice.EndScene
    D3DDevice.Present RECT, ByVal 0, picInv.hWnd, ByVal 0
    Exit Sub
errh:
    LogError "DrawInv: " & D3DX.GetErrorString(Err.Number) & " Desc: " & Err.Description & " #: " & Err.Number
End Sub

' Dibuja un inventario sobre la surface
Private Sub DibujarInvAuxiliar(ancho As Byte, alto As Byte, slotSeleccionado As Byte, inventario() As Inventory)

Dim X As Integer
Dim Y As Integer
Dim i As Integer

For i = 1 To UBound(inventario)
    If inventario(i).GrhIndex > 0 Then
        X = ((i - 1) Mod ancho) * TilePixelWidth
        Y = ((i - 1) \ ancho) * TilePixelHeight
        
        If slotSeleccionado = i Then
            Call Engine.Draw_FilledBox(X, Y, 32, 32, &HFF000000, &HFFB57521)
        End If
        
        Call Engine_GrhDraw.Grh_Render_nocolor(inventario(i).GrhIndex, X, Y)
        
        If inventario(i).Amount > 1 Then
            Call Engine.text_render_graphic(CStr(inventario(i).Amount), CSng(X), Y - 2!, &HD0FFFFFF)
        End If
        
        If inventario(i).Equipped = 1 Then
            Call Engine.Text_Render_ext("+", Y + 20!, X + 20!, 40!, 40!, mzYellow)
        End If
    End If
Next i

End Sub

Sub DibujarInv()

Dim X%
Dim Y%

Dim i As Integer
Dim tt$

Call Engine_GrhDraw.Grh_Render_nocolor(grhFondoInventario, 0, 0)

For i = 1 To MAX_INVENTORY_SLOTS
    If UserInventory(i).GrhIndex > 0 Then
        X = ((i - 1) Mod cantidadColumnas) * 32
        Y = ((i - 1) \ cantidadColumnas) * 32
        
        If itemElegido = i Then
            Call Engine.Draw_FilledBox(X, Y, 32, 32, &HFF000000, &HFFB57521)
        End If
        
        Call Engine_GrhDraw.Grh_Render_nocolor(UserInventory(i).GrhIndex, X, Y)
        
        If UserInventory(i).Amount > 1 Then
            Call Engine.text_render_graphic(CStr(UserInventory(i).Amount), CSng(X), Y - 2!, &HD0FFFFFF)
        End If
        
        If UserInventory(i).Equipped Then
            Call Engine.Text_Render_ext("+", Y + 20!, X + 20!, 40!, 40!, mzYellow)
        End If
    End If
Next i


If dibujar_tooltip_inv Then
    If GetTickCount - inv_tooltip_counter > 1000 Then
        If lineaDosDescripcionObjeto = vbNullString Then
            Call Engine.Text_Render_ext(lineaUnoDescripcionObjeto, frmMain.picInv.Height - 16, frmMain.picInv.width / 2, frmMain.picInv.width, 40!, &HCC6F47, False, True)
        Else
            Call Engine.Text_Render_ext(lineaUnoDescripcionObjeto, frmMain.picInv.Height - 26, frmMain.picInv.width / 2, frmMain.picInv.width, 40!, &HCC6F47, False, True)
            Call Engine.Text_Render_ext(lineaDosDescripcionObjeto, frmMain.picInv.Height - 16, frmMain.picInv.width / 2, frmMain.picInv.width, 40!, &HCC6F47, False, True)
        End If
    End If
End If

If itemElegido = 0 Then Call ItemClick(2, 2, frmMain.picInv)


End Sub

Public Sub Init_Inventario(ByVal ancho As Integer, ByVal alto As Integer, grhFondo As Integer, cantidadColumnas_ As Integer)
    InvRect.left = 0
    InvRect.top = 0
    InvRect.bottom = alto
    InvRect.right = ancho
    grhFondoInventario = grhFondo
    cantidadColumnas = cantidadColumnas_
End Sub
