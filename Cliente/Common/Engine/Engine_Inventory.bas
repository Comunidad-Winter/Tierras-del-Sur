Attribute VB_Name = "Engine_Inventory"
Option Explicit

Dim InventoryOffset As Long             'Number of lines we scrolled down from topmost
Public SelectedItem As Long             'Currently selected item

'Dim InvSurface As DirectDrawSurface7            'DD Surface used to render everything

Dim UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory    'User's inventory

'Dim WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory

#If ConMenuesConextuales = 1 Then
    Dim ItemMenu As Menu    'Menu to be shown as pop up
#End If

Dim last_i As Byte
Dim last_s As Byte
Dim invtl(3) As TLVERTEX
Public dibujar_tooltip_inv As Integer
Dim slots(1 To 6) As Byte

Public inv_tooltip_counter As Integer

Public Sub DrawInventory()
    Dim i As Long
    Dim X!
    Dim Y!
    Dim tt$
    Call GetTexture(9719)
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, invtl(0), TL_size
    For i = 1 To MAX_INVENTORY_SLOTS
        If UserInventory(i).GrhIndex Then
            X = ((i - 1) Mod 4) * 37 + 5
            Y = ((i - 1) \ 4) * 37 + 5
            
            If SelectedItem = i Then
                Call Engine.Draw_FilledBox(X, Y, 32, 32, &H7F000000, &H7FCC0000)
                'Grh_Render_invselslot x, y
            End If
            
            Call Engine.Grh_Render_nocolor(UserInventory(i).GrhIndex, X, Y)
            If UserInventory(i).Amount > 1 Then Call Engine.Text_Render_ext(CStr(UserInventory(i).Amount), Y, X, 40, 40, &HFFFFFFFF)
            If UserInventory(i).Equipped Then
                Call Engine.Text_Render_ext("+", Y + 20!, X + 20!, 40!, 40!, &HFFFFFF00)
            End If
        End If
    Next i

    If dibujar_tooltip_inv And inv_tooltip_counter > 3 Then
        With UserInventory(dibujar_tooltip_inv)
            If Len(.Name) Then
                tt = .Name
                If .MinDef Then tt = tt & vbNewLine & Chr$(255) & " Def: " & Chr$(255) & .MinDef & "/" & .MaxDef
                If .MaxHit Then tt = tt & vbNewLine & Chr$(255) & " Hit: " & Chr$(255) & .MinHit & "/" & .MaxHit
            End If
        End With
        If Len(tt) Then
            
            If inv_tooltip_counter = 4 Then
            Call Engine.Draw_FilledBox(5, 154, 142, 41, &H7F000000, &H9F363636, 2)
            Engine.Text_Render_alpha tt, 155, 9, &HFFFFFFFF, 0, 100
            Else
            Call Engine.Draw_FilledBox(5, 154, 142, 41, &H9F000000, &HBF363636, 2)
            Engine.Text_Render_alpha tt, 155, 9, &HFFFFFFFF, 0, 200
            End If
        End If
    End If
End Sub


Public Sub set_Slots(ByVal slot As Byte, ByVal obj_slot As Byte)
On Error Resume Next
slots(slot) = obj_slot
End Sub

Public Sub reset_slots()
Dim i As Integer
For i = 1 To 6
slots(i) = 0
Next i
End Sub

Public Sub Inventory_init()
'FIXME
'On Error Resume Next
'    frmMain.ImageList1.MaskColor = vbBlack
'    frmMain.ImageList1.UseMaskColor = True
'    init_gui_tl invtl, 0, 0, 199, 200
'    SelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item
End Sub

Public Property Get GrhIndex(ByVal slot As Byte) As Integer
    GrhIndex = UserInventory(slot).GrhIndex
End Property

Public Property Get Amount(ByVal slot As Byte) As Long
    If slot = FLAGORO Then
        Amount = UserGLD
    ElseIf slot >= LBound(UserInventory) And slot <= UBound(UserInventory) Then
        Amount = UserInventory(slot).Amount
    End If
End Property

Public Property Get OBJIndex(ByVal slot As Byte) As Integer
    OBJIndex = UserInventory(slot).OBJIndex
End Property

Public Property Get OBJType(ByVal slot As Byte) As Integer
    OBJType = UserInventory(slot).OBJType
End Property

Public Property Get ItemName(ByVal slot As Byte) As String
    ItemName = UserInventory(slot).Name
End Property

Public Property Get Equipped(ByVal slot As Byte) As Boolean
    Equipped = UserInventory(slot).Equipped
End Property


Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long
'FIXME
'    Dim TempItem As Long
'    Dim temp_x As Long
'    Dim temp_y As Long
'
'    temp_x = x \ 37
'    temp_y = y \ 37
'
'    TempItem = temp_x + (temp_y + InventoryOffset) * (148 \ 37) + 1
'    If TempItem > MAX_INVENTORY_SLOTS Then TempItem = 1
'    'Make sure it's within limits
'    If TempItem <= MAX_INVENTORY_SLOTS Then
'        'Make sure slot isn't empty
'        If UserInventory(TempItem).GrhIndex Then
'            ClickItem = TempItem
'        Else
'            ClickItem = 0
'        End If
'        DrawInventory
'    End If
End Function

Function buscari(gh As Integer) As Integer
'FIXME
'Dim i As Integer
''BUSQUEDA BINARIA?
'' LAS PELOTAS
'' PAJA
'For i = 1 To frmMain.ImageList1.ListImages.count
'    If frmMain.ImageList1.ListImages(i).key = "g" & CStr(gh) Then
'        buscari = i
'        Exit For
'    End If
'Next i
End Function

Public Sub InventoryWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
End Sub



Public Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
End Sub

Public Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
End Sub



