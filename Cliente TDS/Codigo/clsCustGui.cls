VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGui"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************
'Coded by Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'and Fredrik Alexandersson (fredrik@oraklet.zzn.com)
'********************************************************
Option Explicit

Const COMBOBOX_BUTTON_WIDTH = 16    'In pixels
Const COMBOBOX_ITEMS_DISPLAYED = 5  'Must be at least 3

Const CHECKBOX_WIDTH = 15   'In pixels
Const CHECKBOX_HEIGHT = 15

Const OPTIONBUTTON_WIDTH = 10   'In pixels
Const OPTIONBUTTON_HEIGHT = 10

Private Type GUIOptionButton
    text As String
    font_index As Long
    font_color As Long
    top As Long
    left As Long
    width As Long
    height As Long
    value As Boolean
    grh_unselected As Long
    grh_selected As Long
    grh_disabled As Long
    BackgroundColor As Long
    visible As Boolean
    enabled As Boolean
    group As Long
    switch_id As Long
End Type

Private Type GUICheckBox
    text As String
    font_index As Long
    font_color As Long
    top As Long
    left As Long
    width As Long
    height As Long
    value As Boolean
    grh_unchecked As Long
    grh_checked As Long
    BackgroundColor As Long
    visible As Boolean
    enabled As Boolean
    'GUMP specyfic
    switch_id As Long
End Type

Private Type GUIProgressBar
    Percentage As String
    top As Long
    left As Long
    width As Long
    height As Long
    Max As Long
    value As Long
    visible As Boolean
End Type

Private Type GUITextBox
    text As String
    top As Long
    left As Long
    width As Long
    height As Long
    BackgroundColor As Long
    FontColor As Long
    FontColor2 As Long
    grh As Long
    font As Long
    visible As Boolean
    password As Boolean
End Type

Private Type GUIButton
    text As String
    top As Long
    left As Long
    width As Long
    height As Long
    font_color As Long
    Background_Color As Long
    Grh_Unpressed As Long
    Grh_Pressed As Long
    visible As Boolean
    enabled As Boolean
    'GUMP specyfic
    GUMP_type As Long
    button_id As Long
End Type

Private Type GUIImage
    X As Long
    Y As Long
    grh As Long
    height As Long
    width As Long
End Type

Private Type GUILabel
    text As String
    top As Long
    left As Long
    height As Long
    width As Long
    font As Long
    font_color As Long
    Background_Color As Long
    Alignment As FontAlignment
End Type

Private Type GUIComboBox
    text As String
    top As Long
    left As Long
    height As Long
    width As Long
    itemcount As Long
    items() As String
    currentitem As Long
    DropDownButton As GUIButton
    DropDownListOffset As Long
    DropDownListButtons(1 To 2) As GUIButton
    DropDownListScrollButton As GUIButton
    visible As Boolean
    DisplayList As Boolean
End Type

Private Type GUIWindow
    grhs(1 To 9) As Long
    top As Long
    left As Long
    width As Long
    height As Long
    buttoncount As Long
    textscount As Long
    labelcount As Long
    comboboxcount As Long
    progressbarcount As Long
    checkboxcount As Long
    optionbuttoncount As Long
    imagecount As Long
    buttons() As GUIButton
    texts() As GUITextBox
    labels() As GUILabel
    comboboxes() As GUIComboBox
    progressbars() As GUIProgressBar
    checkboxes() As GUICheckBox
    optionbuttons() As GUIOptionButton
    images() As GUIImage
    visible As Boolean
    Moveable As Boolean
    Resizeable As Boolean
    Closeable As Boolean
    'GUMP specyfic
    GUMP As Boolean
    grh_x(1 To 9) As Long
    grh_y(1 To 9) As Long
    grh_width(1 To 9) As Long
    grh_height(1 To 9) As Long
    grh_alphablending(1 To 9) As Long
End Type

Private TextWindow As Long
Private TextActive As Long

Private ClickedControlwindow As Long    'Used with the one below
Private ButtonClicked As Long

Private ComboBoxClicked As Long

'Stores the clicked combobox´s index
Private ComboBoxScrollerClicked As Long
'Stores the item of the dropdown list on which the mouse is set on
Private ComboBoxMouseOnItem As Long

Private WindowPressed As Long

Private WindowsCount As Long

Private Window_Active As Long

Private Windows() As GUIWindow

Private ZOrder() As Long

Public Function Text_Get_Active() As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Text_Get_Active = TextActive
End Function

Public Function Text_Active_Get_Window() As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Text_Active_Get_Window = TextWindow
End Function

Public Function Text_Set_Active(ByVal WindowID As Long, ByVal id As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Text_Check(WindowID, id) Then
        TextWindow = WindowID
        TextActive = id
        Text_Set_Active = True
        Exit Function
    End If
    
    Text_Set_Active = False
End Function

Public Function Windows_Get_Count() As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Windows_Get_Count = WindowsCount
End Function

Public Function Image_Create(ByVal WindowID As Long, ByVal X As Long, ByVal Y As Long, ByVal grh As Long, ByVal height As Long, ByVal width As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Windows(WindowID).imagecount = Windows(WindowID).imagecount + 1
        ReDim Preserve Windows(WindowID).images(0 To Windows(WindowID).imagecount)
        With Windows(WindowID).images(Windows(WindowID).imagecount)
            .grh = grh
            .X = X
            .Y = Y
            .height = height
            .width = width
        End With
        
        Image_Create = Windows(WindowID).imagecount
    Else
        Image_Create = -1
    End If
End Function

Private Function Image_Check(ByVal WindowID As Long, ByVal ImageID As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        If ImageID <= Windows(WindowID).imagecount And ImageID > -1 Then
            Image_Check = True
            Exit Function
        End If
    End If
    
    Image_Check = False
End Function

Public Function Image_Get_Grh(ByVal WindowID As Long, ByVal ImageID As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Image_Check(WindowID, ImageID) Then
        Image_Get_Grh = Windows(WindowID).images(ImageID).grh
    End If
End Function

Public Function Image_Get_Top(ByVal WindowID As Long, ByVal ImageID As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Image_Check(WindowID, ImageID) Then
        Image_Get_Top = Windows(WindowID).images(ImageID).Y
    End If
End Function

Public Function Image_Get_Left(ByVal WindowID As Long, ByVal ImageID As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Image_Check(WindowID, ImageID) Then
        Image_Get_Left = Windows(WindowID).images(ImageID).X
    End If
End Function

Public Function Image_Get_Width(ByVal WindowID As Long, ByVal ImageID As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Image_Check(WindowID, ImageID) Then
        Image_Get_Width = Windows(WindowID).images(ImageID).width
    End If
End Function

Public Function Image_Get_Height(ByVal WindowID As Long, ByVal ImageID As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Image_Check(WindowID, ImageID) Then
        Image_Get_Height = Windows(WindowID).images(ImageID).height
    End If
End Function

Public Function Text_Create(ByVal grh As Long, ByVal WindowID As Long, ByVal text As String, ByVal top As Long, ByVal left As Long, ByVal width As Long, ByVal height As Long, ByVal font As Long, _
                            ByVal FontColor As Long, Optional ByVal FontColor2 As Long = -1, Optional ByVal BackgroundColor As Long = -1, Optional ByVal password As Boolean) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Windows(WindowID).textscount = Windows(WindowID).textscount + 1
        ReDim Preserve Windows(WindowID).texts(0 To Windows(WindowID).textscount)
        With Windows(WindowID).texts(Windows(WindowID).textscount)
            .grh = grh
            .text = text
            .top = top
            .left = left
            .width = width
            .height = height
            .visible = True
            .font = font
            .FontColor = FontColor
            .FontColor2 = FontColor2
            .BackgroundColor = BackgroundColor
            .password = password
        End With
        
        Text_Create = Windows(WindowID).textscount
    Else
        Text_Create = -1
    End If
End Function

Public Function Window_Create(ByVal grh As Long, ByVal top As Long, ByVal left As Long, ByVal width As Long, ByVal height As Long, Optional ByVal Moveable As Boolean = False, _
                            Optional Resizeable As Boolean = False, Optional ByVal Closeable As Boolean = False, Optional ByVal GUMP As Boolean = False) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    WindowsCount = WindowsCount + 1
    ReDim Preserve Windows(0 To WindowsCount) As GUIWindow
    With Windows(WindowsCount)
        .top = top
        .left = left
        .width = width
        .height = height
        .buttoncount = -1
        .labelcount = -1
        .textscount = -1
        .comboboxcount = -1
        .progressbarcount = -1
        .checkboxcount = -1
        .optionbuttoncount = -1
        .imagecount = -1
        .visible = False
        .Moveable = Moveable
        .Resizeable = Resizeable
        .Closeable = .Closeable
        .GUMP = GUMP
        If GUMP Then
            .grhs(1) = grh
            .grhs(2) = grh + 1
            .grhs(3) = grh + 2
            .grhs(4) = grh + 3
            .grhs(5) = grh + 4
            .grhs(6) = grh + 5
            .grhs(7) = grh + 6
            .grhs(8) = grh + 7
            .grhs(9) = grh + 8
            
            'Calculate the size of each grh
            Dim h1 As Long
            Dim h2 As Long
            Dim h3 As Long
            Dim h4 As Long
            Dim h5 As Long
            Dim h6 As Long
            Dim h7 As Long
            Dim h8 As Long
            Dim h9 As Long
            Dim w1 As Long
            Dim w2 As Long
            Dim w3 As Long
            Dim w4 As Long
            Dim w5 As Long
            Dim w6 As Long
            Dim w7 As Long
            Dim w8 As Long
            Dim w9 As Long
            Dim temp_str As String
            Dim temp As Long
            Dim x_factor As Single
            Dim y_factor As Single
            
            'Get height and width of all grhs
            Engine.Grh_Info_Get grh, temp_str, temp, temp, w1, h1, temp
            Engine.Grh_Info_Get grh + 1, temp_str, temp, temp, w2, h2, temp
            Engine.Grh_Info_Get grh + 2, temp_str, temp, temp, w3, h3, temp
            Engine.Grh_Info_Get grh + 3, temp_str, temp, temp, w4, h4, temp
            Engine.Grh_Info_Get grh + 4, temp_str, temp, temp, w5, h5, temp
            Engine.Grh_Info_Get grh + 5, temp_str, temp, temp, w6, h6, temp
            Engine.Grh_Info_Get grh + 6, temp_str, temp, temp, w7, h7, temp
            Engine.Grh_Info_Get grh + 7, temp_str, temp, temp, w8, h8, temp
            Engine.Grh_Info_Get grh + 8, temp_str, temp, temp, w9, h9, temp
            
            'Calculate stretching factor
            x_factor = width / (w1 + w2 + w3)
            y_factor = height / (h1 + h2 + h3)
            
            'Set x and width for all grhs
            .grh_x(1) = .grh_x(4) = .grh_x(7) = left
            
            'First row
            .grh_width(1) = .grh_x(2) = w1 * x_factor
            .grh_width(2) = .grh_x(3) = w2 * x_factor
            .grh_width(3) = w3 * x_factor
            
            'Second row
            .grh_width(4) = .grh_x(5) = w4 * x_factor
            .grh_width(5) = .grh_x(6) = w5 * x_factor
            .grh_width(6) = w6 * x_factor
            
            'Third row
            .grh_width(7) = .grh_x(8) = w7 * x_factor
            .grh_width(8) = .grh_x(9) = w8 * x_factor
            .grh_width(9) = w9 * x_factor
            
            'Set y and height for all grhs
            .grh_y(1) = .grh_y(2) = .grh_y(3) = top
            'First column
            .grh_height(1) = .grh_y(4) = h1 * y_factor
            .grh_height(4) = .grh_y(7) = h4 * y_factor
            .grh_height(7) = h7 * y_factor
            
            'Second column
            .grh_height(2) = .grh_y(5) = h2 * y_factor
            .grh_height(5) = .grh_y(8) = h5 * y_factor
            .grh_height(8) = h8 * y_factor
            
            'Third column
            .grh_height(3) = .grh_y(3) = h3 * y_factor
            .grh_height(6) = .grh_y(6) = h6 * y_factor
            .grh_height(9) = h9 * y_factor
        Else
            .grhs(1) = grh
        End If
    End With
    Window_Create = WindowsCount
    
    'It´s set on top
    Dim LoopC As Long
    
    ReDim Preserve ZOrder(WindowsCount)
    For LoopC = 0 To WindowsCount - 1
        ZOrder(LoopC + 1) = ZOrder(LoopC)
    Next LoopC
    
    ZOrder(0) = WindowsCount
End Function

Public Function Window_Hit_Test(ByVal X As Long, ByVal Y As Long, Optional ByVal right_click As Boolean = False) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim LoopC As Long
    
    If WindowsCount = -1 Then
        Window_Hit_Test = -1
        Exit Function
    End If
    
    'We check in Z-Order
    For LoopC = 0 To UBound(ZOrder)
        If Windows(ZOrder(LoopC)).visible = True Then
            If Windows(ZOrder(LoopC)).top < Y And Windows(ZOrder(LoopC)).height + Windows(ZOrder(LoopC)).top > Y Then
                If Windows(ZOrder(LoopC)).left < X And Windows(ZOrder(LoopC)).left + Windows(ZOrder(LoopC)).width > X Then
                    'If we right clicked, and it's closeable, then we close it
                    If right_click Then
                        If Windows(ZOrder(LoopC)).Closeable Then
                            'Hide the window
                            Window_Show ZOrder(LoopC), False
                            Window_Active = -1
                            Window_Hit_Test = -1
                        End If
                        Exit Function
                    End If
                    Window_Active = ZOrder(LoopC)
                    Window_Hit_Test = ZOrder(LoopC)
                    Window_Set_On_Top ZOrder(LoopC)
                    Exit Function
                End If
            End If
        End If
    Next LoopC
    
    Window_Active = -1
    Window_Hit_Test = -1
End Function

Public Function Window_Show(ByVal id As Long, ByVal show As Boolean) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Windows(id).visible = show
        Window_Show = True
        Exit Function
    End If
    Window_Show = False
End Function

Public Function Window_Set_Movable(ByVal id As Long, ByVal Moveable As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        Windows(id).Moveable = Moveable
        Window_Set_Movable = True
        Exit Function
    End If
    Window_Set_Movable = False
End Function

Public Function Window_Set_Closeable(ByVal id As Long, ByVal Closeable As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        Windows(id).Closeable = Closeable
        Window_Set_Closeable = True
        Exit Function
    End If
    Window_Set_Closeable = False
End Function

Public Function Window_Set_Resizeable(ByVal id As Long, ByVal Resizeable As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        Windows(id).Resizeable = Resizeable
        Window_Set_Resizeable = True
        Exit Function
    End If
    Window_Set_Resizeable = False
End Function

Public Function Render_Visible() As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim m As Long
    Dim text As String
    
    If WindowsCount = -1 Then
        Render_Visible = False
        Exit Function
    End If
    
    For i = UBound(ZOrder) To 0 Step -1
        If Window_Get_Visible(ZOrder(i)) Then
            'Render the Window´s Grh
            If Windows(ZOrder(i)).GUMP Then
                For j = 1 To 9
                    Engine.GUI_Grh_Render_Advance Windows(ZOrder(i)).grhs(j), Windows(ZOrder(i)).grh_x(j), Windows(ZOrder(i)).grh_y(j), Windows(ZOrder(i)).grh_width(j), Windows(ZOrder(i)).grh_height(j), , Windows(ZOrder(i)).grh_alphablending(j)
                Next j
            Else
                Engine.GUI_Grh_Render_Advance Window_Get_Grh(ZOrder(i)), Window_Get_Left(ZOrder(i)), Window_Get_Top(ZOrder(i)), Window_Get_Width(ZOrder(i)), Window_Get_Height(ZOrder(i))
            End If
            
            'Render Images
            If Windows(ZOrder(i)).imagecount > -1 Then
                For j = 0 To Windows(ZOrder(i)).imagecount
                    Engine.GUI_Grh_Render_Advance Image_Get_Grh(ZOrder(i), j), Image_Get_Left(ZOrder(i), j), Image_Get_Top(ZOrder(i), j), Image_Get_Width(ZOrder(i), j), Image_Get_Height(ZOrder(i), j)
                Next j
            End If
            
            'Render Labels
            If Windows(ZOrder(i)).labelcount > -1 Then
                For j = 0 To Windows(ZOrder(i)).labelcount
                    If Label_Get_Background_Color(ZOrder(i), j) > -1 Then
                        Engine.GUI_Box_Filled_Render Window_Get_Left(ZOrder(i)) + Label_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + Label_Get_Top(ZOrder(i), j), Label_Get_Width(ZOrder(i), j), Label_Get_Height(ZOrder(i), j), Label_Get_Background_Color(ZOrder(i), j)
                    End If
                    Engine.GUI_Text_Render Label_Get_Text(ZOrder(i), j), Label_Get_Font(ZOrder(i), j), Window_Get_Left(ZOrder(i)) + Label_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + Label_Get_Top(ZOrder(i), j), Label_Get_Width(ZOrder(i), j), Label_Get_Height(ZOrder(i), j), Label_Get_Font_Color(ZOrder(i), j), Label_Get_Alignment(ZOrder(i), j), True
                Next j
            End If
            
            'Render Buttons
            If Windows(ZOrder(i)).buttoncount > -1 Then
                For j = 0 To Windows(ZOrder(i)).buttoncount
                    If Windows(ZOrder(i)).buttons(j).visible Then
                        If Button_Get_Background_Color(ZOrder(i), j) > -1 Then
                            Engine.GUI_Box_Filled_Render Window_Get_Left(ZOrder(i)) + Button_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + Button_Get_Top(ZOrder(i), j), Button_Get_Width(ZOrder(i), j), Button_Get_Height(ZOrder(i), j), Button_Get_Background_Color(ZOrder(i), j)
                        End If
                        If ButtonClicked = j And ClickedControlwindow = ZOrder(i) Then
                            Engine.GUI_Grh_Render_Advance Button_Get_Grh_Pressed(ZOrder(i), j), Window_Get_Left(ZOrder(i)) + Button_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + Button_Get_Top(ZOrder(i), j), Button_Get_Width(ZOrder(i), j), Button_Get_Height(ZOrder(i), j)
                        Else
                            Engine.GUI_Grh_Render_Advance Button_Get_Grh_Unpressed(ZOrder(i), j), Window_Get_Left(ZOrder(i)) + Button_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + Button_Get_Top(ZOrder(i), j), Button_Get_Width(ZOrder(i), j), Button_Get_Height(ZOrder(i), j)
                        End If
                        If Not Windows(ZOrder(i)).buttons(j).enabled Then
                            Engine.GUI_Grh_Render_Advance Button_Get_Grh_Pressed(ZOrder(i), j), Window_Get_Left(ZOrder(i)) + Button_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + Button_Get_Top(ZOrder(i), j), Button_Get_Width(ZOrder(i), j), Button_Get_Height(ZOrder(i), j)
                        End If
                        If Windows(ZOrder(i)).buttons(j).text <> "" Then
                            Engine.GUI_Text_Render Windows(ZOrder(i)).buttons(j).text, 1, Window_Get_Left(ZOrder(i)) + Button_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + Button_Get_Top(ZOrder(i), j), Button_Get_Width(ZOrder(i), j), Button_Get_Height(ZOrder(i), j), Button_Get_Font_Color(ZOrder(i), j), fa_center
                        End If
                    End If
                Next j
            End If
            
            'Render TextBoxes
            If Windows(ZOrder(i)).textscount > -1 Then
                For j = 0 To Windows(ZOrder(i)).textscount
                    If Windows(ZOrder(i)).texts(j).visible Then
                        If Text_Get_Background_Color(ZOrder(i), j) <> -1 Then
                            Engine.GUI_Box_Filled_Render Window_Get_Left(ZOrder(i)) + Text_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + Text_Get_Top(ZOrder(i), j), Text_Get_Width(ZOrder(i), j), Text_Get_Height(ZOrder(i), j), Text_Get_Background_Color(ZOrder(i), j)
                        End If
                        Engine.GUI_Grh_Render_Advance Text_Get_Grh(ZOrder(i), j), Window_Get_Left(ZOrder(i)) + Text_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + Text_Get_Top(ZOrder(i), j), Text_Get_Width(ZOrder(i), j), Text_Get_Height(ZOrder(i), j)
                        If Windows(ZOrder(i)).texts(j).text <> "" Then
                            text = ""
                            If Text_Get_Password(ZOrder(i), j) Then
                                For m = 1 To Len(Text_Get_Text(ZOrder(i), j))
                                    text = text + "*"
                                Next m
                            Else
                                text = Text_Get_Text(ZOrder(i), j)
                            End If
                            If Text_Get_Font_Color2(ZOrder(i), j) <> -1 Then
                                Engine.GUI_Text_Render text, Windows(ZOrder(i)).texts(j).font, Window_Get_Left(ZOrder(i)) + Text_Get_Left(ZOrder(i), j) + 1, Window_Get_Top(ZOrder(i)) + Text_Get_Top(ZOrder(i), j) + 1, Text_Get_Width(ZOrder(i), j), Text_Get_Height(ZOrder(i), j), Text_Get_Font_Color2(ZOrder(i), j), fa_left, False
                            End If
                            Engine.GUI_Text_Render text, Windows(ZOrder(i)).texts(j).font, Window_Get_Left(ZOrder(i)) + Text_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + Text_Get_Top(ZOrder(i), j), Text_Get_Width(ZOrder(i), j), Text_Get_Height(ZOrder(i), j), Text_Get_Font_Color(ZOrder(i), j), fa_left, False
                        End If
                    End If
                Next j
            End If
            
            'Render ComboBoxes
            If Windows(ZOrder(i)).comboboxcount > -1 Then
                For j = 0 To Windows(ZOrder(i)).comboboxcount
                    If Windows(ZOrder(i)).comboboxes(j).visible Then
                        Engine.GUI_Box_Filled_Render Window_Get_Left(ZOrder(i)) + ComboBox_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + ComboBox_Get_Top(ZOrder(i), j), ComboBox_Get_Width(ZOrder(i), j), ComboBox_Get_Height(ZOrder(i), j), &HFFFFFFFF
                        'Dropdown button
                        Engine.GUI_Box_Filled_Render Window_Get_Left(ZOrder(i)) + ComboBox_Get_Left(ZOrder(i), j) + ComboBox_Get_Width(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + ComboBox_Get_Top(ZOrder(i), j), COMBOBOX_BUTTON_WIDTH, ComboBox_Get_Height(ZOrder(i), j), &HFFC0C0C0
                        'Text
                        Engine.GUI_Text_Render Windows(ZOrder(i)).comboboxes(j).text, 1, Window_Get_Left(ZOrder(i)) + ComboBox_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + ComboBox_Get_Top(ZOrder(i), j), ComboBox_Get_Width(ZOrder(i), j), ComboBox_Get_Height(ZOrder(i), j), &HFF000000, fa_left, False
                        If Windows(ZOrder(i)).comboboxes(j).DisplayList Then
                            'Display box
                            Engine.GUI_Box_Filled_Render Window_Get_Left(ZOrder(i)) + ComboBox_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + ComboBox_Get_Top(ZOrder(i), j) + ComboBox_Get_Height(ZOrder(i), j), ComboBox_Get_Width(ZOrder(i), j), ComboBox_Get_Height(ZOrder(i), j) * COMBOBOX_ITEMS_DISPLAYED, &HFFFFFFFF
                            
                            'Display scrollbar
                            Engine.GUI_Box_Filled_Render Window_Get_Left(ZOrder(i)) + ComboBox_Get_Left(ZOrder(i), j) + ComboBox_Get_Width(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + ComboBox_Get_Top(ZOrder(i), j) + ComboBox_Get_Height(ZOrder(i), j) * 2, COMBOBOX_BUTTON_WIDTH, ComboBox_Get_Height(ZOrder(i), j) * (COMBOBOX_ITEMS_DISPLAYED - 2), &HFFE0E0E0
                            
                            'Display up/down buttons
                            Engine.GUI_Box_Filled_Render Window_Get_Left(ZOrder(i)) + ComboBox_Get_Left(ZOrder(i), j) + ComboBox_Get_Width(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + ComboBox_Get_Top(ZOrder(i), j) + ComboBox_Get_Height(ZOrder(i), j), COMBOBOX_BUTTON_WIDTH, ComboBox_Get_Height(ZOrder(i), j), &HFFC0C0C0
                            Engine.GUI_Box_Filled_Render Window_Get_Left(ZOrder(i)) + ComboBox_Get_Left(ZOrder(i), j) + ComboBox_Get_Width(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + ComboBox_Get_Top(ZOrder(i), j) + ComboBox_Get_Height(ZOrder(i), j) * COMBOBOX_ITEMS_DISPLAYED, COMBOBOX_BUTTON_WIDTH, ComboBox_Get_Height(ZOrder(i), j), &HFFC0C0C0
                            
                            'Display scroll button (only if needed)
                            If Windows(ZOrder(i)).comboboxes(j).DropDownListScrollButton.height < COMBOBOX_ITEMS_DISPLAYED * Windows(ZOrder(i)).comboboxes(j).height Then
                                Engine.GUI_Box_Filled_Render Window_Get_Left(ZOrder(i)) + ComboBox_Get_Left(ZOrder(i), j) + ComboBox_Get_Width(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + ComboBox_Get_Top(ZOrder(i), j) + ComboBox_Get_ScrollButton_Top(ZOrder(i), j), COMBOBOX_BUTTON_WIDTH, ComboBox_Get_ScrollButton_Height(ZOrder(i), j), &HFFC0C0C0
                            End If
                            
                            While k <= Windows(ZOrder(i)).comboboxes(j).itemcount And k < 5
                                'If the mouse is on thios item, the box should be blue, and the text white
                                If ComboBoxMouseOnItem = k Then
                                    'Display box
                                    Engine.GUI_Box_Filled_Render Window_Get_Left(ZOrder(i)) + ComboBox_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + ComboBox_Get_Top(ZOrder(i), j) + ComboBox_Get_Height(ZOrder(i), j) * (k + 1), ComboBox_Get_Width(ZOrder(i), j), ComboBox_Get_Height(ZOrder(i), j), &H80
                                    Engine.GUI_Text_Render Windows(ZOrder(i)).comboboxes(j).items(k + ComboBox_Get_Offset(ZOrder(i), j)), 1, Window_Get_Left(ZOrder(i)) + ComboBox_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + ComboBox_Get_Top(ZOrder(i), j) + ComboBox_Get_Height(ZOrder(i), j) * (k + 1), ComboBox_Get_Width(ZOrder(i), j), ComboBox_Get_Height(ZOrder(i), j), &HFFFFFFFF, fa_left
                                Else
                                    'Display items
                                    Engine.GUI_Text_Render Windows(ZOrder(i)).comboboxes(j).items(k + ComboBox_Get_Offset(ZOrder(i), j)), 1, Window_Get_Left(ZOrder(i)) + ComboBox_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + ComboBox_Get_Top(ZOrder(i), j) + ComboBox_Get_Height(ZOrder(i), j) * (k + 1), ComboBox_Get_Width(ZOrder(i), j), ComboBox_Get_Height(ZOrder(i), j), &HFF000000, fa_left
                                End If
                                
                                k = k + 1
                            Wend
                        End If
                    End If
                Next j
            End If
            
            'Render Progress Bars
            If Windows(ZOrder(i)).progressbarcount > -1 Then
                For j = 0 To Windows(ZOrder(i)).progressbarcount
                    If Windows(ZOrder(i)).progressbars(j).visible Then
                        Engine.GUI_Box_Filled_Render Window_Get_Left(ZOrder(i)) + ProgressBar_Get_Left(ZOrder(i), j) - 1, Window_Get_Top(ZOrder(i)) + ProgressBar_Get_Top(ZOrder(i), j) - 1, ProgressBar_Get_Width(ZOrder(i), j) + 2, ProgressBar_Get_Height(ZOrder(i), j) + 2, &HFFFFFF
                        Engine.GUI_Box_Filled_Render Window_Get_Left(ZOrder(i)) + ProgressBar_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + ProgressBar_Get_Top(ZOrder(i), j), ProgressBar_Get_Value(ZOrder(i), j) * ProgressBar_Get_Width(ZOrder(i), j) / ProgressBar_Get_Max(ZOrder(i), j), ProgressBar_Get_Height(ZOrder(i), j), &H80
                        Engine.GUI_Text_Render ProgressBar_Get_Percentage(ZOrder(i), j), 1, Window_Get_Left(ZOrder(i)) + ProgressBar_Get_Left(ZOrder(i), j) + ProgressBar_Get_Width(ZOrder(i), j) / 2 - 25, Window_Get_Top(ZOrder(i)) + ProgressBar_Get_Top(ZOrder(i), j), 50, ProgressBar_Get_Height(ZOrder(i), j), &HFFC0C0C0, fa_center
                    End If
                Next j
            End If
            
            'Render CheckBoxes
            If Windows(ZOrder(i)).checkboxcount > -1 Then
                For j = 0 To Windows(ZOrder(i)).checkboxcount
                    If Windows(ZOrder(i)).checkboxes(j).visible Then
                        If CheckBox_Get_Background_Color(ZOrder(i), j) <> -1 Then
                            Engine.GUI_Box_Filled_Render Window_Get_Left(ZOrder(i)) + CheckBox_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + CheckBox_Get_Top(ZOrder(i), j), CheckBox_Get_Width(ZOrder(i), j), CheckBox_Get_Height(ZOrder(i), j), CheckBox_Get_Background_Color(ZOrder(i), j), , , , False
                        End If
                        Engine.GUI_Text_Render CheckBox_Get_Text(ZOrder(i), j), CheckBox_Get_Font(ZOrder(i), j), Window_Get_Left(ZOrder(i)) + CheckBox_Get_Left(ZOrder(i), j) + CHECKBOX_WIDTH + 5, Window_Get_Top(ZOrder(i)) + CheckBox_Get_Top(ZOrder(i), j), CheckBox_Get_Width(ZOrder(i), j) - CHECKBOX_WIDTH - 5, CheckBox_Get_Height(ZOrder(i), j), CheckBox_Get_Font_Color(ZOrder(i), j), fa_left
                        If Not CheckBox_Get_Value(ZOrder(i), j) Then
                            If CheckBox_Get_Grh_Unchecked(ZOrder(i), j) > 0 Then
                                Engine.GUI_Grh_Render CheckBox_Get_Grh_Unchecked(ZOrder(i), j), Window_Get_Left(ZOrder(i)) + CheckBox_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + CheckBox_Get_Top(ZOrder(i), j)
                            Else
                                Engine.GUI_Box_Filled_Render Window_Get_Left(ZOrder(i)) + CheckBox_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + CheckBox_Get_Top(ZOrder(i), j), CHECKBOX_WIDTH, CHECKBOX_HEIGHT, &HFFFFFFFF, , , , False
                            End If
                        Else
                            If CheckBox_Get_Grh_Checked(ZOrder(i), j) > 0 Then
                                Engine.GUI_Grh_Render CheckBox_Get_Grh_Checked(ZOrder(i), j), Window_Get_Left(ZOrder(i)) + CheckBox_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + CheckBox_Get_Top(ZOrder(i), j)
                            Else
                                Engine.GUI_Box_Filled_Render Window_Get_Left(ZOrder(i)) + CheckBox_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + CheckBox_Get_Top(ZOrder(i), j), CHECKBOX_WIDTH, CHECKBOX_HEIGHT, &HFFFFFFFF, , , , False
                                Engine.GUI_Text_Render "x", CheckBox_Get_Font(ZOrder(i), j), Window_Get_Left(ZOrder(i)) + CheckBox_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + CheckBox_Get_Top(ZOrder(i), j) - 3, CHECKBOX_WIDTH, CHECKBOX_HEIGHT, &HFF000000, fa_center
                            End If
                        End If
                    End If
                Next j
            End If
            
            'Render OptionButtons
            If Windows(ZOrder(i)).optionbuttoncount > -1 Then
                For j = 0 To Windows(ZOrder(i)).optionbuttoncount
                    If Windows(ZOrder(i)).optionbuttons(j).visible Then
                        If OptionButton_Get_Background_Color(ZOrder(i), j) <> -1 Then
                            Engine.GUI_Box_Filled_Render Window_Get_Left(ZOrder(i)) + OptionButton_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + OptionButton_Get_Top(ZOrder(i), j), OPTIONBUTTON_WIDTH + 5 + OptionButton_Get_Width(ZOrder(i), j), OptionButton_Get_Height(ZOrder(i), j), OptionButton_Get_Background_Color(ZOrder(i), j)
                        End If
                        Engine.GUI_Text_Render OptionButton_Get_Text(ZOrder(i), j), OptionButton_Get_Font(ZOrder(i), j), Window_Get_Left(ZOrder(i)) + OptionButton_Get_Left(ZOrder(i), j) + OPTIONBUTTON_WIDTH + 5, Window_Get_Top(ZOrder(i)) + OptionButton_Get_Top(ZOrder(i), j), OptionButton_Get_Width(ZOrder(i), j) - OPTIONBUTTON_WIDTH - 5, OptionButton_Get_Height(ZOrder(i), j), OptionButton_Get_Font_Color(ZOrder(i), j), fa_left
                        If Not OptionButton_Get_Enabled(ZOrder(i), j) Then
                            If OptionButton_Get_Grh_Disabled(ZOrder(i), j) > 0 Then
                                Engine.GUI_Grh_Render OptionButton_Get_Grh_Disabled(ZOrder(i), j), Window_Get_Left(ZOrder(i)) + OptionButton_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + OptionButton_Get_Top(ZOrder(i), j)
                            Else
                                Engine.GUI_Box_Filled_Render_Advance Window_Get_Left(ZOrder(i)) + OptionButton_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + OptionButton_Get_Top(ZOrder(i), j), OPTIONBUTTON_WIDTH, OPTIONBUTTON_HEIGHT, &HFFC0C0C0, , , , False, General_Convert_Degrees_To_Radians(45)
                            End If
                        ElseIf Not OptionButton_Get_Value(ZOrder(i), j) Then
                            If OptionButton_Get_Grh_Unselected(ZOrder(i), j) > 0 Then
                                Engine.GUI_Grh_Render OptionButton_Get_Grh_Unselected(ZOrder(i), j), Window_Get_Left(ZOrder(i)) + OptionButton_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + OptionButton_Get_Top(ZOrder(i), j)
                            Else
                                Engine.GUI_Box_Filled_Render_Advance Window_Get_Left(ZOrder(i)) + OptionButton_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + OptionButton_Get_Top(ZOrder(i), j), OPTIONBUTTON_WIDTH, OPTIONBUTTON_HEIGHT, &HFFFFFFFF, , , , False, General_Convert_Degrees_To_Radians(45)
                            End If
                        Else
                            If OptionButton_Get_Grh_Selected(ZOrder(i), j) > 0 Then
                                Engine.GUI_Grh_Render OptionButton_Get_Grh_Selected(ZOrder(i), j), Window_Get_Left(ZOrder(i)) + OptionButton_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + OptionButton_Get_Top(ZOrder(i), j)
                            Else
                                Engine.GUI_Box_Filled_Render_Advance Window_Get_Left(ZOrder(i)) + OptionButton_Get_Left(ZOrder(i), j), Window_Get_Top(ZOrder(i)) + OptionButton_Get_Top(ZOrder(i), j), OPTIONBUTTON_WIDTH, OPTIONBUTTON_HEIGHT, &HFFFFFFFF, , , , False, General_Convert_Degrees_To_Radians(45)
                                Engine.GUI_Box_Filled_Render_Advance Window_Get_Left(ZOrder(i)) + OptionButton_Get_Left(ZOrder(i), j) + 2, Window_Get_Top(ZOrder(i)) + OptionButton_Get_Top(ZOrder(i), j) + 2, OPTIONBUTTON_WIDTH - 4, OPTIONBUTTON_HEIGHT - 4, &HFF000000, , , , False, General_Convert_Degrees_To_Radians(45)
                            End If
                        End If
                    End If
                Next j
            End If
        End If
    Next i
    
    Render_Visible = True
End Function

Public Function Window_Get_Button_Count(ByVal id As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Button_Count = Windows(id).buttoncount
    End If
End Function

Public Function Window_Get_Text_Count(ByVal id As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Text_Count = Windows(id).textscount
    End If
End Function

Public Function Window_Get_Label_Count(ByVal id As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Label_Count = Windows(id).labelcount
    End If
End Function

Public Function Window_Get_ComboBox_Count(ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        Window_Get_ComboBox_Count = Windows(id).comboboxcount
    End If
End Function

Public Function Window_Change_Alphablending(ByVal id As Long, ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal height As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        If Windows(id).GUMP Then
            Dim index As Long
            Dim LoopC As Long
            Dim LoopD As Long
            
            index = X + (Y - 1) * 3
            
            'Make usre it's within bounds
            If index > 0 And index < 10 Then
                For LoopD = 0 To height - 1
                    For LoopC = 0 To width - 1
                        If X + LoopC < 4 Then
                            Windows(id).grh_alphablending(index + LoopC) = Not Windows(id).grh_alphablending(index + LoopC)
                        End If
                    Next LoopC
                    index = index + 3
                    If index > 9 Then Exit For
                Next LoopD
                Window_Change_Alphablending = True
            End If
        End If
    End If
End Function

Public Function Window_Get_ProgressBar_Count(ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        Window_Get_ProgressBar_Count = Windows(id).progressbarcount
    End If
End Function

Public Function Window_Get_Image_Count(ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Image_Count = Windows(id).imagecount
    End If
End Function

Public Function Window_Get_Top(ByVal id As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Top = Windows(id).top
    End If
End Function

Public Function Window_Get_Grh(ByVal id As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Grh = Windows(id).grhs(1)
    End If
End Function

Public Function Window_Get_Movable(ByVal id As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Movable = Windows(id).Moveable
    End If
End Function

Public Function Window_Get_Visible(ByVal id As Long) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Visible = Windows(id).visible
    End If
End Function

Public Function Window_Get_Left(ByVal id As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Left = Windows(id).left
    End If
End Function

Public Function Window_Get_Width(ByVal id As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Width = Windows(id).width
    End If
End Function

Public Function Window_Get_Height(ByVal id As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Height = Windows(id).height
    End If
End Function

Public Function Window_Set_Grh(ByVal id As Long, ByVal grh As Long) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        'Allow it only if it's not a GUMP generated window
        If Not Windows(id).GUMP Then
            Windows(id).grhs(1) = grh
            Window_Set_Grh = True
            Exit Function
        End If
    End If
    Window_Set_Grh = False
End Function

Public Function Window_Set_Top(ByVal id As Long, ByVal top As Long) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Windows(id).top = top
        Window_Set_Top = True
        Exit Function
    End If
    Window_Set_Top = False
End Function

Public Function Window_Set_Left(ByVal id As Long, ByVal left As Long) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Windows(id).left = left
        Window_Set_Left = True
        Exit Function
    End If
    Window_Set_Left = False
End Function

Public Function Window_Mouse_Set_Down(ByVal id As Long) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        WindowPressed = id
        
        Window_Set_On_Top id
        
        Window_Mouse_Set_Down = True
        Exit Function
    End If
    Window_Mouse_Set_Down = False
End Function

Public Function Window_Mouse_Get_Pressed() As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    Window_Mouse_Get_Pressed = WindowPressed
End Function

Public Sub Window_Mouse_Set_Up()
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    WindowPressed = -1
    ButtonClicked = -1
    ComboBoxClicked = -1
    ComboBoxScrollerClicked = -1
End Sub

Public Sub Window_Mouse_Reset()
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    WindowPressed = -1
End Sub

Public Function Window_Get_Active() As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If WindowsCount > -1 Then
        Window_Get_Active = Window_Active
    End If
End Function

Public Function Window_Set_Active(ByVal WindowID As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Or WindowID = -1 Then
        Window_Active = WindowID
        Window_Set_Active = True
        Exit Function
    End If
    Window_Set_Active = False
End Function

Public Function Window_Set_Width(ByVal id As Long, ByVal width As Long) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Windows(id).width = width
        Window_Set_Width = True
        Exit Function
    End If
    Window_Set_Width = False
End Function

Public Function Window_Resize(ByVal id As Long, ByVal X As Long, ByVal Y As Long, ByVal width As Long, ByVal height As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        If Windows(id).Resizeable Then
            With Windows(id)
                .top = Y
                .left = X
                .width = width
                .height = height
                If .GUMP Then
                    'Calculate the size of each grh
                    Dim h1 As Long
                    Dim h2 As Long
                    Dim h3 As Long
                    Dim h4 As Long
                    Dim h5 As Long
                    Dim h6 As Long
                    Dim h7 As Long
                    Dim h8 As Long
                    Dim h9 As Long
                    Dim w1 As Long
                    Dim w2 As Long
                    Dim w3 As Long
                    Dim w4 As Long
                    Dim w5 As Long
                    Dim w6 As Long
                    Dim w7 As Long
                    Dim w8 As Long
                    Dim w9 As Long
                    Dim temp_str As String
                    Dim temp As Long
                    Dim x_factor As Single
                    Dim y_factor As Single
                    
                    'Get height and width of all grhs
                    Engine.Grh_Info_Get .grhs(1), temp_str, temp, temp, w1, h1, temp
                    Engine.Grh_Info_Get .grhs(2), temp_str, temp, temp, w2, h2, temp
                    Engine.Grh_Info_Get .grhs(3), temp_str, temp, temp, w3, h3, temp
                    Engine.Grh_Info_Get .grhs(4), temp_str, temp, temp, w4, h4, temp
                    Engine.Grh_Info_Get .grhs(5), temp_str, temp, temp, w5, h5, temp
                    Engine.Grh_Info_Get .grhs(6), temp_str, temp, temp, w6, h6, temp
                    Engine.Grh_Info_Get .grhs(7), temp_str, temp, temp, w7, h7, temp
                    Engine.Grh_Info_Get .grhs(8), temp_str, temp, temp, w8, h8, temp
                    Engine.Grh_Info_Get .grhs(9), temp_str, temp, temp, w9, h9, temp
                    
                    'Calculate stretching factor
                    x_factor = width / (w1 + w2 + w3)
                    y_factor = height / (h1 + h2 + h3)
                    
                    'Set x and width for all grhs
                    .grh_x(1) = .grh_x(4) = .grh_x(7) = X
                    
                    'First row
                    .grh_width(1) = .grh_x(2) = w1 * x_factor
                    .grh_width(2) = .grh_x(3) = w2 * x_factor
                    .grh_width(3) = w3 * x_factor
                    
                    'Second row
                    .grh_width(4) = .grh_x(5) = w4 * x_factor
                    .grh_width(5) = .grh_x(6) = w5 * x_factor
                    .grh_width(6) = w6 * x_factor
                    
                    'Third row
                    .grh_width(7) = .grh_x(8) = w7 * x_factor
                    .grh_width(8) = .grh_x(9) = w8 * x_factor
                    .grh_width(9) = w9 * x_factor
                    
                    'Set y and height for all grhs
                    .grh_y(1) = .grh_y(2) = .grh_y(3) = Y
                    'First column
                    .grh_height(1) = .grh_y(4) = h1 * y_factor
                    .grh_height(4) = .grh_y(7) = h4 * y_factor
                    .grh_height(7) = h7 * y_factor
                    
                    'Second column
                    .grh_height(2) = .grh_y(5) = h2 * y_factor
                    .grh_height(5) = .grh_y(8) = h5 * y_factor
                    .grh_height(8) = h8 * y_factor
                    
                    'Third column
                    .grh_height(3) = .grh_y(3) = h3 * y_factor
                    .grh_height(6) = .grh_y(6) = h6 * y_factor
                    .grh_height(9) = h9 * y_factor
                End If
            End With
            Window_Resize = True
            Exit Function
        End If
    End If
End Function

Public Function Window_Set_Height(ByVal id As Long, ByVal height As Long) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Windows(id).height = height
        Window_Set_Height = True
        Exit Function
    End If
    Window_Set_Height = False
End Function

Public Function Window_Move(ByVal id As Long, ByVal left As Long, ByVal top As Long) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) And ComboBoxScrollerClicked = -1 And Window_Get_Movable(id) Then
        Windows(id).left = left
        Windows(id).top = top
        Window_Move = True
        Exit Function
    End If
    Window_Move = False
End Function

Public Function Button_Hit_Test(ByVal WindowID As Long, ByVal X As Long, ByVal Y As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim i As Long
    
    If Window_Check(WindowID) Then
        Window_Set_On_Top WindowID
        For i = 0 To Windows(WindowID).buttoncount
            If Windows(WindowID).buttons(i).visible And Windows(WindowID).buttons(i).enabled Then
                If Windows(WindowID).top + Windows(WindowID).buttons(i).top < Y And Windows(WindowID).buttons(i).height + Windows(WindowID).buttons(i).top + Windows(WindowID).top > Y Then
                    If Windows(WindowID).left + Windows(WindowID).buttons(i).left < X And Windows(WindowID).left + Windows(WindowID).buttons(i).left + Windows(WindowID).buttons(i).width > X Then
                        Button_Hit_Test = i
                        ButtonClicked = i
                        ClickedControlwindow = WindowID
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
    
    Button_Hit_Test = -1
End Function

Public Function Button_Create(ByVal Grh_Unpressed As Long, ByVal WindowID As Long, ByVal text As String, ByVal FontColor As Long, ByVal top As Long, ByVal left As Long, ByVal width As Long, ByVal height As Long, Optional ByVal Grh_Pressed As Long = 0, Optional ByVal Background_Color As Long = -1, Optional enabled As Boolean = True, Optional button_id As Long = 0, Optional ByVal GUMP_type As Long = 2) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Windows(WindowID).buttoncount = Windows(WindowID).buttoncount + 1
        ReDim Preserve Windows(WindowID).buttons(0 To Windows(WindowID).buttoncount)
        With Windows(WindowID).buttons(Windows(WindowID).buttoncount)
            .Grh_Unpressed = Grh_Unpressed
            .text = text
            .top = top
            .left = left
            .width = width
            .height = height
            .font_color = FontColor
            .Background_Color = Background_Color
            .visible = True
            .enabled = enabled
            .button_id = button_id
            .GUMP_type = GUMP_type
            If Grh_Pressed = 0 Then
                .Grh_Pressed = .Grh_Unpressed
            Else
                .Grh_Pressed = Grh_Pressed
            End If
        End With
        Button_Create = Windows(WindowID).buttoncount
    Else
        Button_Create = -1
    End If
End Function

Public Function Button_Get_Font_Color(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_Font_Color = Windows(WindowID).buttons(id).font_color
    End If
End Function

Public Function Button_Get_Background_Color(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_Background_Color = Windows(WindowID).buttons(id).Background_Color
    End If
End Function

Public Function Button_Get_GUMP_Type(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_GUMP_Type = Windows(WindowID).buttons(id).GUMP_type
    End If
End Function

Public Function Button_Get_Button_ID(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_Button_ID = Windows(WindowID).buttons(id).button_id
    End If
End Function

Public Function Button_Get_Text(ByVal WindowID As Long, ByVal id As Long) As String
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_Text = Windows(WindowID).buttons(id).text
    End If
End Function

Public Function Button_Get_Top(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_Top = Windows(WindowID).buttons(id).top
    End If
End Function

Public Function Button_Get_Left(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_Left = Windows(WindowID).buttons(id).left
    End If
End Function

Public Function Button_Get_Width(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Windows(WindowID).buttoncount > id Or Windows(WindowID).buttoncount = id Then
        Button_Get_Width = Windows(WindowID).buttons(id).width
    End If
End Function

Public Function Button_Get_Height(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_Height = Windows(WindowID).buttons(id).height
    End If
End Function

Public Function Text_Get_Grh(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Grh = Windows(WindowID).texts(id).grh
    End If
End Function

Public Function Text_Get_Font_Color(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Font_Color = Windows(WindowID).texts(id).FontColor
    End If
End Function

Public Function Text_Get_Font_Color2(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Font_Color2 = Windows(WindowID).texts(id).FontColor2
    End If
End Function

Public Function Text_Get_Background_Color(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Background_Color = Windows(WindowID).texts(id).BackgroundColor
    End If
End Function

Public Function Text_Get_Password(ByVal WindowID As Long, ByVal id As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Password = Windows(WindowID).texts(id).password
    End If
End Function

Public Function Text_Get_Font(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Font = Windows(WindowID).texts(id).font
    End If
End Function

Public Function Text_Get_Text(ByVal WindowID As Long, ByVal id As Long) As String
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Text = Windows(WindowID).texts(id).text
    End If
End Function

Public Function Text_Get_Top(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Top = Windows(WindowID).texts(id).top
    End If
End Function

Public Function Text_Get_Left(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Left = Windows(WindowID).texts(id).left
    End If
End Function

Public Function Text_Get_Width(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Width = Windows(WindowID).texts(id).width
    End If
End Function

Public Function Text_Get_Height(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Height = Windows(WindowID).texts(id).height
    End If
End Function

Public Function Button_Get_Grh_Unpressed(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_Grh_Unpressed = Windows(WindowID).buttons(id).Grh_Unpressed
    End If
End Function

Public Function Button_Get_Grh_Pressed(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_Grh_Pressed = Windows(WindowID).buttons(id).Grh_Pressed
    End If
End Function

Public Function Text_Add_Text(ByVal WindowID As Long, ByVal id As Long, ByVal text As String) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan MArtín Sotuyo Dodero
'*********************************************************
    If Text_Check(WindowID, id) Then
        Windows(WindowID).texts(id).text = Windows(WindowID).texts(id).text + text
        Text_Add_Text = True
        Exit Function
    End If
    Text_Add_Text = False
End Function

Public Function Text_Remove_Char(ByVal WindowID As Long, ByVal id As Long, ByVal NumChars As Integer) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan Martín Sotuyo Dodero
'*********************************************************
    If Text_Check(WindowID, id) Then
        If Len(Windows(WindowID).texts(id).text) >= NumChars Then
            Windows(WindowID).texts(id).text = left(Windows(WindowID).texts(id).text, Len(Windows(WindowID).texts(id).text) - NumChars)
        Else
            Windows(WindowID).texts(id).text = ""
        End If
        Text_Remove_Char = True
        Exit Function
    End If
    Text_Remove_Char = False
End Function

Public Function Text_Hit_Test(ByVal WindowID As Long, ByVal X As Long, ByVal Y As Long) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    Dim i As Long
    
    If Window_Check(WindowID) Then
        Window_Set_On_Top WindowID
        For i = 0 To Windows(WindowID).textscount
            If Windows(WindowID).texts(i).visible = True Then
                If Windows(WindowID).top + Windows(WindowID).texts(i).top < Y And Windows(WindowID).texts(i).height + Windows(WindowID).texts(i).top + Windows(WindowID).top > Y Then
                    If Windows(WindowID).left + Windows(WindowID).texts(i).left < X And Windows(WindowID).left + Windows(WindowID).texts(i).left + Windows(WindowID).texts(i).width > X Then
                        Text_Hit_Test = i
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
    
    TextActive = -1
    Text_Hit_Test = -1
End Function

Private Function Window_Check(ByVal WindowID As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If WindowID <= WindowsCount And WindowID > -1 Then
        Window_Check = True
    Else
        Window_Check = False
    End If
End Function

Private Function Button_Check(ByVal WindowID As Long, ByVal ButtonID As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        If ButtonID <= Windows(WindowID).buttoncount And ButtonID > -1 Then
            Button_Check = True
            Exit Function
        End If
    End If
    
    Button_Check = False
End Function

Private Function Text_Check(ByVal WindowID As Long, ByVal TextID As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        If TextID <= Windows(WindowID).textscount And TextID > -1 Then
            Text_Check = True
            Exit Function
        End If
    End If
    
    Text_Check = False
End Function

Private Function ComboBox_Check(ByVal WindowID As Long, ByVal ComboBoxID As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        If ComboBoxID <= Windows(WindowID).comboboxcount And ComboBoxID > -1 Then
            ComboBox_Check = True
            Exit Function
        End If
    End If
    
    ComboBox_Check = False
End Function

Public Function ComboBox_Create(ByVal WindowID As Long, ByVal text As String, ByVal top As Long, ByVal left As Long, ByVal height As Long, ByVal width As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Windows(WindowID).comboboxcount = Windows(WindowID).comboboxcount + 1
        ReDim Preserve Windows(WindowID).comboboxes(0 To Windows(WindowID).comboboxcount)
        With Windows(WindowID).comboboxes(Windows(WindowID).comboboxcount)
            .currentitem = -1
            .itemcount = -1
            .text = text
            .left = left
            .height = height
            .top = top
            .width = width
            .DisplayList = False
            .visible = True
            
            'Buttons aren´t rendered as such, so we don´t store all values
            'Dropdown button
            With .DropDownButton
                .Grh_Pressed = 0    'An arrow Grh should be placed here
                .Grh_Unpressed = 0
            End With
            
            'Up button
            With .DropDownListButtons(1)
                .Grh_Pressed = 0    'An arrow Grh should be placed here
                .Grh_Unpressed = 0
            End With
            
            'Down button
            With .DropDownListButtons(2)
                .Grh_Pressed = 0    'An arrow Grh should be placed here
                .Grh_Unpressed = 0
            End With
            
            'Dropdown scroller button
            With .DropDownListScrollButton
                .left = left + width - COMBOBOX_BUTTON_WIDTH
                .top = top + height * 2 'It starts right below the down button
            End With
        End With
    Else
        ComboBox_Create = -1
    End If
End Function

Public Function ComboBox_Add_Item(ByVal WindowID As Long, ByVal id As Long, ByVal item As String) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ComboBox_Check(WindowID, id) Then
        Windows(WindowID).comboboxes(id).itemcount = Windows(WindowID).comboboxes(id).itemcount + 1
        ReDim Preserve Windows(WindowID).comboboxes(id).items(0 To Windows(WindowID).comboboxes(id).itemcount)
        Windows(WindowID).comboboxes(id).items(Windows(WindowID).comboboxes(id).itemcount) = item
        
        'Update the size of the scroller
        Windows(WindowID).comboboxes(id).DropDownListScrollButton.height = (COMBOBOX_ITEMS_DISPLAYED - 2) * Windows(WindowID).comboboxes(id).height / ((Windows(WindowID).comboboxes(id).itemcount + 1) / COMBOBOX_ITEMS_DISPLAYED)
        If Windows(WindowID).comboboxes(id).DropDownListScrollButton.height < 16 Then
            Windows(WindowID).comboboxes(id).DropDownListScrollButton.height = 16
        End If
        ComboBox_Add_Item = Windows(WindowID).comboboxes(id).itemcount
    Else
        ComboBox_Add_Item = -1
    End If
End Function

Public Function ComboBox_Remove_Item(ByVal WindowID As Long, ByVal id As Long, ByVal itemindex As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ComboBox_Check(WindowID, id) Then
        If itemindex <= Windows(WindowID).comboboxes(id).itemcount Then
            Windows(WindowID).comboboxes(id).itemcount = Windows(WindowID).comboboxes(id).itemcount - 1
            While itemindex < Windows(WindowID).comboboxes(id).itemcount
                Windows(WindowID).comboboxes(id).items(itemindex) = Windows(WindowID).comboboxes(id).items(itemindex + 1)
                itemindex = itemindex + 1
            Wend
            ReDim Preserve Windows(WindowID).comboboxes(id).items(0 To Windows(WindowID).comboboxes(id).itemcount)
            ComboBox_Remove_Item = True
            Exit Function
        End If
    End If
    ComboBox_Remove_Item = False
End Function

Public Function ComboBox_Select_Current_Item(ByVal WindowID As Long, ByVal id As Long, ByVal itemindex As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ComboBox_Check(WindowID, id) Then
        If itemindex <= Windows(WindowID).comboboxes(id).itemcount Then
            Windows(WindowID).comboboxes(id).text = Windows(WindowID).comboboxes(id).items(itemindex)
            Windows(WindowID).comboboxes(id).currentitem = itemindex
        End If
    End If
End Function

Public Function ComboBox_Get_Current_Item(ByVal WindowID As Long, ByVal id As Long) As String
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ComboBox_Check(WindowID, id) Then
        ComboBox_Get_Current_Item = Windows(WindowID).comboboxes(id).text
    End If
End Function

Public Function ComboBox_Get_Item(ByVal WindowID As Long, ByVal id As Long, ByVal itemindex As Long) As String
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ComboBox_Check(WindowID, id) Then
        If itemindex <= Windows(WindowID).comboboxes(id).itemcount Then
            ComboBox_Get_Item = Windows(WindowID).comboboxes(id).items(itemindex)
        End If
    End If
End Function

Public Function ComboBox_Get_Index(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ComboBox_Check(WindowID, id) Then
        ComboBox_Get_Index = Windows(WindowID).comboboxes(id).currentitem
    End If
End Function

Public Function ComboBox_Get_Top(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ComboBox_Check(WindowID, id) Then
        ComboBox_Get_Top = Windows(WindowID).comboboxes(id).top
    End If
End Function

Public Function ComboBox_Get_Left(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ComboBox_Check(WindowID, id) Then
        ComboBox_Get_Left = Windows(WindowID).comboboxes(id).left
    End If
End Function

Public Function ComboBox_Get_Height(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ComboBox_Check(WindowID, id) Then
        ComboBox_Get_Height = Windows(WindowID).comboboxes(id).height
    End If
End Function

Public Function ComboBox_Get_Width(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ComboBox_Check(WindowID, id) Then
        ComboBox_Get_Width = Windows(WindowID).comboboxes(id).width
    End If
End Function

Public Function ComboBox_Get_Item_Count(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ComboBox_Check(WindowID, id) Then
        ComboBox_Get_Item_Count = Windows(WindowID).comboboxes(id).itemcount
    End If
End Function

Public Function ComboBox_Get_Offset(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ComboBox_Check(WindowID, id) Then
        ComboBox_Get_Offset = Windows(WindowID).comboboxes(id).DropDownListOffset
    End If
End Function

Public Function ComboBox_Get_ScrollButton_Height(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ComboBox_Check(WindowID, id) Then
        ComboBox_Get_ScrollButton_Height = Windows(WindowID).comboboxes(id).DropDownListScrollButton.height
    End If
End Function

Public Function ComboBox_Get_ScrollButton_Top(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ComboBox_Check(WindowID, id) Then
        ComboBox_Get_ScrollButton_Top = Windows(WindowID).comboboxes(id).DropDownListScrollButton.top
    End If
End Function

Public Function ComboBox_ScrollButton_Move(ByVal WindowID As Long, ByVal id As Long, ByVal Y As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ComboBoxScrollerClicked > -1 And ComboBox_Check(WindowID, id) Then
        Windows(WindowID).comboboxes(id).DropDownListScrollButton.top = Y
        'Correct Pos
        If Windows(WindowID).comboboxes(id).DropDownListScrollButton.top < Windows(WindowID).comboboxes(id).height * 2 Then
            Windows(WindowID).comboboxes(id).DropDownListScrollButton.top = Windows(WindowID).comboboxes(id).height * 2
        End If
        If Windows(WindowID).comboboxes(id).DropDownListScrollButton.top + Windows(WindowID).comboboxes(id).DropDownListScrollButton.height > Windows(WindowID).comboboxes(id).height * COMBOBOX_ITEMS_DISPLAYED Then
            Windows(WindowID).comboboxes(id).DropDownListScrollButton.top = Windows(WindowID).comboboxes(id).height * COMBOBOX_ITEMS_DISPLAYED - Windows(WindowID).comboboxes(id).DropDownListScrollButton.height
        End If
        'Set the scroll offset
        Windows(WindowID).comboboxes(id).DropDownListOffset = (Windows(WindowID).comboboxes(id).DropDownListScrollButton.top - ComboBox_Get_Height(WindowID, id) * 2) / ComboBox_Get_Height(WindowID, id) / (COMBOBOX_ITEMS_DISPLAYED - 2) * (ComboBox_Get_Item_Count(WindowID, id) + 1)
        
        ComboBox_ScrollButton_Move = True
        Exit Function
    End If
    ComboBox_ScrollButton_Move = False
End Function

Public Function ComboBox_Get_ScrollButton_Clicked(ByVal WindowID As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        ComboBox_Get_ScrollButton_Clicked = ComboBoxScrollerClicked
    End If
End Function

Public Function ComboBox_Hit_Test(ByVal WindowID As Long, ByVal X As Long, ByVal Y As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim i As Long
    Dim hitted As Boolean
    
    ComboBox_Hit_Test = -1
    
    If Window_Check(WindowID) Then
        Window_Active = WindowID
        For i = 0 To Windows(WindowID).comboboxcount
            hitted = False
            If Windows(WindowID).comboboxes(i).visible = True Then
                'Check if the dropdown button was hitted
                If Windows(WindowID).top + Windows(WindowID).comboboxes(i).top < Y And Windows(WindowID).comboboxes(i).height + Windows(WindowID).comboboxes(i).top + Windows(WindowID).top > Y Then
                    If Windows(WindowID).left + Windows(WindowID).comboboxes(i).left + Windows(WindowID).comboboxes(i).width < X And Windows(WindowID).left + Windows(WindowID).comboboxes(i).left + Windows(WindowID).comboboxes(i).width + 16 > X Then
                        ComboBox_Hit_Test = i
                        hitted = True
                        'Change DisplayList value (only if necessary)
                        If ComboBoxClicked <> i Then
                            Windows(WindowID).comboboxes(i).DisplayList = Not Windows(WindowID).comboboxes(i).DisplayList
                            'Reset the offset
                            Windows(WindowID).comboboxes(i).DropDownListOffset = 0
                            'Reset the scroller pos
                            Windows(WindowID).comboboxes(i).DropDownListScrollButton.top = Windows(WindowID).comboboxes(i).height * 2
                            'Set it as clicked
                            ComboBoxClicked = i
                        End If
                        Exit Function
                    End If
                End If
                
                If Windows(WindowID).comboboxes(i).DisplayList Then
                    'Check if a scroll button or the scrollbar was hitted
                    If Windows(WindowID).left + Windows(WindowID).comboboxes(i).left + Windows(WindowID).comboboxes(i).width < X And Windows(WindowID).left + Windows(WindowID).comboboxes(i).left + Windows(WindowID).comboboxes(i).width + COMBOBOX_BUTTON_WIDTH > X Then
                        'Up button
                        If Windows(WindowID).top + Windows(WindowID).comboboxes(i).top + Windows(WindowID).comboboxes(i).height < Y And Windows(WindowID).comboboxes(i).top + Windows(WindowID).top + Windows(WindowID).comboboxes(i).height * 2 > Y Then
                            ComboBox_Hit_Test = i
                            hitted = True
                            Windows(WindowID).comboboxes(i).DropDownListOffset = Windows(WindowID).comboboxes(i).DropDownListOffset - 1
                            If Windows(WindowID).comboboxes(i).DropDownListOffset < 0 Then Windows(WindowID).comboboxes(i).DropDownListOffset = 0
                            'Update the dropdown list scroller pos
                            Windows(WindowID).comboboxes(i).DropDownListScrollButton.top = ComboBox_Get_Height(WindowID, i) * 2 + ComboBox_Get_Offset(WindowID, i) * ComboBox_Get_Height(WindowID, i) * (COMBOBOX_ITEMS_DISPLAYED - 2) / (ComboBox_Get_Item_Count(WindowID, i) + 1)
                        End If
                        'Down button
                        If Windows(WindowID).top + Windows(WindowID).comboboxes(i).top + Windows(WindowID).comboboxes(i).height * COMBOBOX_ITEMS_DISPLAYED < Y And Windows(WindowID).top + Windows(WindowID).comboboxes(i).top + Windows(WindowID).comboboxes(i).height * (COMBOBOX_ITEMS_DISPLAYED + 1) > Y Then
                            ComboBox_Hit_Test = i
                            hitted = True
                            Windows(WindowID).comboboxes(i).DropDownListOffset = Windows(WindowID).comboboxes(i).DropDownListOffset + 1
                            If Windows(WindowID).comboboxes(i).DropDownListOffset > Windows(WindowID).comboboxes(i).itemcount + 1 - COMBOBOX_ITEMS_DISPLAYED Then Windows(WindowID).comboboxes(i).DropDownListOffset = Windows(WindowID).comboboxes(i).itemcount + 1 - COMBOBOX_ITEMS_DISPLAYED
                            'Update the dropdown list scroller pos
                            Windows(WindowID).comboboxes(i).DropDownListScrollButton.top = ComboBox_Get_Height(WindowID, i) * 2 + ComboBox_Get_Offset(WindowID, i) * ComboBox_Get_Height(WindowID, i) * (COMBOBOX_ITEMS_DISPLAYED - 2) / (ComboBox_Get_Item_Count(WindowID, i) + 1)
                        End If
                        'Scrollbar
                        If Windows(WindowID).top + Windows(WindowID).comboboxes(i).top + Windows(WindowID).comboboxes(i).height * 2 < Y And Windows(WindowID).comboboxes(i).height * COMBOBOX_ITEMS_DISPLAYED + Windows(WindowID).top + Windows(WindowID).comboboxes(i).top > Y Then
                            'Check if the scroller button was hitted
                            If Windows(WindowID).top + Windows(WindowID).comboboxes(i).top + Windows(WindowID).comboboxes(i).DropDownListScrollButton.top < Y And Windows(WindowID).top + Windows(WindowID).comboboxes(i).top + Windows(WindowID).comboboxes(i).DropDownListScrollButton.top + Windows(WindowID).comboboxes(i).DropDownListScrollButton.height > Y Then
                                ComboBox_Hit_Test = i
                                hitted = True
                                WindowPressed = WindowID
                                ComboBoxScrollerClicked = i
                            End If
                            'Upside the scroller button
                            If Windows(WindowID).top + Windows(WindowID).comboboxes(i).top + Windows(WindowID).comboboxes(i).DropDownListScrollButton.top > Y And Windows(WindowID).top + Windows(WindowID).comboboxes(i).top + Windows(WindowID).comboboxes(i).height * 2 < Y Then
                                ComboBox_Hit_Test = i
                                hitted = True
                                Windows(WindowID).comboboxes(i).DropDownListOffset = Windows(WindowID).comboboxes(i).DropDownListOffset - COMBOBOX_ITEMS_DISPLAYED
                                If Windows(WindowID).comboboxes(i).DropDownListOffset < 0 Then Windows(WindowID).comboboxes(i).DropDownListOffset = 0
                                Windows(WindowID).comboboxes(i).DropDownListScrollButton.top = ComboBox_Get_Height(WindowID, i) * 2 + ComboBox_Get_Offset(WindowID, i) * ComboBox_Get_Height(WindowID, i) * (COMBOBOX_ITEMS_DISPLAYED - 2) / (ComboBox_Get_Item_Count(WindowID, i) + 1)
                            End If
                            'Below the scroller button
                            If Windows(WindowID).top + Windows(WindowID).comboboxes(i).top + Windows(WindowID).comboboxes(i).DropDownListScrollButton.top + Windows(WindowID).comboboxes(i).DropDownListScrollButton.height < Y And Windows(WindowID).top + Windows(WindowID).comboboxes(i).top + Windows(WindowID).comboboxes(i).height * COMBOBOX_ITEMS_DISPLAYED > Y Then
                                ComboBox_Hit_Test = i
                                hitted = True
                                Windows(WindowID).comboboxes(i).DropDownListOffset = Windows(WindowID).comboboxes(i).DropDownListOffset + COMBOBOX_ITEMS_DISPLAYED
                                If Windows(WindowID).comboboxes(i).DropDownListOffset > Windows(WindowID).comboboxes(i).itemcount + 1 - COMBOBOX_ITEMS_DISPLAYED Then Windows(WindowID).comboboxes(i).DropDownListOffset = Windows(WindowID).comboboxes(i).itemcount + 1 - COMBOBOX_ITEMS_DISPLAYED
                                Windows(WindowID).comboboxes(i).DropDownListScrollButton.top = ComboBox_Get_Height(WindowID, i) * 2 + ComboBox_Get_Offset(WindowID, i) * ComboBox_Get_Height(WindowID, i) * (COMBOBOX_ITEMS_DISPLAYED - 2) / (ComboBox_Get_Item_Count(WindowID, i) + 1)
                            End If
                        End If
                    End If
                    
                    
                    'Check if an item from the list was hitted
                    If Windows(WindowID).top + Windows(WindowID).comboboxes(i).top + Windows(WindowID).comboboxes(i).height < Y And Windows(WindowID).comboboxes(i).height * (COMBOBOX_ITEMS_DISPLAYED + 1) + Windows(WindowID).top + Windows(WindowID).comboboxes(i).top > Y Then
                        If Windows(WindowID).left + Windows(WindowID).comboboxes(i).left < X And Windows(WindowID).left + Windows(WindowID).comboboxes(i).left + Windows(WindowID).comboboxes(i).width > X Then
                            ComboBox_Hit_Test = i
                            hitted = True
                            'Check which item was hitted
                            ComboBox_Select_Current_Item WindowID, i, Fix((Y - Windows(WindowID).top - Windows(WindowID).comboboxes(i).top - Windows(WindowID).comboboxes(i).height) / Windows(WindowID).comboboxes(i).height) + Windows(WindowID).comboboxes(i).DropDownListOffset
                            'Reset all values
                            Windows(WindowID).comboboxes(i).DropDownListOffset = 0
                            Windows(WindowID).comboboxes(i).DisplayList = False
                        End If
                    End If
                End If
            End If
            If hitted = False And ComboBoxScrollerClicked <> i And WindowPressed <> WindowID Then
                Windows(WindowID).comboboxes(i).DisplayList = False
            End If
        Next i
    End If
    
    If ComboBox_Hit_Test = -1 And WindowPressed <> WindowID Then
        ComboBoxScrollerClicked = -1
    End If
End Function

Public Function ComboBox_Mouse_On_Top_Set(ByVal WindowID As Long, ByVal X As Long, ByVal Y As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim i As Long
    
    If Window_Check(WindowID) Then
        For i = 0 To Windows(WindowID).comboboxcount
            If Windows(WindowID).comboboxes(i).visible And Windows(WindowID).comboboxes(i).DisplayList Then
                'Check if mouse is on the list
                If Windows(WindowID).top + Windows(WindowID).comboboxes(i).top + Windows(WindowID).comboboxes(i).height < Y And Windows(WindowID).comboboxes(i).height * (COMBOBOX_ITEMS_DISPLAYED + 1) + Windows(WindowID).top + Windows(WindowID).comboboxes(i).top > Y Then
                    If Windows(WindowID).left + Windows(WindowID).comboboxes(i).left < X And Windows(WindowID).left + Windows(WindowID).comboboxes(i).left + Windows(WindowID).comboboxes(i).width > X Then
                        'Check on which item
                        ComboBoxMouseOnItem = Fix((Y - Windows(WindowID).top - Windows(WindowID).comboboxes(i).top - Windows(WindowID).comboboxes(i).height) / Windows(WindowID).comboboxes(i).height)
                        ComboBox_Mouse_On_Top_Set = True
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
    
    ComboBoxMouseOnItem = -1
    ComboBox_Mouse_On_Top_Set = False
End Function

Public Function ProgressBar_Create(ByVal WindowID As Long, ByVal top As Long, ByVal left As Long, ByVal width As Long, ByVal height As Long, ByVal Max As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Windows(WindowID).progressbarcount = Windows(WindowID).progressbarcount + 1
        ReDim Preserve Windows(WindowID).progressbars(0 To Windows(WindowID).progressbarcount)
        With Windows(WindowID).progressbars(Windows(WindowID).progressbarcount)
            .top = top
            .left = left
            .width = width
            .height = height
            .Max = Max
            'Max can´t be 0
            If .Max = 0 Then .Max = 1
            .value = 0
            .Percentage = "0 %"
            .visible = True
        End With
        ProgressBar_Create = Windows(WindowID).progressbarcount
    Else
        ProgressBar_Create = -1
    End If
End Function

Private Function ProgressBar_Check(ByVal WindowID As Long, ByVal ProgressBarID As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        If ProgressBarID <= Windows(WindowID).progressbarcount And ProgressBarID > -1 Then
            ProgressBar_Check = True
            Exit Function
        End If
    End If
    
    ProgressBar_Check = False
End Function

Public Function ProgressBar_Get_Top(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ProgressBar_Check(WindowID, id) Then
        ProgressBar_Get_Top = Windows(WindowID).progressbars(id).top
    End If
End Function

Public Function ProgressBar_Get_Left(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ProgressBar_Check(WindowID, id) Then
        ProgressBar_Get_Left = Windows(WindowID).progressbars(id).left
    End If
End Function

Public Function ProgressBar_Get_Width(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ProgressBar_Check(WindowID, id) Then
        ProgressBar_Get_Width = Windows(WindowID).progressbars(id).width
    End If
End Function

Public Function ProgressBar_Get_Height(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ProgressBar_Check(WindowID, id) Then
        ProgressBar_Get_Height = Windows(WindowID).progressbars(id).height
    End If
End Function

Public Function ProgressBar_Get_Value(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ProgressBar_Check(WindowID, id) Then
        ProgressBar_Get_Value = Windows(WindowID).progressbars(id).value
    End If
End Function

Public Function ProgressBar_Get_Max(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ProgressBar_Check(WindowID, id) Then
        ProgressBar_Get_Max = Windows(WindowID).progressbars(id).Max
    End If
End Function

Public Function ProgressBar_Get_Percentage(ByVal WindowID As Long, ByVal id As Long) As String
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ProgressBar_Check(WindowID, id) Then
        ProgressBar_Get_Percentage = Windows(WindowID).progressbars(id).Percentage
    End If
End Function

Public Function ProgressBar_Set_Value(ByVal WindowID As Long, ByVal id As Long, ByVal value As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ProgressBar_Check(WindowID, id) Then
        Windows(WindowID).progressbars(id).value = value
        If Windows(WindowID).progressbars(id).value > Windows(WindowID).progressbars(id).Max Then Windows(WindowID).progressbars(id).value = Windows(WindowID).progressbars(id).Max
        Windows(WindowID).progressbars(id).Percentage = str(Fix(Windows(WindowID).progressbars(id).value * 100 / Windows(WindowID).progressbars(id).Max)) & " %"
        ProgressBar_Set_Value = True
        Exit Function
    End If
    ProgressBar_Set_Value = False
End Function

Public Function ProgressBar_Set_Max(ByVal WindowID As Long, ByVal id As Long, ByVal Max As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If ProgressBar_Check(WindowID, id) Then
        Windows(WindowID).progressbars(id).Max = Max
        If Windows(WindowID).progressbars(id).Max < 1 Then Windows(WindowID).progressbars(id).Max = 1
        ProgressBar_Set_Max = True
        Exit Function
    End If
    ProgressBar_Set_Max = False
End Function

Public Function Label_Create(ByVal WindowID As Long, ByVal FontIndex As Long, ByVal FontColor As Long, ByVal text As String, ByVal top As Long, ByVal left As Long, ByVal width As Long, ByVal height As Long, ByVal Alignment As FontAlignment, Optional ByVal BackgroundColor As Long = -1) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Windows(WindowID).labelcount = Windows(WindowID).labelcount + 1
        ReDim Preserve Windows(WindowID).labels(0 To Windows(WindowID).labelcount)
        With Windows(WindowID).labels(Windows(WindowID).labelcount)
            .text = text
            .font = FontIndex
            .font_color = FontColor
            .Alignment = Alignment
            .Background_Color = BackgroundColor
            .top = top
            .left = left
            .height = height
            .width = width
        End With
        Label_Create = Windows(WindowID).labelcount
    Else
        Label_Create = -1
    End If
End Function

Public Function Label_Set_Text(ByVal WindowID As Long, ByVal id As Long, ByVal text As String) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Label_Check(WindowID, id) Then
        Windows(WindowID).labels(id).text = text
        Label_Set_Text = True
        Exit Function
    End If
    Label_Set_Text = False
End Function

Private Function Label_Check(ByVal WindowID As Long, ByVal LabelID As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        If LabelID <= Windows(WindowID).labelcount And LabelID > -1 Then
            Label_Check = True
            Exit Function
        End If
    End If
    Label_Check = False
End Function

Public Function Label_Get_Top(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Label_Check(WindowID, id) Then
        Label_Get_Top = Windows(WindowID).labels(id).top
    End If
End Function

Public Function Label_Get_Left(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Label_Check(WindowID, id) Then
        Label_Get_Left = Windows(WindowID).labels(id).left
    End If
End Function

Public Function Label_Get_Height(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Label_Check(WindowID, id) Then
        Label_Get_Height = Windows(WindowID).labels(id).height
    End If
End Function

Public Function Label_Get_Width(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Label_Check(WindowID, id) Then
        Label_Get_Width = Windows(WindowID).labels(id).width
    End If
End Function

Public Function Label_Get_Text(ByVal WindowID As Long, ByVal id As Long) As String
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Label_Check(WindowID, id) Then
        Label_Get_Text = Windows(WindowID).labels(id).text
    End If
End Function

Public Function Label_Get_Font(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Label_Check(WindowID, id) Then
        Label_Get_Font = Windows(WindowID).labels(id).font
    End If
End Function

Public Function Label_Get_Font_Color(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Label_Check(WindowID, id) Then
        Label_Get_Font_Color = Windows(WindowID).labels(id).font_color
    End If
End Function

Public Function Label_Get_Background_Color(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Label_Check(WindowID, id) Then
        Label_Get_Background_Color = Windows(WindowID).labels(id).Background_Color
    End If
End Function

Public Function Label_Get_Alignment(ByVal WindowID As Long, ByVal id As Long) As FontAlignment
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Label_Check(WindowID, id) Then
        Label_Get_Alignment = Windows(WindowID).labels(id).Alignment
    End If
End Function

Public Function CheckBox_Create(ByVal WindowID As Long, ByVal top As Long, ByVal left As Long, ByVal width As Long, ByVal height As Long, ByVal grh_unchecked As Long, ByVal grh_checked As Long, ByVal text As String, Optional ByVal BackgroundColor As Long = -1, Optional ByVal value As Boolean = False, Optional ByVal font As Long = 1, Optional ByVal font_color As Long, Optional ByVal switch_id As Long = 0) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Windows(WindowID).checkboxcount = Windows(WindowID).checkboxcount + 1
        ReDim Preserve Windows(WindowID).checkboxes(0 To Windows(WindowID).checkboxcount)
        With Windows(WindowID).checkboxes(Windows(WindowID).checkboxcount)
            .top = top
            .left = left
            .width = width
            .height = height
            .value = value
            .font_color = font_color
            .font_index = font
            .grh_unchecked = grh_unchecked
            .grh_checked = grh_checked
            .text = text
            .BackgroundColor = BackgroundColor
            .visible = True
            .enabled = True
            .switch_id = switch_id
        End With
        CheckBox_Create = Windows(WindowID).checkboxcount
    Else
        CheckBox_Create = -1
    End If
End Function

Private Function CheckBox_Check(ByVal WindowID As Long, ByVal CheckBoxID As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        If CheckBoxID <= Windows(WindowID).checkboxcount And CheckBoxID > -1 Then
            CheckBox_Check = True
            Exit Function
        End If
    End If
    
    CheckBox_Check = False
End Function

Public Function CheckBox_Get_Top(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Top = Windows(WindowID).checkboxes(id).top
    End If
End Function

Public Function CheckBox_Get_Switch_ID(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Switch_ID = Windows(WindowID).checkboxes(id).switch_id
    End If
End Function

Public Function CheckBox_Get_Left(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Left = Windows(WindowID).checkboxes(id).left
    End If
End Function

Public Function CheckBox_Get_Width(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Width = Windows(WindowID).checkboxes(id).width
    End If
End Function

Public Function CheckBox_Get_Height(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Height = Windows(WindowID).checkboxes(id).height
    End If
End Function

Public Function CheckBox_Get_Value(ByVal WindowID As Long, ByVal id As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Value = Windows(WindowID).checkboxes(id).value
    End If
End Function

Public Function CheckBox_Get_Enabled(ByVal WindowID As Long, ByVal id As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Enabled = Windows(WindowID).checkboxes(id).enabled
    End If
End Function

Public Function CheckBox_Hit_Test(ByVal WindowID As Long, ByVal X As Long, ByVal Y As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim i As Long
    
    If Window_Check(WindowID) Then
        Window_Set_On_Top WindowID
        For i = 0 To Windows(WindowID).checkboxcount
            If Windows(WindowID).checkboxes(i).visible = True And Windows(WindowID).checkboxes(i).enabled = True Then
                If Windows(WindowID).top + Windows(WindowID).checkboxes(i).top < Y And CHECKBOX_HEIGHT + Windows(WindowID).checkboxes(i).top + Windows(WindowID).top > Y Then
                    If Windows(WindowID).left + Windows(WindowID).checkboxes(i).left < X And Windows(WindowID).left + Windows(WindowID).checkboxes(i).left + CHECKBOX_WIDTH > X Then
                        CheckBox_Hit_Test = i
                        'Change value
                        Windows(WindowID).checkboxes(i).value = Not Windows(WindowID).checkboxes(i).value
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
    
    CheckBox_Hit_Test = -1
End Function

Public Function CheckBox_Get_Font(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Font = Windows(WindowID).checkboxes(id).font_index
    End If
End Function

Public Function CheckBox_Get_Font_Color(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Font_Color = Windows(WindowID).checkboxes(id).font_color
    End If
End Function

Public Function CheckBox_Get_Text(ByVal WindowID As Long, ByVal id As Long) As String
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Text = Windows(WindowID).checkboxes(id).text
    End If
End Function

Public Function CheckBox_Get_Grh_Unchecked(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Grh_Unchecked = Windows(WindowID).checkboxes(id).grh_unchecked
    End If
End Function

Public Function CheckBox_Get_Grh_Checked(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Grh_Checked = Windows(WindowID).checkboxes(id).grh_checked
    End If
End Function

Public Function CheckBox_Get_Background_Color(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Background_Color = Windows(WindowID).checkboxes(id).BackgroundColor
        Exit Function
    End If
    CheckBox_Get_Background_Color = -1
End Function

Public Function CheckBox_Get_Visible(ByVal WindowID As Long, ByVal id As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Visible = Windows(WindowID).checkboxes(id).visible
    End If
End Function

Public Function CheckBox_Set_Value(ByVal WindowID As Long, ByVal id As Long, ByVal value As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        Windows(WindowID).checkboxes(id).value = value
        CheckBox_Set_Value = True
        Exit Function
    End If
    CheckBox_Set_Value = False
End Function

Public Function CheckBox_Set_Visible(ByVal WindowID As Long, ByVal id As Long, ByVal value As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        Windows(WindowID).checkboxes(id).visible = value
        CheckBox_Set_Visible = True
        Exit Function
    End If
    CheckBox_Set_Visible = False
End Function

Public Function CheckBox_Set_Enabled(ByVal WindowID As Long, ByVal id As Long, ByVal value As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        Windows(WindowID).checkboxes(id).enabled = value
        CheckBox_Set_Enabled = True
        Exit Function
    End If
    CheckBox_Set_Enabled = False
End Function

Public Function Button_Set_Unpressed_Grh(ByVal WindowID As Long, ByVal id As Long, ByVal grh_index As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Windows(WindowID).buttons(id).Grh_Unpressed = grh_index
        Button_Set_Unpressed_Grh = True
        Exit Function
    End If
    
    Button_Set_Unpressed_Grh = False
End Function

Public Function Button_Set_Pressed_Grh(ByVal WindowID As Long, id As Long, grh_index As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Windows(WindowID).buttons(id).Grh_Pressed = grh_index
        Button_Set_Pressed_Grh = True
        Exit Function
    End If
    
    Button_Set_Pressed_Grh = False
End Function

Public Function Button_Set_Enabled(ByVal WindowID As Long, ByVal id As Long, ByVal enabled As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Windows(WindowID).buttons(id).enabled = enabled
        Button_Set_Enabled = True
        Exit Function
    End If
    
    Button_Set_Enabled = False
End Function

Public Function Button_Get_Enabled(ByVal WindowID As Long, ByVal id As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_Enabled = Windows(WindowID).buttons(id).enabled
    End If
End Function

Public Function Button_Set_Visible(ByVal WindowID As Long, ByVal id As Long, ByVal visible As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Windows(WindowID).buttons(id).visible = visible
        Button_Set_Visible = True
        Exit Function
    End If
    
    Button_Set_Visible = False
End Function

Public Function Window_Get_CheckBox_Count(ByVal WindowID As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Window_Get_CheckBox_Count = Windows(WindowID).buttoncount
        Exit Function
    End If
    
    Window_Get_CheckBox_Count = -1
End Function

Private Sub Window_Set_On_Top(ByVal WindowID As Long)
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim LoopC As Long
    
    'Move it to the top of the ZOrder list
    LoopC = UBound(ZOrder)
    
    'Find it´s previous index
    Do Until ZOrder(LoopC) = WindowID
        LoopC = LoopC - 1
    Loop
    
    'Move back all indexes untill we hit top of the list
    Do Until LoopC = 0
        ZOrder(LoopC) = ZOrder(LoopC - 1)
        LoopC = LoopC - 1
    Loop
    
    'Place new head of the list
    ZOrder(0) = WindowID
End Sub

Public Function OptionButton_Create(ByVal WindowID As Long, ByVal top As Long, ByVal left As Long, ByVal width As Long, ByVal height As Long, ByVal grh_unselected As Long, ByVal grh_selected As Long, ByVal grh_disabled As Long, ByVal text As String, Optional ByVal group_index As Long = 0, Optional ByVal BackgroundColor As Long = -1, Optional ByVal font As Long = 1, Optional ByVal font_color As Long, Optional ByVal switch_id As Long = 0) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Windows(WindowID).optionbuttoncount = Windows(WindowID).optionbuttoncount + 1
        ReDim Preserve Windows(WindowID).optionbuttons(0 To Windows(WindowID).optionbuttoncount)
        With Windows(WindowID).optionbuttons(Windows(WindowID).optionbuttoncount)
            .top = top
            .left = left
            .width = width
            .height = height
            .value = False
            .font_color = font_color
            .font_index = font
            .grh_unselected = grh_unselected
            .grh_selected = grh_selected
            .grh_disabled = grh_disabled
            .text = text
            .BackgroundColor = BackgroundColor
            .visible = True
            .enabled = True
            .group = group_index
            .switch_id = switch_id
        End With
        
        OptionButton_Create = Windows(WindowID).checkboxcount
        
        'If it´s the first OptionButton of a group, set it´s value to true
        Dim LoopC As Long
        For LoopC = 0 To Windows(WindowID).optionbuttoncount
            If Windows(WindowID).optionbuttons(LoopC).group = group_index And LoopC <> Windows(WindowID).optionbuttoncount Then
                Exit Function
            End If
        Next LoopC
        
        Windows(WindowID).optionbuttons(Windows(WindowID).optionbuttoncount).value = True
    Else
        OptionButton_Create = -1
    End If
End Function

Private Function OptionButton_Check(ByVal WindowID As Long, ByVal OptionButtonID As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        If OptionButtonID <= Windows(WindowID).optionbuttoncount And OptionButtonID > -1 Then
            OptionButton_Check = True
            Exit Function
        End If
    End If
    
    OptionButton_Check = False
End Function

Public Function OptionButton_Get_Top(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Top = Windows(WindowID).optionbuttons(id).top
    End If
End Function

Public Function OptionButton_Get_Switch_ID(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Switch_ID = Windows(WindowID).optionbuttons(id).switch_id
    End If
End Function

Public Function OptionButton_Get_Left(WindowID As Long, id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Left = Windows(WindowID).optionbuttons(id).left
    End If
End Function

Public Function OptionButton_Get_Width(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Width = Windows(WindowID).optionbuttons(id).width
    End If
End Function

Public Function OptionButton_Get_Height(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Height = Windows(WindowID).optionbuttons(id).height
    End If
End Function

Public Function OptionButton_Get_Value(ByVal WindowID As Long, ByVal id As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Value = Windows(WindowID).optionbuttons(id).value
    End If
End Function

Public Function OptionButton_Get_Visible(ByVal WindowID As Long, ByVal id As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Visible = Windows(WindowID).optionbuttons(id).visible
    End If
End Function

Public Function OptionButton_Hit_Test(ByVal WindowID As Long, ByVal X As Long, ByVal Y As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim i As Long
    Dim j As Long
    
    If Window_Check(WindowID) Then
        Window_Set_On_Top WindowID
        For i = 0 To Windows(WindowID).optionbuttoncount
            If Windows(WindowID).optionbuttons(i).visible = True And Windows(WindowID).optionbuttons(i).enabled = True Then
                If Windows(WindowID).top + Windows(WindowID).optionbuttons(i).top < Y And OPTIONBUTTON_HEIGHT + Windows(WindowID).optionbuttons(i).top + Windows(WindowID).top > Y Then
                    If Windows(WindowID).left + Windows(WindowID).optionbuttons(i).left < X And Windows(WindowID).left + Windows(WindowID).optionbuttons(i).left + OPTIONBUTTON_WIDTH > X Then
                        OptionButton_Hit_Test = i
                        'Change value
                        Windows(WindowID).optionbuttons(i).value = True
                        'Look for each option button in the group and set it´s value to false
                        For j = 0 To Windows(WindowID).optionbuttoncount
                            If Windows(WindowID).optionbuttons(j).group = Windows(WindowID).optionbuttons(i).group And i <> j Then
                                Windows(WindowID).optionbuttons(j).value = False
                            End If
                        Next j
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
    
    OptionButton_Hit_Test = -1
End Function

Public Function OptionButton_Get_Font(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Font = Windows(WindowID).optionbuttons(id).font_index
    End If
End Function

Public Function OptionButton_Get_Font_Color(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Font_Color = Windows(WindowID).optionbuttons(id).font_color
    End If
End Function

Public Function OptionButton_Get_Text(ByVal WindowID As Long, ByVal id As Long) As String
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Text = Windows(WindowID).optionbuttons(id).text
    End If
End Function

Public Function OptionButton_Get_Grh_Unselected(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Grh_Unselected = Windows(WindowID).optionbuttons(id).grh_unselected
    End If
End Function

Public Function OptionButton_Get_Grh_Selected(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Grh_Selected = Windows(WindowID).optionbuttons(id).grh_selected
    End If
End Function

Public Function OptionButton_Get_Background_Color(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Background_Color = Windows(WindowID).optionbuttons(id).BackgroundColor
    End If
End Function

Public Function OptionButton_Get_Selected_From_Group(ByVal WindowID As Long, ByVal group_index As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'Returns the index of the selected OptionButton within a group
'*********************************************************
    Dim LoopC As Long
    
    If Window_Check(WindowID) Then
        If Windows(WindowID).optionbuttoncount > -1 Then
            For LoopC = 0 To Windows(WindowID).optionbuttoncount
                If Windows(WindowID).optionbuttons(LoopC).group = group_index And _
                  Windows(WindowID).optionbuttons(LoopC).value = True Then
                    OptionButton_Get_Selected_From_Group = LoopC
                    Exit Function
                End If
            Next LoopC
        End If
    End If
    
    OptionButton_Get_Selected_From_Group = -1
End Function

Public Function OptionButton_Group_Disable(ByVal WindowID As Long, ByVal group_index As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim LoopC As Long
    
    If Windows(WindowID).optionbuttoncount = -1 Then
        OptionButton_Group_Disable = False
        Exit Function
    End If
    
    For LoopC = 0 To Windows(WindowID).optionbuttoncount
        If Windows(WindowID).optionbuttons(LoopC).group = group_index Then
            Windows(WindowID).optionbuttons(LoopC).enabled = False
        End If
    Next LoopC
    
    OptionButton_Group_Disable = True
End Function

Public Function OptionButton_Group_Enable(ByVal WindowID As Long, ByVal group_index As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim LoopC As Long
    
    If Windows(WindowID).optionbuttoncount = -1 Then
        OptionButton_Group_Enable = False
        Exit Function
    End If
    
    For LoopC = 0 To Windows(WindowID).optionbuttoncount
        If Windows(WindowID).optionbuttons(LoopC).group = group_index Then
            Windows(WindowID).optionbuttons(LoopC).enabled = True
        End If
    Next LoopC
    
    OptionButton_Group_Enable = True
End Function

Public Function OptionButton_Set_Visible(ByVal WindowID As Long, ByVal id As Long, ByVal value As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        Windows(WindowID).optionbuttons(id).visible = value
        OptionButton_Set_Visible = True
        Exit Function
    End If
    OptionButton_Set_Visible = False
End Function

Public Function OptionButton_Set_Enabled(ByVal WindowID As Long, ByVal id As Long, ByVal value As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        Windows(WindowID).optionbuttons(id).enabled = value
        OptionButton_Set_Enabled = True
        Exit Function
    End If
    OptionButton_Set_Enabled = False
End Function

Public Function OptionButton_Set_Selected(ByVal WindowID As Long, ByVal id As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim LoopC As Long
    
    If OptionButton_Check(WindowID, id) Then
        Windows(WindowID).optionbuttons(id).value = True
        For LoopC = 1 To Windows(WindowID).optionbuttoncount
            If Windows(WindowID).optionbuttons(LoopC).group = Windows(WindowID).optionbuttons(id).group And LoopC <> id Then
                Windows(WindowID).optionbuttons(LoopC).value = False
            End If
        Next LoopC
        OptionButton_Set_Selected = True
        Exit Function
    End If
    OptionButton_Set_Selected = False
End Function

Public Function Window_Get_OptionButton_Count(ByVal WindowID As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Window_Get_OptionButton_Count = Windows(WindowID).optionbuttoncount
    End If
End Function

Public Function OptionButton_Get_Grh_Disabled(ByVal WindowID As Long, ByVal id As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Grh_Disabled = Windows(WindowID).optionbuttons(id).grh_disabled
    End If
End Function

Public Function OptionButton_Get_Enabled(ByVal WindowID As Long, ByVal id As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Enabled = Windows(WindowID).optionbuttons(id).enabled
    End If
End Function

Public Function Text_Move_Focus() As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Text_Check(GUI.Text_Active_Get_Window, GUI.Text_Get_Active) Then
        If Windows(GUI.Text_Active_Get_Window).textscount > GUI.Text_Get_Active Then
            TextActive = TextActive + 1
        Else
            TextActive = 0
        End If
        
        Text_Move_Focus = True
        Exit Function
    End If
    
    Text_Move_Focus = False
End Function

Public Function Text_Set_Text(ByVal WindowID As Long, ByVal id As Long, ByVal text As String) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Text_Check(WindowID, id) Then
        Windows(WindowID).texts(id).text = text
        Text_Set_Text = True
        Exit Function
    End If
    
    Text_Set_Text = False
End Function

Public Function Label_Hit_Test(ByVal WindowID As Long, ByVal X As Long, ByVal Y As Long) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim i As Long
    
    If Window_Check(WindowID) Then
        Window_Set_On_Top WindowID
        For i = 0 To Windows(WindowID).labelcount
            If Windows(WindowID).top + Windows(WindowID).labels(i).top < Y And Windows(WindowID).labels(i).height + Windows(WindowID).labels(i).top + Windows(WindowID).top > Y Then
                If Windows(WindowID).left + Windows(WindowID).labels(i).left < X And Windows(WindowID).left + Windows(WindowID).labels(i).left + Windows(WindowID).labels(i).width > X Then
                    Label_Hit_Test = i
                    Exit Function
                End If
            End If
        Next i
    End If
    
    Label_Hit_Test = -1
End Function

Private Sub Class_Initialize()
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan Martín Sotuyo Dodero
'*********************************************************
    WindowsCount = -1
    ButtonClicked = -1
    ComboBoxClicked = -1
    ComboBoxScrollerClicked = -1
    ComboBoxMouseOnItem = -1
    ClickedControlwindow = -1
    
    Window_Active = -1
    
    TextWindow = -1
    TextActive = -1
End Sub
