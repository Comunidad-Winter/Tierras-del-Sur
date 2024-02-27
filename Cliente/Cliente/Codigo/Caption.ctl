VERSION 5.00
Begin VB.UserControl Caption 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8895
   ScaleHeight     =   3255
   ScaleWidth      =   8895
   Begin VB.Image IconIm 
      Height          =   375
      Left            =   240
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Caption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Const m_def_CaptionShadowed = True
Const m_def_CaptionAligmend = 0
Const m_def_CaptionGradientTyp = 0
Const m_def_CaptionGradientchangeColor = False
Const m_def_CaptionGradientStart = &HFFFFFF
Const m_def_CaptionGradientEnds = &H404040
Const m_def_Transparent = True
Const m_def_DisplayType = 0
Const m_def_DisplayVertikal = False
Const m_def_CaptionOutlined = True
Const m_def_CaptionINColor = vbWhite
Const m_def_CaptionOUTColor = vbBlack
Const m_def_CaptionTransparent = False
Const m_def_AutoSize = False
Const m_def_Caption = "0"

Dim m_CaptionShadowed As Boolean
Dim m_CaptionAligmend As CaptionAlig
Dim m_CaptionGradientTyp As TypGrad
Dim m_CaptionGradientchangeColor As Boolean
Dim m_CaptionGradientStart As OLE_COLOR
Dim m_CaptionGradientEnds As OLE_COLOR
Dim m_Transparent As Boolean
Dim m_DisplayType As DispTyp
Dim m_DisplayVertikal As Boolean
Dim m_CaptionOutlined As Boolean
Dim m_CaptionINColor As OLE_COLOR
Dim m_CaptionOUTColor As OLE_COLOR
Dim m_CaptionTransparent As Boolean
Dim m_AutoSize As Boolean
Dim m_Caption As String

Public Enum DispTyp
            [Left to Right]
            [Right To Left]
            [Top To Bottom]
            [Bottom To Top]
End Enum

Public Enum CaptionAlig
                OnLeft
                Centered
                OnRight
End Enum

Public Enum TypGrad
            [None]
            [Horizontal]
            [Vertical]
End Enum

Public Enum InBord
            [Flat]
            [Inside]
            [Reized]
            [More Reized]
            [Extern Reized]
            [Out]
            [More Out]
            [Extern Out]
End Enum

Public Enum MousePoin
                    [default]
                    [Arrow]
                    [Cross]
                    [I beam]
                    [Icon]
                    [Resize]
                    [Size NE SW]
                    [Size N S]
                    [Size NW SE]
                    [Size W E]
                    [Up arrow]
                    [Hourglass]
                    [No drop]
                    [Arrow and Hourglass]
                    [Arrow and Question Mark]
                    [Size All]
                    [Custom] = 99
End Enum

Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event Click()
Event DblClick()

Dim BorderPlusL
Dim BorderPlusT
Dim BorderPlusW
Dim BorderPlusH

Const m_def_Borders = 0

Dim m_Borders As InBord
Dim CaptionWork As Boolean

Dim OutLine As Integer
Private Con_LinkKeys As String

Public Property Get Con_LinkKey() As String
  Con_LinkKey = Con_LinkKeys
End Property

Public Property Let Con_LinkKey(NewValue As String)
  Con_LinkKeys = NewValue
End Property

Private Sub IconIm_Click()
    RaiseEvent Click
End Sub

Private Sub IconIm_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub IconIm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x + IconIm.left, y + IconIm.top)
End Sub

Private Sub IconIm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x + IconIm.left, y + IconIm.top)
End Sub

Private Sub IconIm_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x + IconIm.left, y + IconIm.top)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    LookAutoSize
    DrawText
End Property

Public Property Get Borders() As InBord
    Borders = m_Borders
End Property

Public Property Let Borders(ByVal New_Borders As InBord)
    m_Borders = New_Borders
    PropertyChanged "Borders"
    BorderControlCap
        LookAutoSize
'    BorderControlCap
    UserControl_Resize
End Property

Public Property Get Transparent() As Boolean
    Transparent = m_Transparent
End Property
'
Public Property Let Transparent(ByVal New_Transparent As Boolean)
    m_Transparent = New_Transparent
    PropertyChanged "Transparent"
    
    UserControl_Resize
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get AutoSize() As Boolean
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
        LookAutoSize
    DrawText
End Property

Public Property Get font() As font
    Set font = UserControl.font
End Property

Public Property Set font(ByVal New_Font As font)
'If UserControl.Font.Name <> New_Font.Name Then
    Set UserControl.font = New_Font
    PropertyChanged "Font"
    LookAutoSize
    DrawText
'End If
End Property

Public Property Get CaptionINColor() As OLE_COLOR
    CaptionINColor = m_CaptionINColor
End Property

Public Property Let CaptionINColor(ByVal New_CaptionINColor As OLE_COLOR)
    m_CaptionINColor = New_CaptionINColor
    PropertyChanged "CaptionINColor"
    UserControl.ForeColor = m_CaptionINColor
    DrawText
End Property

Public Property Get CaptionOUTColor() As OLE_COLOR
    CaptionOUTColor = m_CaptionOUTColor
End Property

Public Property Let CaptionOUTColor(ByVal New_CaptionOUTColor As OLE_COLOR)
    m_CaptionOUTColor = New_CaptionOUTColor
    PropertyChanged "CaptionOUTColor"
    DrawText
End Property

Public Property Get CaptionTransparent() As Boolean
    CaptionTransparent = m_CaptionTransparent
End Property

Public Property Let CaptionTransparent(ByVal New_CaptionTransparent As Boolean)
    m_CaptionTransparent = New_CaptionTransparent
    PropertyChanged "CaptionTransparent"
    DrawText
End Property

Public Property Get CaptionOutlined() As Boolean
    CaptionOutlined = m_CaptionOutlined
End Property

Public Property Let CaptionOutlined(ByVal New_CaptionOutlined As Boolean)
    m_CaptionOutlined = New_CaptionOutlined
    PropertyChanged "CaptionOutlined"
    OutLine = 0
    If m_CaptionOutlined = True Then OutLine = 30: Call LookAutoSize
    DrawText
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    DrawText
End Property

Public Property Get DisplayVertikal() As Boolean
    DisplayVertikal = m_DisplayVertikal
End Property

Public Property Let DisplayVertikal(ByVal New_DisplayVertikal As Boolean)
    m_DisplayVertikal = New_DisplayVertikal
    PropertyChanged "DisplayVertikal"
    DspH = 0
    DspW = 0
    If m_DisplayVertikal = True Then
        If m_DisplayType = [Left to Right] Or m_DisplayType = [Right To Left] Then
            If m_DisplayType = [Left to Right] Then
                DisplayType = [Top To Bottom]
            Else
                DisplayType = [Bottom To Top]
            End If
        End If
    Else
        If m_DisplayType = [Bottom To Top] Or m_DisplayType = [Top To Bottom] Then
            If m_DisplayType = [Top To Bottom] Then
                DisplayType = [Left to Right]
            Else
                DisplayType = [Right To Left]
            End If
        End If
    End If
    
    LookAutoSize
    DrawText
End Property

Public Property Get DisplayType() As DispTyp
    DisplayType = m_DisplayType
End Property

Public Property Let DisplayType(ByVal New_DisplayType As DispTyp)
    m_DisplayType = New_DisplayType
    PropertyChanged "DisplayType"
    
    If m_DisplayVertikal = True Then
        If m_DisplayType = [Left to Right] Or m_DisplayType = [Right To Left] Then
            DisplayVertikal = False
            DspH = 0
            DspW = 0
        End If
    End If
    If m_DisplayVertikal = False Then
        If m_DisplayType = [Bottom To Top] Or m_DisplayType = [Top To Bottom] Then
        
            DisplayVertikal = True
            DspH = 0
            DspW = 0
        End If
    End If
    LookAutoSize
    DrawText
End Property

Public Function TextHeight(ByVal Str As String) As Single
    If m_CaptionOutlined = True Then
        TextHeight = UserControl.TextHeight(Str) + 30
    Else
        TextHeight = UserControl.TextHeight(Str)
    End If
End Function

Public Function TextWidth(ByVal Str As String) As Single
    If m_CaptionOutlined = True Then
        TextWidth = UserControl.TextWidth(Str) + 30
    Else
        TextWidth = UserControl.TextWidth(Str)
    End If
End Function

Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As StdPicture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
    Set IconIm.MouseIcon = UserControl.MouseIcon
End Property

Public Property Get MousePointer() As MousePoin
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePoin)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
    IconIm.MousePointer = UserControl.MousePointer
End Property

Public Property Get CaptionGradientStart() As OLE_COLOR
    CaptionGradientStart = m_CaptionGradientStart
End Property

Public Property Let CaptionGradientStart(ByVal New_CaptionGradientStart As OLE_COLOR)
    m_CaptionGradientStart = New_CaptionGradientStart
    PropertyChanged "CaptionGradientStart"
    DrawText
End Property

Public Property Get CaptionGradientEnds() As OLE_COLOR
    CaptionGradientEnds = m_CaptionGradientEnds
End Property

Public Property Let CaptionGradientEnds(ByVal New_CaptionGradientEnds As OLE_COLOR)
    m_CaptionGradientEnds = New_CaptionGradientEnds
    PropertyChanged "CaptionGradientEnds"
    DrawText
End Property

Private Sub UserControl_InitProperties()
    m_Caption = ambient.DisplayName
    m_AutoSize = m_def_AutoSize
    Set UserControl.font = ambient.font
    m_CaptionINColor = m_def_CaptionINColor
    m_CaptionOUTColor = m_def_CaptionOUTColor
    m_CaptionTransparent = m_def_CaptionTransparent
    m_CaptionOutlined = m_def_CaptionOutlined
    m_DisplayVertikal = m_def_DisplayVertikal
    m_DisplayType = m_def_DisplayType
    m_Transparent = m_def_Transparent
    m_CaptionGradientStart = m_def_CaptionGradientStart
    m_CaptionGradientEnds = m_def_CaptionGradientEnds
    m_CaptionGradientchangeColor = m_def_CaptionGradientchangeColor
    m_CaptionGradientTyp = m_def_CaptionGradientTyp
    m_CaptionAligmend = m_def_CaptionAligmend
    m_CaptionShadowed = m_def_CaptionShadowed
    IconIm.Width = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
OutLine = 0
IconIm.Width = 0
    IconIm.Stretch = False
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_Borders = PropBag.ReadProperty("Borders", m_def_Borders)
    m_Transparent = PropBag.ReadProperty("Transparent", m_def_Transparent)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    Set UserControl.font = PropBag.ReadProperty("Font", ambient.font)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_CaptionINColor = PropBag.ReadProperty("CaptionINColor", m_def_CaptionINColor)
    m_CaptionOUTColor = PropBag.ReadProperty("CaptionOUTColor", m_def_CaptionOUTColor)
    m_CaptionTransparent = PropBag.ReadProperty("CaptionTransparent", m_def_CaptionTransparent)
    m_CaptionOutlined = PropBag.ReadProperty("CaptionOutlined", m_def_CaptionOutlined)
    m_DisplayVertikal = PropBag.ReadProperty("DisplayVertikal", m_def_DisplayVertikal)
    m_DisplayType = PropBag.ReadProperty("DisplayType", m_def_DisplayType)
    UserControl.ForeColor = m_CaptionINColor
    If m_CaptionOutlined = True Then OutLine = 30
    
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_CaptionGradientStart = PropBag.ReadProperty("CaptionGradientStart", m_def_CaptionGradientStart)
    m_CaptionGradientEnds = PropBag.ReadProperty("CaptionGradientEnds", m_def_CaptionGradientEnds)
    m_CaptionGradientchangeColor = PropBag.ReadProperty("CaptionGradientchangeColor", m_def_CaptionGradientchangeColor)
    m_CaptionGradientTyp = PropBag.ReadProperty("CaptionGradientTyp", m_def_CaptionGradientTyp)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_CaptionAligmend = PropBag.ReadProperty("CaptionAligmend", m_def_CaptionAligmend)
    m_CaptionShadowed = PropBag.ReadProperty("CaptionShadowed", m_def_CaptionShadowed)
    Set IconIm.Picture = PropBag.ReadProperty("CaptionPicture", Nothing)
    
    Set IconIm.MouseIcon = UserControl.MouseIcon
    IconIm.MousePointer = UserControl.MousePointer
    LookAutoSize
    DrawText
End Sub

Private Sub UserControl_Resize()
'BorderControlCap
'IconIm.Move 0 + BorderPlusL, ((UserControl.ScaleHeight / 2) - (IconIm.Height / 2)) + BorderPlusT
If CaptionWork = True Then Exit Sub
DrawText
End Sub

Private Sub LookForDispleyType()
    DspH = 0
    DspW = 0
    If m_DisplayVertikal = True Then
        For SS = 1 To Len(m_Caption)
            DspH = DspH + UserControl.TextHeight(mid(m_Caption, SS, 1)) '/ 1.4
            If DspW < UserControl.TextWidth(mid(m_Caption, SS, 1)) Then DspW = UserControl.TextWidth(mid(m_Caption, SS, 1))
        Next
    End If
End Sub

Private Sub LookAutoSize()
Dim ImWi, ImHei, Aly, Alw
LookForDispleyType
    If m_AutoSize = True Then
        CaptionWork = True
    If IconIm.Picture > 0 Then
        ImWi = IconIm.Width
        ImHei = IconIm.Height
        Alw = IconIm.Width - UserControl.TextWidth(m_Caption)
        Aly = IconIm.Height - UserControl.TextHeight(m_Caption)
        If m_CaptionShadowed = True Then
            Aly = Aly + 45
        End If
    End If
    If m_DisplayVertikal = True Then
        UserControl.size (DspW + (BorderPlusL) + (BorderPlusW) + (OutLine * 2)) + Alw, _
                          (DspH + BorderPlusT) + (BorderPlusH) + OutLine + ImHei
    Else
        UserControl.size (UserControl.TextWidth(m_Caption) + (BorderPlusL) + (BorderPlusW) + OutLine) + ImWi, _
                         (UserControl.TextHeight(m_Caption) + (BorderPlusT) + (BorderPlusH) + OutLine) + Aly
    End If
        CaptionWork = False
    End If
End Sub


Private Sub UserControl_Show()
'IconIm.Move 0, (UserControl.ScaleHeight / 2 - IconIm.Height / 2)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Borders", m_Borders, m_def_Borders)
    Call PropBag.WriteProperty("Transparent", m_Transparent, m_def_Transparent)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
    Call PropBag.WriteProperty("Font", UserControl.font, ambient.font)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("CaptionINColor", m_CaptionINColor, m_def_CaptionINColor)
    Call PropBag.WriteProperty("CaptionOUTColor", m_CaptionOUTColor, m_def_CaptionOUTColor)
    Call PropBag.WriteProperty("CaptionTransparent", m_CaptionTransparent, m_def_CaptionTransparent)
    Call PropBag.WriteProperty("CaptionOutlined", m_CaptionOutlined, m_def_CaptionOutlined)
    Call PropBag.WriteProperty("DisplayVertikal", m_DisplayVertikal, m_def_DisplayVertikal)
    Call PropBag.WriteProperty("DisplayType", m_DisplayType, m_def_DisplayType)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("CaptionGradientStart", m_CaptionGradientStart, m_def_CaptionGradientStart)
    Call PropBag.WriteProperty("CaptionGradientEnds", m_CaptionGradientEnds, m_def_CaptionGradientEnds)
    Call PropBag.WriteProperty("CaptionGradientchangeColor", m_CaptionGradientchangeColor, m_def_CaptionGradientchangeColor)
    Call PropBag.WriteProperty("CaptionGradientTyp", m_CaptionGradientTyp, m_def_CaptionGradientTyp)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("CaptionAligmend", m_CaptionAligmend, m_def_CaptionAligmend)
    Call PropBag.WriteProperty("CaptionShadowed", m_CaptionShadowed, m_def_CaptionShadowed)
    Call PropBag.WriteProperty("CaptionPicture", IconIm.Picture, Nothing)
End Sub

Private Sub DrawText()
Dim PoLeft As Single, PoTop As Single, DD As Single

UserControl.Cls
If m_CaptionGradientTyp <> None Then DrawColor m_CaptionGradientchangeColor

Dim Capp As String
Capp = m_Caption
Dim CuW, CuH

PoLeft = BorderPlusL
PoTop = BorderPlusT
DD = (BorderPlusT) + OutLine

    Select Case m_DisplayType
                    Case 0
                        DrawProcent Capp, PoLeft, PoTop
                        CuW = UserControl.TextWidth(Capp)
                        CuH = UserControl.TextHeight(Capp)
                    Case 1
                        Dim ST As String
                        
                        For SS = 0 To Len(Capp) - 1
                            ST = ST + mid(Capp, Len(Capp) - SS, 1)
                        Next
                        DrawProcent ST, PoLeft, PoTop
                        CuW = UserControl.TextWidth(ST)
                        CuH = UserControl.TextHeight(ST)
                    Case 2
                    PoLeft = (UserControl.ScaleWidth / 2 - DspW / 2)
                        For SS = 1 To Len(Capp)
                            If CuW < UserControl.TextWidth(mid(Capp, SS, 1)) Then CuW = UserControl.TextWidth(mid(Capp, SS, 1))
                            DrawProcent mid(Capp, SS, 1), Fix(UserControl.ScaleWidth / 2 - UserControl.TextWidth(mid(Capp, SS, 1)) / 2), DD '/ 1.41221
                            DD = DD + UserControl.TextHeight(mid(Capp, SS, 1)) '/ 1.4
                            CuH = DD
                        Next
                    
                    Case 3
                    PoLeft = (UserControl.ScaleWidth / 2 - DspW / 2)
                        For SS = 0 To Len(Capp) - 1
                            If CuW < UserControl.TextWidth(mid(Capp, Len(Capp) - SS, 1)) Then CuW = UserControl.TextWidth(mid(Capp, Len(Capp) - SS, 1))
                            DrawProcent mid(Capp, Len(Capp) - SS, 1), Fix(UserControl.ScaleWidth / 2 - UserControl.TextWidth(mid(Capp, Len(Capp) - SS, 1)) / 2), DD ' / 1.41221
                            DD = DD + UserControl.TextHeight(mid(Capp, Len(Capp) - SS, 1)) ' / 1.4
                            CuH = DD
                        Next
    End Select
CurX = BorderPlusL
CurY = BorderPlusT
CurW = UserControl.ScaleWidth - BorderPlusL * 2
CurH = UserControl.ScaleHeight - BorderPlusT * 2
    
End Sub

Private Sub DrawProcent(TextToPrint As String, DspX As Single, DspY As Single)
Dim Captio, AliX, AliY
Captio = TextToPrint
Dim FORC As OLE_COLOR
FORC = UserControl.ForeColor
Dim PosLeft
Dim PosTop
Dim CapOutCol As OLE_COLOR
Dim ImW
If IconIm.Picture > 0 Then ImW = IconIm.Width
    BorderControlCap
If m_DisplayVertikal = False Then
Select Case m_CaptionAligmend
            Case 0
                AliX = 0
                AliY = Fix(Fix(UserControl.ScaleHeight / 2 - UserControl.TextHeight(Captio) / 2)) - DspY
            Case 1
                If m_CaptionShadowed = True Then
                    AliX = Fix(Fix(((UserControl.ScaleWidth - 45) / 2) - (UserControl.TextWidth(Captio) / 2))) - ImW / 2 - DspX
                    AliY = Fix(Fix(((UserControl.ScaleHeight - 30) / 2) - (UserControl.TextHeight(Captio) / 2))) - DspY
                Else
                    AliX = Fix(Fix((UserControl.ScaleWidth / 2) - (UserControl.TextWidth(Captio) / 2))) - ImW / 2 - DspX
                    AliY = Fix(Fix((UserControl.ScaleHeight / 2) - (UserControl.TextHeight(Captio) / 2))) - DspY
                End If
            Case 2
                If m_CaptionShadowed = True Then
                    AliX = Fix(UserControl.ScaleWidth - UserControl.TextWidth(Captio)) - (DspX * 2) - ImW - 45
                    AliY = Fix(Fix(UserControl.ScaleHeight / 2 - UserControl.TextHeight(Captio) / 2)) - DspY - 30
                Else
                    AliX = Fix(UserControl.ScaleWidth - UserControl.TextWidth(Captio)) - (DspX * 2) - ImW
                    AliY = Fix(Fix((UserControl.ScaleHeight / 2) - (UserControl.TextHeight(Captio) / 2))) - DspY
                End If
                
End Select
End If
PosLeft = ImW + DspX + AliX
PosTop = DspY + AliY

CapOutCol = CaptionOUTColor
If UserControl.Enabled = False Then
    CapOutCol = &H808080
Else
    If m_CaptionShadowed = True Then
        UserControl.ForeColor = &H808080
        UserControl.CurrentX = PosLeft + 45
        UserControl.CurrentY = PosTop + 45
        UserControl.Print Captio
    End If
End If
If m_CaptionOutlined = True Then
    If m_DisplayVertikal = False Then
        PosLeft = ImW + AliX + DspX + 15
        PosTop = AliY + DspY + 15
    End If

    UserControl.ForeColor = CapOutCol
    UserControl.CurrentX = PosLeft - 15
    UserControl.CurrentY = PosTop - 15
    UserControl.Print Captio
    UserControl.CurrentX = PosLeft - 15
    UserControl.CurrentY = PosTop
    UserControl.Print Captio
    UserControl.CurrentX = PosLeft
    UserControl.CurrentY = PosTop - 15
    UserControl.Print Captio
    UserControl.CurrentX = PosLeft - 15
    UserControl.CurrentY = PosTop + 15
    UserControl.Print Captio
    UserControl.CurrentX = PosLeft + 15
    UserControl.CurrentY = PosTop - 15
    UserControl.Print Captio
    UserControl.CurrentX = PosLeft + 15
    UserControl.CurrentY = PosTop
    UserControl.Print Captio
    UserControl.CurrentX = PosLeft + 15
    UserControl.CurrentY = PosTop
    UserControl.Print Captio
    UserControl.CurrentX = PosLeft
    UserControl.CurrentY = PosTop + 15
    UserControl.Print Captio
    UserControl.CurrentX = PosLeft + 15
    UserControl.CurrentY = PosTop + 15
    UserControl.Print Captio
End If
If CaptionTransparent = True Then
    UserControl.ForeColor = UserControl.BackColor
Else
    UserControl.ForeColor = CaptionINColor
End If
If UserControl.Enabled = False Then
    UserControl.ForeColor = &HE0E0E0
End If
'IconIm.Move UserControl.ScaleWidth, (UserControl.ScaleHeight)
UserControl.CurrentX = PosLeft
UserControl.CurrentY = PosTop
UserControl.Print Captio
'UserControl.BackColor = vbWhite ' UserControl.Point(0, 0)
If IconIm.Picture > 0 Then
    UserControl.PaintPicture IconIm.Picture, 0 + BorderPlusL, ((UserControl.ScaleHeight / 2) - (IconIm.Height / 2)) + BorderPlusT, IconIm.Width, IconIm.Height
End If

If m_Transparent = True Then
        UserControl.MaskColor = UserControl.BackColor
        UserControl.MaskPicture = UserControl.Image
        UserControl.BackStyle = 0
Else
        UserControl.BackStyle = 1

End If
UserControl.ForeColor = FORC
BorderControlCap
'IconIm.Move 0 + BorderPlusL, ((UserControl.ScaleHeight / 2) - (IconIm.Height / 2)) + BorderPlusT

End Sub
Private Sub BorderControlCap()
'Exit Sub
Dim Scal
Dim UCSw, UCSh
'    UserControl.Cls
    Scal = UserControl.ScaleMode
    ScaleMode = 1 'twip
    UCSw = UserControl.ScaleWidth - ShadowWi
    UCSh = UserControl.ScaleHeight - ShadowHei

If m_Borders = [Flat] Then
    BorderPlusL = 0
    BorderPlusT = 0
    BorderPlusW = 0
    BorderPlusH = 0
    GoTo Exodos
End If
If m_Borders = [Inside] Then
    BorderPlusL = 30
    BorderPlusT = 30
    BorderPlusW = 30
    BorderPlusH = 30
    UserControl.Line (0, 0)-(UCSw, 0), vb3DShadow
    UserControl.Line (0, 15)-(UCSw, 15), vb3DDKShadow
    UserControl.Line (0, 0)-(0, UCSh), vb3DShadow
    UserControl.Line (15, 15)-(15, UCSh), vb3DDKShadow
    
    UserControl.Line (UCSw - 15, 0)-(UCSw - 15, UCSh), vb3DHighlight
    UserControl.Line (UCSw - 30, 15)-(UCSw - 30, UCSh), vb3DLight
    UserControl.Line (15, UCSh - 15)-(UCSw - 15, UCSh - 15), vb3DHighlight
    UserControl.Line (30, UCSh - 30)-(UCSw - 30, UCSh - 30), vb3DLight
    GoTo Exodos
End If
If m_Borders = [Reized] Then
    BorderPlusL = 30
    BorderPlusT = 30
    BorderPlusW = 30
    BorderPlusH = 30
    UserControl.Line (0, 0)-(UCSw - 15, 0), vb3DHighlight
    UserControl.Line (0, 15)-(UCSw - 15, 15), vb3DDKShadow
    UserControl.Line (0, 0)-(0, UCSh - 15), vb3DHighlight
    UserControl.Line (15, 15)-(15, UCSh - 15), vb3DDKShadow
    
    UserControl.Line (UCSw - 30, 0)-(UCSw - 30, UCSh - 15), vb3DHighlight
    UserControl.Line (UCSw - 15, 0)-(UCSw - 15, UCSh), vb3DDKShadow
    UserControl.Line (0, UCSh - 15)-(UCSw, UCSh - 15), vb3DDKShadow
    UserControl.Line (30, UCSh - 30)-(UCSw - 15, UCSh - 30), vb3DHighlight
    GoTo Exodos
End If
If m_Borders = [Out] Then
    BorderPlusL = 15
    BorderPlusT = 15
    BorderPlusW = 15
    BorderPlusH = 15
    UserControl.Line (0, 0)-(UCSw, 0), vb3DDKShadow
    UserControl.Line (0, 0)-(0, UCSh), vb3DDKShadow
    UserControl.Line (UCSw - 15, 0)-(UCSw - 15, UCSh), vb3DDKShadow
    UserControl.Line (15, UCSh - 15)-(UCSw - 15, UCSh - 15), vb3DDKShadow
    GoTo Exodos
End If
If m_Borders = [More Out] Then
    BorderPlusL = 45
    BorderPlusT = 45
    BorderPlusW = 45
    BorderPlusH = 45
    UserControl.Line (0, 0)-(UCSw, 0), vb3DDKShadow
    UserControl.Line (0, 15)-(UCSw, 15), vb3DShadow
    UserControl.Line (0, 30)-(UCSw, 30), vb3DLight

    UserControl.Line (0, 0)-(0, UCSh), vb3DDKShadow
    UserControl.Line (15, 15)-(15, UCSh), vb3DShadow
    UserControl.Line (30, 30)-(30, UCSh), vb3DLight

    UserControl.Line (UCSw - 15, 0)-(UCSw - 15, UCSh), vb3DDKShadow
    UserControl.Line (UCSw - 30, 15)-(UCSw - 30, UCSh), vb3DShadow
    UserControl.Line (UCSw - 45, 30)-(UCSw - 45, UCSh), vb3DLight

    UserControl.Line (15, UCSh - 15)-(UCSw - 15, UCSh - 15), vb3DDKShadow
    UserControl.Line (30, UCSh - 30)-(UCSw - 30, UCSh - 30), vb3DShadow
    UserControl.Line (45, UCSh - 45)-(UCSw - 45, UCSh - 45), vb3DLight
   GoTo Exodos
End If
If m_Borders = [More Reized] Then
    BorderPlusL = 45
    BorderPlusT = 45
    BorderPlusW = 45
    BorderPlusH = 45
    UserControl.Line (0, 0)-(UCSw, 0), vb3DHighlight
    UserControl.Line (0, 15)-(UCSw, 15), vb3DShadow
    UserControl.Line (0, 30)-(UCSw, 30), vb3DDKShadow

    UserControl.Line (0, 0)-(0, UCSh), vb3DHighlight
    UserControl.Line (15, 15)-(15, UCSh), vb3DShadow
    UserControl.Line (30, 30)-(30, UCSh - 45), vb3DDKShadow

    UserControl.Line (UCSw - 15, 0)-(UCSw - 15, UCSh), vb3DDKShadow
    UserControl.Line (UCSw - 30, 15)-(UCSw - 30, UCSh), vb3DShadow
    UserControl.Line (UCSw - 45, 30)-(UCSw - 45, UCSh), vb3DHighlight

    UserControl.Line (0, UCSh - 15)-(UCSw - 15, UCSh - 15), vb3DDKShadow
    UserControl.Line (30, UCSh - 30)-(UCSw - 30, UCSh - 30), vb3DShadow
    UserControl.Line (30, UCSh - 45)-(UCSw - 45, UCSh - 45), vb3DHighlight
   GoTo Exodos
End If
If m_Borders = [Extern Out] Then
    BorderPlusL = 60
    BorderPlusT = 60
    BorderPlusW = 60
    BorderPlusH = 60
    UserControl.Line (0, 0)-(UCSw, 0), vb3DDKShadow
    UserControl.Line (0, 15)-(UCSw, 15), vb3DShadow
    UserControl.Line (0, 30)-(UCSw, 30), vb3DLight
    UserControl.Line (0, 45)-(UCSw, 45), vb3DHighlight

    UserControl.Line (0, 0)-(0, UCSh), vb3DDKShadow
    UserControl.Line (15, 15)-(15, UCSh), vb3DShadow
    UserControl.Line (30, 30)-(30, UCSh), vb3DLight
    UserControl.Line (45, 45)-(45, UCSh), vb3DHighlight

    UserControl.Line (UCSw - 15, 0)-(UCSw - 15, UCSh), vb3DDKShadow
    UserControl.Line (UCSw - 30, 15)-(UCSw - 30, UCSh), vb3DShadow
    UserControl.Line (UCSw - 45, 30)-(UCSw - 45, UCSh), vb3DLight
    UserControl.Line (UCSw - 60, 45)-(UCSw - 60, UCSh), vb3DHighlight

    UserControl.Line (15, UCSh - 15)-(UCSw - 15, UCSh - 15), vb3DDKShadow
    UserControl.Line (30, UCSh - 30)-(UCSw - 30, UCSh - 30), vb3DShadow
    UserControl.Line (45, UCSh - 45)-(UCSw - 45, UCSh - 45), vb3DLight
    UserControl.Line (60, UCSh - 60)-(UCSw - 60, UCSh - 60), vb3DHighlight
End If
If m_Borders = [Extern Reized] Then
    BorderPlusL = 60
    BorderPlusT = 60
    BorderPlusW = 60
    BorderPlusH = 60
    UserControl.Line (0, 0)-(UCSw, 0), vb3DHighlight
    UserControl.Line (0, 15)-(UCSw, 15), vb3DLight
    UserControl.Line (0, 30)-(UCSw, 30), vb3DShadow
    UserControl.Line (0, 45)-(UCSw, 45), vb3DDKShadow

    UserControl.Line (0, 0)-(0, UCSh), vb3DHighlight
    UserControl.Line (15, 15)-(15, UCSh), vb3DLight
    UserControl.Line (30, 30)-(30, UCSh), vb3DShadow
    UserControl.Line (45, 45)-(45, UCSh), vb3DDKShadow

    UserControl.Line (UCSw - 15, 0)-(UCSw - 15, UCSh), vb3DDKShadow
    UserControl.Line (UCSw - 30, 15)-(UCSw - 30, UCSh), vb3DShadow
    UserControl.Line (UCSw - 45, 30)-(UCSw - 45, UCSh), vb3DLight
    UserControl.Line (UCSw - 60, 45)-(UCSw - 60, UCSh), vb3DHighlight

    UserControl.Line (0, UCSh - 15)-(UCSw - 15, UCSh - 15), vb3DDKShadow
    UserControl.Line (15, UCSh - 30)-(UCSw - 30, UCSh - 30), vb3DShadow
    UserControl.Line (30, UCSh - 45)-(UCSw - 45, UCSh - 45), vb3DLight
    UserControl.Line (45, UCSh - 60)-(UCSw - 60, UCSh - 60), vb3DHighlight
End If
Exodos:
ScaleMode = Scal
'UserControl.Refresh
End Sub


Private Sub DrawColor(Colorchange As Boolean)
Dim VR, VG, VB As Single
Dim Color1 As OLE_COLOR, Color2 As OLE_COLOR
Dim DR, DG, DB, DR2, DG2, DB2 As Integer
Dim temp As Long
CurW = UserControl.ScaleWidth - BorderPlusL * 2
CurH = UserControl.ScaleHeight - BorderPlusT * 2

If Colorchange = True Then
    Color2 = m_CaptionGradientStart
    Color1 = m_CaptionGradientEnds
Else
    Color2 = m_CaptionGradientEnds
    Color1 = m_CaptionGradientStart
End If

temp = (Color1 And 255)
DR = temp And 255
temp = Int(Color1 / 256)
DG = temp And 255
temp = Int(Color1 / 65536)
DB = temp And 255
temp = (Color2 And 255)
DR2 = temp And 255
temp = Int(Color2 / 256)
DG2 = temp And 255
temp = Int(Color2 / 65536)
DB2 = temp And 255
If m_CaptionGradientTyp = Vertical Then
                        
    VR = Abs(DR - DR2) / (CurY + CurH)
    VG = Abs(DG - DG2) / (CurY + CurH)
    VB = Abs(DB - DB2) / (CurY + CurH)
    If DR2 < DR Then VR = -VR
    If DG2 < DG Then VG = -VG
    If DB2 < DB Then VB = -VB
    For y = 1 To CurH Step 15
        DR2 = DR + VR * y
        DG2 = DG + VG * y
        DB2 = DB + VB * y
        UserControl.Line (CurX, CurY + y)-(CurX + CurW, CurY + y), RGB(DR2, DG2, DB2)
    Next y
Else
    VR = Abs(DR - DR2) / (CurX + CurW)
    VG = Abs(DG - DG2) / (CurX + CurW)
    VB = Abs(DB - DB2) / (CurX + CurW)
    If DR2 < DR Then VR = -VR
    If DG2 < DG Then VG = -VG
    If DB2 < DB Then VB = -VB
    For x = 1 To CurW Step 15
        DR2 = DR + VR * x
        DG2 = DG + VG * x
        DB2 = DB + VB * x
        UserControl.Line (CurX + x, CurY)-(CurX + x, CurY + CurH), RGB(DR2, DG2, DB2)
    Next x
End If
End Sub

Public Property Get CaptionGradientchangeColor() As Boolean
    CaptionGradientchangeColor = m_CaptionGradientchangeColor
End Property

Public Property Let CaptionGradientchangeColor(ByVal New_CaptionGradientchangeColor As Boolean)
    m_CaptionGradientchangeColor = New_CaptionGradientchangeColor
    PropertyChanged "CaptionGradientchangeColor"
    DrawText
End Property

Public Property Get CaptionGradientTyp() As TypGrad
    CaptionGradientTyp = m_CaptionGradientTyp
End Property

Public Property Let CaptionGradientTyp(ByVal New_CaptionGradientTyp As TypGrad)
    m_CaptionGradientTyp = New_CaptionGradientTyp
    PropertyChanged "CaptionGradientTyp"
    DrawText
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    DrawText
End Property

Public Sub Refresh()
    IconIm.Refresh
    DrawText
    UserControl.Refresh
End Sub

Public Property Get CaptionAligmend() As CaptionAlig
    CaptionAligmend = m_CaptionAligmend
End Property

Public Property Let CaptionAligmend(ByVal New_CaptionAligmend As CaptionAlig)
    m_CaptionAligmend = New_CaptionAligmend
    PropertyChanged "CaptionAligmend"
    DrawText
End Property

Public Property Get CaptionShadowed() As Boolean
    CaptionShadowed = m_CaptionShadowed
End Property

Public Property Let CaptionShadowed(ByVal New_CaptionShadowed As Boolean)
    m_CaptionShadowed = New_CaptionShadowed
    PropertyChanged "CaptionShadowed"
    DrawText
End Property

Public Property Get CaptionPicture() As StdPicture
    Set CaptionPicture = IconIm.Picture
End Property

Public Property Set CaptionPicture(ByVal New_CaptionPicture As StdPicture)
    Set IconIm.Picture = New_CaptionPicture
    PropertyChanged "CaptionPicture"
'BorderControlCap
'IconIm.Move 0 + BorderPlusL, ((UserControl.ScaleHeight / 2) - (IconIm.Height / 2)) + BorderPlusT
    LookAutoSize
    DrawText
End Property









