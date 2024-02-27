Attribute VB_Name = "Engine_UI"
Option Explicit
Option Base 0

Private button_over As Integer

Public Sub init_special_slots()

End Sub

Public Sub Render_GUI()

    Dim lcbk As Long
    lcbk = lColorMod
    lColorMod = D3DTOP_MODULATE
    Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, lColorMod)

    lColorMod = lcbk
    Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, lColorMod)
End Sub

Public Sub GUI_Click(ByVal x As Integer, ByVal y As Integer, ByVal button As Integer)

End Sub

Public Sub GUI_Mouse_Move(ByVal x As Integer, ByVal y As Integer, ByVal button As Integer)

End Sub

Public Sub Handle_Key(KeyCode As Integer, Shift As Integer)

End Sub

Public Sub Handle_KeyP(KeyCode As Integer)

End Sub

Public Sub toggle_render_text_indicator()

End Sub

