Attribute VB_Name = "ME_TRIGGERS"
Option Explicit

Public triggers_count As Integer
Public triggers_names() As String


Public trigger_seleccionado As Integer

Public Sub LoadTriggersRaw(ByRef FileName As String)
    '*****************************************************************
    'Menduz
    '*****************************************************************

    triggers_count = val(GetVar(FileName, "Triggers", "Num"))
    Dim loopc As Integer
    ReDim triggers_names(0 To triggers_count)
    frmMain.lstTriggers.Clear
    For loopc = 0 To triggers_count
        triggers_names(loopc) = GetVar(FileName, "Triggers", CStr("T" & loopc))
        frmMain.lstTriggers.AddItem triggers_names(loopc)
    Next loopc
End Sub

Public Function calcular_trigger_lista() As Long
    Dim i As Integer
    
    calcular_trigger_lista = 0
    
    For i = 0 To frmMain.lstTriggers.ListCount - 1
        If frmMain.lstTriggers.Selected(i) Then
            calcular_trigger_lista = calcular_trigger_lista Or bitwisetable(i)
        End If
    Next i
End Function
