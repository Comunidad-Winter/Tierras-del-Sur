VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Consola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const p_SEGUNDOS = 7

Private p_CANTIDADDIALOGOS As Byte

Private Type t_GuildDlg
    texto As String
    Segundos As Integer
    color As Long
End Type

Private extrat As t_GuildDlg
Private extraalpha As Single

Private yoff As Single

Private p_Dialogos() As t_GuildDlg
Private p_Activo As Boolean

Private usedFont As StdFont

Public Property Let CantidadDialogos(ByVal v As Byte)
    If v > 0 Then
        ReDim Preserve p_Dialogos(1 To v) As t_GuildDlg
    End If
    p_CANTIDADDIALOGOS = v
End Property

''
' Removes all dialogs from screen.

Public Sub RemoveDialogs()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/02/07
'Removes all dialogs
'***************************************************
    Dim i As Long
    
    For i = 1 To p_CANTIDADDIALOGOS
        p_Dialogos(i).texto = vbNullString
    Next i
End Sub

''
' Retrieves the number of dialogs to be displayed on screen.
'
' @return   The number of dialogs to be displayed on screen.

Public Property Get CantidadDialogos() As Byte
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/01/07
'Retrieves the number of dialogs to be displayed on screen
'***************************************************
    CantidadDialogos = p_CANTIDADDIALOGOS
End Property

Public Property Let Activo(ByVal v As Boolean)
    p_Activo = v
End Property

Public Property Get Activo() As Boolean
    Activo = p_Activo
End Property

Private Sub Class_Initialize()
    p_CANTIDADDIALOGOS = 5
    ReDim p_Dialogos(1 To p_CANTIDADDIALOGOS) As t_GuildDlg
    p_Activo = True
End Sub

Public Sub Draw(ByVal x_offs As Single, ByVal y_offs As Single)
Dim i As Byte
Dim YY As Single

If Not p_Activo Then Exit Sub

'Y = 250
'500 600 700 800 900
    If yoff > 0 Then
        yoff = yoff - timerTicksPerFrame
    End If
    
    If yoff < 0 Then yoff = 0
    
    If extraalpha > 0 Then
        extraalpha = extraalpha - timerTicksPerFrame * 25.5555555555556 / 2
    End If
    
    YY = yoff - 5 + y_offs
    
    If extraalpha <= 0 Then
        extraalpha = 0
    Else
        Call Engine.Text_Render_alpha(extrat.texto, YY, x_offs, extrat.color, 0, extraalpha)
    End If
    
    
    
    For i = 1 To p_CANTIDADDIALOGOS
        If Len(p_Dialogos(i).texto) Then
            Call Engine.Text_Render_alpha(p_Dialogos(i).texto, i * 10 + YY, x_offs, p_Dialogos(i).color, 0, minl((i + 1) * (255 / p_CANTIDADDIALOGOS), 255))
        End If
    Next i
End Sub

Public Sub PassTimer()
Dim i As Byte
extrat = p_Dialogos(1)
extraalpha = 255
yoff = yoff + 10
For i = 1 To p_CANTIDADDIALOGOS - 1
    p_Dialogos(i) = p_Dialogos(i + 1)
Next i
p_Dialogos(p_CANTIDADDIALOGOS).Segundos = 0
p_Dialogos(p_CANTIDADDIALOGOS).texto = vbNullString

End Sub




''
' Splits dialogs into lines fitting properly the render area and inserts them

Public Sub PushBackText(ByVal s As String, Optional ByVal color As Long = &HFF00FF00)
    Dim str As String
    Dim tmp As Integer
    On Error Resume Next
    
    
    str = s
    
    ' If it's too long to fit, split it
'    Do While frmMain.TextWidth(STR) > 500
'        tmp = InStrRev(STR, " ")
'
'        If tmp = 0 Then Exit Do
'
'        STR = Left$(STR, tmp - 1)
'    Loop
    
    
    
    
    
    'Show message and continue with the rest
    Call RealPushBackText(str, color)
    
    If LenB(str) <> LenB(s) Then
        Call PushBackText(Right$(s, Len(s) - tmp), color)
    End If
End Sub

Private Sub RealPushBackText(ByVal s As String, Optional ByVal color As Long = &HFFFFFF00)
Dim i As Byte
extrat = p_Dialogos(1)
extraalpha = 2 * 255 / p_CANTIDADDIALOGOS
For i = 1 To p_CANTIDADDIALOGOS - 1
    p_Dialogos(i) = p_Dialogos(i + 1)
Next i
p_Dialogos(p_CANTIDADDIALOGOS).Segundos = 0
p_Dialogos(p_CANTIDADDIALOGOS).texto = s
p_Dialogos(p_CANTIDADDIALOGOS).color = color
yoff = 10
End Sub
