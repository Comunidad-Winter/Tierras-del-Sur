VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vWControlChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private p_CANTIDADDIALOGOS As Byte

Private Type t_GuildDlg
    texto As String
    segundos As Integer
    color As Long
End Type

Private p_Dialogos() As t_GuildDlg
Private p_Activo As Boolean

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
    p_CANTIDADDIALOGOS = 7
    ReDim p_Dialogos(1 To p_CANTIDADDIALOGOS) As t_GuildDlg
    p_Activo = True
End Sub

Public Sub Draw(ByVal x_offs As Single, ByVal y_offs As Single)
Dim i As Byte
Dim alpha As Byte

If Not p_Activo Then Exit Sub
        
For i = 1 To p_CANTIDADDIALOGOS
    If Len(p_Dialogos(i).texto) Then
        alpha = 255
        If p_Dialogos(i).segundos = 2 Then
            alpha = alpha / 2
        ElseIf p_Dialogos(i).segundos = 1 Then
            alpha = alpha / 3
        End If
        Call Engine.Text_Render_alpha(p_Dialogos(i).texto, i * 10 + y_offs, x_offs, p_Dialogos(i).color, 0, alpha)
    End If
Next i
End Sub

Public Sub PassTimer()
Dim i As Byte

For i = 1 To p_CANTIDADDIALOGOS
    If p_Dialogos(i).segundos > 0 Then
        p_Dialogos(i).segundos = p_Dialogos(i).segundos - 1
        
        If p_Dialogos(i).segundos = 0 Then
            p_Dialogos(i).texto = vbNullString
        End If
    End If
Next i

End Sub



''
' Splits dialogs into lines fitting properly the render area and inserts them

Public Sub PushBackText(ByVal s As String, Optional ByVal color As Long = &HFF00FF00, Optional ByVal Duracion As Integer = DURACION_TEXTO)
    Dim Str As String
    Dim Tmp As Integer
    Dim posEnter As Integer
    
    Str = s
        
    If frmMain.TextWidth(Str) > 500 Then
        Tmp = Len(Str)
        ' Vamos sacando palabras hasta que la longitud sea menor a 500
        Do While frmMain.TextWidth(left$(Str, Tmp)) > 500
            Tmp = InStrRev(Str, " ", Tmp - 1)
        Loop
    End If

    
    posEnter = InStr(1, s, vbCrLf)
        
    If (posEnter > 0 And posEnter < Tmp) Or (Tmp = 0 And posEnter > 0) Then
        Str = left$(Str, posEnter)
        Tmp = posEnter + 1
    ElseIf Tmp = 0 Then
        Str = left$(Str, 500)
    Else
        Str = left$(Str, Tmp - 1)
    End If
        
    'Show message and continue with the rest
    Call RealPushBackText(Str, color, Duracion)
    
    If LenB(Str) <> LenB(s) Then
        Call PushBackText(right$(s, Len(s) - Tmp), color)
    End If
End Sub

Private Sub RealPushBackText(ByVal s As String, Optional ByVal color As Long = &HFFFFFF00, Optional ByVal segundos As Integer = DURACION_TEXTO)
Dim i As Byte

For i = 1 To p_CANTIDADDIALOGOS - 1
    p_Dialogos(i) = p_Dialogos(i + 1)
Next i

p_Dialogos(p_CANTIDADDIALOGOS).segundos = segundos
p_Dialogos(p_CANTIDADDIALOGOS).texto = s
p_Dialogos(p_CANTIDADDIALOGOS).color = color

End Sub
