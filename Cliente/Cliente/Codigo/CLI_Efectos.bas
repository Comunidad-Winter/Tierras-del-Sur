Attribute VB_Name = "CLI_Efectos"
Option Explicit

'Lista de cuerpos
Type tFXHECHIZO ':( Missing Scope
    Activo As Boolean
    CharIndex As Integer
    fX As Integer
    Loops As Integer
    Sonido As String
    KillChar As Boolean
    CreateChar As String
End Type

Public FXList() As tFXHECHIZO

Public Sub AddFXList(CharIndex As Integer, fX As Integer, Loops As Integer, Sonido As String)

    Dim AtIndex As Integer
    Dim i As Integer
    Dim EnSlot As Boolean

    For i = 1 To UBound(FXList)
        If FXList(i).Activo = False Then
            EnSlot = True
            Exit For 'loop varying i
        End If
    Next i
    If EnSlot Then
        AtIndex = i
      Else 'ENSLOT = FALSE/0
        AtIndex = UBound(FXList) + 1
        ReDim Preserve FXList(AtIndex)
    End If
    FXList(AtIndex).CharIndex = CharIndex
    FXList(AtIndex).fX = fX
    FXList(AtIndex).Loops = Loops
    FXList(AtIndex).Sonido = Sonido
    FXList(AtIndex).Activo = True

End Sub
