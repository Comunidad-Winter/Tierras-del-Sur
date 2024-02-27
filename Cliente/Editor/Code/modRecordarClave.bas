Attribute VB_Name = "modRecordarClave"
Option Explicit

Public Sub GuardarPassword(ByVal nombre As String, Password As String)

If Trim(nombre) = "" Then Exit Sub

If Not Password = "X" Then
    Call SaveSetting("EditordelMundo", "CurrentRot", UCase$(MD5String(UCase$(nombre))), CryptStr(Password, 1))
ElseIf Password = "X" Then
    If BuscarPassword(nombre) <> "" Then Call DeleteSetting("EditordelMundo", "CurrentRot", UCase$(MD5String(UCase$(nombre))))
End If

End Sub

Public Function BuscarPassword(ByVal nombre As String) As String

    If nombre <> "" Then
        BuscarPassword = DecryptStr(GetSetting("EditordelMundo", "CurrentRot", UCase$(MD5String(UCase$(nombre)))), 1)
    Else
        BuscarPassword = ""
    End If
    
End Function

Public Function SeleccionarPassword(ByVal usuario As String) As String
    
   SeleccionarPassword = BuscarPassword(usuario)

End Function
