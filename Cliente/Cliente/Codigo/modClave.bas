Attribute VB_Name = "modClave"
Option Explicit

#If TDSFacil = 1 Then
    Private Const RegAppName = "ADOOB2"
#Else
    Private Const RegAppName = "ADOOB"
#End If

Public Sub GuardarPassword(ByVal Nombre As String, Password As String)

    If Trim(Nombre) = "" Then Exit Sub
    
    Nombre = MD5String(UCase(Nombre))
    
    Call SaveSetting(RegAppName, "CurrentRot", Nombre, Password)

End Sub

Public Sub EliminarPassword(ByVal Nombre As String)
    If BuscarPassword(Nombre) <> "" Then Call DeleteSetting(RegAppName, "CurrentRot", MD5String(UCase$(Nombre)))
End Sub


Public Function BuscarPassword(Nombre As String)

    If Nombre <> "" Then
        BuscarPassword = DecryptStr(GetSetting(RegAppName, "CurrentRot", MD5String(UCase$(Nombre))), 1)
    Else
        BuscarPassword = ""
    End If
    
End Function
