Attribute VB_Name = "modFotodenuncias"
Option Explicit

'Formato de llegada de la fotodenuncia
'CharIndex > texto

Public Sub reportarFotodenuncia(UsuarioReportador As User, denuncia As String)

    Dim dialogos() As String
    Dim loopDialogo As Integer
    Dim cadena As String
    Dim charIndex As Integer
    Dim UserIndex As Integer
    
    dialogos = Split(denuncia, "$|@")
    
    For loopDialogo = 0 To UBound(dialogos)
        charIndex = val(mid$(dialogos(loopDialogo), 1, InStr(1, dialogos(loopDialogo), ">") - 1))
        UserIndex = ObtengoIndex_CharIndex(charIndex)
        
        If UserIndex > 0 Then 'Si el usuario cerro entre que saco la foto denuncia y la envio.. cagamos
            If loopDialogo = UBound(dialogos) Then
                cadena = cadena & UserList(UserIndex).Name & "> " & mid$(dialogos(loopDialogo), InStr(1, dialogos(loopDialogo), ">") + 1)
            Else
                cadena = cadena & UserList(UserIndex).Name & "> " & mid$(dialogos(loopDialogo), InStr(1, dialogos(loopDialogo), ">") + 1) & "$|@"
            End If
        End If
    Next loopDialogo
    
    conn.Execute "INSERT INTO " & DB_NAME_PRINCIPAL & ".fotodenuncias(Usuario,Texto) values('" & UsuarioReportador.Name & "','" & mysql_real_escape_string(cadena) & "')", , adExecuteNoRecords
    
    UsuarioReportador.Counters.FotoDenuncia = 1
            
End Sub
