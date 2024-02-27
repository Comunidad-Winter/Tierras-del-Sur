Attribute VB_Name = "AdminMacAddress"
Option Explicit
'Tabla en la base de datos: juego_ban_macaddress

Public Sub banMac(MacAddress As String, IDGm As Long, razon As String)

Dim sql As String

sql = "INSERT INTO " & DB_NAME_PRINCIPAL & ".juego_ban_macaddress(MacAddress,Razon,IDGm) VALUES ('" & mysql_real_escape_string(MacAddress) & "','" & mysql_real_escape_string(razon) & "'," & IDGm & ")"
conn.Execute sql, , adExecuteNoRecords

End Sub

Public Sub unBanMac(MacAddress As String)

Dim sql As String

sql = "DELETE FROM " & DB_NAME_PRINCIPAL & ".juego_ban_macaddress WHERE MacAddress = '" & mysql_real_escape_string(MacAddress) & "'"
conn.Execute sql, , adExecuteNoRecords

End Sub

Public Function isMacBaneada(MacAddress As String)
Dim rs As New ADODB.Recordset
Dim sql As String

sql = "SELECT MacAddress FROM " & DB_NAME_PRINCIPAL & ".juego_ban_macaddress WHERE MacAddress = '" & mysql_real_escape_string(MacAddress) & "'"
Set rs = conn.Execute(sql)

'Si hay alguna registro quiere decir que esta baneada.
If rs.EOF Then
    isMacBaneada = False
Else
    isMacBaneada = True
End If
    
rs.Close
Set rs = Nothing

End Function
