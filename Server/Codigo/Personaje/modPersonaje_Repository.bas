Attribute VB_Name = "modPersonaje_Repository"
Option Explicit

' Chequea si el personaje existe en la base de datos
Public Function isNickInapropiado(nick As String) As Boolean
    Dim sql As String
    Dim infoPersonaje As ADODB.Recordset
    
    sql = "SELECT ID FROM " & DB_NAME_PRINCIPAL & ".nicks_inapropiados WHERE nick='" & nick & "'"
    
    Set infoPersonaje = conn.Execute(sql, , adCmdText)
    
    isNickInapropiado = infoPersonaje.EOF = False
End Function

Public Function saveNickInapropiado(nick As String) As Boolean
    Dim sql As String
    Dim resultado As Boolean
    
    sql = "INSERT INTO " & DB_NAME_PRINCIPAL & ".nicks_inapropiados(Nick) VALUES('" & nick & "')"
    resultado = modMySql.ejecutarSQL(sql, 0)
    
    saveNickInapropiado = resultado
End Function
