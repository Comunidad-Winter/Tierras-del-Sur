Attribute VB_Name = "modLogsPersonajes"
Option Explicit

Public Sub LogOroArrojado(personaje As User, cantidad As Long, mapa As Integer, x As Integer, y As Integer)
    Dim sql As String
    
    sql = "INSERT DELAYED " & DB_NAME_PRINCIPAL & ".juego_logs_oro_arrojado(personajeId, cantidad, mapa, x, y) VALUES(" & personaje.id & "," & cantidad & "," & mapa & "," & x & "," & y & ")"
    Debug.Print sql
    conn.Execute sql, , adExecuteNoRecords
End Sub

Public Sub LogSubeNivel(idUsuario As Long, VidaUp As Integer, VidaTotal As Integer, Nivel As Integer)
    Dim sql As String

    sql = "INSERT DELAYED " & DB_NAME_PRINCIPAL & ".juego_logs_desarrollo(IDUsuario,VidaUp,VidaTotal,Nivel) values(" & idUsuario & "," & VidaUp & "," & VidaTotal & "," & Nivel & ")"

    conn.Execute sql, , adExecuteNoRecords
End Sub

Public Sub LogAsignaSkill(idUsuario As Long, Skill As Byte, cantidad As Integer, ip As Currency, Total As Integer)
    Dim sql As String
    
    sql = "INSERT DELAYED " & DB_NAME_PRINCIPAL & ".juego_logs_skills(IDUsuario,Skill,Cantidad,IP,Total) values(" & idUsuario & "," & Skill & "," & cantidad & "," & ip & "," & Total & ")"
    
    conn.Execute sql, , adExecuteNoRecords
End Sub

Public Sub LogCentinelaMysql(idUsuario As Long, codigo As String, tipo_accion As String)
Dim sql As String

sql = "INSERT DELAYED " & DB_NAME_PRINCIPAL & ".juego_logs_centinelas(IDPJ,CODIGO,TIPO_ACCION) values('" & idUsuario & "','" & mysql_real_escape_string(codigo) & "','" & tipo_accion & "')"

conn.Execute sql, , adExecuteNoRecords
End Sub
