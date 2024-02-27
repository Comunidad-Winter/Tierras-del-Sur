Attribute VB_Name = "Logs"
Option Explicit

Public Sub LogCriticEvent(desc As String)
    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
        Print #nfile, Date & " " & Time & " " & desc
    Close #nfile
End Sub

Public Sub logProfilePaquete(desc As String)
    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\logs\profiles-" & Day(Date) & "-" & Hour(Time) & ".log" For Append Shared As #nfile
        Print #nfile, Time & " " & desc
    Close #nfile
End Sub

Public Sub LogProblemaSpawn(desc As String)
    Dim nfile As Integer
    
    Debug.Print desc
    
    #If testeo = 1 Then 'Asi no se pasa ningun error por alto.
        MsgBox desc
    #End If
    
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\spawnFallidos.log" For Append Shared As #nfile
        Print #nfile, Date & " " & Time & " " & desc
    Close #nfile
End Sub

Public Sub LogSQLerror(desc As String)
    Dim nfile As Integer
    
    Debug.Print desc
    
    #If testeo = 1 Then 'Asi no se pasa ningun error por alto.
        MsgBox desc
    #End If
    
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\sqlError.log" For Append Shared As #nfile
        Print #nfile, Date & " " & Time & " " & desc
    Close #nfile
End Sub

'CSEH: Nada
Public Sub LogError(desc As String)
On Error GoTo errhandler

    Dim nfile As Integer
        
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\errores.log" For Append Shared As #nfile
        Print #nfile, Date & " " & Time & " " & desc
        Debug.Print Date & " " & Time & " " & desc
    Close #nfile
    
Exit Sub
errhandler:
End Sub

Public Sub LogMain(desc As String)
    Dim nfile As Integer
        
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\main.log" For Append Shared As #nfile
        Print #nfile, Date & " " & Time & " " & desc
    Close #nfile
End Sub

Public Sub LogEventos(desc As String)
    Dim nfile As Integer
        
    #If testeo = 1 Then 'Asi no se pasa ningun error por alto.
        'MsgBox Desc
    #End If
    
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\eventosTorneos.log" For Append Shared As #nfile
        Print #nfile, Date & " " & Time & " " & desc
    Close #nfile
End Sub

Public Sub LogLenguaje(desc As String)
    Dim nfile As Integer
    
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\lenguaje.log" For Append Shared As #nfile
        Print #nfile, desc
    Close #nfile
End Sub

Public Sub LogBackup(desc As String)
    Dim nfile As Integer
        
    #If testeo = 1 Then 'Asi no se pasa ningun error por alto.
        'MsgBox Desc
    #End If
    
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\backup.log" For Append Shared As #nfile
        Print #nfile, Date & " " & Time & " " & desc
    Close #nfile
End Sub

Public Sub LogTorneos(desc As String)
    Dim nfile As Integer
    
    Debug.Print desc
    
    #If testeo = 1 Then 'Asi no se pasa ningun error por alto.
        'MsgBox Desc
    #End If
    
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\torneos.log" For Append Shared As #nfile
        Print #nfile, Date & " " & Time & " " & desc
    Close #nfile
End Sub
Public Sub LogMultiLogin(desc As String)
    Dim nfile As Integer
    
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\multi-login.log" For Append Shared As #nfile
        Print #nfile, Date & " " & Time & " " & desc
    Close #nfile
End Sub

Public Sub LogDesarrollo(ByVal str As String)
    Dim nfile As Integer
    
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\desarrollo.log" For Append Shared As #nfile
        Print #nfile, Date & " " & Time & " " & str
        Debug.Print Date & " " & Time & " " & str
    Close #nfile
End Sub

Public Sub LogGM(idUsuario As Long, descripcion As String, Optional Comando As String = 0)
Dim sql As String

sql = "INSERT DELAYED " & DB_NAME_PRINCIPAL & ".logs_gms(ID_Usuario,Comando,Descripcion) values('" & idUsuario & "','" & Comando & "','" & mysql_real_escape_string(descripcion) & "')"

conn.Execute sql, , adExecuteNoRecords
End Sub

Public Sub logVentaCasa(ByVal texto As String)
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile
End Sub

Public Sub LogHackAttemp(texto As String)
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile
End Sub

Public Sub LogIP(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\IP.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

End Sub
Public Sub LogCriticalHackAttemp(texto As String)
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile
End Sub

Public Sub LogNuevosRetos(ByVal str1 As String)
    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\retosNUEVOS.log" For Append As #nfile
    Print #nfile, Date & " " & Time & " " & str1
    Close #nfile
End Sub


Public Sub LogHack(texto As String)
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\FalsificacionPaquetes.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile
End Sub

Public Sub LogMacAddressBaneadaIntento(texto As String)
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\LogMacAddressBaneadaIntento.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile
End Sub

Public Sub LogAccionesWeb(texto As String)
    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal

    Open App.Path & "\logs\LogAccionesWeb.log" For Append Shared As #nfile

    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"

    Close #nfile
End Sub

Public Sub LogCambioNick(ByVal str As String)
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CambiosDeNick.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile
End Sub
 
Public Sub LogCentinela(desc As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\centinelas.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & desc
Close #nfile

End Sub

Public Sub LogEstadisticas(texto As String)
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\estadisticas.log" For Append Shared As #nfile
Print #nfile, texto
Close #nfile

End Sub
