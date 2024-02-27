Attribute VB_Name = "CDM"
Option Explicit

Public Enum eEstadoCDM
    Conectando = 1
    conectado = 2
    error = 3
End Enum

Public Enum ePermisosCDM
    lectura = 1
    escritura = 2
End Enum

Public cerebro As clsCDM

Private Const Carpeta = "CDM/"

Public Sub CDM_Iniciar(ControlInet As Inet, ControlTimer As VB.timer, UserAgent As String)

    Set cerebro = New clsCDM
    
    #If Produccion = 1 Then
        Call cerebro.iniciar(ControlInet, ControlTimer, UserAgent, "produccion", app.Path & "/" & Carpeta)
    #ElseIf Produccion = 0 Then
        Call cerebro.iniciar(ControlInet, ControlTimer, UserAgent, "ofi-dev", app.Path & "/" & Carpeta)
    #ElseIf Produccion = 2 Then
        Call cerebro.iniciar(ControlInet, ControlTimer, UserAgent, "pre-produccion", app.Path & "/" & Carpeta)
    #End If
End Sub

