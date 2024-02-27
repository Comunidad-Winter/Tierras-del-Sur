Attribute VB_Name = "HelperTiempo"
Option Explicit

Public Function segundosAHoras(segundos As Long) As String
    
    Dim horas As Integer
    Dim minutos As Byte
    
    horas = segundos \ 3600
    minutos = ((segundos Mod 3600) \ 60)
    
    If horas = 1 Then
        segundosAHoras = segundosAHoras & "una hora"
    ElseIf horas > 1 Then
        segundosAHoras = segundosAHoras & horas & " horas"
    End If
       
    If minutos > 0 Then
        
        If horas > 0 Then segundosAHoras = segundosAHoras & " y "

        If minutos = 1 Then
            segundosAHoras = segundosAHoras & "un minuto"
        Else
            segundosAHoras = segundosAHoras & minutos & " minutos"
        End If
        
    ElseIf minutos = 0 And horas = 0 Then
        segundosAHoras = segundosAHoras & "0 minutos"
    End If
    
End Function

'Esta funcion es como el time() del php
Function timePHP() As Long
    timePHP = DateDiff("s", "01/01/1970 00:00:00", Now, vbMonday, vbFirstFullWeek)
End Function
