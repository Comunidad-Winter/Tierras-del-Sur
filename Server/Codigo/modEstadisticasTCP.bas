Attribute VB_Name = "modEstadisticasTCP"
Option Explicit

'Estadisticas por dia
'Min cantidad usuarios. Max cantidad usuarios. Hora maxima cantidad usuarios.
Type tEstadisticasDiarias
    MaxUsuarios As Integer
    MinUsuarios As Integer
End Type

'Estadisticas por hora
'Paquetes enviados/s: .Bits recibidos/s. Bits enviados/s. Usuarios prom:
Public Type TCPESStats
    BitesEnviadosFraccion As Double 'Cantidad de bites enviados durante la ultima fraccion
    BitesRecibidosFraccion As Double 'Cantidad de bites recibidos durante la ultima fraccion
    PaquetesEnviadosFraccion As Double 'Cantidad de paquetes enviaados durante la ultima fraccion
    UsuariosPromedioFraccion As Double
    UsuariosPremiumPromedioFraccion As Double
    
    BitesEnviadosMinuto As Double 'Cantidad de bites enviados durante la ultima fraccion
    BitesRecibidosMinuto As Double 'Cantidad de bites recibidos durante la ultima fraccion
    PaquetesEnviadosMinuto As Double 'Cantidad de paquetes enviaados durante la ultima fraccion
    PaquetesRecibidosMinuto As Double 'Cantidad de paquetes recibidos durante la ultima fraccion
    
    BitesEnviadosXSEG As Long 'Bites enviados promedio por segundo durante el ultimo minuto
    BitesRecibidosXSEG As Long 'Bites recibidos promedio por segundo durante el ultimo minuto
    PaquetesEnviadosXSeg As Long
End Type

Public TCPESStats As TCPESStats
Public DayStats As tEstadisticasDiarias

Private Const fraccionMinutos As Byte = 30
Private Const fraccionSegundos As Long = fraccionMinutos * 60

'---------------------------------------------------------------------------------------
' Procedure : ActualizaStats
' DateTime  : 18/02/2007 18:59
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub ActualizaStats()


Call LogEstadisticas("Recibido bites: " & TCPESStats.BitesRecibidosMinuto & " Recibido paquetes: " & TCPESStats.PaquetesRecibidosMinuto)

With TCPESStats
        
    'A los datos de la fraccion le sumo lo de los minutos
    .PaquetesEnviadosFraccion = .PaquetesEnviadosFraccion + .PaquetesEnviadosMinuto
    .BitesRecibidosFraccion = .BitesRecibidosFraccion + .BitesRecibidosMinuto
    .BitesEnviadosFraccion = .BitesEnviadosFraccion + .BitesEnviadosMinuto
    .UsuariosPromedioFraccion = .UsuariosPromedioFraccion + NumUsers
    .UsuariosPremiumPromedioFraccion = .UsuariosPremiumPromedioFraccion + NumUsersPremium
    
    'Calculo con los datos del ultimo minuto los correspondientes a los segundos
    .BitesEnviadosXSEG = CLng(.BitesEnviadosMinuto / 60)
    .BitesRecibidosXSEG = CLng(.BitesRecibidosMinuto / 60)
    .PaquetesEnviadosXSeg = CLng(.PaquetesEnviadosMinuto / 60)

    'Reseteo los contadores del minuteo
    .BitesEnviadosMinuto = 0
    .BitesRecibidosMinuto = 0
    .PaquetesEnviadosMinuto = 0
    .PaquetesRecibidosMinuto = 0

End With

End Sub

Public Sub GuardarEstadisticasFraccion(fecha As Date)

Dim sql As String
 
With TCPESStats
   
   'Inserto los datos en la base de datos
    sql = "INSERT DELAYED " & DB_NAME_PRINCIPAL & ".juego_estadisticas_globales(Dia, MbitsEnviados, MbitsRecibidos, PaquetesEnviados, UsuariosOnline, UsuariosOnlinePremium)" & _
            " values('" & Format(Now, "yyyy-mm-dd hh:mm") & "','" & Replace(FormatNumber(.BitesEnviadosFraccion / (fraccionSegundos * 1000000), 4), ",", ".") & "','" & Replace(FormatNumber(.BitesRecibidosFraccion / (fraccionSegundos * 1000000), 4), ",", ".") & "'," & CInt(.PaquetesEnviadosFraccion / fraccionSegundos) & "," & CInt(.UsuariosPromedioFraccion / fraccionMinutos) & "," & CInt(.UsuariosPremiumPromedioFraccion / fraccionMinutos) & ")"

    conn.Execute sql, , adExecuteNoRecords
    
    'Reseteo las estadisticas
    .BitesEnviadosFraccion = 0
    .BitesRecibidosFraccion = 0
    .PaquetesEnviadosFraccion = 0
    .UsuariosPromedioFraccion = 0
    .UsuariosPremiumPromedioFraccion = 0
End With
End Sub

Public Sub enviarEstadisticas(UserIndex As Integer)

    With TCPESStats
        EnviarPaquete Paquetes.mensajeinfo, "Paquetes enviados/s: " & FormatNumber(.PaquetesEnviadosXSeg) & " ", UserIndex, ToIndex
        EnviarPaquete Paquetes.mensajeinfo, "Recibidos/s: " & FormatNumber(.BitesRecibidosXSEG / 1000000, 4) & " Mbits. Enviados/s: " & FormatNumber(.BitesEnviadosXSEG / 1000000, 4) & " Mbits.", UserIndex, ToIndex
        EnviarPaquete Paquetes.mensajeinfo, "Usuarios premium conectados: " & NumUsersPremium, UserIndex, ToIndex
    End With
End Sub

  
