Attribute VB_Name = "modCuentas"
Option Explicit

Public Type tInfoCuenta
    id As Long
    Premium As Boolean
    Estado As String
    bloqueada As Boolean
    mail As String
    pin As String
    segundosTDSF As Long
End Type

#If TDSFacil = 1 Then
Public Sub actualizarDatosCuenta(personaje As User)
    Dim sql                 As String
    Dim segundosPasados     As Long ' Ojo, no poner integer porque a las 8 horas explota
    Dim fechaInicial        As Date
    
    ' Actualizamos la cantidad de segundos que le quedan por jugar
    
    '  Si justo cambio el mes, me quedo con la parte restante
    If Month(personaje.FechaIngreso) = Month(Now) Then
        fechaInicial = personaje.FechaIngreso
    Else
        fechaInicial = DateSerial(Year(Now), Month(Now), 1)
    End If
       
    ' Calculo cuantos segundos estuvo logueado
    segundosPasados = DateDiff("s", fechaInicial, Now)
        
    If personaje.segundosPremium > segundosPasados Then
        personaje.segundosPremium = personaje.segundosPremium - segundosPasados
    Else
        personaje.segundosPremium = 0
    End If
           
    ' Armamos la consulta.
    ' Resto la cantidad de segundos, porque tal vez la cantidad de segundos
    ' fueron actualizados cuando el personaje estaba Online.
    sql = "UPDATE " & DB_NAME_CUENTAS & " SET SEGUNDOS_TDSF = SEGUNDOS_TDSF - " & segundosPasados & _
        " WHERE IDCUENTA=" & personaje.IDCuenta

    ' Ejecutamos
    Call conn.Execute(sql, , adExecuteNoRecords)
    
End Sub
#End If

Public Function obtenerInfoCuentaByMail(correo As String) As tInfoCuenta

Dim sql As String
Dim infoCuenta As ADODB.Recordset

sql = "SELECT cuenta.idcuenta, IF (cuenta.FECHAVENCIMIENTO IS NOT NULL AND CUENTA.FECHAVENCIMIENTO > UNIX_TIMESTAMP(), 'SI', 'NO') AS ESPREMIUM," & _
        " cuenta.ESTADO, cuenta.BLOQUEADA, cuenta.MAIL, cuenta.SEGUNDOS_TDSF" & _
        " FROM " & DB_NAME_CUENTAS & _
        " cuenta WHERE Cuenta.Mail = '" & correo & "'"

Debug.Print sql

'Ejecutamos
Set infoCuenta = conn.Execute(sql, , adCmdText)

If infoCuenta.EOF Then
    obtenerInfoCuentaByMail.id = -1
Else
    'Si no esta definida si esta bloqueada o no. Por algun motivo no esta cargada la info de esta cuenta
    If Not IsNull(infoCuenta!bloqueada) Then
        obtenerInfoCuentaByMail.id = infoCuenta!IDCuenta
         
        obtenerInfoCuentaByMail.mail = infoCuenta!mail
        obtenerInfoCuentaByMail.Premium = (infoCuenta!esPremium = "SI")
        obtenerInfoCuentaByMail.bloqueada = (infoCuenta!bloqueada = "SI")
        obtenerInfoCuentaByMail.Estado = infoCuenta!Estado
        obtenerInfoCuentaByMail.pin = ""
        obtenerInfoCuentaByMail.segundosTDSF = infoCuenta!SEGUNDOS_TDSF
    Else
        obtenerInfoCuentaByMail.id = -2
    End If
End If

'Liberamos
infoCuenta.Close
Set infoCuenta = Nothing
  
End Function
    
    
Public Function obtenerInfoCuentaDeNombreReservado(nombreReservado As String) As tInfoCuenta

Dim sql As String
Dim infoCuenta As ADODB.Recordset

sql = "SELECT reserva.IDCUENTA, IF (cuenta.FECHAVENCIMIENTO IS NOT NULL AND CUENTA.FECHAVENCIMIENTO > UNIX_TIMESTAMP(), 'SI', 'NO') AS ESPREMIUM," & _
        " cuenta.ESTADO, cuenta.BLOQUEADA, cuenta.MAIL, cuenta.PIN" & _
        ", cuenta.SEGUNDOS_TDSF" & _
        " FROM " & DB_NAME_PRINCIPAL & ".nicks_reservados  AS reserva" & _
        " LEFT JOIN " & DB_NAME_CUENTAS & " AS cuenta ON cuenta.IDCUENTA = reserva.IDCUENTA" & _
        " WHERE reserva.Nombre = '" & nombreReservado & "'"

Debug.Print sql
'Ejecutamos
Set infoCuenta = conn.Execute(sql, , adCmdText)

If infoCuenta.EOF Then
    obtenerInfoCuentaDeNombreReservado.id = -1
Else
    'Si no esta definida si esta bloqueada o no. Por algun motivo no esta cargada la info de esta cuenta
    If Not IsNull(infoCuenta!bloqueada) Then
        obtenerInfoCuentaDeNombreReservado.id = infoCuenta!IDCuenta
         
        obtenerInfoCuentaDeNombreReservado.mail = infoCuenta!mail
        obtenerInfoCuentaDeNombreReservado.Premium = (infoCuenta!esPremium = "SI")
        obtenerInfoCuentaDeNombreReservado.bloqueada = (infoCuenta!bloqueada = "SI")
        obtenerInfoCuentaDeNombreReservado.Estado = infoCuenta!Estado
        obtenerInfoCuentaDeNombreReservado.pin = infoCuenta!pin
        obtenerInfoCuentaDeNombreReservado.segundosTDSF = infoCuenta!SEGUNDOS_TDSF
    Else
        obtenerInfoCuentaDeNombreReservado.id = -2
    End If
End If

'Liberamos
infoCuenta.Close
Set infoCuenta = Nothing
  
End Function
    
