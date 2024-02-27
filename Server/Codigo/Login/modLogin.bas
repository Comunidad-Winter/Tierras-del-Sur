Attribute VB_Name = "modLogin"
Option Explicit

Public Sub enviarInformacionCrc(ByRef Usuario As User)
    Dim tempbyte As Byte
    Dim tempbyte2 As Byte
    Dim tempstr As String
    
    'Genero el CRC para el usuario y el Offset de encriptacion
    tempbyte = RandomNumber(0, 5)
    tempbyte2 = RandomNumber(200, 250)

    Usuario.MinPacketNumber = RandomNumber(10, 50)
    Usuario.PacketNumber = Usuario.MinPacketNumber
            
    '1) -- Sin uso ----
    '2) Min Packet Number
    tempstr = Chr$(((tempbyte + 1) Xor 127) Xor 113) & Chr$((Usuario.MinPacketNumber Xor 12) Xor 107)

    EnviarPaquete Paquetes.infoLogin, tempstr, Usuario.UserIndex, ToIndex
End Sub

Private Sub logIntentoIngresoBloqueado(MacAddress As String, ip As Currency, idPersonaje As Long)
    Dim sql As String
    
    sql = "INSERT DELAYED INTO " & DB_NAME_PRINCIPAL & ".juego_logs_ingresos_bloqueados(MACADDRESS, IP, IDPERSONAJE) values('" & MacAddress & "'," & ip & "," & idPersonaje & ")"
    
    Call modMySql.ejecutarSQL(sql)

End Sub

Public Sub ProcesarPaqueteNodo(ByRef Usuario As User, anexo As String)
    Dim hash As String
    Dim datosHash As ADODB.Recordset
    Dim longitudHash As String
    'Obtenemos le hash
       
    longitudHash = Asc(Left$(anexo, 1))
    
    hash = mid$(anexo, 2)
        
    'Obtenemos la informacion de la base de datos para saber si quiere ingresar con un personaje o si
    sql = "SELECT * FROM " & DB_NAME_PRINCIPAL & ".loginserver_permisos WHERE HASH='" & mysql_real_escape_string(hash) & "'"
    
    Set datosHash = conn.Execute(sql, , adCmdText)
    
    '¿Existe este HASH?
    If datosHash.EOF = False Then
        'Cargamos los datos. (MacAddress, IP, Nombre-PC)
        Usuario.ip = datosHash!ip
        Usuario.MacAddress = datosHash!MacAddress
        Usuario.NombrePC = datosHash!NombrePC
        Usuario.CryptOffset = (datosHash!semilla Xor 17) Mod 19
       
        'Chequeamos si la mac address esta baneada.
        If AdminMacAddress.isMacBaneada(Usuario.MacAddress) Then
            'SI: Mensaje de que la MAC esta baneada.
            EnviarPaquete mbox, Chr$(8), Usuario.UserIndex
        
            '¿Es crear?
            Call logIntentoIngresoBloqueado(Usuario.MacAddress, Usuario.ip, IIf(IsNull(datosHash!idPersonaje), 0, datosHash!idPersonaje))
            
            If Not CloseSocket(Usuario.UserIndex) Then Call LogError("Procesar paquete nodo")
        Else
            '¿Es crear o conectar?
            If Not (IsNull(datosHash!idPersonaje) And IsNull(datosHash!clave)) Then
                'Le enviamos la informacion de autentificacion
                Call enviarInformacionCrc(Usuario)
                ' Guardamos el nombre con el cual se hizo la validacion
                Usuario.Name = datosHash!nombrePersonaje
                'Conectar personaje
                Call ConnectUser(Usuario.UserIndex, datosHash!idPersonaje, datosHash!clave)
            Else
                Usuario.FechaIngreso = Now
                'Le enviamos la informacion de autentificacion
                Call enviarInformacionCrc(Usuario)
                'No tengo que hacer nada ya que ahora van a venir los paquetes de tirar dados.
            End If
        End If
    Else
        LogError ("Hash invalido")
        'Hash invalido
        If Not CloseSocket(Usuario.UserIndex) Then LogError ("Hash Invalido Procesar Paquete")
    End If


    Call datosHash.Close
    Set datosHash = Nothing
End Sub

