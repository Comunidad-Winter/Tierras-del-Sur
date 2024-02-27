Attribute VB_Name = "modConectar"
Option Explicit

Private hashActual As String

Public Type retornoInfo
    error As Byte
    errordesc As String
    datos As tDatosConexion
End Type

Private direccionesInnacesibles() As String

Public Sub conectar(datosConexion As tDatosConexion)

    TCP.recibiPaquete = False
    
    If frmMain.Socket1.Connected Or frmMain.Socket1.State = 3 Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
    End If
        
    frmMain.Socket1.HostAddress = datosConexion.ip
    frmMain.Socket1.RemotePort = datosConexion.puerto
    frmMain.Socket1.LocalPort = CInt(Int((64 * Rnd()) + 12620)) '12620-12683
    
    hashActual = datosConexion.hash

    CryptOffs = (datosConexion.semilla Xor 17) Mod 19
    
    frmMain.Socket1.Connect

End Sub

Private Function encriptarIP(ip As String) As String
    
    Dim segmentos() As String
    Dim azar1 As Byte, azar2 As Byte
    Dim loopSegmento As Integer
    Dim ultimo As Byte
    
    segmentos = Split(ip, ".", 4, vbBinaryCompare)
    
    'Genero dos numeros al azar del 10 al 255
    azar1 = Int(RandomNumber(100, 255))
    azar2 = Int(RandomNumber(65, 79))
    
    'El primero lo guardo XOR 55
    encriptarIP = Chr$(azar1)
    'Al segundo lo guardo  + 36 XOR El Primero.
    encriptarIP = encriptarIP & Chr$((azar2 + 36) Xor azar1)
    'Al ultimo le hago XOR del Azar1 y XOR DEL Azar2
    ultimo = (CByte(segmentos(3)) Xor azar1) Xor azar2
    
    encriptarIP = encriptarIP & Chr$(ultimo)
        
   'Al tercero le hago XOR del ultimo
    For loopSegmento = 2 To 0 Step -1
        ultimo = CByte(segmentos(loopSegmento)) Xor ultimo
        encriptarIP = encriptarIP & Chr$(CByte(ultimo))
    Next
    

End Function
Public Sub enviarHash()
    Dim TempStr As String
    Dim ip As String
    Dim longitud As Integer
    
    ip = encriptarIP(frmMain.Socket1.HostAddress)
    
    longitud = Len(ip) + Len(hashActual)
    
    TempStr = ip & Chr$(Len(hashActual)) & hashActual 'Longitud  hash + Hash
   
    TempStr = Chr$(Len(TempStr) + 1) & Chr$(1) & TempStr 'Longitud del paquete + Numero Paquete + Hash
    
    Debug.Print "Enviamos Hash"
    
    Call frmMain.Socket1.Write(TempStr, Len(TempStr)) 'Enviamos
End Sub

Public Function conectarPersonaje() As retornoInfo
    Dim infoConexion As tLoginServerRespuesta
    Dim servidor As Byte
    Dim MD5exe As String
    Dim macaddress As String
    
    #If TDSFacil = 1 Then
        servidor = 2
    #Else
        servidor = 1
    #End If

    #If LOCALHOST = 0 And testeo = 1 Then
       servidor = 4
    #End If
        
    MD5exe = generarMD5
    macaddress = UserMac
    
    infoConexion = modLogin.iniciarConexionPersonaje(servidor, UserName, UserPassword, macaddress, GetIdentificacionPC(), MD5exe)
    
    If infoConexion.error = 0 Then
        conectarPersonaje.error = 0
        conectarPersonaje.errordesc = ""
        conectarPersonaje.datos = infoConexion.datosConexion
    Else
        conectarPersonaje.error = infoConexion.error
        conectarPersonaje.errordesc = infoConexion.errordesc
    End If
        
        
End Function

Public Function conectarParaCrear() As retornoInfo
    Dim infoConexion As tLoginServerRespuesta
    Dim servidor As Byte
    Dim MD5exe As String
    Dim macaddress As String
    
    #If TDSFacil = 1 Then
        servidor = 2
    #Else
        servidor = 1
    #End If

    #If LOCALHOST = 0 And testeo = 1 Then
       servidor = 4
    #End If

    MD5exe = generarMD5
    macaddress = UserMac
    
    infoConexion = modLogin.iniciarConexionCrear(servidor, macaddress, GetIdentificacionPC(), MD5exe)
    
    If infoConexion.error = 0 Then
        conectarParaCrear.error = 0
        conectarParaCrear.errordesc = ""
        conectarParaCrear.datos = infoConexion.datosConexion
    Else
        conectarParaCrear.error = infoConexion.error
        conectarParaCrear.errordesc = infoConexion.errordesc
    End If
        
    
End Function

Private Function generarMD5() As String
'generarMD5 = "28d764a391f98fd1e6277ec649ba0d3f"
'generarMD5 = "09fad505e654d082bd621854533f5a29"
'generarMD5 = "f9aed88194e4b0e178afa47b7cfd1e5a"
 'nerarMD5 = "02c7c54a059169e8a753f8d39f15ce04"
'Exit Function

#If testeo = 1 Then
    generarMD5 = "A"
#Else
  '  generarMD5 = "6e6da7d1dfd70460370075679f9f5c70"
    generarMD5 = MD5String(MD5File(app.Path & "\" & app.EXEName & "." & "exe") & "PRPEPE9")
    '#If TDSFacil = 1 Then
    '   generarMD5 = MD5String(MD5File(app.Path & "\Tierras del Sur Innova.exe") & "PRPEPE9")
    '#Else
    '    generarMD5 = MD5String(MD5File(app.Path & "\" & "TDS" & ".exe") & "PRPEPE9")
    '#End If
#End If

End Function
