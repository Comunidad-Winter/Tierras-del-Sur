Attribute VB_Name = "modLogin"
'**********************************************************
'*  Este modulo se conecta con alguno de los Logins
'* servers para saber a que dirección IP y puerto se
' debe comunicar el usuario
'**********************************************************
Option Explicit

Public Type tDatosConexion
    ip As String
    puerto As Integer
    hash As String
    semilla As Byte
End Type

Public Type tLoginServerRespuesta
    error As Byte
    errordesc As String
    datosConexion As tDatosConexion
End Type

Private Type tDirLoginServer
    direccion As String
    'tipo de conexiones. Web / Socket
    probado As Boolean
End Type

Private loginServers() As tDirLoginServer 'Login servers habilitados
Private loginServersErrores() As String 'Error que devuelve cada login server

Private idLoginServerPreferido As Byte ' Login Server que nos funciono

Public erroresDescripcion() As String ' Descripciones DoEvents los errores

Private Enum eLoginErrores
    cdclRespuestaInentendible = 101
End Enum

Public Const JUEGO_DESACTUALIZADO = 8
Public Const ERROR_CONEXION = 7 'No tiene internet

Private Const ERROR_DESCONOCIDO = 6 'Esto lo tienen que compartir tanto el cliente como el login

' Direcciones IPS a las cuales no se pudo conectar el usuario
Private ipsNoAccesibles() As String



Public Sub iniciarLogins()

Dim loopLoginServer As Byte

' Por defecto no tiene ningun loginserver especial
idLoginServerPreferido = 0

'Cargamos los logins servers

#If testeo = 0 Then
    ReDim loginServers(1 To 3) As tDirLoginServer
    ReDim loginServersErrores(LBound(loginServers) To UBound(loginServers)) As String

    Call resetLoginServerPrueba

   loginServers(1).direccion = "https://ld7.tierrasdelsur.cc" ' NORMAL
   loginServers(2).direccion = "https://ld.tierrasdelsur.cc" ' Se cae USA
   loginServers(3).direccion = "https://ld8.tierrasdelsur.cc" ' Se cae cloudflare / USA
    
#Else
    #If LOCALHOST = 1 Then
        ReDim loginServers(1 To 1) As tDirLoginServer
        ReDim loginServersErrores(LBound(loginServers) To UBound(loginServers)) As String
       
       loginServers(1).direccion = "127.0.0.1:12000" ' NORMAL
    #Else
        ReDim loginServers(1 To 3) As tDirLoginServer
        ReDim loginServersErrores(LBound(loginServers) To UBound(loginServers)) As String
        
        loginServers(1).direccion = "http://ld7.tierrasdelsur.cc" ' NORMAL
        loginServers(2).direccion = "http://ld8.tierrasdelsur.cc" ' Se cae USA
        loginServers(3).direccion = "http://ld.tdsx.com.ar" ' Se cae cloudflare / USA
    #End If
    
    Call resetLoginServerPrueba
#End If

'Cargamos la lista de errores que puede retornar el login server
ReDim erroresDescripcion(1 To 8) As String

erroresDescripcion(1) = "Clave del personaje incorrecta."
erroresDescripcion(2) = "El personaje se encuentra online."
erroresDescripcion(3) = "No es posible conectarse al Mundo en estos momentos. Chequea el estado del juego en la página web."
erroresDescripcion(4) = "No es posible conectarse al Mundo en estos momentos. Chequea el estado del juego en la página web."
erroresDescripcion(5) = "Tu cuenta y tus personajes se encuentran bloqueada. Consultá más información en la web. "
erroresDescripcion(6) = "No es posible acceder a Tierras del Sur en estos momentos, por favor intente más tarde."
erroresDescripcion(7) = "Verifique tener acceso a internet. No es posible acceder a Tierras del Sur en estos momentos, por favor intente más tarde."
erroresDescripcion(8) = "El juego se encuentra desactualizado."

' Lista de direcciones IPS no bloqueadas
ReDim ipsNoAccesibles(0) As String
ipsNoAccesibles(0) = ""

Call EnlaceWeb.iniciarConector
End Sub


Public Sub agregarIpNoAccesible(ip As String)

If Not ipsNoAccesibles(UBound(ipsNoAccesibles)) = "" Then
    ReDim Preserve ipsNoAccesibles(UBound(ipsNoAccesibles) + 1) As String
End If

ipsNoAccesibles(UBound(ipsNoAccesibles)) = ip

Debug.Print "Agregada IP no accesible:" & ip
End Sub

Private Function obtenerLoginServerSinProbar() As Byte
    
    Dim loopLoginServer As Byte
    
    For loopLoginServer = 1 To UBound(loginServers)
    
        If loginServers(loopLoginServer).probado = False Then
            obtenerLoginServerSinProbar = loopLoginServer
            Exit Function
            
        End If
    Next loopLoginServer

    obtenerLoginServerSinProbar = 0
End Function

Private Sub resetLoginServerPrueba()

    Dim loopLoginServer As Byte
    
    For loopLoginServer = LBound(loginServers) To UBound(loginServers)
        loginServers(loopLoginServer).probado = False
    Next loopLoginServer
    
End Sub

Private Sub resetLoginServerErrores()

    Dim loopLoginServer As Byte
    
    For loopLoginServer = LBound(loginServersErrores) To UBound(loginServersErrores)
        loginServersErrores(loopLoginServer) = ""
    Next loopLoginServer
    
End Sub

Private Function obtenerListaErrores() As String

    Dim loopLoginServer As Byte
    
    For loopLoginServer = LBound(loginServersErrores) To UBound(loginServersErrores)
        obtenerListaErrores = obtenerListaErrores & loginServersErrores(loopLoginServer)
    Next loopLoginServer
    
End Function


Private Function consultarDatosConexionLoginServer(ByVal consulta As String) As tLoginServerRespuesta

    Dim error As Byte
    Dim resultado As String
    Dim loginServer As Byte
    Dim conectado As Boolean

    Dim resultadoJSON As Object
    Dim esJsonValido As Boolean
    Dim puedeNoTenerInternet As Boolean
    
    On Error GoTo hErr
    
    Dim checksum As String
    
    'Generamos el checksum
    
    ' Resteo los login servers probados
    Call resetLoginServerPrueba
    Call resetLoginServerErrores
    
    ' Me conecto al loginserver preferido o al primero
    If idLoginServerPreferido = 0 Then
        loginServer = obtenerLoginServerSinProbar
    Else
        loginServer = idLoginServerPreferido
    End If
    
    conectado = False
    puedeNoTenerInternet = True
    
    'Mientras no me haya podido conectar y me queden loginservers por probar...
    Do While loginServer > 0 And conectado = False
                
        'Tratamos de obtener los datos del login server
        loginServers(loginServer).probado = True
        error = 0
        
        Debug.Print "CONECTANDO a " & loginServers(loginServer).direccion
         
        resultado = EnlaceWeb.ejecutarConsulta(loginServers(loginServer).direccion, "TDSExternalUser", consulta, error)

        If error > 0 Then ' Se produjo una falla al intentar conectarse
            'Si es una falla tecnica pruebo con el siguiente nodo
            loginServersErrores(loginServer) = CStr(error) & "A" & resultado
            
            If error = eEnlanceWebErrores.ejcDireccionInalcanzable Then
                puedeNoTenerInternet = (puedeNoTenerInternet And True)
            Else 'Otro tipo de error queire decir que llego a conectarse
                puedeNoTenerInternet = False
            End If
        Else ' Me respondio un login server
            puedeNoTenerInternet = False 'Lo descartamos porque tuve respuesta
            
            esJsonValido = False
            
            Set resultadoJSON = JSON.parse(resultado)
                        
            ' Lo que se devolvio es algo correcto?
            If Not resultadoJSON Is Nothing Then
                esJsonValido = True
                'Ningun problema tecnico?
                If resultadoJSON.item("error_tecnico") = False Then
                    idLoginServerPreferido = loginServer
                    conectado = True
                
                    If resultadoJSON.item("habilitado") = 1 Then
                        'OK:
                        'IP, PUERTO Y HASH
                        consultarDatosConexionLoginServer.datosConexion.ip = resultadoJSON.item("ip")
                        consultarDatosConexionLoginServer.datosConexion.puerto = resultadoJSON.item("puerto")
                        consultarDatosConexionLoginServer.datosConexion.hash = resultadoJSON.item("hash")
                        consultarDatosConexionLoginServer.datosConexion.semilla = resultadoJSON.item("semilla")
                                                
                        consultarDatosConexionLoginServer.error = 0
                    Else
                        'NO OK
                        ' Bloqueo de IP / MAC ADDRESS / PAIS por ban
                        ' No hay Nodos disponibles para esta persona
                        ' Datos de autentificacion incorrectos
                        consultarDatosConexionLoginServer.error = resultadoJSON.item("razon")
                    End If
                End If
            End If 'Por ahora no discrimino si hubo una falla en la respuesta
            
            If Not esJsonValido Then
                'Por algun motivo no se pudo entender correctamente la respuesta
                loginServersErrores(loginServer) = "B" & CStr(eLoginErrores.cdclRespuestaInentendible)
            End If
        End If
        
        ' Si no me pude conectar, pruebo con otro
        If conectado = False Then
            loginServer = obtenerLoginServerSinProbar
        End If
            
    Loop

    If conectado = False Then
        'No queda ningun otro login server
        'Posibles razones: a) No tengo internet b) todos los logins servers estan caidos.
        If puedeNoTenerInternet Then
            consultarDatosConexionLoginServer.error = ERROR_CONEXION
            consultarDatosConexionLoginServer.errordesc = "X4A"
            ' Si me puedo conectar a google. Tengo Internet (los logins servers estan caidos). Sino no tengo internet.
        Else
            consultarDatosConexionLoginServer.error = ERROR_DESCONOCIDO
            consultarDatosConexionLoginServer.errordesc = obtenerListaErrores()
        End If
            'c) Se pudo conectar pero ningun login server le termino de dar la info. Login Severs congestionados.
            
    End If
 
Exit Function

hErr:
consultarDatosConexionLoginServer.error = ERROR_DESCONOCIDO
consultarDatosConexionLoginServer.errordesc = "Z" & Err.Number
    
End Function

Private Function generarListaIpsNoAdmitidas()
    Dim ipsNoAdmitidas As String
    Dim loopIp As Integer
    
    ipsNoAdmitidas = ""
    For loopIp = LBound(ipsNoAccesibles) To UBound(ipsNoAccesibles)
        If Not ipsNoAccesibles(loopIp) = "" Then
        
            If Not ipsNoAdmitidas = "" Then ipsNoAdmitidas = ipsNoAdmitidas & ","
        
            ipsNoAdmitidas = ipsNoAdmitidas & ipsNoAccesibles(loopIp)
        End If
    Next
    
    generarListaIpsNoAdmitidas = ipsNoAdmitidas
End Function
' Devuelve una estructura valida si es posible
Public Function iniciarConexionPersonaje(servidor As Byte, UserName As String, UserPassword As String, Mac As String, pc As String, MD5exe As String) As tLoginServerRespuesta
    Dim consulta As String
    Dim key As Byte

    key = CByte(Int(Rnd() * 255))
    
    consulta = "{'r' : " & key & ", 'accion':'INGRESAR', 's':" & servidor & ",'u': '" & UserName & "','p':'" & UserPassword & "', 'm':'" & Mac & "', 'pc':'" & pc & "', 'md':'" & MD5exe & "', 'ipse' :'" & generarListaIpsNoAdmitidas & "'}"
    
    Debug.Print consulta
    Call intercambiarComillas(consulta)
    
    iniciarConexionPersonaje = consultarDatosConexionLoginServer(consulta)
End Function

Public Function iniciarConexionCrear(servidor As Byte, Mac As String, pc As String, MD5exe As String) As tLoginServerRespuesta
    
    Dim consulta As String
    Dim key As Byte
    
    key = CByte(Int(Rnd() * 255))
    
    consulta = "{'r' : " & key & ", 'accion':'CREAR', 's':" & servidor & ",'m':'" & Mac & "', 'pc':'" & pc & "', 'md':'" & MD5exe & "', 'ipse' :'" & generarListaIpsNoAdmitidas & "'}"
    
    Debug.Print consulta
    Call intercambiarComillas(consulta)
    'Call eliminarEspacios(consulta)
    
    iniciarConexionCrear = consultarDatosConexionLoginServer(consulta)
    
End Function

Private Sub eliminarEspacios(ByRef cadena As String)
    cadena = Replace$(cadena, " ", "")
End Sub
Private Sub intercambiarComillas(ByRef cadena As String)
    cadena = Replace$(cadena, "'", Chr(34))
End Sub


Public Sub informarFalla(conexion As tDatosConexion)
End Sub


