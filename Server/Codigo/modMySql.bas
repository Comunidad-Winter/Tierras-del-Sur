Attribute VB_Name = "modMySql"
Option Explicit

' Conexion a la base/s de datos
Public conn As ADODB.Connection 'Conexion
Public sql As String 'String Sql general

Public constr As String

' Conexion a las base de datos
Private Const DB_SERVER = "localhost"
Private Const DB_PORT = "3306"

Public Const DB_NAME_CUENTAS = "web.cuentas_cache"


#If SERVER_PRUEBAS = 1 Then
    Public Const DB_NAME_PRINCIPAL = "tds_balance"
    Private Const DB_USER = "usr_balance"
    Private Const DB_PASS = "857kGPQdXnuTvAthvPvH"
#End If

#If testeo = 1 Then
    Public Const DB_NAME_PRINCIPAL = "tds_alta"
    Private Const DB_USER = "root"
    Private Const DB_PASS = "1161846173Tomas"
#End If

Private Type tCola
    UserIndex As Integer
    idPersonaje As Long
    Password As String
    tok As Long
End Type

Public transaccionEnEspera As Boolean
Public cargarPersonajeIndexEspera As Integer
Public cargarPersonajeTokEnEspera As Long

Private slotColaLibre As Byte
Private slotAProcesar As Byte
Private ultimoTock As Long
Private fechaUltimaConsulta As Long

Private Const TAMANIO_COLA_ESPERA As Byte = 10 ' ESTO ES UN BUFFER CIRCULAR.
Private Const TIEMPO_MINIMO_SIN_CONSULTA As Long = 300000  ' 5 Minutos


Private colaEsperaLogin(1 To TAMANIO_COLA_ESPERA) As tCola

Public Sub procesarCargaColaEspera()
Dim UserIndex As Integer

    'Tengo algo que procesar?t
    If colaEsperaLogin(slotAProcesar).UserIndex > 0 Then
    
        UserIndex = colaEsperaLogin(slotAProcesar).UserIndex
    
        colaEsperaLogin(slotAProcesar).UserIndex = 0
        
        transaccionEnEspera = False
        
        cargarPersonajeIndexEspera = UserIndex
        cargarPersonajeTokEnEspera = colaEsperaLogin(slotAProcesar).tok
                
        Call solicitarInfoPersonaje(UserIndex, colaEsperaLogin(slotAProcesar).idPersonaje, colaEsperaLogin(slotAProcesar).Password)
        
        'Paso al siguiente slot
         slotAProcesar = (slotAProcesar Mod TAMANIO_COLA_ESPERA) + 1
    Else
        'No tengo nada que procesar
        transaccionEnEspera = False
    End If

End Sub

Public Sub enviarPingBaseDeDatos()

    If GetTickCount() < fechaUltimaConsulta + TIEMPO_MINIMO_SIN_CONSULTA Then
        Exit Sub
    End If
    
    Call LogDesarrollo("Enviando PING")
    
    'Comeinzo la transaccion.
    transaccionEnEspera = True
    fechaUltimaConsulta = GetTickCount
        
    'Es importante esto dado que se usa un FOR UPDATE
    frmMysqlAuxiliar.cargadorPersonajes.BeginTrans
    
    sql = "SELECT 1 AS fakeload FROM " & DB_NAME_PRINCIPAL & ".usuarios LIMIT 1"
    
    'A quien le vamos a dar el personaje?
    cargarPersonajeIndexEspera = -1
    cargarPersonajeTokEnEspera = -1
    
    Call frmMysqlAuxiliar.cargadorPersonajes.Execute(sql, , adCmdText Or adAsyncExecute)
End Sub

Public Function solicitarInfoPersonaje(UserIndex As Integer, idPersonaje As Long, passwordPersonaje As String) As Boolean

Dim sql As String

ultimoTock = ultimoTock + 1

If ultimoTock > 100000 Then
    ultimoTock = 1
End If

If transaccionEnEspera = False Then

    'Comeinzo la transaccion.
    transaccionEnEspera = True
    fechaUltimaConsulta = GetTickCount
        
    'Es importante esto dado que se usa un FOR UPDATE
    frmMysqlAuxiliar.cargadorPersonajes.BeginTrans
    
    sql = "SELECT SQL_NO_CACHE usr.* , gms.Privilegio as Privilegios, " & _
        "IF (cuenta.FECHAVENCIMIENTO IS NOT NULL AND CUENTA.FECHAVENCIMIENTO > UNIX_TIMESTAMP(), 'SI', 'NO') AS ESPREMIUM, " & _
        "cuenta.ESTADO, cuenta.BLOQUEADA, cuenta.SEGUNDOS_TDSF " & _
        "FROM " & DB_NAME_PRINCIPAL & ".usuarios AS usr " & _
        "LEFT JOIN " & DB_NAME_PRINCIPAL & ".juego_gms AS gms ON usr.ID = gms.IDUsuario " & _
        "LEFT JOIN " & DB_NAME_CUENTAS & " AS cuenta ON cuenta.IDCuenta = usr.IDCuenta " & _
        "WHERE usr.ID = " & idPersonaje & _
        " AND usr.passwordB = '" & mysql_real_escape_string(passwordPersonaje) & _
        "' FOR UPDATE"
    
    'A quien le vamos a dar el personaje?
    cargarPersonajeIndexEspera = UserIndex
    cargarPersonajeTokEnEspera = ultimoTock
    
    UserList(UserIndex).TokSolicitudDePersonaje = ultimoTock
    
    Call frmMysqlAuxiliar.cargadorPersonajes.Execute(sql, , adCmdText Or adAsyncExecute)
    
    solicitarInfoPersonaje = True
    
Else
    'Tengo espacio en la cola?
    'Como es un buffer circular si el userindex es mayor a 0 quiere decir que no tengo espacio
    If colaEsperaLogin(slotColaLibre).UserIndex = 0 Then
    
        'Guardo los datos
        colaEsperaLogin(slotColaLibre).UserIndex = UserIndex
        colaEsperaLogin(slotColaLibre).idPersonaje = idPersonaje
        colaEsperaLogin(slotColaLibre).Password = passwordPersonaje
        colaEsperaLogin(slotColaLibre).tok = ultimoTock
          
        'Avanzo al siguiente slot
            ' Cuando sea 9.   10 Mod 10 = 0
        slotColaLibre = (slotColaLibre Mod TAMANIO_COLA_ESPERA) + 1
        
        UserList(UserIndex).TokSolicitudDePersonaje = ultimoTock
            
        solicitarInfoPersonaje = True
    Else
        'No se puede encolar la solicitud
        solicitarInfoPersonaje = False
        Call LogError("No hay mas espacios para guardar solicitudes.")
    End If
    'No puedo ejecutar
End If ' Encolo la solicitud

End Function


'CSEH: Nada
Public Sub iniciarConexionBaseDeDatos()

On Error GoTo iniciarConexionBaseDeDatos_Err

Dim TempInt As Integer

Set conn = New ADODB.Connection

Call LogDesarrollo("Iniciando conexion con la base de datos")

'Conectamos el general
constr = "Provider=MSDASQL.1;Password=;Persist Security Info=True;User ID=;Extended Properties=" & Chr$(34) & "DRIVER={MySQL ODBC 8.0 ANSI Driver};DESC=;DATABASE=" & DB_NAME_PRINCIPAL & ";SERVER=" & DB_SERVER & ";UID=" & DB_USER & ";PASSWORD=" & DB_PASS & ";PORT=" & DB_PORT & ";OPTION=16387;STMT=;" & Chr$(34)

conn.Open constr

'Cargamos el formulario en donde tenemos el objeto que nos permite cargar los personajes
Load frmMysqlAuxiliar

Set frmMysqlAuxiliar.cargadorPersonajes = New ADODB.Connection

frmMysqlAuxiliar.cargadorPersonajes.Open constr

' Cola de carga de personajes
Dim loopC As Byte

For loopC = 1 To UBound(colaEsperaLogin)
    colaEsperaLogin(loopC).idPersonaje = 0
    colaEsperaLogin(loopC).Password = ""
    colaEsperaLogin(loopC).UserIndex = 0
Next loopC

slotColaLibre = 1
slotAProcesar = 1
transaccionEnEspera = False
cargarPersonajeIndexEspera = 0

    Exit Sub

iniciarConexionBaseDeDatos_Err:
    LogError Err.Description & " en modMySql.iniciarConexionBaseDeDatos"
    End
End Sub

Public Sub reConectar()
    Call LogDesarrollo("Re conectando a la base de datos.")
    
    conn.Close
    Set conn = Nothing
    Set conn = New ADODB.Connection
    conn.Open constr
    
    Call LogDesarrollo("Re conectado a la base de datos.")
End Sub

'CSEH: Nada
Public Function ejecutarSQL(sql As String, Optional ByVal max_intentos As Byte = 4) As Boolean
Dim intentos As Byte

On Error Resume Next

intentos = 0

    Do
        
        If Not Err.Number = 0 Then
            Call modMySql.reConectar
        End If

        Err.Number = 0
 
        conn.Execute sql, , adExecuteNoRecords
        intentos = intentos + 1
    Loop While intentos < max_intentos And Err.Number <> 0
    
    If intentos = max_intentos And Err.Number <> 0 Then
        ejecutarSQL = False 'No se pudo ejecutar correctamente la sentencia
        Call LogSQLerror("ERROR al ejecutar: " & sql & ". Descripcion: " & Err.Description)
    Else
        ejecutarSQL = True
    End If

End Function

Public Function mysql_real_escape_string(cadena As String) As String

cadena = Replace(cadena, "\", "\\")
cadena = Replace(cadena, Chr(34), "\" + Chr(34))
cadena = Replace(cadena, Chr(39), "\" + Chr(39))

mysql_real_escape_string = cadena
End Function
