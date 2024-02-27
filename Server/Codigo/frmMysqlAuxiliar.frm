VERSION 5.00
Begin VB.Form frmMysqlAuxiliar 
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2235
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMysqlAuxiliar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Use the WithEvents keyword to designate that events
' can be handled by this Connection object
Public WithEvents cargadorPersonajes As ADODB.Connection
Attribute cargadorPersonajes.VB_VarHelpID = -1

Private Sub cargadorPersonajes_Disconnect(adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
Logs.LogError ("Se desconecto de la base de datos el modulo de carga de personajes")
End Sub

' Note how the object name, objConn, is incorporated into the event Sub name
Public Sub cargadorPersonajes_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)

Dim UserIndex As Integer
Dim estaCorrecto As Boolean

If pRecordset Is Nothing Then Exit Sub
    
If RecordsAffected = 1 Then

    If cargarPersonajeIndexEspera = -1 Then
        Call Logs.LogDesarrollo("Se ejecuta correctamente el PING a la base de datos en el modulo de usuarios.")
        estaCorrecto = True ' No tengo que cerrar el Socket
        pRecordset.Close
        Set pRecordset = Nothing
    Else
        Debug.Print "PERSONAJE ENCONTRADO: " & pRecordset!nickb
        
        If UserList(cargarPersonajeIndexEspera).TokSolicitudDePersonaje = cargarPersonajeTokEnEspera Then
            estaCorrecto = TCP.conectarPersonaje(pRecordset, cargarPersonajeIndexEspera)
        Else
            'Al pedo cargamos el personaje. Esto no deberia suceder, deberia ser rapido la carga del personaje
            LogError ("Se tardo demasiado en cargar un personaje y el usuario cerro.")
            estaCorrecto = True ' Para que no desconecte al usuario
        End If
    End If
Else

    'Esto es raro que suceda ya que la clave la valida el MUNDO
    Debug.Print "PERSONAJE INEXISTENTE O CLAVE INVALIDA O TIME OUT"
    
    If Not cargarPersonajeIndexEspera = -1 Then
        If UserList(cargarPersonajeIndexEspera).TokSolicitudDePersonaje = cargarPersonajeTokEnEspera Then
        
            estaCorrecto = False
            
            If pError Is Nothing Then
                EnviarPaquete mbox, Chr$(1), cargarPersonajeIndexEspera
            Else
                Call LogError("Error al cargar personaje " & pError)
                EnviarPaquete mbox, Chr$(14) & "Se ha producido un error. Por favor intente ingresar nuevamente en un minuto.", cargarPersonajeIndexEspera, ToIndex
            End If
        Else
            estaCorrecto = True ' Para que no desconecte al usuario
        End If
    End If ' TODO. Sucedio un error al hacer el PING.
        
End If

' Commit la transaccion (Se abre el candado)
frmMysqlAuxiliar.cargadorPersonajes.CommitTrans

' Ingreso correctamente?
If estaCorrecto = False Then
    If Not CloseSocket(cargarPersonajeIndexEspera) Then LogError ("Execute Complete")
End If

'Proceso los que estaban esperando cargar su personaje
Call modMySql.procesarCargaColaEspera

End Sub

Private Sub cargadorPersonajes_InfoMessage(ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "info mensaje?"
End Sub

