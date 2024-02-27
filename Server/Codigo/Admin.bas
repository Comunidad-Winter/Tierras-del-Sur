Attribute VB_Name = "Admin"
Option Explicit

Public Type tMotd
    texto As String
    formato As String
End Type

Public MaxLines As Integer
Public MOTD() As tMotd

Public Type tAPuestas
    Ganancias As Long
    Perdidas As Long
    Jugadas As Long
End Type

Public apuestas As tAPuestas
Public DebugSocket As Boolean

Private cantidadOnlineUltimaActualizacionDB As Integer

Public servidorAtacado As Boolean




'---------------------------------------------------------------------------------------
' Procedure : WorldSave
' DateTime  : 18/02/2007 19:03
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub WorldSave()

Dim loopX As Integer
Dim j As Integer, k As Integer

' Aivamos que se inicia el WordSave
EnviarPaquete Paquetes.MensajeSimple, Chr$(30), 0, ToAll

Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales

' Cantidad de mapas a guardar
For j = 1 To NumMaps
    If MapInfo(j).BackUp = 1 Then k = k + 1
Next j

FrmStat.ProgressBar1.min = 0
FrmStat.ProgressBar1.max = k
FrmStat.ProgressBar1.value = 0

' Guardamos los mapas
For loopX = 1 To NumMaps
    If MapInfo(loopX).BackUp = 1 Then
        Call SaveMapData(loopX)
        ' Barra de progreso
        FrmStat.ProgressBar1.value = FrmStat.ProgressBar1.value + 1
    End If
Next loopX

' Ocultamos la barra de prograso
FrmStat.Visible = False

' Avisamos que el WorldSave termino
EnviarPaquete Paquetes.MensajeSimple, Chr$(31), 0, ToAll

End Sub

'---------------------------------------------------------------------------------------
'Esta funcion actualiza la cantidad de usuarios online que hay en el juego en la base de datos
'para que pueda ser consultando desde la página web
Public Sub actualizarOnlinesDB(Optional ByVal forzar As Boolean = False)

'Solo actualizo la base de datos si la cantidad difiere de lo ultimo actualizado
'Actualizo si o si cuando hay 0 onlines para no perder conexion con la base de datos
If cantidadOnlineUltimaActualizacionDB <> NumUsers Or forzar Or NumUsers = 0 Then
    cantidadOnlineUltimaActualizacionDB = NumUsers
    sql = "UPDATE " & DB_NAME_PRINCIPAL & ".online set NumeroB=" & NumUsers & " WHERE CantidadB = 'Numero'"
    conn.Execute sql, , adExecuteNoRecords
End If

End Sub

Public Sub servidorComienzaAtaque()
    
    Admin.servidorAtacado = True
    
    frmMain.estadoServidor.Caption = "Atacado"
    frmMain.estadoServidor.ForeColor = vbRed
    
    LogDesarrollo ("El servidor comienza a ser atacado")
    
    Call Admin.liberarTodosSlots
End Sub

Public Sub servidorTerminaAtaque()
    
    Admin.servidorAtacado = False
    
    frmMain.estadoServidor.Caption = "Norma"
    frmMain.estadoServidor.ForeColor = vbGreen
    
    Call LogDesarrollo("Termina el ataque al servidor")
    
    Call Admin.liberarTodosSlots
End Sub

Public Sub liberarTodosSlots()

Dim i As Integer

For i = 1 To MaxUsers
    If Not UserList(i).ConnID = INVALID_SOCKET And UserList(i).flags.UserLogged Then
        Call CloseSocket(i)
    End If
Next i

End Sub

