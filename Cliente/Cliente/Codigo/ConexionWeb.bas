Attribute VB_Name = "ConexionWeb"
Option Explicit

#If NuevaVersion Then
Public Const PUERTO_WEB = 7011
#ElseIf TDSFacil Then
Public Const PUERTO_WEB = 7010
#Else
Public Const PUERTO_WEB = 7009
#End If


'Private Const MAX_CONEXIONES = 10
'Public CantidadConexiones As Integer


'Public Sub IniciarComponente()
'Dim i As Integer

'For i = 0 To MAX_CONEXIONES
'    If i > 0 Then Load frmMain.SocketC(i)
'    frmMain.SocketC(i).AddressFamily = 2
'    frmMain.SocketC(i).Protocol = 6
'    frmMain.SocketC(i).SocketType = 1
'    frmMain.SocketC(i).LocalPort = PUERTO_WEB + 100
'    frmMain.SocketC(i).Binary = True
'    frmMain.SocketC(i).BufferSize = 128
'    frmMain.SocketC(i).Blocking = False
'    frmMain.SocketC(i).Interval = 500000
'Next

'frmMain.SocketC(0).listen

'CantidadConexiones = 0

'End Sub
'Private Function ObtenerSlot() As Integer
'Dim slot As Integer
'Dim encontrado As Boolean
'encontrado = False

'slot = 1
'Do While slot <= MAX_CONEXIONES And Not encontrado
'    If Not frmMain.SocketC(slot).Connected Then
'        encontrado = True
'    Else
'        slot = slot + 1
'    End If
'Loop

'ObtenerSlot = slot

'End Function
'Public Sub AceptarConexion(SocketId As Integer)
'On Error GoTo AceptarConexion_Error

'Dim slot As Integer
'Dim datos As String

'slot = ObtenerSlot()
'debug.Print slot
'Debug.Print frmMain.SocketC(1).PeerAddress

'If slot < MAX_CONEXIONES Then

 '   frmMain.SocketC(slot).accept = SocketId
 '   CantidadConexiones = CantidadConexiones + 1

'    frmMain.nuevoSocket = CantidadConexiones
   
 '   If frmMain.SocketC(slot).PeerAddress = "127.0.0.1" Then
         'proceso lo que se me envio
'        datos = CantidadConexiones
'        frmMain.SocketC(slot).Write datos, Len(datos)
'        frmMain.SocketC(slot).Flush
'    End If
 '   CerrarConexion (slot)
'Else
'    frmMain.SocketC(0).Disconnect
'    frmMain.SocketC(0).listen
'End If

'Exit Sub

'AceptarConexion_Error:

'LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure AceptarConexion of Módulo ConexionWeb" & slot
'
'End Sub

'Public Sub CerrarConexion(slot As Integer)
'frmMain.SocketC(slot).Disconnect
'End Sub

'Public Sub verEstados()
'frmMain.List.Clear
'Dim i As Integer
'For i = 0 To MAX_CONEXIONES
'    frmMain.List.AddItem frmMain.SocketC(i).Connected
'Next
'End Sub
