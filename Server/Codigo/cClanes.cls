VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cClanes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'************
' Objetivo de la clase Clanes:
'   Contiene el conjunto de clanes. Contiene la logica de lo que solo respecta
'   al conjunto de clanes.
'   No envia información a los usuarios
'   Colaboradores: clase Clan

Option Explicit
Private CantidadClanes As Integer
Private clanes() As cClan
Private listaClanes As String

'Un iterador simple para recorrer la lista desde afuera
Private iteradorPos As Integer

Private Sub Class_Initialize()
    CantidadClanes = 0
    ReDim clanes(1 To 1)
End Sub

Public Function ExisteClan(ByVal NombreClan As String) As Boolean

If ObtenerIDClan(NombreClan) = 0 Then
    ExisteClan = False
Else
    ExisteClan = True
End If

End Function
Public Sub NuevoClan(clan As cClan)
    If clan.id > UBound(clanes) Then
        ReDim Preserve clanes(1 To clan.id)
    End If
    
    Set clanes(clan.id) = clan

    'Solo cuento los clanes activos
    If clan.getEstado = eEstadoClan.Activo Then
        CantidadClanes = CantidadClanes + 1
        actualizarListaClanesString
    End If
End Sub
Private Function ObtenerIDClan(ByVal nombre As String) As Long
    Dim i As Long

    nombre = UCase$(nombre)
    For i = 1 To UBound(clanes)
        If Not clanes(i) Is Nothing Then
            If UCase(clanes(i).getNombre()) = nombre Then
                ObtenerIDClan = i
                Exit Function
            End If
        End If
    Next
    
    ObtenerIDClan = 0
End Function
Public Sub cargar()
    Dim sql As String
    Dim info As ADODB.Recordset
    Dim clan As cClan
        
    sql = "SELECT ID FROM " & DB_NAME_PRINCIPAL & ".clanes"

    Set info = conn.Execute(sql)

    Do While (info.EOF = False)

        Set clan = New cClan

        Call clan.IniciarDesdeDB(info!id)

        Call Me.NuevoClan(clan)

        info.MoveNext
    Loop
    
    info.Close
    Set info = Nothing

    actualizarListaClanesString

End Sub

'*****************************************************************************************************
'                                   SETTERS Y GETTERS
'*****************************************************************************************************
Public Function getClan(id As Long) As cClan
    If id > UBound(clanes) Then
        Set getClan = Nothing
    Else
        Set getClan = clanes(id)
    End If
End Function
Public Function getClanPorNombre(ByVal nombre As String) As cClan
    Dim id As Long
    
    id = ObtenerIDClan(nombre)
    If id = 0 Then
        Set getClanPorNombre = Nothing
    Else
        Set getClanPorNombre = clanes(id)
    End If
End Function
Public Function getCantidad() As Integer
    getCantidad = CantidadClanes
End Function

Public Function getClanesString() As String
    getClanesString = listaClanes
End Function

Private Sub actualizarListaClanesString()
    generarClanesString
End Sub

Public Sub cambiarNombreClan(clan As cClan, nombre As String)
    Call clan.setNombre(nombre)
    actualizarListaClanesString
End Sub
'Genera los nombres de los clanes separados por una coma
Private Sub generarClanesString()
    Dim i As Integer
    
    listaClanes = ""
    For i = 1 To UBound(clanes())
        If Not clanes(i) Is Nothing Then
            'No muestro los clanes disueltos
            If clanes(i).getEstado = eEstadoClan.Activo Then
                listaClanes = listaClanes & clanes(i).getNombre & ","
            End If
        End If
    Next i
'Debug.Print listaClanes
End Sub

Public Sub disolverclan(clan As cClan)
    clan.disolver
    CantidadClanes = CantidadClanes - 1
    Call actualizarListaClanesString
End Sub

Public Sub ReanudarClan(clan As cClan)
    clan.reanudar
    CantidadClanes = CantidadClanes + 1
    Call actualizarListaClanesString
End Sub

Public Function obtenerUltimoClanDisueltoPorUsuario(idUsuario As Long) As cClan
'Ver alguna forma mas optima...
    Dim i As Integer
    
    Dim ultimaFecha As Date
    Dim clanNumero As Integer
    
    ultimaFecha = 0
    
    For i = UBound(clanes()) To 1 Step -1
        If Not clanes(i) Is Nothing Then
            'No muestro los clanes disueltos
            If clanes(i).getEstado = eEstadoClan.Disuelto Then
                If clanes(i).getIDLider = idUsuario Then
                    If ultimaFecha < clanes(i).getFechaDisolucion Then
                        ultimaFecha = clanes(i).getFechaDisolucion
                        clanNumero = i
                    End If
                End If
            End If
        End If
    Next i
    
    If clanNumero > 0 Then
        Set obtenerUltimoClanDisueltoPorUsuario = clanes(clanNumero)
    Else
        Set obtenerUltimoClanDisueltoPorUsuario = Nothing
    End If
End Function

'ITerador
Public Sub iteradorIniciar()
    iteradorPos = 1
End Sub

Public Function iteradorObtener() As cClan
    Dim encontrado As Boolean
   
    encontrado = False
    
    Do While Not encontrado
        If Not clanes(iteradorPos) Is Nothing Then
        
            If clanes(iteradorPos).getEstado = Activo Then
                encontrado = True
            End If
            
        End If
        iteradorPos = iteradorPos + 1
    Loop
    
    Set iteradorObtener = clanes(iteradorPos - 1)
    
End Function
