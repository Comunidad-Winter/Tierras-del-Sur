VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CondicionEventoNoRepClan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iCondicionEvento

Private Type infoClanRegistrado
    IDUsuarioLider As Long
    IDClan As Long
End Type

Private clanesAnterioresInfo() As infoClanRegistrado
Private cantidadClanesRegistrados As Byte

Private Sub iCondicionEvento_asentarIngresoEquipo(integrantesIndex() As Integer)
    Dim posicionLibre As Byte
    
    cantidadClanesRegistrados = cantidadClanesRegistrados + 1
    
    For posicionLibre = 1 To UBound(clanesAnterioresInfo)
        If clanesAnterioresInfo(posicionLibre).IDClan = 0 Then
            clanesAnterioresInfo(posicionLibre).IDClan = UserList(integrantesIndex(1)).ClanRef.id
            clanesAnterioresInfo(posicionLibre).IDUsuarioLider = UserList(integrantesIndex(1)).id
            Exit For
        End If
    Next
    
    
End Sub

Private Function iCondicionEvento_obtenerDescripcion() As String
    iCondicionEvento_obtenerDescripcion = "No vale repetir clan."
End Function

Private Function iCondicionEvento_puedeIngresarEquipo(integrantesIndex() As Integer) As Boolean

    Dim UserIndex As Integer
    Dim loopC As Byte
    Dim loopClanes As Byte
    Dim clanActual As Long
    
    For loopC = 1 To UBound(integrantesIndex)
    
        UserIndex = integrantesIndex(loopC)
        
        If UserList(UserIndex).ClanRef Is Nothing Then
            iCondicionEvento_puedeIngresarEquipo = False
            Exit Function
        Else
            If clanActual = 0 Then
            
                clanActual = UserList(UserIndex).ClanRef.id
                
                For loopClanes = 1 To UBound(clanesAnterioresInfo)
                    If clanesAnterioresInfo(loopClanes).IDClan = clanActual Then
                        iCondicionEvento_puedeIngresarEquipo = False
                        Exit Function
                    End If
                Next
            ElseIf Not clanActual = UserList(UserIndex).ClanRef.id Then
                iCondicionEvento_puedeIngresarEquipo = False
                Exit Function
            End If
            
        End If
        
    Next loopC
    
    iCondicionEvento_puedeIngresarEquipo = True
    
End Function

Private Sub iCondicionEvento_removerEquipoInscripto(integrantesid() As Long)
    Dim posicionLibre As Byte
    
    cantidadClanesRegistrados = cantidadClanesRegistrados - 1
    
    For posicionLibre = 1 To UBound(clanesAnterioresInfo)
        If clanesAnterioresInfo(posicionLibre).IDUsuarioLider = integrantesid(1) Then
            clanesAnterioresInfo(posicionLibre).IDClan = 0
            clanesAnterioresInfo(posicionLibre).IDUsuarioLider = 0
            Exit For
        End If
    Next
            
End Sub

Public Sub iCondicionEvento_setMaximaMemoria(slots As Byte)
    ReDim clanesAnterioresInfo(1 To slots) As infoClanRegistrado
    cantidadClanesRegistrados = 0
End Sub
