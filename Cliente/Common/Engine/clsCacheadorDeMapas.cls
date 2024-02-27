VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMemMapManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type MAPA_ENTY
    MapNum As Integer
    UltimoAcceso As Long
    stored_map As clsMapCache
End Type

Private pMapas() As MAPA_ENTY

Private pMaxMapas As Integer
Private pCantidadMapas As Integer

Private Const MAX_DE_MAPAS_TOTALES As Integer = 1600 'Son los mapas del disco.
Private Tabla(MAX_DE_MAPAS_TOTALES) As Integer

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Class_Terminate()
On Error Resume Next
    Dim i As Long
    
    For i = 1 To pCantidadMapas
        Set pMapas(i).stored_map = Nothing
    Next i
    
    Erase pMapas
End Sub

Public Function Getstored_map(ByVal MapNum As Integer, ByVal ActMap As Integer) As Boolean
Dim Index As Integer

'paso1:
'    cuando cambio de mapa me fijo si el mapa actual esta en la lista
'        si está GOTO paso2
'        si no está lo guardo
'paso2:
'    busco si el nuevo mapa está en la lista
'        si está lo cargo y hago return TRUE
'        si no ésta hago return FALSE y se va a guardar en la memoria en el próximo cambio de mapa
        
        'Busco el mapa actual en la lista
        If ActMap = MapNum Then
            Getstored_map = False
            Exit Function
        End If
        If Tabla(ActMap) = 0 And ActMap > 0 Then
            'Guardo el mapa actual
            If pCantidadMapas >= pMaxMapas Then
                'Sacamos el que hace más que no usamos, y utilizamos el slot
                Index = CrearMapa(ActMap, BorraMenosUsado())
            Else
                Index = CrearMapa(ActMap)
            End If
            Tabla(ActMap) = Index
        End If
        
        'Busco el nuevo mapa en la lista
        If Tabla(MapNum) <> 0 Then
            With pMapas(Tabla(MapNum))
                .UltimoAcceso = GetTickCount
                If .stored_map.Cargar = False Then
                    Call Borrar(MapNum)
                    Getstored_map = False
                    Debug.Print "MAPA CACHEADO INCORRECTAMENTE(?)"
                Else
                    Getstored_map = True
                    Debug.Print "MAPA OK!"
                End If
            End With
        Else
            Getstored_map = False
            Debug.Print "NO SE ENCONTRO REGISTRO DEL MAPA"
        End If
End Function

Public Function Init(ByVal MaxMapas As Long) As Boolean
    pMaxMapas = MaxMapas
    
    If pMaxMapas < 1 Then
        Exit Function
    End If
    
    pCantidadMapas = 0
    
    Init = True
End Function

Public Sub BorrarTodo()
    Dim i As Long
    
    For i = 1 To pCantidadMapas
        Set pMapas(i).stored_map = Nothing
    Next i
    
    ReDim pMapas(0)
    pCantidadMapas = 0
End Sub

Public Function Borrar(ByVal MapNum As Integer) As Integer
    Dim indice As Long
    Debug.Print "BORRANDO MAPA "; MapNum; "DE LA MEMORA"
    
    'Obtenemos el ìndice
    indice = Tabla(MapNum)
    If indice = 0 Then Exit Function
    Tabla(MapNum) = 0
    'Lo eliminamos
    With pMapas(indice)
        Set .stored_map = Nothing
    End With
    
    'Actualizamos el número de gráficos
    pCantidadMapas = pCantidadMapas - 1
    
    'Movemos para atrás el resto de la lista
    For indice = indice To pCantidadMapas
        pMapas(indice) = pMapas(indice + 1)
    Next indice
    
    'Redimencionamos la lista
    ReDim Preserve pMapas(1 To pCantidadMapas)
End Function


Private Function CrearMapa(ByVal MapNum As Integer, Optional ByVal Index As Integer = -1) As Integer
    If Index < 0 Then
        Index = pCantidadMapas + 1
        ReDim Preserve pMapas(1 To Index)
    End If
    
    If Index = 0 Then Index = pCantidadMapas
    
    With pMapas(Index)
        .MapNum = MapNum
        .UltimoAcceso = GetTickCount
        
        Set .stored_map = New clsMapCache
        .stored_map.Guardar
    End With
    
    pCantidadMapas = pCantidadMapas + 1
    CrearMapa = Index
End Function

Private Function BorraMenosUsado() As Integer
    Dim valor As Long
    Dim i As Long
    'Inicializamos todo
    valor = GetTickCount() 'pMapas(1).UltimoAcceso
    BorraMenosUsado = 1
    'Buscamos cual es el que lleva más tiempo sin ser utilizado
    For i = 1 To pCantidadMapas
        If pMapas(i).UltimoAcceso < valor Then
            valor = pMapas(i).UltimoAcceso
            BorraMenosUsado = i
        End If
    Next i
    'Disminuimos el contador
    pCantidadMapas = pCantidadMapas - 1
    'Borramos la stored_map
    
    Set pMapas(BorraMenosUsado).stored_map = Nothing
    Tabla(pMapas(BorraMenosUsado).MapNum) = 0
End Function

Public Property Get MaxMapas() As Integer
    MaxMapas = pMaxMapas
End Property

Public Property Let MaxMapas(ByVal vNewValue As Integer)
    pMaxMapas = vNewValue
End Property

Public Property Get CantidadMapas() As Integer
    CantidadMapas = pCantidadMapas
End Property




