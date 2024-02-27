VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_id As Long 'Es un identificador unico generado externamente
Private m_mainKey As String 'Es el nombre de la sección
 
'Clase que nos ayuda persitir la seccion
Private m_iniFile As cIniManager

' Items que tiene la seccion
Private m_items() As cItem
Private m_itemCount As Long

'Obtiene todos los items de la seccion
Public Function getItems() As cItem()
    getItems = m_items
End Function

'Obtiene un item especifico en base a la posicion que tiene en la lista
Public Function getItem(ByVal i As Long) As cItem
    Set getItem = m_items(i)
End Function

'Obtiene un item en base a su CLAVE
Public Function getItemByName(ByVal nombre As String) As cItem
    Dim loopItem As Integer
    
    For loopItem = LBound(m_items) To UBound(m_items)
        If UCase$(m_items(loopItem).getKey) = nombre Then
            Set getItemByName = m_items(loopItem)
            Exit Function
        End If
    Next
    
    Set getItemByName = Nothing
End Function

'Inicializa la seccion
Public Sub init(ByRef file As cFileJSON, ByRef iniFile As cIniManager, ByVal id As Long, ByRef mainKey As String)
    Dim key As String, value As String
    Dim items() As cItem, i As Long, item As cItem
    
    'Nos quedamos con nuestro ayudante de persistencia
    Set m_iniFile = iniFile
    
    ' Datos de la seccion
    m_id = id
    m_mainKey = mainKey

    'Estructura de la seccion
    items = file.getItems()
        
    m_itemCount = UBound(items)
        
    ReDim Preserve m_items(LBound(items) To UBound(items))
    
    'Cargamos los items
    For i = LBound(items) To UBound(items)
        'Obtenemos la clave
        key = items(i).getKey
        'Obtenemos el valor persistido
        value = m_iniFile.getValue(m_mainKey, key)
        'Creamos el item y le ponemos la informacion
        Set m_items(i) = New cItem
        m_items(i).copy items(i)
        m_items(i).setValue value
    Next i
End Sub

Public Sub saveInFile()
    Dim i As Long
    Dim value As String
    Dim key As String
    
    'Recorremos cada uno de los item
    For i = 1 To m_itemCount
        'Obtenemos la clave
        key = m_items(i).getKey()

        'Obtenemos el valor
        value = m_items(i).getValue()
        'Lo persistimos
        m_iniFile.ChangeValue m_mainKey, key, value
    Next i
End Sub

'Retorna el nombre de la seccion
Public Function getName() As String
    getName = m_mainKey
End Function

'Retorna el identificador unico de la seccion
Public Function getID() As Long
    getID = m_id
End Function
