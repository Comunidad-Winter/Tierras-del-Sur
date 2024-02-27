VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileJSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_JSON As Dictionary

Private m_items() As cItem
Private m_itemCount As Long

Public Function getItems() As cItem()
    getItems = m_items
End Function

Public Function init(strFile As String) As Boolean
On Error GoTo errHandler:

    Dim handle As Integer
    Dim strData As String
    Dim json_ As Dictionary
    
    Dim element As Variant
    Dim value As Object
     
    Dim dic As Dictionary
   
    Dim items() As Variant
    
    handle = FreeFile

    'Abro el archivo. Y lo leo todo deuna
    Open strFile For Input As #handle
    
    strData = Input$(LOF(handle), handle)
    
    Close #handle

    'Esto devuelve un Dictionary o un Collection
    Set json_ = JSON.parse(strData)
    
    'Blanqueamos
    strData = ""
   
    ReDim m_items(1 To json_.Count)
   
    m_itemCount = 0

    For Each element In json_
        m_itemCount = m_itemCount + 1

        Set value = json_.item(element)

        If TypeName(value) <> "Dictionary" Then
            GoTo errHandler
        Else

            Set dic = value

            If dic.Count <> 3 Then
                GoTo errHandler
            Else
                'Creo el item
                Set m_items(m_itemCount) = New cItem
                
                items = dic.items

                m_items(m_itemCount).init CStr(element), CStr(items(0)), CStr(items(1)), items(2), m_itemCount, 0
            End If
       End If
    Next
    
    Set dic = Nothing
    Set value = Nothing
    Set json_ = Nothing

    init = True
   Exit Function

errHandler:
    Debug.Print "Error " & Err.Number & " al cargar JSON " & strFile & " : " & Err.Description
    init = False
    
End Function
