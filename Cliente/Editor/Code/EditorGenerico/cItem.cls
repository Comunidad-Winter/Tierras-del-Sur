VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********** Parametros generale
Private m_Key As String 'CLAVE. Ejemplo "GrhIndex"
Private m_descKey As String 'Forma legible de la clave Ejemplo: "Grafico"

Private m_keyIndex As Integer 'Indice de argumento en la clave
Private m_itemIndex As Integer 'Indice de item en la representacion

Private m_desk As String 'Descripcion de ayuda acerca del item
Private m_type As ItemType 'Tipo de item que es

' *********** Para valores Simples
Private m_values As Collection 'Lista de valores posibles. Cada valor es un par (ID, Descripcion)

'     Para cadena y numeros
Private m_minValue As Long 'Minima cantidad de caracteres / valor
Private m_maxValue As Long 'Maxima cantidad de caracteres / valor
'     Para enumerado
Private m_combinado As Boolean '¿Pueden tener variosvalores?
'       Para enumerado dinamico
Private m_fuente As String 'Si es un enumerado dinamico, de donde salen los valores posibles?

' *********** Para valores Mixtos.
Private m_items() As cItem ' Items que a su vez tiene el item
Private m_itemCount As Long

Private m_value As String ' Representacion en String que tiene el item

Public Function init(key As String, humano As String, desk As String, value, ByVal itemIndex As Integer, ByVal keyIndex As Integer) As Boolean
On Error GoTo errHandler
    m_Key = key
    m_desk = desk
    m_keyIndex = keyIndex
    m_itemIndex = itemIndex
    m_descKey = humano
    

    'En esta instancia parseo el valor para definir el tipo de item
    
    Select Case TypeName(value)
        Case "Dictionary"
            If Not parseSingleValue(value) Then GoTo errHandler
        
        Case "Collection"
            If Not parseMixedValue(value) Then GoTo errHandler

    End Select
    
    init = True
    Exit Function
    
errHandler:
    init = False
    MsgBox "Error en cItem.init()"
End Function

Private Function parseSingleValue(value) As Boolean
On Error GoTo errHandler
    m_itemCount = -1 'Con esto obtengo que la clase sepa que es el UNICO valor
            
    Dim keys, element, tmpStr As String
    
    keys = value.keys
    
    For Each element In keys
        Select Case element
            Case "tipo"
                tmpStr = value.item(element)
                Select Case UCase$(tmpStr)
                    Case "ENUMERADO": m_type = ItemType.e_Enumerado
                    Case "CADENA": m_type = ItemType.e_Cadena
                    Case "NUMERICO": m_type = ItemType.e_Numerico
                    Case "ENUMERADO_DINAMICO": m_type = ItemType.e_EnumeradoDinamico
                End Select
                
            Case "min"
                 tmpStr = value.item(element)
                 
                If IsNumeric(tmpStr) Then
                    m_minValue = val(tmpStr)
                Else
                    m_minValue = modEnumerandosDinamicos.obtenerValorConstante(tmpStr)
                End If
            Case "max"
                tmpStr = value.item(element)
            
                If IsNumeric(tmpStr) Then
                    m_maxValue = val(tmpStr)
                Else
                    m_maxValue = modEnumerandosDinamicos.obtenerValorConstante(tmpStr)
                End If
                
            Case "combinado"
                m_combinado = (UCase$(value.item(element)) = "TRUE")
            
            Case "valores"
                Set m_values = value.item(element)
                
            Case "fuente"
                m_fuente = value.item(element)
  
        End Select
        
        
    Next
    
    parseSingleValue = True
    Exit Function
    
errHandler:
    parseSingleValue = False
    MsgBox "Error en cItem.parseSingleValue()"
End Function

' Tipo de dato compuesto
Private Function parseMixedValue(value) As Boolean
On Error GoTo errHandler

    Dim dic As Dictionary, items, element
    
    'Por cada elemento que componga a este item
    
    ReDim Preserve m_items(1 To value.Count)
    
    For Each element In value
        m_itemCount = m_itemCount + 1
        
        If TypeName(element) <> "Dictionary" Then
            GoTo errHandler
        Else
            If element.Count <> 3 Then
                GoTo errHandler
            Else
                Set m_items(m_itemCount) = New cItem
                items = element.items
                
                m_items(m_itemCount).init m_Key, CStr(items(0)), CStr(items(1)), items(2), m_itemIndex, m_itemCount
                
            End If
        End If
    Next
    
    m_type = ItemType.e_MixedValue
    
    parseMixedValue = True
    Exit Function
    
errHandler:
    parseMixedValue = False
    MsgBox "Error en cItem.parseMixedValue()"
End Function

Public Function getItems() As cItem()
    getItems = m_items
End Function

Public Function getItemIndex() As Long
    getItemIndex = m_itemIndex
End Function

Public Function getType() As ItemType
    getType = m_type
End Function

Public Function getKey() As String
    getKey = m_Key
End Function

Public Function getDesk() As String
    getDesk = m_desk
End Function

Public Function getKeyIndex() As Integer
    getKeyIndex = m_keyIndex
End Function

Public Function getCombinado() As Boolean
    getCombinado = m_combinado
End Function

Public Function getValues() As Collection
    Set getValues = m_values
End Function

Public Function getFuente() As String
    getFuente = m_fuente
End Function

Public Function getMinValue() As Long
    getMinValue = m_minValue
End Function

Public Function getMaxValue() As Long
    getMaxValue = m_maxValue
End Function

' #######################
'#  cItem adaption  #
' #######################

Public Sub setValue(ByRef value As String)
    m_value = value

    If m_itemCount <> -1 Then
        'Mixed value
        Dim strArr() As String
        strArr = Split(value, "-")
        
        Dim i As Long
        For i = 1 To IIf(m_itemCount > UBound(strArr) + 1, UBound(strArr) + 1, m_itemCount)
            m_items(i).setValue strArr(i - 1)
        Next i
    End If
End Sub

Public Sub setValueByIndex(ByRef value As String, ByVal Index As Integer)
    If (Index = 0) Then
        setValue value
        Exit Sub
    End If
    
    If (Index > m_itemCount And Index < 1) Then Exit Sub
    
    m_items(Index).setValue value
    m_value = StringChangeField(Index, m_value, value, SEPARE_ITEM_CHAR)
    
End Sub

Public Function getHumanoKey() As String
    getHumanoKey = m_descKey
End Function
Public Function getValue() As String
    getValue = m_value
End Function

Public Function getValueByIndex(ByVal Index As Integer) As String
    If (m_itemCount = -1) Or (Index >= m_itemCount And Index < 1) Then Exit Function
    
    getValueByIndex = m_items(Index).getValue()
End Function

Public Function isValidValue(ByRef value As String, Optional keyIndex As Integer = 0) As Boolean
    If keyIndex = 0 Then
        If m_type = ItemType.e_Numerico Then
            isValidValue = IsOnlyNumber(value) And _
                    (m_maxValue >= val(value)) And _
                    (m_minValue <= val(value))
        ElseIf getType() = ItemType.e_Cadena Then
            isValidValue = (Len(value) <= m_maxValue) And _
                      (Len(value) >= m_minValue)
        Else
            isValidValue = True
        End If
    Else
        If m_items(keyIndex).isValidValue(value) Then
            isValidValue = True
        Else
            isValidValue = False
        End If
    End If
End Function

Public Sub copy(ByRef item As cItem)
    m_Key = item.getKey()
    m_keyIndex = item.getKeyIndex()
    m_itemIndex = item.getItemIndex()
    m_descKey = item.getHumanoKey
    
    m_desk = item.getDesk()
    m_type = item.getType()
        
    Set m_values = item.getValues
    m_minValue = item.getMinValue()
    m_maxValue = item.getMaxValue()
    
    m_combinado = item.getCombinado()
    m_fuente = item.getFuente()
    
    If m_type = ItemType.e_MixedValue Then
        Dim items() As cItem
        Dim i As Long
        
        items = item.getItems()
        
        m_itemCount = UBound(items)
        ReDim m_items(1 To m_itemCount)
        
        For i = 1 To m_itemCount
            Set m_items(i) = New cItem
            m_items(i).copy items(i)
        Next i
    Else
        m_itemCount = -1
    End If
End Sub

' ##################
'# Common functions #
' ##################

Private Function IsOnlyNumber(ByRef myStr As String) As Boolean
    Dim i As Long, c As Long
    For i = 1 To Len(myStr)
        c = Asc(mid$(myStr, i, 1))
        If c < 48 Or c > 57 Then
            IsOnlyNumber = False
            Exit Function
        End If
    Next i
    
    IsOnlyNumber = True
End Function

Private Function StringChangeField(ByVal Pos As Integer, ByRef Text As String, ByRef changeFieldText As String, ByVal SepASCII As Byte) As String
    Dim rightStr As String
    Dim leftStr As String
    Dim delimiter As String * 1
    
    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim ii As Long
    
    delimiter = Chr$(SepASCII)
    
Again:
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
        If CurrentPos = 0 Then Exit For
    Next i
    
    If i = Pos Then
        leftStr = Left$(Text, lastPos)
        StringChangeField = leftStr & changeFieldText
    ElseIf i < Pos + 1 Then
        For ii = i To Pos - 1
            Text = Text & delimiter
        Next ii
        
        GoTo Again
    Else
        leftStr = Left$(Text, lastPos)
        rightStr = Right$(Text, Len(Text) - CurrentPos + 1)
        
        StringChangeField = leftStr & changeFieldText & rightStr
    End If
End Function

