Attribute VB_Name = "modFormLoader"

Option Explicit

Public Type TDAList
    picCont As Object    'Contenedor de mainframe
    mainFrame As Object  'Contenedor de todos los elementos interactuables
    
    textboxes As Object  'Tipo texto
    numerico As Object   'Tipo numerico
    lists As Object      'Enumerados combinados
    combos As Object     'Enumerandos no combinados
    labels As Object     'Titulos
    Frames As Object     'Tipo Mixto
    
    lblSaveStatus As Object
End Type


'Inicializa el TDA
Public Sub InitTDA(ByRef list As TDAList, ByRef picCont As Object, ByRef mainFrame As Object, ByRef textboxes As Object, _
                        ByRef lists As Object, ByRef combos As Object, ByRef labels As Object, ByRef Frames As Object, ByRef lblSaveStatus As Object, ByRef numerico As Object)
    With list
        Set .picCont = picCont
        Set .mainFrame = mainFrame
        Set .textboxes = textboxes
        Set .lists = lists
        Set .combos = combos
        Set .labels = labels
        Set .Frames = Frames
        Set .lblSaveStatus = lblSaveStatus
        Set .numerico = numerico
    End With
End Sub


Public Sub LoadTDAFromSingleItem(ByRef list As TDAList, ByRef frameIndex As Frame, ByRef item As cItem, ByRef x As Long, ByRef y As Long)
    Dim labelIndex As Long
    Dim textBoxIndex As Long
    Dim listBoxIndex As Long
    Dim controlWidth As Long
    Dim strCaption As String
    Dim iVar As Collection, itemList As String, i As Long, element
    Dim idElemento As Long
    Dim representacionElemento As String
            
    'Titulo
    strCaption = item.getHumanoKey
    
    'Cargamos el label
    labelIndex = LoadLabelInTDA(frameIndex, list, strCaption, x, y, item.getDesk())
    
    'Calculamos el Y nuevo
    y = y + yp(list.labels(labelIndex).Height)
    
    'Lo ponemos donde va
    list.labels(labelIndex).Left = x
    list.labels(labelIndex).tag = frameIndex ' Para la busqueda
    
    'Nos fijiamos que tipo de item es, y en base a eso creamos el elemento correspondiente
    Select Case item.getType()
        Case ItemType.e_Numerico
            textBoxIndex = LoadTextBoxNumericoInTDA(frameIndex, list, item.getKey(), x, y, item.getDesk())
            y = y + yp(list.numerico(textBoxIndex).Height) + 10
            
            list.numerico(textBoxIndex).tag = item.getItemIndex() & Chr(SEPARE_TAG_CHAR) & item.getKeyIndex()
            
            'Propiedades
            list.numerico(textBoxIndex).MinValue = item.getMinValue
            list.numerico(textBoxIndex).MaxValue = item.getMaxValue
    
            controlWidth = list.numerico(textBoxIndex).Width
            
        Case ItemType.e_Cadena
            textBoxIndex = LoadTextBoxInTDA(frameIndex, list, item.getKey(), x, y, item.getDesk())
            y = y + yp(list.textboxes(textBoxIndex).Height) + 10
            
            list.textboxes(textBoxIndex).tag = item.getItemIndex() & Chr(SEPARE_TAG_CHAR) & item.getKeyIndex()
            
            list.textboxes(textBoxIndex).MinLength = item.getMinValue
            list.textboxes(textBoxIndex).MaxLength = item.getMaxValue

            controlWidth = list.textboxes(textBoxIndex).Width
            
        Case ItemType.e_Enumerado
            If item.getCombinado() Then
                listBoxIndex = LoadListboxInTDA(frameIndex, list, x, y, item.getDesk())
                list.lists(listBoxIndex).tag = item.getItemIndex() & Chr(SEPARE_TAG_CHAR) & item.getKeyIndex()
                
                y = y + yp(list.lists(listBoxIndex).Height) + 10
            
                controlWidth = list.lists(listBoxIndex).Width
            Else
                listBoxIndex = LoadComboboxInTDA(frameIndex, list, x, y, item.getDesk())
                list.combos(listBoxIndex).tag = item.getItemIndex() & Chr(SEPARE_TAG_CHAR) & item.getKeyIndex()
                
                y = y + yp(list.combos(listBoxIndex).Height) + 10
            
                controlWidth = list.combos(listBoxIndex).Width
            End If
            
            
            For Each iVar In item.getValues()
                i = 0
                For Each element In iVar
                    i = i + 1
                    If i = 1 Then
                        idElemento = element
                    Else
                        representacionElemento = element
                    End If
                Next
                
                If item.getCombinado() Then
                    list.lists(listBoxIndex).AddItem idElemento & " # " & representacionElemento
                Else
                    list.combos(listBoxIndex).addString idElemento, idElemento & " # " & representacionElemento
                End If
            Next
            
        Case ItemType.e_EnumeradoDinamico
            If item.getCombinado() Then
                listBoxIndex = LoadListboxInTDA(frameIndex, list, x, y, item.getDesk())
                list.lists(listBoxIndex).tag = item.getItemIndex() & Chr(SEPARE_TAG_CHAR) & item.getKeyIndex()
                
                y = y + yp(list.lists(listBoxIndex).Height) + 10
            
                controlWidth = list.lists(listBoxIndex).Width
            Else
                listBoxIndex = LoadComboboxInTDA(frameIndex, list, x, y, item.getDesk())
                list.combos(listBoxIndex).tag = item.getItemIndex() & Chr(SEPARE_TAG_CHAR) & item.getKeyIndex()
                
                y = y + yp(list.combos(listBoxIndex).Height) + 10
            
                controlWidth = list.combos(listBoxIndex).Width
            End If
            
            Dim enumerados() As eEnumerado
            
            enumerados = obtenerEnumeradosDinamicos(item.getFuente())
            
            For i = LBound(enumerados) To UBound(enumerados)
                If item.getCombinado() Then
                    list.lists(listBoxIndex).AddItem enumerados(i).valor & " # " & enumerados(i).nombre
                Else
                    list.combos(listBoxIndex).addString enumerados(i).valor, enumerados(i).valor & " # " & enumerados(i).nombre
                End If
            Next i
            
    End Select
    
    If controlWidth > 0 Then
        list.labels(labelIndex).Width = controlWidth
    End If
    
    x = x + xp(list.labels(labelIndex).Width) + 10
End Sub

Public Sub LoadTDAFromMixedItem(ByRef list As TDAList, ByRef frameIndex As Frame, ByRef item As cItem, ByRef x As Long, ByRef y As Long)
    Dim iFrame As Long
    Dim items() As cItem
    Dim fX As Long, fY As Long 'posiciones en el frame
    Dim i As Long
    Dim maxXPos As Long
    
    iFrame = LoadFrameInTDA(frameIndex, list, item.getHumanoKey, x, y, item.getDesk())
    list.Frames(iFrame).Height = px(20)
    
    items = item.getItems
    fY = 15
    For i = 1 To UBound(items)
        fX = 10
        If items(i).getType = ItemType.e_MixedValue Then
            LoadTDAFromMixedItem list, list.Frames(iFrame), items(i), fX, fY
        Else
            LoadTDAFromSingleItem list, list.Frames(iFrame), items(i), fX, fY
        End If
        
        If maxXPos < fX Then
            maxXPos = fX
        End If
            
    Next i
    
    y = y + fY + 10
    list.Frames(iFrame).Width = px(maxXPos)
    list.Frames(iFrame).Height = py(fY)
End Sub

Public Sub LoadTDAFromItems(ByRef list As TDAList, ByRef items() As cItem)
    Dim i As Long, x As Long, y As Long
    
    y = 1
    For i = 1 To UBound(items)
        x = 1
        
        If items(i).getType = ItemType.e_MixedValue Then
            LoadTDAFromMixedItem list, list.mainFrame, items(i), x, y
        Else
            LoadTDAFromSingleItem list, list.mainFrame, items(i), x, y
        End If
    Next i
    
    If list.mainFrame.Height < px(y + yp(list.picCont.Height)) Then
        list.mainFrame.Height = px(y) + list.picCont.Height
    End If
End Sub

Function LoadLabelInTDA(ByRef parent As Frame, ByRef list As TDAList, ByRef caption As String, ByVal x As Long, ByVal y As Long, ByRef tag As String) As Long
    Dim i As Long
    i = list.labels.UBound + 1
    load list.labels(i)
    
    With list.labels(i)
        .top = px(y)
        .Left = py(x)
        
        .AutoSize = True
          
        .caption = caption
        .ToolTipText = tag
        
        .visible = True
        
        Set .Container = parent
    End With
    
    LoadLabelInTDA = i
End Function

Function LoadTextBoxInTDA(ByRef parent As Frame, ByRef list As TDAList, ByRef caption As String, ByVal x As Long, ByVal y As Long, ByRef tag As String) As Long
    Dim i As Long
    i = list.textboxes.UBound + 1
    load list.textboxes(i)
    
    With list.textboxes(i)
        Set .Container = parent
        
        .top = px(y)
        .Left = py(x)
        
        .ToolTipText = tag
               
        .visible = True
        
        .ZOrder 1
    End With
    
    LoadTextBoxInTDA = i
End Function

Function LoadTextBoxNumericoInTDA(ByRef parent As Frame, ByRef list As TDAList, ByRef caption As String, ByVal x As Long, ByVal y As Long, ByRef tag As String) As Long
    Dim i As Long
    i = list.numerico.UBound + 1
    load list.numerico(i)
    
    With list.numerico(i)
        Set .Container = parent
        
        .top = px(y)
        .Left = py(x)
        
        .ToolTipText = tag
        
        .visible = True
        
        .ZOrder 1
    End With
    
    LoadTextBoxNumericoInTDA = i
End Function

Function LoadListboxInTDA(ByRef parent As Frame, ByRef list As TDAList, ByVal x As Long, ByVal y As Long, ByRef tag As String) As Long
    Dim i As Long
    i = list.lists.UBound + 1
    load list.lists(i)
    
    With list.lists(i)
        Set .Container = parent
        
        .top = px(y)
        .Left = py(x)
        
        .ToolTipText = tag
        
        .visible = True
        
        .ZOrder 1
    End With
    
    LoadListboxInTDA = i
End Function

Function LoadComboboxInTDA(ByRef parent As Frame, ByRef list As TDAList, ByVal x As Long, ByVal y As Long, ByRef tag As String) As Long
    Dim i As Long
    i = list.combos.UBound + 1
    load list.combos(i)
    
    With list.combos(i)
        Set .Container = parent
        
        .top = px(y)
        .Left = py(x)
        
        .ToolTipText = tag
        
        .visible = True
        .ZOrder 1
    End With
    
    LoadComboboxInTDA = i
End Function

Function LoadFrameInTDA(ByRef parent As Frame, ByRef list As TDAList, ByRef caption As String, ByVal x As Long, ByVal y As Long, ByRef tag As String) As Long
    Dim i As Long
    i = list.Frames.UBound + 1
    load list.Frames(i)
    
    With list.Frames(i)
        Set .Container = parent
        
        .top = px(y)
        .Left = py(x)
        
        .caption = caption
        
        .ToolTipText = tag
        
        .visible = True
        
        .ZOrder 1
    End With
    
    LoadFrameInTDA = i
End Function

Function px(ByVal v As Long) As Long
    px = v * Screen.TwipsPerPixelX
End Function

Function py(ByVal v As Long) As Long
    py = v * Screen.TwipsPerPixelY
End Function
Function xp(ByVal v As Long) As Long
    xp = v / Screen.TwipsPerPixelX
End Function

Function yp(ByVal v As Long) As Long
    yp = v / Screen.TwipsPerPixelY
End Function

'################ Muestra de valores ######################

Private Function GetControlInTDA(ByRef list As TDAList, ByRef item As cItem) As Integer
   Dim control
    Dim searchStr As String
    searchStr = item.getItemIndex() & Chr(SEPARE_TAG_CHAR) & item.getKeyIndex()
    
    Select Case item.getType()
        Case ItemType.e_Cadena
            For Each control In list.textboxes
                If control.tag = searchStr Then
                    GetControlInTDA = control.Index
                    Exit Function
                End If
            Next
        
        Case ItemType.e_Numerico
        
            For Each control In list.numerico
                If control.tag = searchStr Then
                    GetControlInTDA = control.Index
                    Exit Function
                End If
            Next
        
        Case ItemType.e_Enumerado, ItemType.e_EnumeradoDinamico
            If item.getCombinado() Then
                For Each control In list.lists
                    If control.tag = searchStr Then
                        GetControlInTDA = control.Index
                        Exit Function
                    End If
                Next
            Else
                For Each control In list.combos
                    If control.tag = searchStr Then
                        GetControlInTDA = control.Index
                        Exit Function
                    End If
                Next
            End If
            
        Case ItemType.e_MixedValue
            GetControlInTDA = -1 'No deberìa pasar, ya que controlo antes
    End Select
End Function
Private Sub SetSingleItemInTDA(ByRef list As TDAList, ByRef item As cItem)
    Dim controlIndex As Long, i As Long, ii As Long
    Dim vArr() As String, vSearch As String, v As String
    
    controlIndex = GetControlInTDA(list, item)
    v = item.getValue()
    
    Select Case item.getType()
        Case ItemType.e_Cadena
            list.textboxes(controlIndex).Text = v
        
        Case ItemType.e_Numerico
            list.numerico(controlIndex).value = val(v)

        Case ItemType.e_Enumerado, ItemType.e_EnumeradoDinamico
        
            If item.getCombinado Then
                vArr = Split(v, Chr(SEPARE_VALUE_CHAR))
                For ii = LBound(vArr) To UBound(vArr)
                    v = vArr(ii)
                    For i = 0 To list.lists(controlIndex).ListCount - 1
                        vSearch = Left$(list.lists(controlIndex).list(i), InStr(list.lists(controlIndex).list(i), "#") - 2)
                        If vSearch = CStr(v) Then
                            list.lists(controlIndex).Selected(i) = True
                            Exit For
                        End If
                    Next i
                Next ii
            Else
                If Not item.getValue = "" Then
                    Call list.combos(controlIndex).seleccionarID(item.getValue)
                End If
            End If
    End Select
End Sub

Private Sub SetMixedItemInTDA(ByRef list As TDAList, ByRef item As cItem)
    Dim items() As cItem
    Dim i, singleItem As cItem
    
    items = item.getItems()
    For Each i In items
        Set singleItem = i
        
        SetSingleItemInTDA list, singleItem
    Next
End Sub

Public Sub SetSectionInTDA(ByRef list As TDAList, ByRef section As cSection)
    Dim items() As cItem, i, item As cItem
    items = section.getItems()
    
    For Each i In items
        Set item = i

        If item.getType <> ItemType.e_MixedValue Then
            SetSingleItemInTDA list, item
        Else
            SetMixedItemInTDA list, item
        End If
    Next
End Sub

'Resetea todos los campos
Public Sub SetNullValuesInTDA(ByRef list As TDAList)
    Dim control, i As Long
    
    For Each control In list.textboxes
        control.Text = ""
    Next
    
    For Each control In list.combos
        Call control.desseleccionar
    Next
    
    For Each control In list.numerico
        control.value = 0
    Next
    
    For Each control In list.lists
        For i = 0 To control.ListCount - 1
            control.Selected(i) = False
        Next i
    Next
End Sub

Public Function SetSectionFromTDA(ByRef list As TDAList, ByRef section As cSection) As Boolean
On Error GoTo err_setSection:
    Dim control, i As Long
    Dim itemValue As cItem
    Dim itemIndex As Integer
    Dim keyIndex As Integer
    Dim newValue As String, arrValue As String
    
    For Each control In list.textboxes
        If control.Index > 0 Then
            itemIndex = val(ReadField(1, control.tag, SEPARE_TAG_CHAR))
            keyIndex = val(ReadField(2, control.tag, SEPARE_TAG_CHAR))

            If itemIndex > 0 Then
                If Not section.getItem(itemIndex).isValidValue(control.Text, keyIndex) Then
                    If Not list.lblSaveStatus Is Nothing Then
                        list.lblSaveStatus.caption = section.getItem(itemIndex).getKey() & " : "
                        If section.getItem(itemIndex).getType = ItemType.e_Cadena Then
                            list.lblSaveStatus.caption = list.lblSaveStatus.caption & "La cadena debe tener entre " & _
                                    section.getItem(itemIndex).getMinValue() & " y " & _
                                    section.getItem(itemIndex).getMaxValue() & " caracteres."
                        Else
                            list.lblSaveStatus.caption = list.lblSaveStatus.caption & "El valor debe estar comprendido " & _
                                    section.getItem(itemIndex).getMinValue() & " y " & _
                                    section.getItem(itemIndex).getMaxValue() & "."
                        End If
                    End If
                    
                    Call control.SetFocus
                    
                    SetSectionFromTDA = False
                    Exit Function
                End If
            End If
        End If
    Next
    
    'Ahora si actualizamos
    For Each control In list.textboxes
        itemIndex = val(ReadField(1, control.tag, SEPARE_TAG_CHAR))
        keyIndex = val(ReadField(2, control.tag, SEPARE_TAG_CHAR))
                
        If itemIndex > 0 Then
            section.getItem(itemIndex).setValueByIndex control.Text, keyIndex
        End If
    Next
    
    For Each control In list.numerico
        itemIndex = val(ReadField(1, control.tag, SEPARE_TAG_CHAR))
        keyIndex = val(ReadField(2, control.tag, SEPARE_TAG_CHAR))
        If itemIndex > 0 Then
            section.getItem(itemIndex).setValueByIndex control.value, keyIndex
        End If
    Next
    
    For Each control In list.combos
        itemIndex = val(ReadField(1, control.tag, SEPARE_TAG_CHAR))
        keyIndex = val(ReadField(2, control.tag, SEPARE_TAG_CHAR))
        
        If itemIndex > 0 Then
            newValue = control.obtenerIDValor
            section.getItem(itemIndex).setValueByIndex newValue, keyIndex
        End If
    Next
    
    For Each control In list.lists
        itemIndex = val(ReadField(1, control.tag, SEPARE_TAG_CHAR))
        keyIndex = val(ReadField(2, control.tag, SEPARE_TAG_CHAR))
        
        If itemIndex > 0 Then
            newValue = ""
            
            For i = 0 To control.ListIndex
                If control.Selected(i) Then
                    newValue = newValue & RTrim$(LTrim$(ReadField(1, control.list(i), 35))) & Chr(SEPARE_VALUE_CHAR)
                End If
            Next i
            
            If Len(newValue) > 0 Then newValue = Left$(newValue, Len(newValue) - 1)
            
            section.getItem(itemIndex).setValueByIndex newValue, keyIndex
        End If
    Next
    
    If Not list.lblSaveStatus Is Nothing Then
        list.lblSaveStatus.caption = "Seccion guardada."
    End If
    
    SetSectionFromTDA = True
    Exit Function

err_setSection:
    Debug.Print "Error " & Err.Description & "en SetSectionFromTDA"
End Function
