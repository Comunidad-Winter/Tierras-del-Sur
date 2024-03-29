VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 0

Private Type ChildNode
    key As String
    value As String
End Type

Private Type MainNode
    Name As String
    values() As ChildNode
    numValues As Integer
End Type

Private fileData() As MainNode

Private MainNodes As Long

''
'Default constructor. Does nothing.

Private Sub Class_Initialize()
'
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'
'
End Sub

''
'Destroy every array and deallocates al memory.
'

Private Sub Class_Terminate()
'
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'
'
    Dim i As Long
    
    'Clean up
    If MainNodes Then
        For i = 1 To MainNodes - 1
            Erase fileData(i).values
        Next i
        
        Erase fileData
    End If
    
    MainNodes = 0
End Sub

''
'Loads a INI file so it's values can be read. Must be called before being able to use GetValue.
'
'file Complete path of the INI file to be loaded.
'@see      GetValue

Public Sub Initialize(ByVal file As String)
'
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 27/07/2006
'Opens the requested file and loads it's data into memory
'
    Dim handle As Integer
    Dim text As String
    Dim Pos As Long
    
    'Prevent memory losses if we are attempting to reload a file....
    Call Class_Terminate
    
    'Get a free handle and start reading line by line until the end
    handle = FreeFile
    
    Open file For Input As handle

   'frmMain.Text2.Text = frmMain.Text2.Text & vbNewLine & file
    Do Until EOF(handle)
        Line Input #handle, text
        
        'Is it null??
        If Len(text) Then
            'If it starts with '['it is a main node or nothing (GetPrivateProfileStringA works this way), otherwise it's a value
            If left$(text, 1) = "[" Then
                'If it has an ending ']'it's a main node, otherwise it's nothing
                Pos = InStr(2, text, "]")
                If Pos Then
                'INSERT INTO `Items` (`ID` ,`Name` ,`Valor`) VALUES ('0', 'NAME', '0');

                    'Add a main node
                    ReDim Preserve fileData(MainNodes) As MainNode
                    
                    fileData(MainNodes).Name = UCase$(Trim$(mid$(text, 2, Pos - 2)))
                    'frmMain.Text2.Text = frmMain.Text2.Text & ";" & vbNewLine & "INSERT INTO `Items` SET `ID`='" & fileData(MainNodes).name & "'"
                    MainNodes = MainNodes + 1
                End If
            Else
                'So it's a value. Check if it has a '=', otherwise it's nothing
                Pos = InStr(2, text, "=")
                If Pos Then
                    'Is it under any main node??
                    If MainNodes Then
                        With fileData(MainNodes - 1)
                            'Add it to the main node's value
                            ReDim Preserve .values(.numValues) As ChildNode
                            
                            
                            .values(.numValues).value = right$(text, Len(text) - Pos)
                            .values(.numValues).key = UCase$(left$(text, Pos - 1))
                            
                            'If .values(.numValues).key = "VALOR" Or .values(.numValues).key = "NAME" Then
                            'frmMain.Text2.Text = frmMain.Text2.Text & ",`" & .values(.numValues).key & "`='" & .values(.numValues).Value & "'"
                            'End If
                            .numValues = .numValues + 1
                        End With
                    End If
                End If
            End If
        End If
    Loop
    
    Close handle
    
    Dim i As Long
    
    If MainNodes Then
        'Sort main nodes to allow binary search
        Call SortMainNodes(0, MainNodes - 1)
        
        'Sort values of each node to allow binary search
        For i = 0 To MainNodes - 1
            If fileData(i).numValues Then _
                Call SortChildNodes(fileData(i), 0, fileData(i).numValues - 1)
        Next i
    End If
    
End Sub

Public Sub Initialize_raw(raw As String)
    Dim text As String
    Dim Pos As Long
    Dim finalizar As Boolean
    Dim i As Long
    Call Class_Terminate
    

    Do Until finalizar
        text = GetNextLine(raw, , finalizar)
        If finalizar = False Then
            If Len(text) Then
                If left$(text, 1) = "[" Then
                    Pos = InStr(2, text, "]")
                    If Pos Then
                        ReDim Preserve fileData(MainNodes) As MainNode
                        fileData(MainNodes).Name = UCase$(Trim$(mid$(text, 2, Pos - 2)))
                        MainNodes = MainNodes + 1
                    End If
                Else
                    Pos = InStr(2, text, "=")
                    If Pos Then
                        If MainNodes Then
                            With fileData(MainNodes - 1)
                                ReDim Preserve .values(.numValues) As ChildNode
                                .values(.numValues).value = right$(text, Len(text) - Pos)
                                .values(.numValues).key = UCase$(left$(text, Pos - 1))
                                .numValues = .numValues + 1
                            End With
                        End If
                    End If
                End If
            End If
        Else
            Exit Do
        End If
    Loop
    

    If MainNodes Then
        'Sort main nodes to allow binary search
        Call SortMainNodes(0, MainNodes - 1)
        
        'Sort values of each node to allow binary search
        For i = 0 To MainNodes - 1
            If fileData(i).numValues Then _
                Call SortChildNodes(fileData(i), 0, fileData(i).numValues - 1)
        Next i
    End If
    
End Sub

Private Function GetNextLine(ByRef sText As String, Optional ByVal reset As Boolean = False, Optional ByRef final As Boolean = False) As String
Static lLineStart As Long
Dim lLineEnd As Long
Dim lLength As Long

If right$(sText, 2) <> vbCrLf Then
    sText = sText & vbCrLf
End If
If lLineStart = 0 Then lLineStart = 1
If reset = True Then lLineStart = 1
lLineStart = InStr(lLineStart, sText, vbCrLf)
lLineStart = lLineStart + 2

If lLineStart < Len(sText) Then
    lLineEnd = InStr(lLineStart, sText, vbCrLf)
    lLength = lLineEnd - lLineStart
    GetNextLine = mid$(sText, lLineStart, lLength)
Else
    GetNextLine = vbNullString
    lLineStart = 1
    final = True
End If
End Function


''
'Sorts all child nodes within the given MainNode alphabetically by their keys. Uses quicksort.
'
'Node The MainNode whose values are to be sorted.
'first The first index to consider when sorting.
'last The last index to be considered when sorting.

Private Sub SortChildNodes(ByRef Node As MainNode, ByVal First As Integer, ByVal Last As Integer)
'
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Sorts the list of values in a given MainNode using quicksort,
'this allows the use of Binary Search for faster searches
'
    Dim min As Long      'First item in the list
    Dim max As Long      'Last item in the list
    Dim COMP As String      'Item used to compare
    Dim temp As ChildNode
    Dim total As Long
    min = First
    max = Last
    
    With Node

        COMP = .values((min + max) \ 2).key
    
        Do While min <= max
            Do While .values(min).key < COMP And min < Last
                min = min + 1
            Loop
            Do While .values(max).key > COMP And max > First
                max = max - 1
            Loop
            If min <= max Then
                temp = .values(min)
                .values(min) = .values(max)
                .values(max) = temp
                min = min + 1
                max = max - 1
            End If
        Loop
    End With
    
    If First < max Then SortChildNodes Node, First, max
    If min < Last Then SortChildNodes Node, min, Last
End Sub

''
'Sorts all main nodes in the loaded INI file alphabetically by their names. Uses quicksort.
'
'first The first index to consider when sorting.
'last The last index to be considered when sorting.

Private Sub SortMainNodes(ByVal First As Integer, ByVal Last As Integer)
'
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Sorts the MainNodes list using quicksort,
'this allows the use of Binary Search for faster searches
'
    Dim min As Integer      'First item in the list
    Dim max As Integer      'Last item in the list
    Dim COMP As String      'Item used to compare
    Dim temp As MainNode
    
    min = First
    max = Last
    
    COMP = fileData((min + max) \ 2).Name
    
    Do While min <= max
        Do While fileData(min).Name < COMP And min < Last
            min = min + 1
        Loop
        Do While fileData(max).Name > COMP And max > First
            max = max - 1
        Loop
        If min <= max Then
            temp = fileData(min)
            fileData(min) = fileData(max)
            fileData(max) = temp
            min = min + 1
            max = max - 1
        End If
    Loop
    
    If First < max Then SortMainNodes First, max
    If min < Last Then SortMainNodes min, Last
End Sub

''
'Searches for a given key within a given main section and if it exists retrieves it's value, otherwise a null string
'
'Main The name of the main section in which we will be searching.
'key The key of the value we are looking for.
'@returns  The value asociated with the given key under the requeted main section of the INI file or a null string if it's not found.

Public Function GetValue(ByVal Main As String, ByVal key As String) As String
Attribute GetValue.VB_UserMemId = 0
'
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Returns a value if the key and main node exist, or a nullstring otherwise
'
    Dim i As Long
    Dim j As Long
    
    'Search for the main node
    i = FindMain(UCase$(Main))
    
    If i >= 0 Then
        'If valid, binary search among keys
        j = FindKey(fileData(i), UCase$(key))
        
        'If we found it we return it
        If j >= 0 Then GetValue = fileData(i).values(j).value
    End If
End Function

Public Function GetValueOrDefault(ByVal Main As String, ByVal key As String, ByVal default As String) As String
    Dim currentValue As String
    
    currentValue = GetValue(Main, key)
    
    If Not currentValue = "" Then
        GetValueOrDefault = currentValue
    Else
        GetValueOrDefault = default
    End If
End Function

Public Function GetValueOrDefaultInt(ByVal Main As String, ByVal key As String, ByVal default As Integer) As Integer
    Dim currentValue As String
    
    GetValueOrDefaultInt = CInt(GetValueOrDefault(Main, key, default))
End Function

''
'Searches for a given key within a given main node and returns the index in which it's stored or the negation of the index in which it should be if not found.
'
'Node The MainNode among whose value entries we will be searching.
'key The key of the value we are looking for.
'@returns  The index in which the value with the key we are looking for is stored or the negation of the index in which it should be if not found.

Private Function FindKey(ByRef Node As MainNode, ByVal key As String) As Long
'
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Returns the index of the value which key matches the requested one,
'or the negation of the position were it should be if not found
'
    Dim min As Long
    Dim max As Long
    Dim mid As Long
    
    min = 0
    max = Node.numValues - 1
    
    Do While min <= max
        mid = (min + max) \ 2
        
        If Node.values(mid).key < key Then
            min = mid + 1
        ElseIf Node.values(mid).key > key Then
            max = mid - 1
        Else
            'We found it
            FindKey = mid
            Exit Function
        End If
    Loop
    
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    FindKey = Not mid
End Function

''
'Searches for a main section with the given name within the loaded INI file and returns the index in which it's stored or the negation of the index in which it should be if not found.
'
'name The name of the MainNode we are looking for.
'@returns  The index in which the main section we are looking for is stored or the negation of the index in which it should be if not found.

Private Function FindMain(ByVal Name As String) As Long
'
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Returns the index of the MainNode which name matches the requested one,
'or the negation of the position were it should be if not found
'
    Dim min As Long
    Dim max As Long
    Dim mid As Long
    
    min = 0
    max = MainNodes - 1
    
    Do While min <= max
        mid = (min + max) \ 2
        
        If fileData(mid).Name < Name Then
            min = mid + 1
        ElseIf fileData(mid).Name > Name Then
            max = mid - 1
        Else
            'We found it
            FindMain = mid
            Exit Function
        End If
    Loop
    
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    FindMain = Not mid
End Function

''
'Checks wether a given key exists or not.
'
'name    The name of the element whose existance is being checked.
'@returns  True if the key exists, false otherwise.

Public Function KeyExists(ByVal Name As String) As Boolean
'
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 04/01/2008
'Returns true of the key exists, false otherwise.
'
    KeyExists = FindMain(Name) >= 0
End Function

Private Function AddMain(ByVal Name As String) As Boolean
    If KeyExists(Name) = False Then
        ReDim Preserve fileData(MainNodes) As MainNode
        fileData(MainNodes).Name = UCase$(Name)
        AddMain = True
        MainNodes = MainNodes + 1
    End If
End Function

Private Function AddKey(ByVal Main As String, ByVal Name As String, ByVal value As String) As Boolean
    Dim i As Long
    i = KeyExists(Main)
    If i Then
        With fileData(i)
            ReDim Preserve .values(.numValues) As ChildNode
            .values(.numValues).value = value
            .values(.numValues).key = UCase$(Name)
            .numValues = .numValues + 1
        End With
        AddKey = True
    End If
End Function

Public Function SetVal(ByVal Main As String, ByVal key As String, ByVal value As String)
    Dim i As Long
    Dim j As Long
    
    'Search for the main node
    i = FindMain(UCase$(Main))
    
    If i < 0 Then
        If AddMain(Main) Then
            Call SortMainNodes(0, MainNodes - 1)
        End If
        i = FindMain(UCase$(Main))
    End If
    
    If i >= 0 Then
        j = FindKey(fileData(i), UCase$(key))
        
        If j >= 0 Then
            fileData(i).values(j).value = value
        Else
            If AddKey(UCase$(Main), key, value) Then
            If fileData(i).numValues Then _
                    Call SortChildNodes(fileData(i), 0, fileData(i).numValues - 1)
            End If
        End If
    End If
End Function

Public Function GetRAW() As String
    Dim i As Long
    Dim j As Long
    
    Dim b As String
    
    If MainNodes Then
        For i = 0 To MainNodes - 1
            With fileData(i)
                b = vbCrLf & "[" & .Name & "]" & vbCrLf
                If .numValues Then
                    For j = 0 To .numValues - 1
                        b = .values(j).key & "=" & .values(j).value & vbCrLf
                    Next j
                End If
            End With
        Next i
    End If
    
    GetRAW = b
End Function


