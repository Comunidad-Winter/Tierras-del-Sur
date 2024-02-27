VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIniManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

''
'Structure that contains all info under a tag in a INI file.
'Such tags are indicated with the "[" and "]" characters.
'
' @param    name String containing the text within the "[" and "]" characters.
'It's the key used when searching for a main section of the INI data.
' @param    values Array of ChildNodes, each containing a value entry along with it's key.
' @param    numValues Number of entrys in the main node.

Private Type MainNode
    Name As String
    values() As ChildNode
    numValues As Integer
End Type

''
'Containts all Main sections of the loaded INI file
Private fileData() As MainNode

''
'Stores the total number of main sections in the loaded INI file
Private MainNodes As Long

''
'Default constructor. Does nothing.

Private Sub Class_Initialize()
'**************************************************************
'Author: Juan Mart?n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'
'**************************************************************
End Sub

''
'Destroy every array and deallocates al memory.
'

Private Sub Class_Terminate()
'**************************************************************
'Author: Juan Mart?n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'
'**************************************************************
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
' @param    file Complete path of the INI file to be loaded.
' @see      GetValue

Public Sub Initialize(ByVal file As String)
'**************************************************************
'Author: Juan Mart?n Sotuyo Dodero
'Last Modify Date: 27/07/2006
'Opens the requested file and loads it's data into memory
'**************************************************************
    Dim handle As Integer
    Dim text As String
    Dim Pos As Long
    
    'Prevent memory losses if we are attempting to reload a file....
    Call Class_Terminate
    
    'Get a free handle and start reading line by line until the end
    handle = FreeFile
    
    Open file For Input As handle
    
    Do Until EOF(handle)
        Line Input #handle, text
        
        'Is it null??
        If Len(text) Then
            'If it starts with '[' it is a main node or nothing (GetPrivateProfileStringA works this way), otherwise it's a value
            If left$(text, 1) = "[" Then
                'If it has an ending ']' it's a main node, otherwise it's nothing
                Pos = InStr(2, text, "]")
                If Pos Then
                    'Add a main node
                    ReDim Preserve fileData(MainNodes) As MainNode
                    
                    fileData(MainNodes).Name = UCase$(Trim$(mid$(text, 2, Pos - 2)))
                    
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

''
'Sorts all child nodes within the given MainNode alphabetically by their keys. Uses quicksort.
'
' @param    Node The MainNode whose values are to be sorted.
' @param    first The first index to consider when sorting.
' @param    last The last index to be considered when sorting.

Private Sub SortChildNodes(ByRef node As MainNode, ByVal First As Integer, ByVal Last As Integer)
'**************************************************************
'Author: Juan Mart?n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Sorts the list of values in a given MainNode using quicksort,
'this allows the use of Binary Search for faster searches
'**************************************************************
    Dim min As Integer      'First item in the list
    Dim max As Integer      'Last item in the list
    Dim COMP As String      'Item used to compare
    Dim Temp As ChildNode
    
    min = First
    max = Last
    
    With node
        COMP = .values((min + max) \ 2).key
    
        Do While min <= max
            Do While .values(min).key < COMP And min < Last
                min = min + 1
            Loop
            Do While .values(max).key > COMP And max > First
                max = max - 1
            Loop
            If min <= max Then
                Temp = .values(min)
                .values(min) = .values(max)
                .values(max) = Temp
                min = min + 1
                max = max - 1
            End If
        Loop
    End With
    
    If First < max Then SortChildNodes node, First, max
    If min < Last Then SortChildNodes node, min, Last
End Sub

''
'Sorts all main nodes in the loaded INI file alphabetically by their names. Uses quicksort.
'
' @param    first The first index to consider when sorting.
' @param    last The last index to be considered when sorting.

Private Sub SortMainNodes(ByVal First As Integer, ByVal Last As Integer)
'**************************************************************
'Author: Juan Mart?n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Sorts the MainNodes list using quicksort,
'this allows the use of Binary Search for faster searches
'**************************************************************
    Dim min As Long      'First item in the list
    Dim max As Long      'Last item in the list
    Dim COMP As String      'Item used to compare
    Dim Temp As MainNode
    
    min = First
    max = Last
        
    COMP = fileData((min + max) \ 2).Name
    
    Do While min <= max
        Do While val(fileData(min).Name) < val(COMP) And min < Last
            min = min + 1
        Loop
        Do While val(fileData(max).Name) > val(COMP) And max > First
            max = max - 1
        Loop
        If min <= max Then
            Temp = fileData(min)
            fileData(min) = fileData(max)
            fileData(max) = Temp
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
' @param    Main The name of the main section in which we will be searching.
' @param    key The key of the value we are looking for.
' @returns  The value asociated with the given key under the requeted main section of the INI file or a null string if it's not found.

Public Function getValue(ByVal Main As String, ByVal key As String) As String
'**************************************************************
'Author: Juan Mart?n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Returns a value if the key and main node exist, or a nullstring otherwise
'**************************************************************
    Dim i As Long
    Dim j As Long
    
    'Search for the main node
    i = FindMain(UCase$(Main))
    
    If i >= 0 Then
        'If valid, binary search among keys
        j = FindKey(fileData(i), UCase$(key))
        
        'If we found it we return it
        If j >= 0 Then getValue = fileData(i).values(j).value
    End If
End Function

''
'Searches for a given key within a given main section and if it exists retrieves it's value, otherwise a null string
'
' @param    Main The name of the main section in which we will be searching.
' @param    key The key of the value we are looking for.
' @returns  The value asociated with the given key under the requeted main section of the INI file or a null string if it's not found.

Public Sub ChangeValue(ByVal Main As String, ByVal key As String, ByVal value As String)
'**************************************************************
'Author: ZaMa
'Last Modify Date: 27/05/2009
'If the key and main node exist, changes the value
'**************************************************************
    Dim i As Long
    Dim j As Long
    
    'Search for the main node
    i = FindMain(UCase$(Main))
    
    If i < 0 Then i = AddNode(Main)
    
    'If valid, binary search among keys
    j = FindKey(fileData(i), UCase$(key))
    
    If j < 0 Then
        Call AddKey(Main, key, value)
    Else
        'If we found it we change it
        fileData(i).values(j).value = value
    End If
End Sub

''
'Searches for a given key within a given main node and returns the index in which it's stored or the negation of the index in which it should be if not found.
'
' @param    Node The MainNode among whose value entries we will be searching.
' @param    key The key of the value we are looking for.
' @returns  The index in which the value with the key we are looking for is stored or the negation of the index in which it should be if not found.

Private Function FindKey(ByRef node As MainNode, ByVal key As String) As Long
'**************************************************************
'Author: Juan Mart?n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Returns the index of the value which key matches the requested one,
'or the negation of the position were it should be if not found
'**************************************************************
    Dim min As Long
    Dim max As Long
    Dim mid As Long
    
    min = 0
    max = node.numValues - 1
    
    Do While min <= max
        mid = (min + max) \ 2
        
        If node.values(mid).key < key Then
            min = mid + 1
        ElseIf node.values(mid).key > key Then
            max = mid - 1
        Else
            'We found it
            FindKey = mid
            Exit Function
        End If
    Loop
    
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    FindKey = Not min
End Function

''
'Searches for a main section with the given name within the loaded INI file and returns the index in which it's stored or the negation of the index in which it should be if not found.
'
' @param    name The name of the MainNode we are looking for.
' @returns  The index in which the main section we are looking for is stored or the negation of the index in which it should be if not found.

Private Function FindMain(ByVal Name As String) As Long
'**************************************************************
'Author: Juan Mart?n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Returns the index of the MainNode which name matches the requested one,
'or the negation of the position were it should be if not found
'**************************************************************
    Dim min As Long
    Dim max As Long
    Dim mid As Long
    
    min = 0
    max = MainNodes - 1
    
    Do While min <= max
        mid = (min + max) \ 2
        
        If val(fileData(mid).Name) < val(Name) Then
            min = mid + 1
        ElseIf val(fileData(mid).Name) > val(Name) Then
            max = mid - 1
        Else
            'We found it
            FindMain = mid
            Exit Function
        End If
    Loop
    
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    FindMain = Not min
End Function

''
'Checks wether a given key exists or not.
'
' @param    name    The name of the element whose existance is being checked.
' @returns  True if the key exists, false otherwise.

Public Function KeyExists(ByVal Name As String) As Boolean
'**************************************************************
'Author: Juan Mart?n Sotuyo Dodero
'Last Modify Date: 04/01/2008
'Returns true of the key exists, false otherwise.
'**************************************************************
    KeyExists = FindMain(UCase$(Name)) >= 0
End Function

Private Function AddNode(ByRef Name As String) As Long
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 10/10/2010
'
'**************************************************************
Dim i As Long
Dim MainPos As Long

Name = UCase$(Name)

'Add a main node
ReDim Preserve fileData(MainNodes) As MainNode

If MainNodes Then
    i = MainNodes - 1
    MainPos = Not FindMain(Name)
    
    Do While i >= MainPos
        fileData(i + 1) = fileData(i)
        i = i - 1
    Loop
    
    fileData(MainPos).numValues = 0
    Erase fileData(MainPos).values
End If

fileData(MainPos).Name = Name

MainNodes = MainNodes + 1

AddNode = MainPos
End Function

Private Function AddKey(ByRef Main As String, ByRef key As String, Optional ByRef value As String = vbNullString) As Long
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 10/10/2010
'
'**************************************************************
Dim MainPos As Long
Dim KeyPos As Long
Dim i As Long

Main = UCase$(Main)
key = UCase$(key)

MainPos = FindMain(Main)

If MainPos < 0 Then MainPos = AddNode(Main)

With fileData(MainPos)
    'Add it to the main node's value
    ReDim Preserve .values(.numValues) As ChildNode
    
    If .numValues > 0 Then
        i = .numValues - 1
        KeyPos = Not FindKey(fileData(MainPos), key)

        Do While i >= KeyPos
            .values(i + 1) = .values(i)
            i = i - 1
        Loop
    End If
    
    .values(KeyPos).key = key
    .values(KeyPos).value = value
    
    .numValues = .numValues + 1
End With

AddKey = KeyPos
End Function

Public Sub DumpFile(ByRef file As String)
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 10/10/2010
'
'**************************************************************
Dim hFile As Integer
Dim i As Long
Dim j As Long

hFile = FreeFile()

Open file For Output As hFile

    For i = 0 To MainNodes - 1
        With fileData(i)
            Print #hFile, "[" & .Name & "]"
            
            For j = 0 To .numValues - 1
                Print #hFile, .values(j).key & "=" & .values(j).value
            Next j
            
            Print #hFile, ""
        End With
    Next i
    
Close hFile
End Sub

Public Sub EraseMain(ByVal Main As String)
    Dim i As Long
    
    i = FindMain(Main)
    
    If i >= 0 Then
        If i <> MainNodes - 1 Then
            fileData(i) = fileData(MainNodes - 1)
        End If
        
        MainNodes = MainNodes - 1
        
        If i > 0 Or MainNodes > 1 Then
            ReDim Preserve fileData(MainNodes) As MainNode
        Else
            ReDim fileData(MainNodes) As MainNode
        End If
        
        SortMainNodes 0, MainNodes - 1
    End If
End Sub

'Obtiene una lista con los nombres de todas las secciones
Public Function GetMainKeys() As String()
    Dim retString() As String
    Dim i As Long
    
    If MainNodes > 0 Then
            ReDim retString(MainNodes - 1)
            For i = 0 To MainNodes - 1
                retString(i) = fileData(i).Name
            Next i
    End If
    GetMainKeys = retString
End Function

'Obtiene la cantidad de elementos
Public Function getCount() As Long
    getCount = MainNodes
End Function

' Obtiene el nombre de la ultima seccion
Public Function getNameLastSection() As String
    getNameLastSection = fileData(MainNodes - 1).Name
End Function

' Guarda una sección en un archivo determinado
Public Function seccionAArchivo(id As Long, archivo As String)

Dim numeroSeccion As Long
Dim j As Integer

numeroSeccion = FindMain(CStr(id))

'¿Existe?
If (numeroSeccion >= 0) Then
    ' Lo guardamos
    With fileData(numeroSeccion)
        For j = 0 To .numValues - 1
            Call WriteVar(archivo, .Name, .values(j).key, .values(j).value)
        Next j
    End With
Else
    Call WriteVar(archivo, id, "---ELIMINADO---", "----FORCE-DELETE----")
End If
End Function

Friend Function getKeys(id As Integer) As ChildNode()

    Dim numeroSeccion As Integer
    
    numeroSeccion = FindMain(id)
    
    If Not (numeroSeccion = Not 0) Then
        getKeys = fileData(numeroSeccion).values
    End If

End Function
' Guarda una sección en un archivo determinado
Public Function copiar(ini As cIniManager)

    Dim mainKeys() As String
    Dim loopMain As Integer
    Dim valores() As ChildNode
    Dim loopKey As Integer
    
    mainKeys = ini.GetMainKeys
    
    For loopMain = LBound(mainKeys) To UBound(mainKeys)
        
        ' Si existe, lo elimino
        If KeyExists(mainKeys(loopMain)) Then
            Call EraseMain(mainKeys(loopMain))
        End If
        
        ' Obtengo los nuevos valores
        valores = ini.getKeys(CInt(mainKeys(loopMain)))
        
        ' ¿Tengo un solo valor? ¿Y es el indicador de forzar el eliminar?
        If UBound(valores) = 0 Then
            If valores(0).key = "---ELIMINADO---" And valores(0).value = "----FORCE-DELETE----" Then
                ' Ya elimine la sección antes, no hago nada mas
                GoTo continue
            End If
        End If

        Call AddNode(mainKeys(loopMain))
    
        For loopKey = LBound(valores) To UBound(valores)
            Call AddKey(mainKeys(loopMain), valores(loopKey).key, valores(loopKey).value)
        Next loopKey

continue:
    Next loopMain
End Function


