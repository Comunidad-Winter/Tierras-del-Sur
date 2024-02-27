Attribute VB_Name = "HelperFiles"
'ARCHIVO COMPARTIDO POR TODOS LOS PROGRAMAS

Option Explicit

Public Function getFileSize(archivo As String) As Long

getFileSize = FileLen(archivo)

End Function
Public Function GetFileVersion(archivo As String) As String
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    On Error GoTo error:
    GetFileVersion = fso.GetFileVersion(archivo)
    
    If GetFileVersion = "" Then GetFileVersion = "0.0.0.0"

    Exit Function
error:
    GetFileVersion = "0.0.0.0"
End Function

' Ejemplo
' Absoluto 1: C:\SVN\
' Absoluto 2: C:\SVN\MapEditor\Editor.exe


Public Function obtenerPathRelativo(ByVal base As String, ByVal absoluto As String) As String
    
    base = Replace$(base, "/", "\")
    absoluto = Replace$(absoluto, "/", "\")
    
    If right$(base, 1) = "\" Then base = mid$(base, 1, Len(base) - 1)
    
    obtenerPathRelativo = Replace$(absoluto, base, ".")

End Function


Public Function PathGetParent(ByVal sFolder As String, Optional lParentIndex As Long = 1) As String
    Dim asFolders() As String
    Dim sPathSep As String
    Dim lThisFolder As Long
    
    If Len(sFolder) > 0 Then
        'Determine the path seperator
        If InStr(1, sFolder, "/") > 0 Then
            sPathSep = "/"
        Else
            sPathSep = "\"
        End If
        
        If right$(sFolder, 1) <> sPathSep Then
            sFolder = sFolder & sPathSep
        End If
        
        asFolders = Split(sFolder, sPathSep)
        'Get the requested parent folder
        For lThisFolder = 0 To UBound(asFolders) - lParentIndex - 1
            PathGetParent = PathGetParent & asFolders(lThisFolder) & sPathSep
        Next
    End If
End Function

Public Function ProccessPath(tmpPath As String) As String
    If Len(tmpPath) Then
        If left$(tmpPath, 2) = ".." Then _
            tmpPath = PathGetParent(app.Path) & right$(tmpPath, Len(tmpPath) - 3)
        If left$(tmpPath, 1) = "." Then _
            tmpPath = app.Path & right$(tmpPath, Len(tmpPath) - 1)
    Else
        tmpPath = app.Path
    End If
        
    If right$(tmpPath, 1) <> "\" Then tmpPath = tmpPath & "\"
    
    ProccessPath = tmpPath
End Function

Public Function GetPathIni(ByVal file As String, ByVal Main As String, ByVal key As String, Optional ByVal default As String = vbNullString) As String
'MZ
    Dim tmpPath As String
    
    If default = vbNullString Then default = app.Path
    
    tmpPath = GetVar(file, Main, key)
    

    GetPathIni = ProccessPath(tmpPath)
End Function

Public Function FileExist(ByVal file As String, Optional ByVal FileType As VbFileAttribute = vbNormal) As Boolean
'Marce On error resume next

    FileExist = (Dir$(file, FileType) <> "")
End Function

Public Function generarRandomNameFile(ByVal carpeta As String, longitud As String, extension As String) As String
    Dim Nombre As String
    Dim loopCaracter As Byte
    

    Do
        Nombre = ""
        
        For loopCaracter = 1 To longitud
            Nombre = Nombre & Chr$(97 + Round(Rnd() * 25))
        Next
       
    Loop While FileExist(carpeta & "\" & Nombre & "." & extension)
    

   generarRandomNameFile = carpeta & "\" & Nombre & "." & extension
End Function

Public Function FolderExist(ByVal folder As String) As Boolean
    FolderExist = folder <> "" And (Dir$(folder, vbDirectory) <> "")
End Function

' Obtiene el nombre de un archivo que se encuentra dentro del path absoluto
Public Function getNameFileInPath(Nombre As String) As String
    getNameFileInPath = right$(Nombre, Len(Nombre) - InStrRev(Nombre, "\"))
End Function

Public Function getPathFileInPath(Nombre As String) As String
    getPathFileInPath = mid$(Nombre, 1, InStrRev(Nombre, "\"))
End Function

Public Function LeerArchivo(archivo As String) As String

Dim handle As Integer
On Error GoTo hayError:

handle = 0
handle = FreeFile

Open archivo For Input As #handle
LeerArchivo = StrConv(InputB(LOF(handle), #handle), vbUnicode)
Close #handle

Exit Function
hayError:
If handle > 0 Then Close #handle
LeerArchivo = vbNullString
End Function
