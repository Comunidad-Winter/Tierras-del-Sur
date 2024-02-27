Attribute VB_Name = "modFolderBrowse"
Option Explicit

Function Seleccionar_Carpeta(Optional Titulo As String, _
                        Optional Path_Inicial As Variant) As String

On Local Error GoTo errFunction
    
    Dim objShell As Object
    Dim objFolder As Object
    Dim o_Carpeta As Object
    
    ' Nuevo objeto Shell.Application
    Set objShell = CreateObject("Shell.Application")
    
    'Marce On error resume next
    'Abre el cuadro de diálogo para seleccionar
    Set objFolder = objShell.BrowseForFolder( _
                            0, _
                            Titulo, _
                            0, _
                            Path_Inicial)
    
    ' Devuelve solo el nombre de carpeta
    If objFolder Is Nothing Then
        Seleccionar_Carpeta = ""
        Exit Function
    End If
    
    Set o_Carpeta = objFolder.self
    
    ' Devuelve la ruta completa seleccionada en el diálogo
    Seleccionar_Carpeta = o_Carpeta.Path

    If InStr(1, Seleccionar_Carpeta, "\") = 0 Then
        Seleccionar_Carpeta = vbNullString
    End If
    
Exit Function
'Error
errFunction:
    MsgBox Err.Description, vbCritical
    Seleccionar_Carpeta = vbNullString

End Function
