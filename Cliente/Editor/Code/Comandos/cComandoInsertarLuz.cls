VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComandoInsertarLuz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Implements iComando
Private luz As tLuzPropiedades
Private luzBackup As tLuzPropiedades

Private posX As Byte
Private posY As Byte

Friend Sub crear(x As Integer, y As Integer, luz_ As tLuzPropiedades)
    luz = luz_
    posX = x
    posY = y
End Sub

Private Sub iComando_desHacer()

    If MapData(posX, posY).luz > 0 Then
        DLL_Luces.Quitar MapData(posX, posY).luz
        MapData(posX, posY).luz = 0
    End If
    
    If luzBackup.LuzRadio > 0 Then
        MapData(posX, posY).luz = _
        DLL_Luces.crear(posX, posY, luzBackup.LuzColor.r, luzBackup.LuzColor.g, luzBackup.LuzColor.b, luzBackup.LuzRadio, luzBackup.LuzBrillo, luzBackup.LuzTipo, luzBackup.luzInicio, luzBackup.luzFin)
    End If
    
End Sub

Private Sub iComando_hacer()
    'Saco la actual
    If MapData(posX, posY).luz Then
        DLL_Luces.Get_Light MapData(posX, posY).luz, posX, posY, luzBackup.LuzColor.r, luzBackup.LuzColor.g, luzBackup.LuzColor.b, luzBackup.LuzRadio, luzBackup.LuzBrillo, luzBackup.LuzTipo, luzBackup.luzInicio, luzBackup.luzFin
        DLL_Luces.Quitar MapData(posX, posY).luz
        MapData(posX, posY).luz = 0
    End If
   
    If luz.LuzRadio > 0 Then
        MapData(posX, posY).luz = _
        DLL_Luces.crear(posX, posY, luz.LuzColor.r, luz.LuzColor.g, luz.LuzColor.b, luz.LuzRadio, luz.LuzBrillo, luz.LuzTipo, luz.luzInicio, luz.luzFin)

        Debug.Print "Generada la luz"; MapData(posX, posY).luz, "Count"; DLL_Luces.Count
    End If
End Sub

Private Function iComando_obtenerNombre() As String
    If luz.LuzRadio = 0 Then
        iComando_obtenerNombre = "Borrar Luz" & " en (" & posX & "," & posY & ")"
    Else
        iComando_obtenerNombre = "Insertar Luz" & " en (" & posX & "," & posY & ")"
    End If
End Function
