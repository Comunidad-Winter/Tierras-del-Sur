VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComandoInsertarAntorcha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Implements iComando
Private luz As tLuzPropiedades
Private luzBackup As tLuzPropiedades

Private UserIndex As Integer

Friend Sub crear(UserIndex_ As Integer, luz_ As tLuzPropiedades)
    luz = luz_
    UserIndex = UserIndex_
End Sub

Private Sub iComando_desHacer()

    If CharList(UserIndex).luz > 0 Then
        DLL_Luces.Quitar CharList(UserIndex).luz
        CharList(UserIndex).luz = 0
    End If
    
    If luzBackup.LuzRadio > 0 Then
        CharList(UserIndex).luz = _
        DLL_Luces.crear(CharList(UserIndex).Pos.x, CharList(UserIndex).Pos.y, luzBackup.LuzColor.r, luzBackup.LuzColor.g, luzBackup.LuzColor.b, luzBackup.LuzRadio, luzBackup.LuzBrillo, luzBackup.LuzTipo, luzBackup.luzInicio, luzBackup.luzFin)
    End If
    
End Sub

Private Sub iComando_hacer()
    'Saco la actual
    Dim posx As Byte
    Dim posy As Byte
    
    If CharList(UserIndex).luz Then
        DLL_Luces.Get_Light CharList(UserIndex).luz, posx, posy, luzBackup.LuzColor.r, luzBackup.LuzColor.g, luzBackup.LuzColor.b, luzBackup.LuzRadio, luzBackup.LuzBrillo, luzBackup.LuzTipo, luzBackup.luzInicio, luzBackup.luzFin
        DLL_Luces.Quitar CharList(UserIndex).luz
         CharList(UserIndex).luz = 0
    End If
   
    If luz.LuzRadio > 0 Then
        CharList(UserIndex).luz = _
        DLL_Luces.crear(CharList(UserIndex).Pos.x, CharList(UserIndex).Pos.y, luz.LuzColor.r, luz.LuzColor.g, luz.LuzColor.b, luz.LuzRadio, luz.LuzBrillo, luz.LuzTipo, luz.luzInicio, luz.luzFin)

        Debug.Print "Generada la luz"; CharList(UserIndex).luz, "Count"; DLL_Luces.count
    End If
End Sub

Private Function iComando_obtenerNombre() As String
    If luz.LuzRadio = 0 Then
        iComando_obtenerNombre = "Poner Antorcha"
    Else
        iComando_obtenerNombre = "Sacar Antorcha"
    End If
End Function

