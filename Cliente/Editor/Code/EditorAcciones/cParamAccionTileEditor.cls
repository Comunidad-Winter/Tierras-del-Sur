VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cParamAccionTileEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eTIPO_PARAMETRO
    ALFANUMERICO
    NUMERICO
End Enum

Private nombre As String
Private valor As String
Private textoAyuda As String
Private tipo As eTIPO_PARAMETRO
Private maximoValorAtributo As Long
Private minimoValorAtributo As Long

Private Sub Class_Initialize()
    nombre = "INDEFINIDO"
    textoAyuda = ""
    valor = ""
    tipo = eTIPO_PARAMETRO.ALFANUMERICO
    maximoValorAtributo = 289
    minimoValorAtributo = 0
End Sub

Public Sub crear(nombreParametro As String, valorInicial As String, tipo_ As eTIPO_PARAMETRO, minimoValorAtributo_ As Long, maximoValorAtributo_ As Long)
    nombre = nombreParametro
    valor = valorInicial
    tipo = tipo_
    minimoValorAtributo = minimoValorAtributo_
    maximoValorAtributo = maximoValorAtributo_
End Sub

Public Function setAyuda(texto As String) As String
    textoAyuda = texto
End Function
Public Function GetNombre() As String
    GetNombre = nombre
End Function

Public Function getValor() As String
    getValor = valor
End Function
Public Function getAyuda() As String
    If Len(textoAyuda) = 0 Then
        getAyuda = getAyudaDefault()
    Else
        getAyuda = textoAyuda
    End If
End Function

Private Function getAyudaDefault() As String
    If tipo = NUMERICO Then
        getAyudaDefault = "Número entre " & minimoValorAtributo & " y " & maximoValorAtributo & "."
    ElseIf tipo = ALFANUMERICO Then
        getAyudaDefault = "Texto entre " & minimoValorAtributo & " y " & maximoValorAtributo & " caracteres."
    End If
End Function
Public Function getTIPO() As eTIPO_PARAMETRO
    getTIPO = tipo
End Function

Public Function setValor(valor_ As String) As Boolean
    valor_ = Trim$(valor_) 'Le saco los espacios adelante y atras
    setValor = esValido(valor_)
    If setValor = True Then valor = valor_
End Function

Public Function Clonar() As cParamAccionTileEditor
    Set Clonar = New cParamAccionTileEditor
    Call Clonar.crear(nombre, valor, tipo, minimoValorAtributo, maximoValorAtributo)
    Call Clonar.setAyuda(textoAyuda)
End Function

Public Function esValido(valor_ As String) As Boolean
    If tipo = eTIPO_PARAMETRO.NUMERICO Then
        esValido = validarNumero(valor_)
    Else
        esValido = validarAlfaNumerico(valor_)
    End If
End Function

Private Function validarNumero(valor_ As String) As Boolean
    If IsNumeric(valor_) Then
        validarNumero = (valor_ >= minimoValorAtributo And valor_ <= maximoValorAtributo)
    End If
End Function

Private Function validarAlfaNumerico(valor_ As String) As Boolean
    validarAlfaNumerico = (Len(valor_) >= minimoValorAtributo And Len(valor_) <= minimoValorAtributo)
End Function


'/*************************************************************/
'/*********** FUNCIONES DE PERSISTENCIA ***********************/

Public Function persistir(archivoDestino As Integer)
        '1) Nombre
        Put archivoDestino, , CInt(Len(nombre))
        Put archivoDestino, , nombre
        '2) Texto Ayuda
        Put archivoDestino, , CInt(Len(textoAyuda))
        Put archivoDestino, , textoAyuda
        '3) Valor
        Put archivoDestino, , CInt(Len(valor))
        Put archivoDestino, , valor
        '4) Tipo
        Put archivoDestino, , CByte(tipo)
        '5) Minimo Valor Atributo
        Put archivoDestino, , minimoValorAtributo
        '6) Maximo Valor Atributo
        Put archivoDestino, , maximoValorAtributo
End Function

Public Function Cargar(archivoFuente As Integer)
Dim longitud As Integer
Dim byteAuxiliar As Byte
        
        '1) Nombre
        Get archivoFuente, , longitud
        nombre = Space$(longitud)
        Get archivoFuente, , nombre
        '2) Texto ayuda
        Get archivoFuente, , longitud
        textoAyuda = Space$(longitud)
        Get archivoFuente, , textoAyuda
        '3) Valor
        Get archivoFuente, , longitud
        valor = Space$(longitud)
        Get archivoFuente, , valor
        '4) Tipo
        Get archivoFuente, , byteAuxiliar
        tipo = byteAuxiliar
        '5) Minimo Valor Atributo
        Get archivoFuente, , minimoValorAtributo
        '6) Maximo Valor Atributo
        Get archivoFuente, , maximoValorAtributo
End Function

