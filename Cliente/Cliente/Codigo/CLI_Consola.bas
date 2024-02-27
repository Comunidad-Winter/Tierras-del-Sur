Attribute VB_Name = "CLI_Consola"
Option Explicit

Private Type Historial
    text As String
    red As Integer
    green As Integer
    blue As Integer
    bold As Boolean
    italic As Boolean
    bCrLf As Boolean
End Type

Private HistorialConsola(0 To 100) As Historial

Private posicionActual As Integer

Private Type linkData
    texto As String
    url As String
End Type

Private Type MensajeConsola
    texto As String
    linkData As linkData
End Type

Public Sub AddtoRichTextBoxHistorico(text As String, Optional red As Integer = -1, Optional green As Integer, Optional blue As Integer, Optional bold As Boolean, Optional italic As Boolean, Optional bCrLf As Boolean)
    With frmConsola.ConsolaFlotante
        If (Len(.text)) > 6000 Then
            .SelStart = 0
            .SelLength = InStr(800, .text, vbCrLf)
            .SelText = ""
        End If
        
        .SelStart = Len(.text)
        .SelLength = 0
        
        .SelBold = IIf(bold, True, False)
        .SelItalic = IIf(italic, True, False)
        .SelUnderline = False
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        .SelText = IIf(bCrLf, text, text & vbCrLf)
    End With
End Sub
Public Sub AddtoRichTextBox(RichTextBox As RichTextBox, text As String, Optional red As Integer = -1, Optional green As Integer, Optional blue As Integer, Optional bold As Boolean, Optional italic As Boolean, Optional bCrLf As Boolean, Optional ByVal segundos As Integer = DURACION_TEXTO)
    
    Dim i As Integer
    
    If InStr(text, "[URL=") > 0 Then
        Dim Data As MensajeConsola
        Data = removerYObtenerLink(text)
        text = Trim$(Data.texto)
        
        Call CLI_CurrentLInk.setLink(Data.linkData.texto, Data.linkData.url)
     End If
    
    ' Mensaje vacio, no lo agregamos.
    If LenB(text) = 0 Then
        Exit Sub
    End If
    
    If frmConsola.ConsolaFlotante.Visible = False Then
    
        Consola.PushBackText remplazarCaracteresNoAdmitidos(text), D3DColorXRGB(red, green, blue), segundos
    
    End If
    
    
    If posicionActual > 99 Then
        posicionActual = 0
    End If
    
'    Call agregarElemento(HistorialConsola, text, red, green, blue, bold, italic, bCrLf)
    HistorialConsola(posicionActual).text = text
    HistorialConsola(posicionActual).red = red
    HistorialConsola(posicionActual).green = green
    HistorialConsola(posicionActual).blue = blue
    HistorialConsola(posicionActual).bold = bold
    HistorialConsola(posicionActual).italic = italic
    HistorialConsola(posicionActual).bCrLf = bCrLf
    
    posicionActual = posicionActual + 1
    
    If frmConsola.ConsolaFlotante.Visible = True Then
    
        AddtoRichTextBoxHistorico text, red, green, blue, bold, italic, bCrLf
    
    End If
    
End Sub

Public Sub CargarConsola()
Dim i As Integer
    If frmConsola.Visible = False Then
        For i = (posicionActual) To UBound(HistorialConsola)
            If LenB(HistorialConsola(i).text) <> 0 Then
                AddtoRichTextBoxHistorico HistorialConsola(i).text, HistorialConsola(i).red, HistorialConsola(i).green, HistorialConsola(i).blue, HistorialConsola(i).bold, HistorialConsola(i).italic, HistorialConsola(i).bCrLf
            End If
        Next
        
        For i = LBound(HistorialConsola) To posicionActual - 1
            AddtoRichTextBoxHistorico HistorialConsola(i).text, HistorialConsola(i).red, HistorialConsola(i).green, HistorialConsola(i).blue, HistorialConsola(i).bold, HistorialConsola(i).italic, HistorialConsola(i).bCrLf
        Next
    End If
End Sub

Private Function remplazarCaracteresNoAdmitidos(ByRef mensaje As String) As String
    remplazarCaracteresNoAdmitidos = Replace$(mensaje, vbTab, "  ")
End Function

Private Function removerYObtenerLink(ByRef mensaje As String) As MensajeConsola
Dim textoLink As String
Dim url As String

Dim posInicioLink As Integer
Dim posFinLink As Integer

Dim ultimaPos As Integer
Dim linkData As linkData

'Busco si hay almenos una url

ultimaPos = 1

posInicioLink = InStr(1, mensaje, "[URL=", vbTextCompare)

Do While posInicioLink > 0

    'Agrego el texto anterior al link
    If posInicioLink - ultimaPos > 0 Then
        removerYObtenerLink.texto = removerYObtenerLink.texto & mid$(mensaje, ultimaPos, posInicioLink - ultimaPos)
    End If
    
    'Obtengo el texto del link entre el = y el ]
    posInicioLink = posInicioLink + 5
    ultimaPos = InStr(posInicioLink, mensaje, "]", vbTextCompare)
    
    If ultimaPos > 0 Then '¿Esta bien armado el link?
    
        textoLink = mid$(mensaje, posInicioLink, ultimaPos - posInicioLink)
        
        'Obtengo la URL
        posFinLink = InStr(ultimaPos, mensaje, " ", vbTextCompare) - 1
        
        If (posFinLink = -1) Then posFinLink = Len(mensaje)
        
        
        url = mid$(mensaje, ultimaPos + 1, posFinLink - ultimaPos)
        
        linkData.url = url
        linkData.texto = textoLink
    End If
    
    posInicioLink = InStr(posFinLink, mensaje, "[URL=", vbTextCompare)
    ultimaPos = posFinLink + 1
Loop

If Len(mensaje) - posFinLink > 0 Then
    removerYObtenerLink.texto = removerYObtenerLink.texto & mid$(mensaje, posFinLink + 1, Len(mensaje) - posFinLink)
End If

removerYObtenerLink.linkData = linkData

End Function


