Attribute VB_Name = "ConsolaInteligente"
Option Explicit

Private Const EM_CHARFROMPOS& = &HD7
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' Return the word the mouse is over.
Public Function RichWordOver(rch As RichTextBox, x As Single, y As Single) As String
Dim pt As POINTAPI
Dim pos As Integer
Dim start_pos As Integer
Dim end_pos As Integer
Dim ch As String
Dim escapo As Boolean
    ' Convert the position to pixels.
    pt.x = x \ Screen.TwipsPerPixelX
    pt.y = y \ Screen.TwipsPerPixelY

    ' Get the character number
    pos = SendMessage(rch.hWnd, EM_CHARFROMPOS, 0&, pt)
    
    ' Busco hacia la izquierda para buscar si es un link
    For start_pos = pos To 1 Step -1
    
        If start_pos <= 5 Then Exit Function
        
        ch = mid$(rch.text, start_pos - 4, 5)
        
        'Espacio o salto de linea paro
        If Right$(ch, 1) = (" ") Or Right$(ch, 1) = Chr(13) Then
            If Not escapo Then
                Exit Function
            Else
                escapo = False
            End If
        ElseIf (ch = "[URL]") Then
            Exit For
        ElseIf Right$(ch, 1) = Chr(255) Then
            escapo = True
        End If
        
    Next start_pos
    
    start_pos = start_pos - 5

    ' Busco el link
    For end_pos = start_pos To 1 Step -1
        ch = mid$(rch.text, end_pos, 1)

        If ch = " " Then Exit For
        
        If ch = Chr(13) Then
            end_pos = end_pos + 1
            Exit For
        End If
    Next end_pos

    end_pos = end_pos + 1

    If start_pos >= end_pos Then _
        RichWordOver = mid$(rch.text, end_pos, start_pos - end_pos + 1)
End Function

Public Sub addoToConsolaInteligente(rch As RichTextBox, mensaje As String, red As Integer, blue As Integer, green As Integer, negrita As Boolean, cursiva As Boolean)

Dim textoLink As String
Dim url As String

Dim posInicioLink As Integer
Dim posFinLink As Integer

Dim ultimaPos As Integer
'Busco si hay almenos una url

ultimaPos = 1

posInicioLink = InStr(1, mensaje, "[URL=", vbTextCompare)

Do While posInicioLink > 0

    'Agrego el texto anterior al link
    If posInicioLink - ultimaPos > 0 Then
        Call AddtoRichTextBox(rch, mid$(mensaje, ultimaPos, posInicioLink - ultimaPos), red, green, blue, negrita, cursiva, True)
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
        
        Call AddLINK(rch, textoLink, url, red, green, blue)
    End If
    
    posInicioLink = InStr(posFinLink, mensaje, "[URL=", vbTextCompare)
    ultimaPos = posFinLink + 1
Loop

'Agrego lo que quedo despues del link
If Len(mensaje) - posFinLink > 0 Then
    Call AddtoRichTextBox(rch, mid$(mensaje, posFinLink + 1, Len(mensaje) - posFinLink), red, green, blue, negrita, cursiva)
Else
    Call AddtoRichTextBox(rch, "", red, green, blue)
End If
End Sub


Public Sub AddLINK(RichTextBox As RichTextBox, texto As String, url As String, red As Integer, green As Integer, blue As Integer)
On Error GoTo hayError:
    Dim i As Integer
    
    With RichTextBox
    
        If (Len(.text)) > 2000 Then
            .SelStart = 0
            .SelLength = InStr(800, .text, vbCrLf)
            .SelText = ""
        End If
        
        .SelStart = Len(RichTextBox.text)
        .SelLength = 0
        .SelRTF = "{\rtf1\ansi " + "\v " + url + "[URL]" + "\v0}"
        .SelBold = True
        .SelItalic = True
        .SelUnderline = True
       
        .SelColor = RGB(red, green, blue)
       
        For i = 1 To Len(texto)
            If mid$(texto, i, 1) = " " Then
                .SelText = " "
                .SelRTF = "{\rtf1\ansi " + "\v " + Chr(255) + "\v0}"
                
                .SelBold = True
                .SelItalic = True
                .SelUnderline = True
                .SelColor = RGB(red, green, blue)
                .SelStart = Len(RichTextBox.text)
                .SelLength = 1
            Else
                .SelStart = Len(RichTextBox.text)
                .SelLength = 1
                .SelText = mid$(texto, i, 1)
            End If
        Next
    End With
    
hayError:
End Sub
