Attribute VB_Name = "modContrato"
Option Explicit

Public Function contratosAceptados() As Boolean

Dim contrato As String

' ¿Ya lo acepto?
If ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("CONTRATO") = Environ("USERNAME") & "/" & Environ("USERDOMAIN") Then
    contratosAceptados = True
    Exit Function
End If

'No, no lo acepto
contrato = DBPath & "confidencialidad.rtf"

' ¿Existe?
If Not FileExist(contrato, vbNormal) Then
    contratosAceptados = False
    Exit Function
End If

' ¿Es valido?
If Not MD5File(contrato) = "e6b84c64f08af9f6ac7bba1be2211f20" Then
    Call MsgBox("Por favor, descarga manualmente la ultima versión del Editor del Mundo.", vbExclamation, "Tierras del Sur")
    contratosAceptados = False
    Exit Function
End If

' Lo cargo
Call frmContrato.rtbContrato.LoadFile(contrato)

' Remplazo los elementos
frmContrato.rtbContrato.TextRTF = Replace$(frmContrato.rtbContrato.TextRTF, "\{DIA\}", Day(Now))
frmContrato.rtbContrato.TextRTF = Replace$(frmContrato.rtbContrato.TextRTF, "\{MES\}", Month(Now))
frmContrato.rtbContrato.TextRTF = Replace$(frmContrato.rtbContrato.TextRTF, "\{ANO\}", year(Now))
frmContrato.rtbContrato.TextRTF = Replace$(frmContrato.rtbContrato.TextRTF, "\{DESTINATARIO-NOMBRE\}", UCase$(cerebro.Usuario.PersonaNombre))
frmContrato.rtbContrato.TextRTF = Replace$(frmContrato.rtbContrato.TextRTF, "\{DESTINATARIO-CORREO\}", "- " & cerebro.Usuario.Correo & " -")

' Mostramos
frmContrato.Show vbModal

' Devolvemos la devolucion
contratosAceptados = frmContrato.aceptoContrato

If contratosAceptados Then
    Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("CONTRATO", Environ("USERNAME") & "/" & Environ("USERDOMAIN"))
End If

Unload frmContrato

End Function



