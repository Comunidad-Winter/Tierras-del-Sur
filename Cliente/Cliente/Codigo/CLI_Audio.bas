Attribute VB_Name = "CLI_Audio"
Option Explicit

Public CurMidi As Integer

Public Sub actualizarVolumen(volumen As Single)
    Sonido_CambiarVolumen_Ambiente (volumen)
End Sub
Public Sub toogleMusica()

    Musica = Not Musica
    
    If Musica Then
        Call activarMusica
    Else
        Call desactivarMusica
    End If
                    
End Sub


Public Sub activarMusica()

    Musica = True

    If CurMidi > 0 Then Call Sonido_Play_Ambiente(CurMidi)
                    
End Sub


Public Sub desactivarMusica()

    Musica = False
    
    If CurMidi > 0 Then Call Sonido_Stop_Ambiente(CurMidi)
                    
End Sub

