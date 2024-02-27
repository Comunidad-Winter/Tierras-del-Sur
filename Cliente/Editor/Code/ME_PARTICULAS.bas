Attribute VB_Name = "ME_PARTICULAS"
Option Explicit

'TODO, no usan.
Public particula_editada As Integer
Public particula_testeo As Integer
Public emisor_testeo As Integer
Public emisor_editado As Integer
Public emisor_casteado(2) As Integer

Public part_totales As Integer

Enum Caracteristicas
    Crece = &H1
    movimiento_sinoudal = &H2
                                ' X:     Y:
    spd_trig = &H4
        
    acc_trig = &H100
        acc_SS = &H200

    pos_trig = &H2000
        pos_SS = &H4000
        
    eCaracteristicas_ForceDWORD = &H7FFFFFFF
End Enum


Public TipoEditorParticulas As Boolean

Public Sub VBC2RGBC(ByVal lColor As Long, ByRef col As RGBCOLOR)
col.r = (lColor And 255)
col.g = ((lColor \ 256) And 255)
col.b = ((lColor \ 65536) And 255)
End Sub

Public Sub VBC2MZC(ByVal lColor As Long, lRed As Single, lGreen As Single, lBlue As Single)
lRed = (lColor And 255) / 255
lGreen = ((lColor \ 256) And 255) / 255
lBlue = ((lColor \ 65536) And 255) / 255
End Sub

