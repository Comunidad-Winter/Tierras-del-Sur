Attribute VB_Name = "Engine_Desencadenadores"
Option Explicit

'Qué son los desencadenadores?!?!?
' Son acciones genéricas del engine. Por ejemplo. Crear una partícula.

'Cómo funcionan?
' Con cadenas de texto. Es como mandar un paquete al servidor. Se interpreta de la misma forma.

'Para qué sirve?
' Para guardar videos. Para crear una particula cuando otra muere. Para enviar datos al engine desde el servidor. Para que explote un barril.
' Para emitir sonidos cuando se realizen ciertas acciones.
'!Para combinarse con los tileEvents. EJ: cuando se pise una tile que caiga un rayo.
'   Entonces el server lo que hace es enviar el "RAYO" a todos los clientes del area, sin hardcodearlo.

'!Sincronizacion de clientes y servidor:
'   Mejor explicado en el GDOC: https://docs.google.com/document/d/1D3WrMWhPx1Wiu_8IzEVExJ3-q3m4GMgpx99jkdv_OCA/edit?hl=es




'----------------------------------------------------------------------------------------------------------------

Public Enum eDesencadenadores
    CrearParticula = 1
End Enum

Public Function ObtenerDesencadenador(Serializado As String) As Engine_IDesencadenador
    Dim Tipo As Byte
    Tipo = AscB(Serializado)
    Select Case Tipo
        Case eDesencadenadores.CrearParticula
            ObtenerDesencadenador = New desencadenadorCrearParticula
            ObtenerDesencadenador.Unserialize Serializado
    End Select
End Function

Public Function EncolarDesencadenador(Desencadenador As Engine_IDesencadenador, Optional ByVal TickSincronizado As Long = 0)
    'If TickSincronizado > ObtenerTickSincronizado() Then
        'Desencadenador.Tick = TickSincronizado
        'ListaDesencadenadores.Agregar Desencadenador
        'ListaDesencadenadores.ReOrdenar
    'Else
        Desencadenador.Ejecutar
    'End if
End Function

Public Function ProcesarDesencadenadores()
'    Dim TickSincronia As Long
'    TickSincronia = ObtenerTickSincronizado
'
'    Do While lista.Count > 0
'        If ListaDesencadenadores.Ultimo.Tick > TickSincronia Then Exit Do
'        ListaDesencadenadores.Ultimo.Ejecutar
'        ListaDesencadenadores.QuitarUltimo
'    Wend
End Function

