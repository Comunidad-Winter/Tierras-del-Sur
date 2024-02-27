VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComandoRemplazarTileSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iComando

Private texturaOriginal As Integer
Private comienzoOriginal As Integer

Private texturaNueva As Integer
Private comienzoNueva As Integer

Public Sub crear(ByVal texturaOriginal_ As Integer, ByVal comienzoOriginal_ As Integer, ByVal texturaNueva_ As Integer, ByVal comienzoNueva_)
    texturaOriginal = texturaOriginal_
    comienzoOriginal = comienzoOriginal_
    texturaNueva = texturaNueva_
    comienzoNueva = comienzoNueva_
End Sub

Private Sub iComando_desHacer()
    Call remplazar(texturaNueva, texturaOriginal)
End Sub

Private Sub iComando_hacer()
    Call remplazar(texturaOriginal, texturaNueva)
End Sub

Private Function iComando_obtenerNombre() As String
    iComando_obtenerNombre = "Remplazar " & Tilesets(texturaOriginal).nombre & " por " & Tilesets(texturaNueva).nombre
End Function

Private Sub remplazar(ByVal texturaBuscada As Integer, ByVal texturaRemplazo As Integer)
    
    Dim loopX As Integer
    Dim loopY As Integer
    Dim remplazar As Boolean
    Dim seccionX As Integer
    Dim seccionY As Integer
    
    If Not comienzoOriginal = -1 Then
        seccionX = (comienzoOriginal Mod 16) \ 8
        seccionY = comienzoOriginal \ 128
    End If
    
    
    For loopX = SV_Constantes.X_MINIMO_USABLE To SV_Constantes.X_MAXIMO_USABLE
        For loopY = SV_Constantes.Y_MINIMO_USABLE To SV_Constantes.Y_MAXIMO_USABLE
            '¿Es la textura que busco?
            If mapdata(loopX, loopY).tile_texture = texturaBuscada Then
            
                remplazar = True
                
                ' ¿No especifica seccion?
                If comienzoOriginal > -1 Then
                    '¿Esta en la misma seccion?
                    If Not ((CInt((mapdata(loopX, loopY).tile_number Mod 16) \ 8)) = seccionX And CInt(mapdata(loopX, loopY).tile_number \ 128) = seccionY) Then
                        remplazar = False
                    End If
                End If
                
                ' Si corresponde, remplazo
                If remplazar Then
                    mapdata(loopX, loopY).tile_texture = texturaRemplazo
                    mapdata(loopX, loopY).tile_number = (mapdata(loopX, loopY).tile_number - comienzoOriginal) + comienzoNueva
                End If
            End If
        Next loopY
    Next loopX
    
    
   ' Call actualizarEfectoPisada
        
End Sub
