VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComandoInsertarMotania"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iComando

Private caja As RECT

Private alturasPieViejas() As Integer
Private alturasViejas() As AUDT
Private alturasPieNuevas() As Integer
Private alturasNuevas() As AUDT

Private toolInterna As Integer

Private modificaAlturaPie As Boolean

Public Sub crear(ByVal x As Integer, ByVal y As Integer, ByVal radio As Integer)
    toolInterna = ME_Tools.mt_select

    Dim radio_ As Integer
    radio_ = IIf(toolInterna = mtools.mt_pie, 0, radio) + 1
    
    caja.top = maxl(y - radio_, Y_MINIMO_VISIBLE)
    caja.Left = maxl(x - radio_, X_MINIMO_VISIBLE)
    caja.Bottom = minl(y + radio_, Y_MAXIMO_VISIBLE)
    caja.Right = minl(x + radio_, X_MAXIMO_VISIBLE)
    
    modificaAlturaPie = toolInterna = mtools.mt_pie Or frmMain.modifica_alt_pie.Value = vbChecked
    
    Dim tX As Integer, tY As Integer
    
    If toolInterna <> mtools.mt_pie Then
        ReDim alturasViejas(caja.Left To caja.Right, caja.top To caja.Bottom)
        ReDim alturasNuevas(caja.Left To caja.Right, caja.top To caja.Bottom)
        
        For tX = caja.Left To caja.Right
            For tY = caja.top To caja.Bottom
                alturasViejas(tX, tY) = hMapData(tX, tY)
            Next tY
        Next tX
    End If
    
    If modificaAlturaPie Then
        ReDim alturasPieViejas(caja.Left To caja.Right, caja.top To caja.Bottom)
        ReDim alturasPieNuevas(caja.Left To caja.Right, caja.top To caja.Bottom)
        
        For tX = caja.Left To caja.Right
            For tY = caja.top To caja.Bottom
                alturasPieViejas(tX, tY) = AlturaPie(tX, tY)
            Next tY
        Next tX
    End If
End Sub

Public Sub leerNuevas()

    Dim tX As Integer, tY As Integer
    
    If toolInterna <> mtools.mt_pie Then
        For tX = caja.Left To caja.Right
            For tY = caja.top To caja.Bottom
                alturasNuevas(tX, tY) = hMapData(tX, tY)
            Next tY
        Next tX
    End If
    
    If modificaAlturaPie Then
        For tX = caja.Left To caja.Right
            For tY = caja.top To caja.Bottom
                alturasPieNuevas(tX, tY) = AlturaPie(tX, tY)
            Next tY
        Next tX
    End If
End Sub

Private Sub iComando_desHacer()

    'If centroX = 0 And centroY = 0 And radio_ = 0 Then Exit Sub
    
    Dim tX As Integer, tY As Integer
    
    If toolInterna <> mtools.mt_pie Then
        For tX = caja.Left To caja.Right
            For tY = caja.top To caja.Bottom

                hMapData(tX, tY) = alturasViejas(tX, tY)
                
                With hMapData(tX, tY)
                    .h = (.hs(0) + .hs(1) + .hs(2) + .hs(3)) / 4
                    
                    Alturas(tX, tY) = .h
                    
                    If modificaAlturaPie Then '(.h > AlturaPie(x, Y) And AlturaPie(x, Y) = 0) Or AlturaPie(x, Y) = d Then
                        alturasPieViejas(tX, tY) = .h
                    End If
                    
                    If .h Then
                        If (.hs(0) > (.hs(1) + .hs(2)) / 2 And (.hs(1) + .hs(2)) / 2 > .hs(3)) Or (.hs(3) > (.hs(1) + .hs(2)) / 2 And (.hs(1) + .hs(2)) / 2 > .hs(0)) Then
                            MapData(tX, tY).tile_orientation = 1
                        Else
                            MapData(tX, tY).tile_orientation = 0
                        End If
                    End If
                End With
            Next tY
        Next tX
    End If
    
    If modificaAlturaPie Then
        For tX = caja.Left To caja.Right
            For tY = caja.top To caja.Bottom
                AlturaPie(tX, tY) = alturasPieViejas(tX, tY)
                Alturas(tX, tY) = AlturaPie(tX, tY)
            Next tY
        Next tX
    End If
    
    Backup_HM
    
    Compute_Mountain
End Sub

Private Sub iComando_hacer()
    'If centroX = 0 And centroY = 0 And radio_ = 0 Then Exit Sub
    
    Dim tX As Integer, tY As Integer
    
    If toolInterna <> mtools.mt_pie Then
        For tX = caja.Left To caja.Right
            For tY = caja.top To caja.Bottom
                hMapData(tX, tY) = alturasNuevas(tX, tY)
                
                With hMapData(tX, tY)
                    .h = (.hs(0) + .hs(1) + .hs(2) + .hs(3)) / 4
                    
                    Alturas(tX, tY) = .h
                    
                    If modificaAlturaPie Then '(.h > AlturaPie(x, Y) And AlturaPie(x, Y) = 0) Or AlturaPie(x, Y) = d Then
                        alturasPieNuevas(tX, tY) = .h
                    End If
                    
                    If .h Then
                        If (.hs(0) > (.hs(1) + .hs(2)) / 2 And (.hs(1) + .hs(2)) / 2 > .hs(3)) Or (.hs(3) > (.hs(1) + .hs(2)) / 2 And (.hs(1) + .hs(2)) / 2 > .hs(0)) Then
                            MapData(tX, tY).tile_orientation = 1
                        Else
                            MapData(tX, tY).tile_orientation = 0
                        End If
                    End If
                End With
            Next tY
        Next tX
    End If
    
    If modificaAlturaPie Then
        For tX = caja.Left To caja.Right
            For tY = caja.top To caja.Bottom
                AlturaPie(tX, tY) = alturasPieNuevas(tX, tY)
            Next tY
        Next tX
    End If
    
    Backup_HM
    
    Compute_Mountain
End Sub

Private Function iComando_obtenerNombre() As String
    If toolInterna = mtools.mt_pie Then
        iComando_obtenerNombre = "Editar altura del pie"
    Else
        iComando_obtenerNombre = "Editar montaña" & IIf(modificaAlturaPie, " && altura del pie", "")
    End If
End Function

