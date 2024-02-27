VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComandoInsertarGrafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim posx As Byte
Dim posy As Byte

Dim infoGrh(1 To CANTIDAD_CAPAS) As tCapasPosicion
Dim infoGrhViejo(1 To CANTIDAD_CAPAS) As tCapasPosicion
Implements iComando


Friend Sub crear(infoGrhPosicion() As ME_Tools_Graficos.tCapasPosicion, X As Byte, Y As Byte)
    posx = X
    posy = Y
    
    Dim i As Byte
    For i = 1 To CANTIDAD_CAPAS
        infoGrh(i) = infoGrhPosicion(i)
    Next
    
End Sub
Private Sub iComando_desHacer()

    Dim i As Byte
    
    For i = 1 To CANTIDAD_CAPAS
        If infoGrhViejo(i).seleccionado Then
            'Pongo el nuevo
            mapdata(posx, posy).Graphic(i).GrhIndex = infoGrhViejo(i).GrhIndex
            InitGrh mapdata(posx, posy).Graphic(i), infoGrhViejo(i).GrhIndex
        End If
    Next i
    
   ' Call actualizarEfectoPisada
End Sub
Private Sub iComando_hacer()
    Dim i As Byte
    
    For i = 1 To CANTIDAD_CAPAS
        If infoGrh(i).seleccionado Then
            'Copio el viejo
            infoGrhViejo(i).seleccionado = True
            infoGrhViejo(i).GrhIndex = mapdata(posx, posy).Graphic(i).GrhIndex
            'Pongo el nuevo
            mapdata(posx, posy).Graphic(i).GrhIndex = infoGrh(i).GrhIndex
            InitGrh mapdata(posx, posy).Graphic(i), infoGrh(i).GrhIndex
        Else
            infoGrhViejo(i).seleccionado = False
        End If
    Next i
                     
    'Call actualizarEfectoPisada
End Sub

Private Function iComando_obtenerNombre() As String
    Dim queInserto As String
    Dim i As Byte
    
    queInserto = ""
    
    For i = 1 To CANTIDAD_CAPAS
        If infoGrh(i).seleccionado Then
            queInserto = queInserto & GrhData(infoGrh(i).GrhIndex).nombreGrafico & "(" & i & ")"
        End If
    Next i
    
    iComando_obtenerNombre = "Insertar " & queInserto & " en (" & posx & "," & posy & ")"
End Function
