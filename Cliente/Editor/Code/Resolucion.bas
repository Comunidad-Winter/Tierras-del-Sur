Attribute VB_Name = "modPantalla"
Public PixelesPorTile As D3DVECTOR2
Public TilesPantalla As Position ' Cantidad de tiles que muestra el render

Public mostrarBarraHerramientas As Boolean

Public Sub Pantalla_Iniciar()

TilesPantalla.x = CByte(val(ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("TilesAncho")))
TilesPantalla.y = CByte(val(ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("TilesAlto")))

mostrarBarraHerramientas = (ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("MostrarBarraHerramientas") = "SI")

If TilesPantalla.x = 0 Then TilesPantalla.x = 32
If TilesPantalla.y = 0 Then TilesPantalla.y = 20

End Sub

Public Sub Pantalla_Guardar()

Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("TilesAncho", CStr(TilesPantalla.x))
Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("TilesAlto", CStr(TilesPantalla.y))

Call ME_Configuracion_Usuario.actualizarPreferenciaWorkSpace("MostrarBarraHerramientas", IIf(mostrarBarraHerramientas, "SI", "NO"))

End Sub

Public Sub Pantalla_AcomodarElementos()

Dim AnchoFinal As Integer
Dim AltoFinal As Integer
Dim Offset As Integer

AnchoFinal = modPantalla.TilesPantalla.x * 32
AltoFinal = modPantalla.TilesPantalla.y * 32
    
frmMain.width = AnchoFinal * Screen.TwipsPerPixelX
frmMain.height = AltoFinal * Screen.TwipsPerPixelY
    
Offset = frmMain.ScaleWidth - AnchoFinal
Offset = frmMain.ScaleHeight - AltoFinal

frmMain.width = (AnchoFinal - Offset) * Screen.TwipsPerPixelX
frmMain.height = (AltoFinal - Offset) * Screen.TwipsPerPixelY

End Sub

