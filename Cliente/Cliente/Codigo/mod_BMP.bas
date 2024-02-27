Attribute VB_Name = "mod_BMP"
''''''''''''''''''''''''''''''''''''''''''''''
''''ESTE MOD ES TIPO BMP MAN PERO MEJOR :P''''
''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''MADE BY EL YIND'''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''
'''''''''javier_podavini@hotmail.com''''''''''
''''''''''''''''''''''''''''''''''''''''''''''
'''''''''"SI TE PICA EL CULO RASCATE"'''''''''
''''''''''''''''''''''''''''''''''''''''''''''
'MODIFACO POR MARCHE

Private Type InfoHe
EmpiezaByte As Long
CantidaddeBytes As Long
End Type


Private CabezalInterace() As InfoHe
Private CabezalGraficos() As InfoHe
Private CabezalMapas() As InfoHe


Public Type tGP
    file As Integer
    Offset As Long
    Height As Long
    Width As Long
    FileSizeBMP As Long
End Type
'Public GPdata() As tGP
'Public GPdataBMP() As tGP
Public UsarBinario As Boolean
Private Dibujitos(1 To 15000) As DirectDrawSurface7
Private DibujitosC(1 To 15000) As Boolean
Public Function DameBMP(Num As Integer) As DirectDrawSurface7
Call CargarSurface(Num)

Set DameBMP = Dibujitos(Num)
End Function
'BMPMAN FEO... para que borrar los graficos. si los cargaste es porque lo usaste feo.
Public Function CargarSurface(Num As Integer) As Boolean
On Error GoTo ErrHandler

If DibujitosC(Num) Then Exit Function
Dim ddsd As DDSURFACEDESC2, ddck As DDCOLORKEY
ddsd.lFlags = DDSD_CAPS
If UseMemVideo Then
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Else
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
End If
ddck.high = 0: ddck.low = 0
'If UsarBinario Then
    ExtraerGraficos (Num)
    'Call ExtractData(App.Path & "\Graficos\Graficos.tds", GPdataBMP(Num).Offset, GPdataBMP(Num).FileSizeBMP)
    Set Dibujitos(Num) = DirectDraw.CreateSurfaceFromFile(App.Path & "\Graficos\Temp.bmp", ddsd) 'creo la surface
    Call Kill(App.Path & "\Graficos\Temp.bmp")
'Else
 '   Set Dibujitos(Num) = DirectDraw.CreateSurfaceFromFile(App.Path & "\Graficos\" & Num & ".bmp", ddsd) 'creo la surface
'End If
Dibujitos(Num).SetColorKey DDCKEY_SRCBLT, ddck
CargarSurface = True
DibujitosC(Num) = True
Exit Function
ErrHandler:
CargarSurface = False
End Function
Public Sub LiberarGraficos()
Dim i As Integer
For i = 1 To NumeroDeBMPs
    If Not Dibujitos(i) Is Nothing Then
        Set Dibujitos(i) = Nothing
    End If
Next i
End Sub
'Sub ExtractData(strFileName As String, lngOffset As Long, FileSizeBMP As Long)
''on error Resume Next
'Dim intBMPFile As Integer
'Dim intFreeFile As Integer
'Dim I As Integer
'Dim FileEnSi() As Byte
'intBMPFile = FreeFile()

'Open strFileName For Binary Access Read Lock Write As intBMPFile
 '   ReDim FileEnSi(FileSizeBMP)
  '  Get intBMPFile, lngOffset, FileEnSi
   ' intFreeFile = FreeFile()
   ' Open App.Path & "\Temp.bmp" For Binary Access Read Write Lock Write As intFreeFile
    '    Put intFreeFile, , FileEnSi
   ' Close intFreeFile
'Close intBMPFile
'End Sub
'Sub CargarGPData()
'on error Resume Next
'Dim intFreeFile As Integer
'Dim Nums As Integer
'intFreeFile = FreeFile()
'Dim LoopC As Integer
'Open App.Path & "\Graficos\LG.tds" For Binary As intFreeFile
 '   Get intFreeFile, , MiCabecera
  '  Get intFreeFile, , Nums
   ' ReDim GPdata(Nums)
   ' For LoopC = 1 To Nums

    '    Get intFreeFile, , GPdata(LoopC)
     '   If GPdata(LoopC).file > NumBMP Then
      '      NumBMP = GPdata(LoopC).file
       '     ReDim Preserve GPdataBMP(NumBMP)
       ' End If
       '     GPdataBMP(GPdata(LoopC).file) = GPdata(LoopC)
    'Next LoopC
'Close intFreeFile
'End Sub
'Public Sub Unir(Archivo As String, NumBMP As Integer)
'Dim bytSecond() As Byte
'Dim a As Long
'Dim mlngLocationFirst As Long
'intFreeFile = FreeFile
'Open Archivo For Binary Access Read Lock Write As intFreeFile
 '   a = LOF(intFreeFile) - 1
  '  ReDim bytSecond(a)
   ' Get intFreeFile, , bytSecond()
'Close intFreeFile
'intFreeFile = FreeFile
'Open App.Path & "\Graficos\Graficos.tds" For Binary Access Read Write Lock Write As intFreeFile
'    mlngLocationFirst = LOF(intFreeFile) - 1
 '   Put intFreeFile, mlngLocationFirst, bytSecond()
'Close intFreeFile

'Dim Encontro As Boolean
'For a = 1 To UBound(GPdata)
   ' If GPdata(a).file = NumBMP Then
   '     Encontro = True
  '      Exit For
 '   End If
'Next a
'If Encontro Then
    'GPdata(a).file = NumBMP
    'GPdata(a).FileSizeBMP = FileLen(Archivo)
   ' GPdata(a).Height = 0
  '  GPdata(a).Offset = mlngLocationFirst
 '   GPdata(a).Width = 0
'Else
   ' ReDim Preserve GPdata(UBound(GPdata) + 1)
   ' GPdata(UBound(GPdata)).file = NumBMP
   ' GPdata(UBound(GPdata)).FileSizeBMP = FileLen(Archivo)
   ' GPdata(UBound(GPdata)).Height = 0
  '  GPdata(UBound(GPdata)).Offset = mlngLocationFirst
 '   GPdata(UBound(GPdata)).Width = 0
'End If
'intFreeFile = FreeFile()
'Open App.Path & "\Graficos\LG.tds" For Binary As intFreeFile
   ' Put intFreeFile, , MiCabecera
   ' Put intFreeFile, , CInt(UBound(GPdata))
  '  For LoopC = 1 To UBound(GPdata)
 '       Put intFreeFile, , GPdata(LoopC)
'    Next LoopC
'Close intFreeFile
'End Sub
Public Sub ExtraerImagen(Imagen As Integer)
archivoB = FreeFile()
Archivo = FreeFile() + 1
Open App.Path & "\Graficos\Interface.TDS" For Binary As archivoB
        ReDim Data(CabezalInterace(Imagen).CantidaddeBytes - 1) As Byte
        Get archivoB, CabezalInterace(Imagen).EmpiezaByte, Data
        
        Open App.Path & "\Graficos\temp.jpg" For Binary As Archivo
        Put Archivo, , Data
        Close Archivo
Close archivoB
End Sub

Public Sub CargarHeads()
Dim cantidad As Integer
Dim UltimoByte As Long

archivoB = FreeFile()
    
    '//////// INTERFACE
    Open App.Path & "\Graficos\Interface.TDS" For Binary As archivoB
    Get archivoB, 1, cantidad
    Get archivoB, , UltimoByte
    ReDim CabezalInterace(cantidad)
    Get archivoB, UltimoByte, CabezalInterace
    Close archivoB
    '/////// MAPAS
    Open App.Path & "\Mapas\Mapas.TDS" For Binary As archivoB
    Get archivoB, 1, cantidad
    Get archivoB, , UltimoByte
    ReDim CabezalMapas(cantidad)
    Get archivoB, UltimoByte, CabezalMapas
    Close archivoB
    '////// GRAFICOS
    Open App.Path & "\Graficos\Graficos.TDS" For Binary As archivoB
    Get archivoB, 1, cantidad
    Get archivoB, , UltimoByte
    ReDim CabezalGraficos(cantidad)
    Get archivoB, UltimoByte, CabezalGraficos
    Close archivoB
    '///////////////////////
    
End Sub
Public Sub ExtraerMapa(Numero As Integer)
archivoB = FreeFile()
Archivo = FreeFile() + 1
Open App.Path & "\Mapas\Mapas.TDS" For Binary As archivoB
        ReDim Data(CabezalMapas(Numero).CantidaddeBytes - 1) As Byte
        Get archivoB, CabezalMapas(Numero).EmpiezaByte, Data
        
        Open App.Path & "\Mapas\temp.map" For Binary As Archivo
        Put Archivo, , Data
        Close Archivo
Close archivoB
End Sub

Public Sub ExtraerGraficos(Numero As Integer)
archivoB = FreeFile()
Archivo = FreeFile() + 1
Open App.Path & "\Graficos\Graficos.TDS" For Binary As archivoB
        ReDim Data(CabezalGraficos(Numero).CantidaddeBytes - 1) As Byte
        Get archivoB, CabezalGraficos(Numero).EmpiezaByte, Data
        
        Open App.Path & "\Graficos\temp.bmp" For Binary As Archivo
        Put Archivo, , Data
        Close Archivo
Close archivoB
End Sub

Public Sub DibujaGrh(Grh As Integer)
Dim sR As RECT, DR As RECT

sR.left = 0
sR.top = 0
sR.right = 32
sR.bottom = 32

DR.left = 0
DR.top = 0
DR.right = 32
DR.bottom = 32
Call DrawGrhtoHdc(frmComerciarUsu.Picture1.hWnd, frmComerciarUsu.Picture1.hdc, Grh, sR, DR)

End Sub

Sub DrawGrhtoHdc(hWnd As Long, hdc As Long, Grh As Integer, SourceRect As RECT, destRect As RECT)
If Grh <= 0 Then Exit Sub
SecundaryClipper.SetHWnd hWnd
Call CargarSurface(GrhData(Grh).FileNum)
Call Dibujitos(GrhData(Grh).FileNum).BltToDC(hdc, SourceRect, destRect)
End Sub
