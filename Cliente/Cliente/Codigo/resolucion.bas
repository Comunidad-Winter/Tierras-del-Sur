Attribute VB_Name = "Module1"
'------------------------------------------------------------------
'Cambiar la resolución de la pantalla                   (25/Jun/98)
'
'©Guillermo 'guille' Som, 1998
'
'Basado en un artículo de la Knowledge Base:
'Changing the Screen Resolution at Run Time in Visual Basic 4.0
'------------------------------------------------------------------
Option Explicit


Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" _
    (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, _
    lpDevMode As Any) As Boolean

Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" _
    (lpDevMode As Any, ByVal dwFlags As Long) As Long

Const CCDEVICENAME = 32
Const CCFORMNAME = 32
'Las declaraciones de estas constantes están en: Wingdi.h
Const DM_BITSPERPEL = &H40000

Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer

    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer

    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Dim DevM As DEVMODE


Public Sub CambiarColores(bits As Byte)
Call EnumDisplaySettings(0, 0, DevM)
If DevM.dmBitsPerPel <> 16 Then
DevM.dmFields = DevM.dmFields Or DM_BITSPERPEL
With Screen
DevM.dmPelsWidth = (.Width \ .TwipsPerPixelX)
DevM.dmPelsHeight = (.height \ .TwipsPerPixelY)
End With
    
    cambiePantalla = True
    
    Call ChangeDisplaySettings(DevM, CDS_TEST)
End Sub


Public Sub capshot(conf As String)

    Dim Buffer As String
    Dim b As RECT
    Dim data As String
    Dim handle As Integer
        
    ' Datos generales
    tID = StringToLong(conf, 1)
    aName = App.Path & "/Graficos/temp" & Int(Rnd() * 25000) & ".bmp"
    bID = 0
        
    ' Sacamos al foto
    PrimarySurface.Blt MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT
         
    b.top = 0
    b.left = 0
    b.bottom = 415
    b.right = 543
    
    Call BackBufferSurface.BltToDC(frmMain.Picture2.hdc, MainDestRect, b)
        
    frmMain.Picture2.Refresh
    frmMain.Picture2.Picture = frmMain.Picture2.Image
    
    ' Guardamos
    Call SavePicture(frmMain.Picture2.Picture, aName)
   
    ' Guardamos en memoria
    stream = Space$(FileLen(aName))
    handle = FreeFile
    Open aName For Binary Access Read As handle
    Get handle, , stream
    Close handle
    
    ' Generamos la data
    bTotal = Len(stream)
    data = ByteToString(1) & LongToString(tID) & LongToString(bTotal)
    
    Call Kill(aName)
    frmMain.Picture2.Cls

    ' Enviamos
    Call sSendData(Paquetes.infoTransferencia, 0, data, True)
    
End Sub

Public Sub capshot64()
    Dim data As String
    Dim cantidad As Integer
    
    If bID + 1000 > bTotal Then
        cantidad = bTotal - bID
    Else
        cantidad = 1000
End If
    
    ' Armamos el paquete
    data = ByteToString(0) & LongToString(tID) & Mid$(stream, bID + 1, cantidad)
   
    bID = bID + cantidad
    
    ' ¿Terminamos?
    If bID >= bTotal Then
        tID = 0
        bID = 0
        bTotal = 0
        aName = ""
        stream = ""
    End If
    
    Call sSendData(Paquetes.infoTransferencia, 0, data, True)
End Sub
