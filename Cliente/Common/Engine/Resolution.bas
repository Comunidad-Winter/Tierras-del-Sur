Attribute VB_Name = "Engine_Resolution"
' ARCHIVO COMPARTIDO
Option Explicit

Private Const CCDEVICENAME As Long = 32
Private Const CCFORMNAME As Long = 32
Private Const DM_BITSPERPEL As Long = &H40000
Private Const DM_PELSWIDTH As Long = &H80000
Private Const DM_PELSHEIGHT As Long = &H100000
Private Const DM_DISPLAYFREQUENCY As Long = &H400000
Private Const CDS_TEST As Long = &H4
Private Const ENUM_CURRENT_SETTINGS As Long = -1

Private Type typDevMODE
    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type

Private oldDepth As Integer
Public PDepth As Integer
Private oldFrequency As Long

Public oldResHeight As Long, oldResWidth As Long
Public bNoResChange As Boolean

Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long
Public resolucionActual As Integer

Public pixelesAncho As Integer
Public pixelesAlto As Integer

Public Const RESOLUCION_43 = 1
Public Const RESOLUCION_169 = 2

Public Sub setResolucionJuego(numero As Integer)
    
    Select Case numero
        Case RESOLUCION_43
            pixelesAlto = 768
            pixelesAncho = 1024
        Case RESOLUCION_169
            pixelesAlto = 720
            pixelesAncho = 1280
    End Select

    resolucionActual = numero
End Sub


Public Sub SetResolutionPantalla(Optional ByVal pCambiarResolucion As Boolean = True, Optional ByVal ancho As Integer, Optional ByVal alto As Integer)
On Error GoTo errh:

    Dim lRes As Long
    Dim MidevM As typDevMODE
    
    If pCambiarResolucion Then
        pCambiarResolucion = (oldResWidth <> ancho Or oldResHeight <> alto)
    End If

    If pCambiarResolucion Then
    
        lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, MidevM)
    
        oldResWidth = Screen.width \ Screen.TwipsPerPixelX
        oldResHeight = Screen.height \ Screen.TwipsPerPixelY
        oldDepth = MidevM.dmBitsPerPel
        PDepth = oldDepth
        oldFrequency = MidevM.dmDisplayFrequency
    
        With MidevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT ' Or DM_BITSPERPEL
            .dmPelsWidth = ancho
            .dmPelsHeight = alto
        End With
        
        lRes = ChangeDisplaySettings(MidevM, CDS_TEST)
        bNoResChange = True
    End If

errh:

End Sub

Public Sub ResetResolution()
    Dim typDevM As typDevMODE
    Dim lRes As Long
    
    If bNoResChange Then
    
        lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, typDevM)
        
        With typDevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
            .dmPelsWidth = oldResWidth
            .dmPelsHeight = oldResHeight
        End With
        
        lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
    End If
End Sub
