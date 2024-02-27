Attribute VB_Name = "Engine_Declares"
Option Explicit

Public Type position
    X As Integer ' this should be a Byte
    Y As Integer ' this too
End Type

    
    



'Posicion en el Mundo
Public Type WorldPos
    map As Integer
    X As Integer ' ésto deberia ser un Byte
    Y As Integer ' ésto también.
End Type

Public Const INFINITE_LOOPS As Integer = -1



' GLOBALES MUAJAJA

Public hay_fogata_viewport  As Boolean
Public fogata_pos           As position

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function GetActiveWindow Lib "user32" () As Long


' COLORES ###########################################
Public Const mzRed          As Long = &HFFFF0000
Public Const mzGreen        As Long = &HFF00FF00
Public Const mzBlue         As Long = &HFF0000FF
Public Const mzWhite        As Long = &HFFFFFFFF
Public Const mzBlack        As Long = &HFF000000
Public Const mzColorApu     As Long = &HFFFFFFC0
Public Const mzYellow       As Long = &HFFFFFF00 '&HFFFFFFC0
Public Const mzColorMagic   As Long = &HFF00C0B9
Public Const mzPInk         As Long = &HFFD27CD8
Public Const mzCTalkMuertos As Long = &HF0787878
Public Const mzInterfaceColor1 As Long = &HFFE1D49F
Public Const mzInterfaceColor2 As Long = &HFF534926
'/COLORES ###########################################

' TEXTURAS ##########################################
Public Const TexturaSombra  As Integer = 1124
Public Const TexturaSangre  As Integer = 1126

Public Const TexturaTexto   As Integer = 1123
Public Const TexturaTexto2  As Integer = 1125

Public Const LightTextureHorizontal As Integer = 1922
Public Const LightTextureVertical As Integer = 1927
Public Const LightTextureFloor As Integer = 1923
Public Const LightTextureWall As Integer = 1924
'/TEXTURAS ##########################################

'/LUCES #############################################
Public Const LightBackbufferSize As Integer = 512 ' Tamaño en pixels del backbuffer de luces
Public Const LucesEnPantallaMax As Integer = 1000
'/LUCES #############################################

' PARTICULAS ########################################
Public Const PARTICULAS_LLUVIA As Integer = 5
Public Const PARTICULAS_NIEVE  As Integer = 9
'/PARTICULAS ########################################

Public Const bTRUE         As Byte = 255
Public Const bFALSE        As Byte = 0

Public Const Max_Int_Val   As Integer = 32767 ' (2 ^ 16) / 2 - 1

Public Const TILES_WIDTH = 21
Public Const TILES_HEIGHT = 21

#Const LUCES_VIEJAS = True 'NUEVAS == MMX

Public Function IsAppActive() As Boolean
    IsAppActive = (GetActiveWindow <> 0)
End Function
