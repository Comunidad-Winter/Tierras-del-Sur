Attribute VB_Name = "Api"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez
'********************Misery_Ezequiel 28/05/05********************'
Option Explicit

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const EM_SETREADONLY = &HCF

Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long

'[MatuX] : 24 de Marzo del 2002
    Declare Function SetWindowPos& Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long)
'[END]

Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1

Type typDevMODE
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
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
(ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
(ByVal hWnd As Long, ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
'Requires Windows 2000 or later:
Private Const WS_EX_LAYERED = &H80000

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2

'********************Misery_Ezequiel 28/05/05********************'
Public Sub MakeWindowTransparent(ByVal hWnd As Long, ByVal alphaAmount As Byte)
Dim lStyle As Long
lStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
lStyle = lStyle Or WS_EX_LAYERED
SetWindowLong hWnd, GWL_EXSTYLE, lStyle
SetLayeredWindowAttributes hWnd, 0, alphaAmount, LWA_ALPHA
End Sub

