Attribute VB_Name = "Mod_Declaraciones"
Option Explicit

'Sonidos
Public Const SND_CLICK As Integer = 300
Public Const SND_PASOS1 As Integer = 23
Public Const SND_PASOS2 As Integer = 24
Public Const SND_NAVEGANDO As Integer = 50
Public Const SND_OVER As Integer = 301
Public Const SND_DICE As Integer = 302
Public Const SND_LLUVIAINEND As Integer = 303
Public Const SND_LLUVIAOUTEND As Integer = 304
Public Const SND_PASOS3 As Integer = 305
Public Const SND_PASOS4 As Integer = 306
Public Const SND_FUEGO As Integer = 307

Public Type tColor
    r As Byte
    g As Byte
    b As Byte
End Type

Public ColoresPJ(0 To 50) As tColor

'Objetos

Public Const MAX_NPC_INVENTORY_SLOTS As Byte = 50


Public Const Fogata_grh As Integer = 1521

Public Const NUMCIUDADES As Byte = 5
Public Const NUMSKILLS As Byte = 21
Public Const NUMATRIBUTOS As Byte = 5
Public Const NUMCLASES As Byte = 16
Public Const NUMRAZAS As Byte = 5



'
'********** FUNCIONES API ***********
'

Public Declare Function GetTickCount Lib "kernel32" () As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Para ejecutar el Internet Explorer para el manual
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public console_alpha As Boolean

Public antilag As Boolean
