VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDialogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'********************Misery_Ezequiel 28/05/05********************'
Option Explicit

Private Const MAX_DIALOGOS = 300
Private Const MAXLONG = 15
Private Type tDialogo
    Formato()  As String
    Formateado As Boolean
    Texto     As String
    Delay     As Integer
    X         As Integer
    Y         As Integer
    Vida      As Integer
    UserIndex As Integer
    Color     As Long
End Type
'Vector que contiene los DIALOGOS
Private Dialogos(1 To MAX_DIALOGOS) As tDialogo
'Tiempo de vida de cada mensaje
Private Delay As Integer
'Apunta a el ultimo mensaje
Public UltimoDialogo As Integer
'Contiene la cantidad de mensajes activos
Public CantidadDialogos As Integer

Private Function PrimerIndiceVacio() As Integer
On Error Resume Next
Dim I As Integer, j As Integer
Dim flag As Boolean

I = 1
j = 1
Do While I <= MAX_DIALOGOS And Not flag
     If (Dialogos(I).Vida < Dialogos(j).Vida) Then
         j = I
     End If
    flag = (Dialogos(I).Vida = 0)
    If Not flag Then I = I + 1
Loop
If flag Then
    PrimerIndiceVacio = I
Else
    PrimerIndiceVacio = j
End If
End Function

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?Formatea la cadena S¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Private Function DarFormato(s As String, I As Integer, k As Integer, anti As Integer, Dialgi As Integer)
On Error Resume Next
If anti + I <= Len(s) + 1 Then
    If ((I >= MAXLONG) And Mid$(s, anti + I, 1) = " ") Or (anti + I = Len(s)) Then
        Dialogos(Dialgi).Formato(k) = Mid(s, anti, I + 1)
        k = k + 1
        anti = anti + I + 1
        I = 0
    Else
        I = I + 1
    End If
    Call DarFormato(s, I, k, anti, Dialgi)
End If
End Function

Private Function HayEspacio(s As String) As Boolean
On Error Resume Next
If InStr(1, s, " ") <> 0 Then
    HayEspacio = True
Else
    HayEspacio = False
End If
End Function

Private Sub DarleFormatoTXT(Dialgi As Integer)
On Error Resume Next
 'Le damos formato al txt
  If (Len(Dialogos(Dialgi).Texto) > MAXLONG) And HayEspacio(Dialogos(Dialgi).Texto) Then
                Dim j As Integer
                ReDim Dialogos(Dialgi).Formato(0 To (Len(Dialogos(Dialgi).Texto) \ (MAXLONG \ 2)))
                Dim I As Integer, k As Integer, anti As Integer
                anti = 1
                k = 0
                I = 0
                Call DarFormato(Dialogos(Dialgi).Texto, I, k, anti, Dialgi)
                I = 0
                Do While Dialogos(Dialgi).Formato(I) <> "" And I < UBound(Dialogos(Dialgi).Formato)
                 I = I + 1
                Loop
                ReDim Preserve Dialogos(Dialgi).Formato(0 To I - 1)
                Dialogos(Dialgi).Formateado = True
  End If
End Sub

Public Sub CrearDialogo(txt As String, ByVal User As Integer, Color As Long)
On Error Resume Next
Dim MiUserIndex As Integer
Dim IndiceLibre As Integer
If BuscarUserIndex(User, MiUserIndex) Then
    Dialogos(MiUserIndex).Vida = 0
    Dialogos(MiUserIndex).Color = Color
    Dialogos(MiUserIndex).Formateado = False
    Dialogos(MiUserIndex).Texto = txt
    Dialogos(MiUserIndex).Delay = Delay
    Dialogos(MiUserIndex).UserIndex = User
    Dialogos(MiUserIndex).Vida = Delay
    DarleFormatoTXT MiUserIndex
Else
    IndiceLibre = PrimerIndiceVacio
    Dialogos(IndiceLibre).Color = Color
    Dialogos(IndiceLibre).Formateado = False
    Dialogos(IndiceLibre).Texto = txt
    Dialogos(IndiceLibre).Delay = Delay
    Dialogos(IndiceLibre).UserIndex = User
    Dialogos(IndiceLibre).Vida = Delay
    DarleFormatoTXT IndiceLibre
    If UltimoDialogo > IndiceLibre Then
        UltimoDialogo = IndiceLibre
    End If
End If
CantidadDialogos = CantidadDialogos + 1
End Sub

Private Function BuscarUserIndex(User As Integer, MiUser As Integer) As Boolean
On Error Resume Next
If CantidadDialogos > 0 Then
    Dim I As Integer
    Dim flag As Boolean
    I = 1
    Do While I <= MAX_DIALOGOS And Not flag
           flag = (Dialogos(I).UserIndex = User)
           If Not flag Then I = I + 1
    Loop
    
    If flag Then
            MiUser = I
            BuscarUserIndex = True
    Else
            BuscarUserIndex = False
    End If
End If
End Function

Public Sub Update_Dialog_Pos(X As Integer, Y As Integer, Index As Integer)
On Error Resume Next
Dim MiUserIndex As Integer
If BuscarUserIndex(Index, MiUserIndex) Then
  If Dialogos(MiUserIndex).Vida > 0 Then
        If Dialogos(MiUserIndex).Formateado Then
            Dialogos(MiUserIndex).X = X - (MAXLONG \ 2) * 3
            Dialogos(MiUserIndex).Y = Y - (UBound(Dialogos(MiUserIndex).Formato) * 3)
        Else
            Dialogos(MiUserIndex).X = X - (Len(Dialogos(MiUserIndex).Texto) \ 2) * 3
            Dialogos(MiUserIndex).Y = Y
        End If
        If Dialogos(MiUserIndex).Vida > 0 Then
           Dialogos(MiUserIndex).Vida = Dialogos(MiUserIndex).Vida - 1
        End If
        If Dialogos(MiUserIndex).Vida < 1 Then
            If MiUserIndex = UltimoDialogo Then
                
                Do While Dialogos(UltimoDialogo).Vida = 0
                    UltimoDialogo = UltimoDialogo - 1
                    If UltimoDialogo < 1 Then
                        UltimoDialogo = 0
                        CantidadDialogos = 0
                        Exit Sub
                    End If
                Loop
            End If
            CantidadDialogos = CantidadDialogos - 1
        End If
  End If
End If
End Sub
Public Sub DrawText(lngXPos As Integer, lngYPos As Integer, strText As String, lngColor As Long)
On Error Resume Next
   If strText <> "" Then
        BackBufferSurface.SetFontTransparency True
        BackBufferSurface.SetForeColor vbBlack
        BackBufferSurface.SetFont frmMain.Font
        BackBufferSurface.DrawText lngXPos - 2, lngYPos - 1, strText, False
        BackBufferSurface.SetFontTransparency True
        BackBufferSurface.SetForeColor lngColor
        BackBufferSurface.SetFont frmMain.Font
        BackBufferSurface.DrawText lngXPos, lngYPos, strText, False
   End If
End Sub

Public Sub MostrarTexto()
On Error Resume Next
Dim I As Integer
'For i = 1 To CantidadDialogos
I = 1
Do While I <= CantidadDialogos And I <= MAX_DIALOGOS
    If Dialogos(I).Vida > 0 Then
        If Dialogos(I).Formateado Then
            Dim j As Integer, desp As Integer
            desp = (-frmMain.Font.Size) * UBound(Dialogos(I).Formato) - 5
            For j = 0 To UBound(Dialogos(I).Formato)
             DrawText Dialogos(I).X, Dialogos(I).Y + desp, Dialogos(I).Formato(j), Dialogos(I).Color
             desp = desp + (frmMain.Font.Size) + 5
            Next
         Else
            DrawText Dialogos(I).X, Dialogos(I).Y, Dialogos(I).Texto, Dialogos(I).Color
         End If
    Else
      Dialogos(I).Texto = ""
      Dialogos(I).UserIndex = 0
      Dialogos(I).Formateado = False
      Dialogos(I).Delay = 0
    End If
    I = I + 1
Loop
'Next
End Sub

Public Sub QuitarDialogo(ByVal CharIndex As Integer)
On Error Resume Next
Dim I As Integer
If BuscarUserIndex(CharIndex, I) Then
      Dialogos(I).Vida = 0
      Dialogos(I).Texto = ""
      Dialogos(I).UserIndex = 0
      Dialogos(I).Formateado = False
      Dialogos(I).Delay = 0
      CantidadDialogos = CantidadDialogos - 1
      If I = UltimoDialogo Then
        Do While Dialogos(I).Vida = 0
                I = I - 1
                If I < 1 Then
                    UltimoDialogo = 0
                    CantidadDialogos = 0
                    Exit Sub
                End If
            Loop
      End If
End If
End Sub

Public Sub BorrarDialogos()
On Error Resume Next
Dim I As Integer
For I = 1 To MAX_DIALOGOS
      Dialogos(I).Vida = 0
      Dialogos(I).Texto = ""
      Dialogos(I).UserIndex = 0
      Dialogos(I).Formateado = False
      Dialogos(I).Delay = 0
Next I
End Sub

Private Sub Class_Initialize()
Delay = 300
UltimoDialogo = 0
CantidadDialogos = 0
End Sub
'********************Misery_Ezequiel 28/05/05********************'
