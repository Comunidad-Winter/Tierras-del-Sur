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
Option Explicit

Private Const MAX_DIALOGOS = 300
Private Const MAXLONG = 15
Private Type tDialogo
    Formato()  As String
    Formateado As Boolean
    texto     As String
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
Dim i As Integer, j As Integer
Dim flag As Boolean

i = 1
j = 1
Do While i <= MAX_DIALOGOS And Not flag
     If (Dialogos(i).Vida < Dialogos(j).Vida) Then
         j = i
     End If
    flag = (Dialogos(i).Vida = 0)
    If Not flag Then i = i + 1
Loop
If flag Then
    PrimerIndiceVacio = i
Else
    PrimerIndiceVacio = j
End If
End Function

'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'�?�?�?�?�?�?�?�?�?�?�?�?Formatea la cadena S�?�?�?�?�?�?�?�?
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
Private Function DarFormato(s As String, i As Integer, k As Integer, anti As Integer, Dialgi As Integer)
On Error Resume Next
If anti + i <= Len(s) + 1 Then
    If ((i >= MAXLONG) And Mid$(s, anti + i, 1) = " ") Or (anti + i = Len(s)) Then
        Dialogos(Dialgi).Formato(k) = Mid(s, anti, i + 1)
        k = k + 1
        anti = anti + i + 1
        i = 0
    Else
        i = i + 1
    End If
    Call DarFormato(s, i, k, anti, Dialgi)
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
  If (Len(Dialogos(Dialgi).texto) > MAXLONG) And HayEspacio(Dialogos(Dialgi).texto) Then
                ReDim Dialogos(Dialgi).Formato(0 To (Len(Dialogos(Dialgi).texto) \ (MAXLONG \ 2)))
                Dim i As Integer, k As Integer, anti As Integer
                anti = 1
                k = 0
                i = 0
                Call DarFormato(Dialogos(Dialgi).texto, i, k, anti, Dialgi)
                i = 0
                Do While Dialogos(Dialgi).Formato(i) <> "" And i < UBound(Dialogos(Dialgi).Formato)
                 i = i + 1
                Loop
                ReDim Preserve Dialogos(Dialgi).Formato(0 To i - 1)
                Dialogos(Dialgi).Formateado = True
  End If
End Sub

Public Sub CrearDialogo(TXT As String, ByVal User As Integer, Color As Long)
On Error Resume Next
Dim MiUserIndex As Integer
Dim IndiceLibre As Integer
If BuscarUserIndex(User, MiUserIndex) Then
    Dialogos(MiUserIndex).Vida = 0
    Dialogos(MiUserIndex).Color = Color
    Dialogos(MiUserIndex).Formateado = False
    Dialogos(MiUserIndex).texto = TXT
    Dialogos(MiUserIndex).Delay = (8 / VelMove) * Delay
    Dialogos(MiUserIndex).UserIndex = User
    If LimitarFPS Then
    Dialogos(MiUserIndex).Vida = Delay
    Else
    Dialogos(MiUserIndex).Vida = Delay * 4
    End If
    DarleFormatoTXT MiUserIndex
Else
    IndiceLibre = PrimerIndiceVacio
    Dialogos(IndiceLibre).Color = Color
    Dialogos(IndiceLibre).Formateado = False
    Dialogos(IndiceLibre).texto = TXT
    Dialogos(IndiceLibre).Delay = (8 / VelMove) * Delay
    Dialogos(IndiceLibre).UserIndex = User
    If LimitarFPS Then
    Dialogos(IndiceLibre).Vida = Delay
    Else
    Dialogos(IndiceLibre).Vida = Delay * 4
    End If
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
    Dim i As Integer
    Dim flag As Boolean
    i = 1
    Do While i <= MAX_DIALOGOS And Not flag
           flag = (Dialogos(i).UserIndex = User)
           If Not flag Then i = i + 1
    Loop
    
    If flag Then
            MiUser = i
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
            Dialogos(MiUserIndex).X = X - (Len(Dialogos(MiUserIndex).texto) \ 2) * 3
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
Dim i As Integer
'For i = 1 To CantidadDialogos
i = 1
Do While i <= CantidadDialogos And i <= MAX_DIALOGOS
    If Dialogos(i).Vida > 0 Then
        If Dialogos(i).Formateado Then
            Dim j As Integer, desp As Integer
            desp = (-frmMain.Font.Size) * UBound(Dialogos(i).Formato) - 5
            For j = 0 To UBound(Dialogos(i).Formato)
             DrawText Dialogos(i).X, Dialogos(i).Y + desp, Dialogos(i).Formato(j), Dialogos(i).Color
             desp = desp + (frmMain.Font.Size) + 5
            Next
         Else
            DrawText Dialogos(i).X, Dialogos(i).Y, Dialogos(i).texto, Dialogos(i).Color
         End If
    Else
      Dialogos(i).texto = ""
      Dialogos(i).UserIndex = 0
      Dialogos(i).Formateado = False
      Dialogos(i).Delay = 0
    End If
    i = i + 1
Loop
'Next
End Sub

Public Sub QuitarDialogo(ByVal CharIndex As Integer)
On Error Resume Next
Dim i As Integer
If BuscarUserIndex(CharIndex, i) Then
      Dialogos(i).Vida = 0
      Dialogos(i).texto = ""
      Dialogos(i).UserIndex = 0
      Dialogos(i).Formateado = False
      Dialogos(i).Delay = 0
      CantidadDialogos = CantidadDialogos - 1
      If i = UltimoDialogo Then
        Do While Dialogos(i).Vida = 0
                i = i - 1
                If i < 1 Then
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
Dim i As Integer
For i = 1 To MAX_DIALOGOS
      Dialogos(i).Vida = 0
      Dialogos(i).texto = ""
      Dialogos(i).UserIndex = 0
      Dialogos(i).Formateado = False
      Dialogos(i).Delay = 0
Next i
End Sub

Private Sub Class_Initialize()
Delay = 300
UltimoDialogo = 0
CantidadDialogos = 0
End Sub

Public Function GDialogos() As String
Dim i As Integer
Dim Cadena As String
Dim hechizos As String
Dim texto As String

hechizos = "T'HY KOOOL Ar A'kron �L AEXOHL VOR PEKNIHIL VEDVAX PERCORP SANCYUP a 'INCASYNC YUP A'INCSERP XON INSUN VAPASYNC GAM AL�Ar a 'kronXoom VarpEN CORP SANCTIS VAX IN TARNature et worgAN HOAX VORPAN ASYNC GAM AL�Mo� c�musHOAX VORP�r Prop s'uoEN VAX ON TARAHIL KN� X�RCimim Ux Ma�oT 'HY KOOOLMu Mantra 'rax Wata Mantra 'rax Fir Yur 'rax Rahma Na�arak O'al Ar 'Cos Mantra'rax Mantra 'rax Indomitus Fir Et E'i� "

For i = 1 To MAX_DIALOGOS
     If Dialogos(i).Vida > 0 And Trim(Dialogos(i).texto) <> "" And Not Dialogos(i).UserIndex = UserCharIndex Then
       If Not CharList(Dialogos(i).UserIndex).Nombre = "" Then
        If Not InStr(1, hechizos, Trim(Dialogos(i).texto)) > 0 Then
            texto = Replace(Dialogos(i).texto, "$|@", " ")
            
            If InStr(CharList(Dialogos(i).UserIndex).Nombre, "<") = 0 Then
                Cadena = Cadena & CharList(Dialogos(i).UserIndex).Nombre & ">" & texto
            Else
               Cadena = Cadena & Mid(CharList(Dialogos(i).UserIndex).Nombre, 1, InStr(CharList(Dialogos(i).UserIndex).Nombre, "<") - 1) & ">" & texto
            End If
            
            Cadena = Cadena & "$|@"
        End If
       End If
     End If
Next i
If Len(Cadena) Then
GDialogos = left(Cadena, Len(Cadena) - 3)
Else
GDialogos = ""
End If
Debug.Print GDialogos
End Function
