Attribute VB_Name = "modHexaStrings"
'Argentum Online 0.9.0.4
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
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
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

'Modulo realizado por Gonzalo Larralde(CDT) <gonzalolarralde@yahoo.com.ar>
'Para la conversion a caracteres de cadenas MD5 y de
'semi encriptaci�n de cadenas por ascii table offset

Option Explicit

Public Function hexMd52Asc(ByVal md5 As String) As String
    Dim I As Integer, l As String
    md5 = UCase$(md5)
    If Len(md5) Mod 2 = 1 Then md5 = "0" & md5
    For I = 1 To Len(md5) \ 2
        l = Mid$(md5, (2 * I) - 1, 2)
        hexMd52Asc = hexMd52Asc & Chr$(hexHex2Dec(l))
    Next I
End Function

Public Function hexHex2Dec(ByVal hex As String) As Long
    Dim I As Integer, l As String
    For I = 1 To Len(hex)
        l = Mid$(hex, I, 1)
        Select Case l
            Case "A": l = 10
            Case "B": l = 11
            Case "C": l = 12
            Case "D": l = 13
            Case "E": l = 14
            Case "F": l = 15
        End Select
        
        hexHex2Dec = (l * 16 ^ ((Len(hex) - I))) + hexHex2Dec
    Next I
End Function

Public Function txtOffset(ByVal Text As String, ByVal off As Integer) As String
    Dim I As Integer, l As String
    For I = 1 To Len(Text)
        l = Mid$(Text, I, 1)
        txtOffset = txtOffset & Chr$((Asc(l) + off) Mod 256)
    Next I
End Function
