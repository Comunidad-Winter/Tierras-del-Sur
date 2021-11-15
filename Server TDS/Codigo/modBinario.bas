Attribute VB_Name = "modBinario"
'Argentum Online 0.9.0.4
'
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

'Modulo realizado por Gonzalo Larralde(CDT) <gonzalolarralde@yahoo.com.ar>
'con ayuda de Alejandro Santos(AlejoLP)
'Para la conversion a caracteres de números

Option Explicit

Public Function binCBytes(ByVal numero As Long) As Integer
    If numero <= 255 Then
        binCBytes = 1
    ElseIf numero <= 65534 Then
        binCBytes = 2
    ElseIf numero <= 2147483647 Then
        binCBytes = 4
    End If
End Function

Public Function binBin2Dec(ByVal strnum As String) As Long
    Dim i As Integer
    For i = 1 To Len(strnum)
        binBin2Dec = CStr(CByte(Mid(strnum, i, 1)) * 2 ^ (Len(strnum) - i)) + binBin2Dec
    Next i
End Function

Public Function binDec2Bin(ByVal num As Long) As String
    Dim i As Integer: i = binCBytes(num)
    Do
        binDec2Bin = num Mod 2 & binDec2Bin
        num = num \ 2
    Loop While Not num = 0
    
    'Relleno de 0
    For i = 0 To (i * 8) - 1 - Len(binDec2Bin)
        binDec2Bin = "0" & binDec2Bin
    Next i
End Function

Public Function binBin2Asc(ByVal binnum As String) As String
    Dim cbytes As Integer, i As Integer
    cbytes = Len(binnum) / 8
    For i = 0 To cbytes - 1
        binBin2Asc = binBin2Asc & Chr(binBin2Dec(Mid(binnum, i * 8 + 1, 8)))
    Next i
End Function

Public Function binAsc2Bin(ByVal ascnum As String) As String
    Dim i As Integer
    For i = 1 To Len(ascnum)
        binAsc2Bin = binAsc2Bin & binDec2Bin(Asc(Mid(ascnum, i, 1)))
    Next i
End Function

Public Function binAsc2Dec(ByVal ascnum As String) As Long
    binAsc2Dec = binBin2Dec(binAsc2Bin(ascnum))
End Function

Public Function binDec2Asc(ByVal numero As Long, Optional tipo As Integer) As String
    binDec2Asc = binBin2Asc(binDec2Bin(numero))
    
    Dim i As Integer
    Select Case tipo
        Case vbInteger
            If Len(binDec2Asc) < 2 Then For i = 1 To 2 - Len(binDec2Asc): binDec2Asc = Chr(0) & binDec2Asc: Next i
        Case vbLong
            If Len(binDec2Asc) < 4 Then For i = 1 To 4 - Len(binDec2Asc): binDec2Asc = Chr(0) & binDec2Asc: Next i
    End Select
End Function

