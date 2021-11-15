Attribute VB_Name = "modLecturaSecuencial"
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
'Para la lectura secuencial de una cadena

Option Explicit

Private Type tlecCadena
    lecCadena As String
    lecPosicion As Integer
End Type

Private lecLecturaSecuencial() As tlecCadena, lecIniciado As Boolean

Public Function lecCambiarCadena(cadena As String) As Integer
    If Not lecIniciado Then ReDim lecLecturaSecuencial(0): lecIniciado = True
    
    lecCambiarCadena = lecCadenaLibre
    
    lecLecturaSecuencial(lecCambiarCadena).lecCadena = cadena
    lecLecturaSecuencial(lecCambiarCadena).lecPosicion = 1
End Function

Public Function lecLeer(cadena As Integer, Optional initpos As Integer, Optional longitud As Integer) As String
    If Not lecIniciado Then ReDim lecLecturaSecuencial(0): lecIniciado = True
    
    If cadena > UBound(lecLecturaSecuencial) Or lecLecturaSecuencial(cadena).lecCadena = "" Then lecLeer = -1: Exit Function
    If initpos > 0 And initpos < Len(lecLecturaSecuencial(cadena).lecCadena) Then lecLecturaSecuencial(cadena).lecPosicion = initpos
    If lecLecturaSecuencial(cadena).lecPosicion + longitud = Len(lecLecturaSecuencial(cadena).lecCadena) + 2 Then lecLeer = -1: Exit Function
    If longitud = 0 Then longitud = Len(lecLecturaSecuencial(cadena).lecCadena) - lecLecturaSecuencial(cadena).lecPosicion
    
    lecLeer = Mid$(lecLecturaSecuencial(cadena).lecCadena, lecLecturaSecuencial(cadena).lecPosicion, longitud)
    lecLecturaSecuencial(cadena).lecPosicion = lecLecturaSecuencial(cadena).lecPosicion + longitud
End Function

Public Function lecCerrarCadena(cadena As Integer) As Boolean
    If Not lecIniciado Then ReDim lecLecturaSecuencial(0): lecIniciado = True
    
    If cadena > UBound(lecLecturaSecuencial) Then lecCerrarCadena = False: Exit Function
    If cadena = UBound(lecLecturaSecuencial) Then ReDim Preserve lecLecturaSecuencial(UBound(lecLecturaSecuencial) - 1): Exit Function
    lecLecturaSecuencial(cadena).lecCadena = ""
    lecLecturaSecuencial(cadena).lecPosicion = 0
End Function

Private Function lecCadenaLibre() As Integer
    Dim i As Integer
    For i = 1 To UBound(lecLecturaSecuencial)
        If lecLecturaSecuencial(i).lecCadena = "" Then
            lecCadenaLibre = i
            Exit Function
        End If
    Next i
    lecCadenaLibre = UBound(lecLecturaSecuencial) + 1
    ReDim Preserve lecLecturaSecuencial(lecCadenaLibre)
End Function


