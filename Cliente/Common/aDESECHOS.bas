Attribute VB_Name = "aDESECHOS"
'Public Sub amigar_colores(Optional ByVal SX As Byte = 1, Optional ByVal SY As Byte = 1, Optional ByVal EX As Byte = 101, Optional ByVal EY As Byte = 101)
'Dim x As Single, y As Single
'Dim suma%
'    For y = SY + 1 To EY - 1
'        For x = SX + 1 To EX - 1
'            suma = cColorData(x - 1, y - 1).Color(2).r
'            suma = suma + cColorData(x, y - 1).Color(0).r
'            suma = suma + cColorData(x, y).Color(1).r
'            suma = suma + cColorData(x - 1, y).Color(3).r
'            suma = (suma / 4) '* hLightData(X - 1, y - 1,2)
'            cColorData(x - 1, y - 1).Color(2).r = suma
'            cColorData(x, y - 1).Color(0).r = suma
'            cColorData(x, y).Color(1).r = suma
'            cColorData(x - 1, y).Color(3).r = suma
'
'            suma = cColorData(x - 1, y - 1).Color(2).g
'            suma = suma + cColorData(x, y - 1).Color(0).g
'            suma = suma + cColorData(x, y).Color(1).g
'            suma = suma + cColorData(x - 1, y).Color(3).g
'            suma = (suma / 4) '* hLightData(X - 1, y - 1,2)
'            cColorData(x - 1, y - 1).Color(2).g = suma
'            cColorData(x, y - 1).Color(0).g = suma
'            cColorData(x, y).Color(1).g = suma
'            cColorData(x - 1, y).Color(3).g = suma
'
'            suma = cColorData(x - 1, y - 1).Color(2).b
'            suma = suma + cColorData(x, y - 1).Color(0).b
'            suma = suma + cColorData(x, y).Color(1).b
'            suma = suma + cColorData(x - 1, y).Color(3).b
'            suma = (suma / 4) '* hLightData(X - 1, y - 1,2)
'            cColorData(x - 1, y - 1).Color(2).b = suma
'            cColorData(x, y - 1).Color(0).b = suma
'            cColorData(x, y).Color(1).b = suma
'            cColorData(x - 1, y).Color(3).b = suma
'        Next x
'    Next y
'End Sub

'Public Sub map_render_light()
'On Error GoTo enda:
'    Dim y                       As Integer
'    Dim x                       As Integer
'    Dim alpha                   As Single
'    Dim cr As Integer, cb As Integer, cg As Integer
'    Dim i As Long
'    If minX = 0 Then Exit Sub
'
'    For y = minY To maxY
'        For x = minX To maxX
'            With MapData(x, y)
'                i = 0
'                If .last_light(i) <> Engine_Landscape.last_light_calculate Then
'                    cr = bcMapData(x, y).c(i).r + (1 - Intensidad_Del_Terreno_Float(x, y).f(i)) * cColorData(x, y).Color(i).r * color_mod_day.r
'                    cg = bcMapData(x, y).c(i).g + (1 - Intensidad_Del_Terreno_Float(x, y).f(i)) * cColorData(x, y).Color(i).g * color_mod_day.g
'                    cb = bcMapData(x, y).c(i).b + (1 - Intensidad_Del_Terreno_Float(x, y).f(i)) * cColorData(x, y).Color(i).b * color_mod_day.b
'                    .light_value(i) = D3DColorXRGB(cr, cg, cb)
'                    .last_light(i) = Engine_Landscape.last_light_calculate
'
'                    MapData(x, y + 1).light_value(1) = .light_value(i)
'                    MapData(x - 1, y + 1).light_value(3) = .light_value(i)
'                    MapData(x - 1, y).light_value(2) = .light_value(i)
'
'                    MapData(x, y + 1).last_light(1) = .last_light(i)
'                    MapData(x - 1, y + 1).last_light(3) = .last_light(i)
'                    MapData(x - 1, y).last_light(2) = .last_light(i)
'                End If
'
''                i = 1
''                If .last_light(i) <> Engine_Landscape.last_light_calculate Then
''                    cr = bcMapData(x, y).c(i).r + (1 - Intensidad_Del_Terreno_Float(x, y).f(i)) * cColorData(x, y).Color(i).r * color_mod_day.r
''                    cg = bcMapData(x, y).c(i).g + (1 - Intensidad_Del_Terreno_Float(x, y).f(i)) * cColorData(x, y).Color(i).g * color_mod_day.g
''                    cb = bcMapData(x, y).c(i).b + (1 - Intensidad_Del_Terreno_Float(x, y).f(i)) * cColorData(x, y).Color(i).b * color_mod_day.b
''
''                    .light_value(i) = D3DColorXRGB(cr, cg, cb)
''                    .last_light(i) = Engine_Landscape.last_light_calculate
''
''                    MapData(x, y - 1).light_value(0) = .light_value(i)
''                    MapData(x - 1, y - 1).light_value(2) = .light_value(i)
''                    MapData(x - 1, y).light_value(3) = .light_value(i)
''
''                    MapData(x, y - 1).last_light(0) = .last_light(i)
''                    MapData(x - 1, y - 1).last_light(2) = .last_light(i)
''                    MapData(x - 1, y).last_light(3) = .last_light(i)
''                End If
''
''                i = 2
''                If .last_light(i) <> Engine_Landscape.last_light_calculate Then
''                    cr = bcMapData(x, y).c(i).r + (1 - Intensidad_Del_Terreno_Float(x, y).f(i)) * cColorData(x, y).Color(i).r * color_mod_day.r
''                    cg = bcMapData(x, y).c(i).g + (1 - Intensidad_Del_Terreno_Float(x, y).f(i)) * cColorData(x, y).Color(i).g * color_mod_day.g
''                    cb = bcMapData(x, y).c(i).b + (1 - Intensidad_Del_Terreno_Float(x, y).f(i)) * cColorData(x, y).Color(i).b * color_mod_day.b
''                    .light_value(i) = D3DColorXRGB(cr, cg, cb)
''                    .last_light(i) = Engine_Landscape.last_light_calculate
''
''                    MapData(x, y + 1).light_value(3) = .light_value(i)
''                    MapData(x + 1, y + 1).light_value(1) = .light_value(i)
''                    MapData(x + 1, y).light_value(0) = .light_value(i)
''
''                    MapData(x, y + 1).last_light(3) = .last_light(i)
''                    MapData(x + 1, y + 1).last_light(1) = .last_light(i)
''                    MapData(x + 1, y).last_light(0) = .last_light(i)
''                End If
''
''                i = 3
''                If .last_light(i) <> Engine_Landscape.last_light_calculate Then
''                    cr = bcMapData(x, y).c(i).r + (1 - Intensidad_Del_Terreno_Float(x, y).f(i)) * cColorData(x, y).Color(i).r * color_mod_day.r
''                    cg = bcMapData(x, y).c(i).g + (1 - Intensidad_Del_Terreno_Float(x, y).f(i)) * cColorData(x, y).Color(i).g * color_mod_day.g
''                    cb = bcMapData(x, y).c(i).b + (1 - Intensidad_Del_Terreno_Float(x, y).f(i)) * cColorData(x, y).Color(i).b * color_mod_day.b
''                    .light_value(i) = D3DColorXRGB(cr, cg, cb)
''                    .last_light(i) = Engine_Landscape.last_light_calculate
''
''                    MapData(x, y - 1).light_value(2) = .light_value(i)
''                    MapData(x + 1, y - 1).light_value(0) = .light_value(i)
''                    MapData(x + 1, y).light_value(1) = .light_value(i)
''
''                    MapData(x, y - 1).last_light(2) = .last_light(i)
''                    MapData(x + 1, y - 1).last_light(0) = .last_light(i)
''                    MapData(x + 1, y).last_light(1) = .last_light(i)
''                End If
'            End With
'        Next x
'    Next y
'
'Exit Sub
'enda:
'If Err.Number = 10 Then
'LogError "ERROR 10 EN mrl"
'End If
'End Sub

'
'Private Function Interpolate(a As Single, B As Single, V As Single)
'Interpolate = a + (V * (B - a))
'End Function

'Private Sub text_render_graphic(t$, X!, Y!, Optional ByVal color = &HFFFFFFFF)
'    'Dim i As Integer
'    Dim char As Byte
'    Dim lenght&
'    lenght = Len(t)
'    If lenght = 0 Then Exit Sub
'
'    Dim th!, tw!, tX!, tY!, left!, W!, H!
'    Dim ind&, TempStr$()
'    Dim TLV() As TLVERTEX
'
'    ReDim TLV((Len(t) * 4) - 1)
'
'    Call GetTexture(9718) '//tehoma shadow
'    Call GetTextureDimension(9718, th, tw)
'    W = tw / 16
'    H = th / 16
'    left = Round(X)
'    Y = Round(Y)
''    For i = 1 To lenght
''        char = AscB(mid$(t, i, 1))
''
''        ind = (i - 1) * 4
''
'''            tX = (char Mod W) * 16 '(tw / 16))
'''            tY = (char \ H) * 16 '(th / 16))
'''            TLV(ind) = Geometry_Create_TLVertex(left, Y + H, color, tX / tw, (H + tY) / th)
'''            TLV(ind + 1) = Geometry_Create_TLVertex(left, Y, color, tX / tw, tY / th)
'''            left = left + W
'''            TLV(ind + 2) = Geometry_Create_TLVertex(left, Y + H, color, (W + tX) / tw, (H + tY) / th)
'''            TLV(ind + 3) = Geometry_Create_TLVertex(left, Y, color, (W + tX) / tw, tY / th)
'''
''        With arr_font_255_uv(char)
''            TLV(ind) = Geometry_Create_TLVertex(left, Y + H, color, .X, .W)
''            TLV(ind + 1) = Geometry_Create_TLVertex(left, Y, color, .X, .Y)
''
''            TLV(ind + 2) = Geometry_Create_TLVertex(left + W, Y + H, color, .Z, .W)
''            TLV(ind + 3) = Geometry_Create_TLVertex(left + W, Y, color, .Z, .Y)
''            left = left + 9
''        End With
''    Next i
'
'
'Dim Count As Integer
'Dim Ascii() As Byte
'Dim Row As Integer
'Dim u As Single
'Dim v As Single
'Dim i As Long
'Dim j As Long
'Dim KeyPhrase As Byte
'Dim TempColor As Long
'Dim ResetColor As Byte
'Dim SrcRect As RECT
'Dim v2 As D3DVECTOR2
'Dim v3 As D3DVECTOR2
'Dim YOffset As Single
'
'TempColor = color
'        TempStr = Split(t, vbCrLf)
'        For i = 0 To UBound(TempStr)
'        If Len(TempStr(i)) > 0 Then
'            YOffset = i * Font_Default.CharHeight
'            Count = 0
'
'            'Convert the characters to the ascii value
'            Ascii() = StrConv(TempStr(i), vbFromUnicode)
'
'            'Loop through the characters
'            For j = 1 To Len(TempStr(i))
'
'                'Check for a key phrase
'                If Ascii(j - 1) = 124 Then 'If Ascii = "|"
'                    KeyPhrase = (Not KeyPhrase)  'TempColor = ARGB 255/255/0/0
'                    If KeyPhrase Then TempColor = -65536 Else ResetColor = 1
'                Else
'
'
'                        'Copy from the cached vertex array to the temp vertex array
'                        CopyMemory TLV(ind), Font_Default.HeaderInfo.CharVA(Ascii(j - 1)).Vertex(0), TL_size * 4
'
'                        'Set up the verticies
'                        TLV(ind).v.X = X + Count
'                        TLV(ind).v.Y = Y + YOffset
'
'                        TLV(ind + 1).v.X = TLV(ind + 1).v.X + X + Count
'                        TLV(ind + 1).v.Y = TLV(ind).v.Y
'
'                        TLV(ind + 2).v.X = TLV(ind).v.X
'                        TLV(ind + 2).v.Y = TLV(ind + 2).v.Y + TLV(ind).v.Y
'
'                        TLV(ind + 3).v.X = TLV(ind + 1).v.X
'                        TLV(ind + 3).v.Y = TLV(ind + 2).v.Y
'
'                        'Set the colors
'                        TLV(ind).color = TempColor
'                        TLV(ind + 1).color = TempColor
'                        TLV(ind + 2).color = TempColor
'                        TLV(ind + 3).color = TempColor
'                        ind = ind + 4
'                    'Shift over the the position to render the next character
'                    Count = Count + Font_Default.HeaderInfo.CharWidth(Ascii(j - 1))
'
'                End If
'
'                'Check to reset the color
'                If ResetColor Then
'                    ResetColor = 0
'                    TempColor = color
'                End If
'
'            Next j
'
'        End If
'    Next i
'
'    Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_MODULATE)
'    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, ind - 3, TLV(0), TL_size
'    Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, lColorMod)
'End Sub

'Option Explicit
'
'Private Declare Sub CopyMemory Lib "kernel32" _
'    Alias "RtlMoveMemory" (Destination As Any, _
'    Source As Any, ByVal Length As Long)
'Private Declare Function GetProcessHeap Lib "kernel32" () As Long
'Private Declare Function HeapAlloc Lib "kernel32" _
'    (ByVal hHeap As Long, ByVal dwFlags As Long, _
'     ByVal dwBytes As Long) As Long
'Private Declare Function HeapFree Lib "kernel32" _
'    (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
'Private Declare Sub CopyMemoryWrite Lib "kernel32" Alias _
'    "RtlMoveMemory" (ByVal Destination As Long, _
'    Source As Any, ByVal Length As Long)
'Private Declare Sub CopyMemoryRead Lib "kernel32" Alias _
'    "RtlMoveMemory" (Destination As Any, _
'    ByVal Source As Long, ByVal Length As Long)
'
'Private Sub modmem()
'    Dim ptr As Long   'int * ptr;
'
'    Dim hHeap As Long
'    hHeap = GetProcessHeap()
'    ptr = HeapAlloc(hHeap, 0, 2) 'an integer in Visual Basic is 2 bytes
'
'    If ptr <> 0 Then
'    'memory was allocated
'
'    'do stuff
'
'        Dim i As Integer
'        i = 10
'        CopyMemoryWrite ptr, i, 2 ' an intger is two bytes
'
'        Dim j As Integer
'        CopyMemoryRead j, ptr, 2
'        MsgBox "The adress of ptr is " & CStr(ptr) & _
'            vbCrLf & "and the value is " & CStr(j)
'        HeapFree GetProcessHeap(), 0, ptr
'    End If
'End Sub
'


'Public Function txtOffset(ByVal Text As String, ByVal off As Integer) As String
'    Dim i As Integer, l As String
'    For i = 1 To Len(Text)
'        l = mid(Text, i, 1)
'        txtOffset = txtOffset & Chr((Asc(l) + off) Mod 256)
'    Next i
'End Function
'
'Public Sub Grh_Render_relieve(ByVal GrhIndex As Long, ByVal tLeft As Single, ByVal tTop As Single, ByVal map_x As Byte, ByVal map_y As Byte, ByVal flip As Byte)
''*********************************************
''Author: menduz
''*********************************************
'    Dim tBottom!, tRight! ', tTop!, tLeft!
'    Dim ll As Long
'
'    If GrhIndex = 0 Then Exit Sub
'    Call GetTexture(GrhData(GrhIndex).FileNum) '
'
'    With GrhData(GrhIndex)
'        tBottom = tTop + .pixelHeight
'        tRight = tLeft + .pixelWidth
'
'
'        ll = 0 '&H7F7F7F7F
'
'        If .hardcor = 0 Then Init_grh_tutv GrhIndex
'        'If frmMain.caca.value Then flip = Not flip 'flip = 0
'        If flip Then
'        '01
'        '23
'
''            Call cTLVertex(temp_verts(0), dest_rect.Left, dest_rect.Top - hMapData(map_x, map_y).plus(1), MapData(map_x, map_y).light_value(1) Or ll, .tu(1), .tv(1))
''            Call cTLVertex(temp_verts(1), dest_rect.Right, dest_rect.Top - hMapData(map_x, map_y).plus(3), MapData(map_x, map_y).light_value(3) Or ll, .tu(3), .tv(3))
''            Call cTLVertex(temp_verts(2), dest_rect.Left, dest_rect.Bottom - hMapData(map_x, map_y).plus(0), MapData(map_x, map_y).light_value(0) Or ll, .tu(0), .tv(0))
''            Call cTLVertex(temp_verts(3), dest_rect.Right, dest_rect.Bottom - hMapData(map_x, map_y).plus(2), MapData(map_x, map_y).light_value(2) Or ll, .tu(2), .tv(2))
'
'            tVerts(0).v.X = tLeft
'            tVerts(0).v.Y = tTop - hMapData(map_x, map_y).plus(1)
'            tVerts(0).color = MapData(map_x, map_y).light_value(1) Or ll
'            tVerts(0).tu = .tu(1)
'            tVerts(0).tv = .tv(1)
'
'            tVerts(1).v.X = tRight
'            tVerts(1).v.Y = tTop - hMapData(map_x, map_y).plus(3)
'            tVerts(1).color = MapData(map_x, map_y).light_value(3) Or ll
'            tVerts(1).tu = .tu(3)
'            tVerts(1).tv = .tv(3)
'
'            tVerts(2).v.X = tLeft
'            tVerts(2).v.Y = tBottom - hMapData(map_x, map_y).plus(0)
'            tVerts(2).color = MapData(map_x, map_y).light_value(0) Or ll
'            tVerts(2).tu = .tu(0)
'            tVerts(2).tv = .tv(0)
'
'            tVerts(3).v.X = tRight
'            tVerts(3).v.Y = tBottom - hMapData(map_x, map_y).plus(2)
'            tVerts(3).color = MapData(map_x, map_y).light_value(0) Or ll
'            tVerts(3).tu = .tu(2)
'            tVerts(3).tv = .tv(2)
'        Else
'        '13
'        '02
'
''            Call cTLVertex(temp_verts(0), dest_rect.Left, dest_rect.Bottom - hMapData(map_x, map_y).plus(0), MapData(map_x, map_y).light_value(0) Or ll, .tu(0), .tv(0))
''            Call cTLVertex(temp_verts(1), dest_rect.Left, dest_rect.Top - hMapData(map_x, map_y).plus(1), MapData(map_x, map_y).light_value(1) Or ll, .tu(1), .tv(1))
''            Call cTLVertex(temp_verts(2), dest_rect.Right, dest_rect.Bottom - hMapData(map_x, map_y).plus(2), MapData(map_x, map_y).light_value(2) Or ll, .tu(2), .tv(2))
''            Call cTLVertex(temp_verts(3), dest_rect.Right, dest_rect.Top - hMapData(map_x, map_y).plus(3), MapData(map_x, map_y).light_value(3) Or ll, .tu(3), .tv(3))
'            tVerts(0).v.X = tLeft
'            tVerts(0).v.Y = tBottom - hMapData(map_x, map_y).plus(0)
'            tVerts(0).color = MapData(map_x, map_y).light_value(0) Or ll
'            tVerts(0).tu = .tu(0)
'            tVerts(0).tv = .tv(0)
'
'            tVerts(1).v.X = tLeft
'            tVerts(1).v.Y = tTop - hMapData(map_x, map_y).plus(1)
'            tVerts(1).color = MapData(map_x, map_y).light_value(1) Or ll
'            tVerts(1).tu = .tu(1)
'            tVerts(1).tv = .tv(1)
'
'            tVerts(2).v.X = tRight
'            tVerts(2).v.Y = tBottom - hMapData(map_x, map_y).plus(2)
'            tVerts(2).color = MapData(map_x, map_y).light_value(2) Or ll
'            tVerts(2).tu = .tu(2)
'            tVerts(2).tv = .tv(2)
'
'            tVerts(3).v.X = tRight
'            tVerts(3).v.Y = tTop - hMapData(map_x, map_y).plus(3)
'            tVerts(3).color = MapData(map_x, map_y).light_value(3) Or ll
'            tVerts(3).tu = .tu(3)
'            tVerts(3).tv = .tv(3)
'        End If
'    End With
''D3DDevice.SetTransform D3DPT_TRIANGLESTRIP, 2, tVerts(0), TL_size
'    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tVerts(0), TL_size
'End Sub
'
