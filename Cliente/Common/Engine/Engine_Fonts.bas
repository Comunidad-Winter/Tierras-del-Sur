Attribute VB_Name = "Engine_Fonts"
Option Explicit


Private Type CharVA
    vertex As Box_Vertex
End Type

Private Type VFH
    BitmapWidth As Long         'Size of the bitmap itself
    BitmapHeight As Long
    CellWidth As Long           'Size of the cells (area for each character)
    CellHeight As Long
    BaseCharOffset As Byte      'The character we start from
    CharWidth(0 To 255) As Byte 'The actual factual width of each character
    CharVA(0 To 255) As CharVA
End Type

Private Type CustomFont
    HeaderInfo As VFH           'Holds the header information
    texture As Direct3DTexture8 'Holds the texture of the text
    RowPitch As Integer         'Number of characters per row
    RowFactor As Single         'Percentage of the texture width each character takes
    ColFactor As Single         'Percentage of the texture height each character takes
    CharHeight As Byte          'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
End Type


Public Font_Default As CustomFont   'Describes our custom font "default"

Public Type tFont
    texture As Integer
    Font As CustomFont
End Type

Public Fonts(1 To 4) As tFont

Sub Engine_Init_FontSettings()
   Call LoadFont("\f97181.dat", Font_Default)
   
   Call LoadFont("\f97181.dat", Fonts(1).Font)
   Fonts(1).texture = 1125
   
   Call LoadFont("\f9718.dat", Fonts(2).Font)
   Fonts(2).texture = 3404
   
   Call LoadFont("\p.dat", Fonts(3).Font)
   Fonts(3).texture = 3405
   
   Call LoadFont("\h1.dat", Fonts(4).Font)
   Fonts(4).texture = 3406
End Sub

Public Function FontGetTextWidth(ByVal text As String, ByVal Font As Byte) As Integer
    Dim i As Integer
    If LenB(text) = 0 Then Exit Function
    For i = 1 To Len(text)
        FontGetTextWidth = FontGetTextWidth + Fonts(Font).Font.HeaderInfo.CharWidth(Asc(mid$(text, i, 1)))
    Next i
End Function

Private Sub LoadFont(file As String, dst As CustomFont)
Dim filenum As Byte
Dim LoopChar As Long
Dim Row As Single
Dim u As Single
Dim v As Single

    '*** Default font ***

    'Load the header information
    filenum = FreeFile
    Open IniPath & file For Binary As #filenum
        Get #filenum, , dst.HeaderInfo
    Close #filenum
    
    'Calculate some common values
    dst.CharHeight = dst.HeaderInfo.CellHeight - 4
    dst.RowPitch = dst.HeaderInfo.BitmapWidth \ dst.HeaderInfo.CellWidth
    dst.ColFactor = dst.HeaderInfo.CellWidth / dst.HeaderInfo.BitmapWidth
    dst.RowFactor = dst.HeaderInfo.CellHeight / dst.HeaderInfo.BitmapHeight
    
    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
        
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - dst.HeaderInfo.BaseCharOffset) \ dst.RowPitch
        u = ((LoopChar - dst.HeaderInfo.BaseCharOffset) - (Row * dst.RowPitch)) * dst.ColFactor
        v = Row * dst.RowFactor

        'Set the verticies
        With dst.HeaderInfo.CharVA(LoopChar)
            .vertex.color0 = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .vertex.rhw0 = 1
            .vertex.tu0 = u
            .vertex.tv0 = v
            .vertex.x0 = 0
            .vertex.y0 = 0
            .vertex.Z0 = 0
            
            .vertex.Color1 = D3DColorARGB(255, 0, 0, 0)
            .vertex.rhw1 = 1
            .vertex.tu1 = u + dst.ColFactor
            .vertex.tv1 = v
            .vertex.x1 = dst.HeaderInfo.CellWidth
            .vertex.y1 = 0
            .vertex.Z1 = 0
            
            .vertex.Color2 = D3DColorARGB(255, 0, 0, 0)
            .vertex.rhw2 = 1
            .vertex.tu2 = u
            .vertex.tv2 = v + dst.RowFactor
            .vertex.x2 = 0
            .vertex.y2 = dst.HeaderInfo.CellHeight
            .vertex.z2 = 0
            
            .vertex.color3 = D3DColorARGB(255, 0, 0, 0)
            .vertex.rhw3 = 1
            .vertex.tu3 = u + dst.ColFactor
            .vertex.tv3 = v + dst.RowFactor
            .vertex.x3 = dst.HeaderInfo.CellWidth
            .vertex.y3 = dst.HeaderInfo.CellHeight
            .vertex.Z3 = 0
        End With
        
    Next LoopChar
End Sub
