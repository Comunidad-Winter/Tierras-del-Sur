Attribute VB_Name = "Engine_Renderer"
Option Explicit

Private tBox As Box_Vertex


'#################### BUFFFFEEEEEERRRRRRRRRR ####################

Private Const batch_max                     As Long = (INDEX_BUFFER_SIZE / 4)
Private batch_buffer(0 To batch_max * TL_size) As Byte
Private batch_count                         As Long
Private batch_texture                       As Integer
Private batch_blend_mode                    As Long
Private batch_triangulos                    As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private pIB As Direct3DIndexBuffer8

'Dibuja y "vacia" el buffer de poligonos
Public Sub batch_render()
    If batch_count Then
        Set_Blend_Mode batch_blend_mode
        Call GetTexture(batch_texture)
        'D3DDevice.DrawIndexedPrimitiveUP D3DPT_TRIANGLELIST, 0, Vertices, Triangulos, StaticIndexBuffer(0), D3DFMT_INDEX16, vOut(0), TL_size
        'D3DDevice.DrawIndexedPrimitiveUP D3DPT_TRIANGLESTRIP, 0, batch_count, batch_triangulos, StaticIndexBuffer(0), D3DFMT_INDEX16, batch_buffer(0), TL_size
        Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, batch_triangulos, batch_buffer(0), TL_size)
        batch_count = 0
        batch_triangulos = 0
    End If
End Sub

'PUSH en el buffer de poligonos
Public Sub batch_add(ByRef vertex As TLVERTEX, ByVal cantidad As Long, ByVal texture As Integer, ByVal blend As Byte)
    If (batch_count >= (batch_max - 1) Or batch_texture <> texture Or batch_blend_mode <> blend) Then
        batch_render
        batch_blend_mode = blend
        batch_texture = texture
    End If
    Call CopyMemory(batch_buffer(batch_count * TL_size), vertex, cantidad * TL_size)
    batch_count = batch_count + cantidad
    batch_triangulos = batch_triangulos + cantidad / 2
End Sub

'PUSH en el buffer de poligonos
Public Sub batch_add_box(ByRef vertex As Box_Vertex, ByVal texture As Integer, ByVal blend As Byte)
    If (batch_count >= (batch_max - 1) Or batch_texture <> texture Or batch_blend_mode <> blend) Then
        batch_render
        batch_blend_mode = blend
        batch_texture = texture
    End If
    Call CopyMemory(batch_buffer(batch_count * TL_size), vertex, BV_size)
    batch_count = batch_count + 4
    batch_triangulos = batch_triangulos + 2
End Sub

Public Sub batch_init(ByRef max As Integer)
'
'    Dim j As Long, i As Long
'    Dim indices() As Integer
'    Dim indxbuffsize As Long
'
'    'pVB = D3DDevice.CreateVertexBuffer(max * TL_size, D3DUSAGE_WRITEONLY, FVF, D3DPOOL_DEFAULT)
'
'    Set pIB = D3DDevice.CreateIndexBuffer(max * 16, D3DUSAGE_WRITEONLY, D3DFMT_INDEX16, D3DPOOL_DEFAULT)
'
'    ReDim indices(4 * max)
'    j = 0
'    For i = 0 To max - 1
'        indices(j) = 4 * i + 0: j = j + 1
'        indices(j) = 4 * i + 1: j = j + 1
'        indices(j) = 4 * i + 2: j = j + 1
'        indices(j) = 4 * i + 3: j = j + 1
'    Next
'
'    ' Set the data on the d3d buffer
'    D3DIndexBuffer8SetData pIB, 0, indxbuffsize, 0, indices(0)
End Sub

'#################### /BUFFFFEEEEEERRRRRRRRRR ####################


