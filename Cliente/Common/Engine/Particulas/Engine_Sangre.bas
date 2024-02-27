Attribute VB_Name = "Engine_Sangre"
Option Explicit

Private Const Gravedad As Single = 4
Private Const Sangre_grh As Integer = 1126 '263
Private Const Xvar As Single = 2 'Variación X cuando nace la partícula
Private Const Yvar As Single = 3.2 'Variación Y cuando nace la partícula
Private Const sangre_agrandamiento As Single = 0.15 'Variación Y cuando nace la partícula
Private Const Vida_gota As Integer = 900 'cada gota vive medio segundo antes de renacer

Public Type sangre_particula
    viva As Boolean
    MuereEnTick As Long
    NaceEnTick As Long
    ModificadorY As Integer 'altura en la que cae
    ModificadorY2 As Integer 'altura en la que cae
    x As Single
    y As Single
    vX As Single
    vY As Single
    Tamaño As Single
    alpha As Single
End Type

Private Type Sangres_fx
    viva As Boolean

    Start As Long
    end As Long
    lenght As Long
    
    src_char As Integer
    
    src_x_map As Integer
    src_y_map As Integer

    fuente_x As Integer
    fuente_y As Integer

    cantidad As Integer
    Particulas() As sangre_particula
    
    wind As Byte 'N=0 E=1 S=2 O=3
    
    last_act As Currency
    
    Altura As Byte
    
    col As D3DCOLORVALUE
End Type

Private Sangre_Libre As Integer

Private Sangres() As Sangres_fx
Private Ultima_Sangre As Integer

Public Sub Initialize_Sangre()
    ReDim Sangres(0)
    ReDim Sangres(0).Particulas(0)
    Ultima_Sangre = 0
    Sangre_Libre = -1
End Sub

Public Sub Sangre_Crear(Char As Integer, cantidad As Integer, Duracion As Long, Altura As Byte)
    Dim tick As Long
    Dim actual As Integer
    tick = GetTimer

    actual = Sangre_Obtener_Libre
    
    With Sangres(actual)
        .viva = True
        .src_char = Char
        .src_x_map = CharList(Char).Pos.x
        .src_y_map = CharList(Char).Pos.y
        .Start = tick
        .end = tick + Duracion
        .lenght = Duracion
        .cantidad = cantidad
        .Altura = Altura
        .fuente_x = CharList(Char).Pos.x
        .fuente_y = CharList(Char).Pos.y
        
        If .src_x_map > 0 Then
            Call Long2RGB(ResultColorArray(.src_x_map, .src_y_map), .col.r, .col.g, .col.b)
        Else
            .col.r = 1
            .col.g = 0
            .col.b = 0
        End If
        
        If UBound(.Particulas()) <> cantidad Then
            ReDim .Particulas(cantidad) As sangre_particula
        End If
 
        Call Sangre_Init_ID(actual)
    End With
End Sub

Private Function Sangre_Obtener_Libre() As Integer
    Dim i As Integer
    
    Sangre_Obtener_Libre = -1
    
    If Sangre_Libre <> -1 Then
        If Sangres(Sangre_Libre).viva = False Then Sangre_Obtener_Libre = Sangre_Libre
    End If
    
    If Sangre_Obtener_Libre = -1 Then
        For i = 0 To Ultima_Sangre
            If Sangres(i).viva = False Then
                Sangre_Obtener_Libre = i
                Exit For
            End If
        Next i
    End If
    
    If Sangre_Obtener_Libre = -1 Then
        Ultima_Sangre = Ultima_Sangre + 1
        ReDim Preserve Sangres(Ultima_Sangre)
        ReDim Sangres(Ultima_Sangre).Particulas(0) As sangre_particula
        Sangre_Obtener_Libre = Ultima_Sangre
    End If
    
    Debug.Print "SE ABRIO"; Sangre_Obtener_Libre; "PARA SANGRE"
End Function

Public Sub Sangre_Render()
    Dim i As Integer
    Dim conteo As Integer
    For i = 0 To Ultima_Sangre 'nunca va a ser un numero grande, en la pantalla no entra tanta sangre mueejjej
        If Sangres(i).viva = True Then
            Sangre_Render_ID i
            conteo = conteo + 1
        End If
    Next i

    If conteo = 0 Then
        ReDim Sangres(0)
        ReDim Sangres(0).Particulas(0)
        Ultima_Sangre = 0
        Sangre_Libre = -1
    End If
End Sub

Private Sub Sangre_Render_ID(id As Integer)
    Dim tick As Long
    Dim factor As Single
    Dim i As Integer
    Dim MY As Integer
    Dim Matar As Byte
    Dim cant As Integer
    Dim mox As Integer
    Dim moy As Integer
    Dim TLlist() As PARTVERTEX
    mox = offset_map_part.x
    moy = offset_map_part.y
    
    If Ultima_Sangre >= id And id >= 0 Then
        If Sangres(id).viva = True Then
        
            tick = GetTimer
            
            If Sangres(id).end >= tick Then
            
                factor = timerElapsedTime / 2

                Matar = 0
                
                ReDim TLlist(0 To Sangres(id).cantidad)
                For i = 0 To Sangres(id).cantidad
                    With Sangres(id).Particulas(i)
                        If tick > .MuereEnTick Then
                            .viva = False
                        End If
                        If .viva Then
                            If tick > .NaceEnTick Then
                                If .y < .ModificadorY2 Then
                                    .x = .x + (factor * .vX)
                                    .vY = .vY + factor * 0.04 ' Gravedad * 0.01
                                    .y = .y + (factor * .vY)
                                

                                    If .Tamaño < 64 Then _
                                        .Tamaño = .Tamaño + factor * 0.3 '2 * factor_agrandamiento
                                End If
                                TLlist(cant).Tamanio = .Tamaño
                                .alpha = .alpha - factor * 0.001
                                If .alpha < 0 Then
                                    .viva = False
                                    .alpha = 0
                                End If
                                
                                TLlist(cant).Tamanio = .Tamaño
                                TLlist(cant).v.x = .x + mox
                                TLlist(cant).v.y = .y + moy + 16
                                'Debug.Print TLlist(cant).v.x; TLlist(cant).v.y
                                TLlist(cant).Color = D3DColorMake(Sangres(id).col.r, Sangres(id).col.g, Sangres(id).col.b, .alpha)
                                TLlist(cant).rhw = 1
                                cant = cant + 1
                            Else
                                .x = (CharList(Sangres(id).src_char).Pos.x) * 32 + CharList(Sangres(id).src_char).MoveOffsetX + 16 + Rnd * 5 - Rnd * 5
                                .y = (CharList(Sangres(id).src_char).Pos.y) * 32 + CharList(Sangres(id).src_char).MoveOffsetY - Sangres(id).Altura + Rnd * 5 - Rnd * 5
                                .ModificadorY2 = (CharList(Sangres(id).src_char).Pos.y) * 32 + CharList(Sangres(id).src_char).MoveOffsetY - .ModificadorY
                            End If
                            Matar = 1

                        End If
                        
                    End With
                Next i

                If Matar = 0 Then
                    Sangres(id).viva = False
                    Sangre_Libre = id
                Else
                    If cant > 0 Then
                        DibujarParticulas TLlist, cant, Sangre_grh, 0
                    End If
                End If
            Else
                Sangres(id).viva = False
                Sangre_Libre = id
            End If
        Else
            Sangre_Libre = id
        End If
    End If
End Sub

Private Sub Sangre_Init_ID(id As Integer)
    Dim tick As Long
    tick = GetTimer
    Dim i As Integer
    Dim MX As Integer
    Dim MY As Integer
    Dim MMX As Single
    Dim MMY As Single
    Dim Tmp As Integer
    MMX = 0
    MMY = 0
    
    Select Case (CharList(Sangres(id).src_char).heading - 1) 'Sangres(ID).wind
        Case 0
            MY = 1
            MX = 0
            MMX = Rnd
            MMY = -1
        Case 1
            MY = 0
            MX = 1
            MMY = Rnd
        Case 2
            MY = 1
            MX = 0
            MMX = Rnd
        Case 3
            MY = 0
            MX = -1
            MMY = Rnd
    End Select
                
    For i = 0 To Sangres(id).cantidad
        With Sangres(id).Particulas(i)
            .vX = MX - (Rnd * MX) + ((Rnd * MMX) - (Rnd * MMX))
            .vY = 0
            .x = CharList(Sangres(id).src_char).MoveOffsetX
            .y = CharList(Sangres(id).src_char).MoveOffsetY - Sangres(id).Altura
            .ModificadorY = (Rnd * 40) - (Rnd * 40)
            .ModificadorY2 = (CharList(Sangres(id).src_char).Pos.y) * 32 + CharList(Sangres(id).src_char).MoveOffsetY + .ModificadorY
            Tmp = Rnd * 450
            .NaceEnTick = tick + Tmp
            .MuereEnTick = .NaceEnTick + Sangres(id).end - Sangres(id).Start + (Rnd * Vida_gota) + (Rnd * Vida_gota) + (Rnd * Vida_gota)
            
            If .MuereEnTick > Sangres(id).end Then
                Sangres(id).end = .MuereEnTick
            End If
            
            .viva = True
            .Tamaño = 1
            .alpha = 1
        End With
    Next i
    
End Sub


