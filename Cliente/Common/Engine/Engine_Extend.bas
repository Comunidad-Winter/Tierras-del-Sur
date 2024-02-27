Attribute VB_Name = "Engine_Extend"
' ESTE ARCHIVO ESTA COMPARTIDO POR TODOS LOS PROGRAMAS.

''
' @require Engine.bas
' @require Engine_Landscape.bas
' @require Engine_Landscape_Water.bas
' @require Engine_Particles.bas

Option Explicit

Public user_screen_pos As D3DVECTOR2



'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As position
    #If esMe = 1 Then
        Nombre As String
    #End If
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
    
    #If esMe = 1 Then
        Nombre As String
    #End If
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    Grh As Integer
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
    
    #If esMe = 1 Then
        Nombre As String
    #End If
    
    '[ANIM ATAK]
    WeaponAttack As Byte
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
    ShieldAttack As Byte
    
    #If esMe = 1 Then
        Nombre As String
    #End If
    
End Type
'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

Public Type tLuzPropiedades
    LuzMouse As Integer
    LuzColor As RGBCOLOR
    LuzRadio As Byte
    LuzTipo As Integer
    luzInicio As Byte
    luzFin As Byte
    LuzBrillo As Byte
End Type


Public Type tIndiceArma
    Walk(1 To 4) As Integer
End Type

Public Type tIndiceEscudo
    Walk(1 To 4) As Integer
End Type

Public Type tIndiceCuerpo
    body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Integer
    offsetX As Single
    offsetY As Single
    particula As Integer
    wav As Integer
    
    #If esMe = 1 Then
        Nombre As String
    #End If
    
End Type

Public Type tIndiceEntidad
    #If esMe = 1 Then
        Nombre As String
    #End If
    
    Graficos() As Integer
    Sonidos() As Integer
    Particulas() As Integer
    SonidosAlPegar() As Integer
    
    tipo As Byte
    Vida As Integer
    CrearAlMorir As Integer
    Proyectil As Byte
    
    luz As tLuzPropiedades
End Type


Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public EntidadesIndexadas() As tIndiceEntidad

Public NumWeaponAnims As Integer
Public NumEscudosAnims As Integer
Public Type tArmas
    texture(5) As Integer
    textures As Byte
    num As Integer
End Type

Public Type Arma_act
    num         As Integer
    TEX_FLAGS   As Byte
    mano        As Byte
End Type

Public Enum ePersonajeFlags
    muerto = 1
    invisible = 2               ' Invisible basica, con intermitencia
    Oculto = 4
    tieneClan = 8
    active = 16
    invisibleTotal = 64         ' No tiene intermitencia
End Enum

Public Enum eAlineaciones
    indefinido = 0
    Neutro = 1
    Real = 2
    caos = 3
End Enum

Public Type Char
    Alineacion As eAlineaciones
    
    flags As Byte

    CharIndex As Integer
    
    active As Byte
    heading As E_Heading

    alpha As Single
    alphacounter As Single
    
    'sangre_fx As Integer
    
    Color As Long
    Color2 As Long
    
    alpha_sentido As Boolean
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    pie As Boolean
    muerto As Boolean
    priv As Byte
    
    'MANp As Byte
    'VIDp As Byte
    
    iHead As Integer
    iBody As Integer
    pelo As Integer
    barba As Integer
    ropaInterior As Integer
    body As BodyData
    Head As HeadData
    casco As HeadData
    arma As WeaponAnimData
    escudo As ShieldAnimData
    UsandoArma As Boolean
    
    hit_color As Long
    
    hit As Integer
    hit_act As Byte
    hit_off As Single
    
    fX As Grh
    FxIndex As Integer
    
    Nombre As String
    Clan As String
    
    Pos As position
    MPPos As position
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    Particle_group(0 To 1) As Engine_Particle_Group
    
    luz As Integer
    
    Velocidad As position
    
    'spd As D3DVECTOR2
    'ace As D3DVECTOR2
    vec As D3DVECTOR2
    'do_onda As Byte
    
    DirY As Integer
    Offset_Altura_Inicial As Single
    Offset_Altura_Final As Single
    OffY As Single
    
    attaking As Byte
    invh As Byte
    invheading As E_Heading
    
    center_text As Integer
    
    rcrc As Long
    
    armaz(0 To 1) As Arma_act
    
    mu As Single
    
    NickLabel As clsGUIText
    NickClan As clsGUIText
    
    entidad As Integer
End Type

Public Const MaxChar As Integer = 10000

Public CharList(1 To MaxChar) As Char

'TODO MARCE Revisar esto. �Es necesario que sea del tama�o de la parte visible?
Public CharMap(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE) As Integer

Public lista_armas(30) As tArmas

Public weapon_array() As Box_Vertex
Public fix_heading(1 To 4) As Integer

Public Nombres As Boolean 'Muestra o no los nombres de los personajes

Dim sombras(0) As PARTVERTEX 'Esto se usa para dibujar las olitas y la sombra del pj

Private NumChars As Integer

Public Function ActivateChar(caracter As Char) As Boolean
                   
     If caracter.CharIndex <> UserCharIndex And UserCharIndex > 0 Then
        'Si la distancia es mayor al rango que se visualiza entonces se "borra" el char. Se desactiva su display
        If Abs(CharList(UserCharIndex).Pos.X - caracter.Pos.X) > ARangoX Or Abs(CharList(UserCharIndex).Pos.Y - caracter.Pos.Y) > ARangoY Then
            'Lo desactivamos
            caracter.active = 0
            'Terminamos
            ActivateChar = False
            Exit Function
        End If
    End If
    
    caracter.active = 1
    'Actualiamos la neuva posicion ya sea mia o del otro charindex
    CharMap(caracter.Pos.X, caracter.Pos.Y) = caracter.CharIndex
    ActivateChar = True
End Function

Public Sub DeactivateChar(caracter As Char)
    caracter.active = 0
    CharMap(caracter.Pos.X, caracter.Pos.Y) = 0
End Sub

'Public Sub CharSetPos(ByVal CharIndex As Integer, ByVal x As Integer, ByVal y As Integer)
'    Dim bkX As Integer
'    Dim bkY As Integer
'
'    With CharList(CharIndex).pos
'        bkX = .x
'        bkY = .y
        
'        If CharList(CharIndex).active Then
'            If esPosicionJugable(bkX, bkY) And CharMap(bkX, bkY) = CharIndex Then CharMap(bkX, bkY) = 0
'            If esPosicionJugable(x, y) Then CharMap(x, y) = CharIndex
'        End If
        
'        .x = x
'        .y = y
'    End With
'End Sub



Public Sub Engine_SortIntArray(TheArray() As Integer, TheIndex() As Integer, ByVal LowerBound As Integer, ByVal UpperBound As Integer)
    Dim indxt As Long   'Stored index
    Dim swp As Integer  'Swap variable
    Dim i As Integer    'Subarray Low  Scan Index
    Dim j As Integer    'Subarray High Scan Index

    'Start the loop
    For j = LowerBound + 1 To UpperBound
        indxt = TheIndex(j)
        swp = TheArray(indxt)
        For i = j - 1 To LowerBound Step -1
            If TheArray(TheIndex(i)) <= swp Then Exit For
            TheIndex(i + 1) = TheIndex(i)
        Next i
        TheIndex(i + 1) = indxt
    Next j

End Sub

Public Sub Init_weapons()
Dim X&, Y&
ReDim weapon_array(0 To 7, 0 To 7)
For Y = 0 To 7
    For X = 0 To 7
        With weapon_array(X, Y)
            .tu0 = (X * 64) / 512
            .tv0 = ((Y + 1) * 64) / 512
            
            .tu1 = .tu0
            .tv1 = (Y * 64) / 512
            
            .tu2 = ((X + 1) * 64) / 512
            .tv2 = .tv0
            
            .tu3 = .tu2
            .tv3 = .tv1
            .rhw0 = 1
            .rhw1 = 1
            .rhw2 = 1
            .rhw3 = 1
        End With
    Next X
Next Y
fix_heading(1) = 1
fix_heading(2) = 4
fix_heading(3) = 2
fix_heading(4) = 3
For Y = 0 To 30
    lista_armas(Y).texture(0) = 14000 + Y
    lista_armas(Y).textures = 1
Next Y
End Sub


Public Sub Render_Armas(ByVal CharIndex As Integer, ByVal dest_x As Integer, ByVal dest_y As Integer)
'*********************************************
'Author: menduz
'*********************************************
    Dim dest_x2 As Integer
    Dim dest_y2 As Integer
    Dim map_x As Integer
    Dim map_y As Integer
    Dim TGRH As Box_Vertex
    Dim tex_Act As Byte
    Dim ArmAct As Integer
    Dim armaX As Arma_act
    Dim tBox As Box_Vertex
    Dim anim As Integer
    Dim i%, f%
    
    If CharIndex = 0 Then Exit Sub
    tex_Act = 1
    map_x = CharList(CharIndex).Pos.X
    map_y = CharList(CharIndex).Pos.Y
    
    
    
    dest_x = dest_x - 16
    dest_y = dest_y - 32
    dest_y2 = dest_y + 64
    dest_x2 = dest_x + 64
    
    If CharList(CharIndex).attaking Then
        anim = (CharList(CharIndex).arma.WeaponWalk(CharList(CharIndex).heading).FrameCounter Mod 8) - 1
        TGRH = weapon_array(fix_heading(CharList(CharIndex).heading) + 3, anim)
    Else
        anim = (CharList(CharIndex).body.Walk(CharList(CharIndex).heading).FrameCounter Mod 8) - 1
        TGRH = weapon_array(anim, fix_heading(CharList(CharIndex).heading) - 1)
    End If
    
    With tBox
        .x0 = dest_x
        .y0 = dest_y2
        .x1 = .x0
        .y1 = dest_y
        .x2 = dest_x2
        .y2 = .y0
        .x3 = .x2
        .y3 = .y1
        .rhw0 = 1
        .rhw1 = 1
        .rhw2 = 1
        .rhw3 = 1
    End With
    
    For f = 0 To 1
        armaX = CharList(CharIndex).armaz(f)
        
        If armaX.num > 0 Then
            
                If armaX.mano Then
                    With tBox
                        .tu0 = TGRH.tu2
                        .tv0 = TGRH.tv2
                        .tu1 = TGRH.tu3
                        .tv1 = TGRH.tv3
                        .tu2 = TGRH.tu0
                        .tv2 = TGRH.tv0
                        .tu3 = TGRH.tu1
                        .tv3 = TGRH.tv1
                    End With
                    Colorear_TBOX_Flip tBox, map_x, map_y
                Else
                    With tBox
                        .tu0 = TGRH.tu0
                        .tv0 = TGRH.tv0
                        .tu1 = TGRH.tu1
                        .tv1 = TGRH.tv1
                        .tu2 = TGRH.tu2
                        .tv2 = TGRH.tv2
                        .tu3 = TGRH.tu3
                        .tv3 = TGRH.tv3
                    End With
                    Colorear_TBOX tBox, map_x, map_y
                End If

            Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(lista_armas(armaX.num).texture(0))
            Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
            
            If armaX.TEX_FLAGS > 0 Then
                With tBox
                    .color0 = -1&
                    .Color1 = .color0
                    .Color2 = .color0
                    .color3 = .color0
                End With
                For i = 1 To 5
                    If BS_Byte_Get(armaX.TEX_FLAGS, i) Then
                        Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(lista_armas(armaX.num).texture(i))
                        Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
                        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size
                    End If
                Next i
            End If
        End If
    Next f
End Sub

Public Sub actualizarPersonajeArmadura(Rdata As String)
    Dim TempInt As Integer
    TempInt = STI(Rdata, 1)
    CharList(TempInt).iBody = STI(Rdata, 3)
    CharList(TempInt).body = BodyData(STI(Rdata, 3))
    CharList(TempInt).ropaInterior = STI(Rdata, 5)
     
    If (CharList(TempInt).iBody = 8 Or CharList(TempInt).iBody = 145) And Not CharList(TempInt).Nombre = "" Then
        CharList(TempInt).muerto = True
    Else
        CharList(TempInt).muerto = False
    End If
End Sub

Public Sub actualizarPersonaje(Rdata As String)
    Dim TempInt As Integer
    Dim TempByte As Byte
    
    TempInt = STI(Rdata, 1)
    CharList(TempInt).iBody = STI(Rdata, 3)
    CharList(TempInt).body = BodyData(STI(Rdata, 3))
    CharList(TempInt).Head = HeadData(STI(Rdata, 5))
    CharList(TempInt).heading = StringToByte(Rdata, 7)
    CharList(TempInt).invheading = CharList(TempInt).heading
                
    If StringToByte(Rdata, 8) > 0 Then
        CharList(TempInt).arma = WeaponAnimData(StringToByte(Rdata, 8))
        CharList(TempInt).escudo = ShieldAnimData(StringToByte(Rdata, 9))
    End If
             
    'CharList(tempint).FxLoopTimes = STI(Rdata, 11)
    CharList(TempInt).casco = CascoAnimData(StringToByte(Rdata, 13))
    CharList(TempInt).pelo = STI(Rdata, 14)
    CharList(TempInt).barba = STI(Rdata, 16)
    CharList(TempInt).ropaInterior = STI(Rdata, 18)
    
    TempByte = StringToByte(Rdata, 10)
    'CharList(tempint).fx = TempByte
    SetCharacterFx TempInt, TempByte, STI(Rdata, 11)
    
    If (CharList(TempInt).iBody = 8 Or CharList(TempInt).iBody = 145) And Not CharList(TempInt).Nombre = "" Then
        CharList(TempInt).muerto = True
    Else
        CharList(TempInt).muerto = False
    End If
        
End Sub
'
Public Sub MakeChar(ByVal CharIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal arma As Integer, ByVal escudo As Integer, ByVal casco As Integer) ':(�Missing Scope

    'Apuntamos al ultimo Char
    NumChars = NumChars + 1

    If arma = 0 Then arma = 2
    If escudo = 0 Then escudo = 2
    If casco = 0 Then casco = 2

    CharList(CharIndex).CharIndex = CharIndex
    CharList(CharIndex).iHead = Head
    CharList(CharIndex).iBody = body
    CharList(CharIndex).Head = HeadData(Head)
    CharList(CharIndex).body = BodyData(body)
    CharList(CharIndex).arma = WeaponAnimData(arma)
    
    '[ANIM ATAK]
    CharList(CharIndex).arma.WeaponAttack = 0
    CharList(CharIndex).escudo.ShieldAttack = 0
    CharList(CharIndex).escudo = ShieldAnimData(escudo)
    CharList(CharIndex).casco = CascoAnimData(casco)
    
    'Reset moving stats
    CharList(CharIndex).Moving = 0
    CharList(CharIndex).MoveOffsetX = 0
    CharList(CharIndex).MoveOffsetY = 0
    
    'Cabeza
    If heading = 0 Then heading = E_Heading.SOUTH
    CharList(CharIndex).heading = heading
    CharList(CharIndex).invheading = CharList(CharIndex).heading
    
    'Lado del Arma
    If CharList(CharIndex).invh Then
        If CharList(CharIndex).heading = E_Heading.EAST Then
            CharList(CharIndex).invheading = E_Heading.WEST
          ElseIf CharList(CharIndex).heading = E_Heading.WEST Then 'NOT CHARLIST(CHARINDEX).HEADING...
            CharList(CharIndex).invheading = E_Heading.EAST
        End If
    End If
    
    'Reset moving stats
    CharList(CharIndex).Moving = 0
    CharList(CharIndex).MoveOffsetX = 0
    CharList(CharIndex).MoveOffsetY = 0
    
    If CharList(CharIndex).Velocidad.X = 0 Then
        CharList(CharIndex).Velocidad.X = ScrollPixelsPerFrameX
        CharList(CharIndex).Velocidad.Y = ScrollPixelsPerFrameY
    End If
 
    ' char_act_color CharIndex

    If CharList(CharIndex).luz Then
        DLL_Luces.Quitar CharList(CharIndex).luz
        CharList(CharIndex).luz = 0
    End If
        
     If (CharList(CharIndex).iBody = 8 Or CharList(CharIndex).iBody = 145) And Not CharList(CharIndex).Nombre = "" Then
        CharList(CharIndex).muerto = True
    Else
        CharList(CharIndex).muerto = False
    End If

    CharList(CharIndex).Pos.X = X
    CharList(CharIndex).Pos.Y = Y
    
    If CharIndex = UserCharIndex Then
        UserPos.X = X
        UserPos.Y = Y
    End If
    
    
    Call ActivateChar(CharList(CharIndex))

   Exit Sub

MakeChar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MakeChar of M�dulo Engine_Extend"
End Sub


Sub ResetCharInfo(ByVal CharIndex As Integer)
    With CharList(CharIndex)
        .FxIndex = 0
        
        .Moving = 0
        .muerto = False
        
        .Nombre = ""
    
        .Clan = ""
        .center_text = 0
        .pie = False

        .UsandoArma = False
        .luz = 0
        .flags = 0
        .Alineacion = indefinido
        
        Set .NickClan = Nothing
        Set .NickLabel = Nothing
    End With
    
    DeactivateChar CharList(CharIndex)
    
    If Not CharList(CharIndex).NickLabel Is Nothing Then
        Set CharList(CharIndex).NickLabel = Nothing
    End If
    
    If Not CharList(CharIndex).NickClan Is Nothing Then
        Set CharList(CharIndex).NickClan = Nothing
    End If

End Sub

Public Sub Engine_MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X As Integer
    Dim Y As Integer
    Dim tx As Integer
    Dim ty As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
        
        Case E_Heading.EAST
            X = 1
        
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            X = -1
    End Select
    
    'Fill temp pos
    tx = UserPos.X + X
    ty = UserPos.Y + Y
    
    'Check to see if its out of bounds
    If tx < X_MINIMO_VISIBLE Or tx > X_MAXIMO_VISIBLE Or ty < Y_MINIMO_VISIBLE Or ty > Y_MAXIMO_VISIBLE Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = tx
        AddtoUserPos.Y = Y
        UserPos.Y = ty
        UserMoving = 1
        
        bTecho = CBool(mapdata(UserPos.X, UserPos.Y).trigger And eTriggers.BajoTecho)
    End If
End Sub

Public Sub Engine_MoveScreen2pos(ByVal NX As Byte, ByVal NY As Byte)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X!
    Dim Y!
    Dim tx!
    Dim ty!
    Dim addx As Integer
    Dim addy As Integer
    Dim nHeading As E_Heading
    'Figure out which way to move

    addx = NX - UserPos.X
    addy = NY - UserPos.Y

    tx = NX
    ty = NY

    'Check to see if its out of bounds
    If tx < MinXBorder Or tx > MaxXBorder Or ty < MinYBorder Or ty > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = addx
        UserPos.X = tx
        AddtoUserPos.Y = addy
        UserPos.Y = ty
        'UserMoving = 1

        bTecho = CBool(mapdata(UserPos.X, UserPos.Y).trigger And eTriggers.BajoTecho)
    End If
End Sub


Public Function Char_Pos_Get(ByVal char_index As Integer, ByRef map_x As Integer, ByRef map_y As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'*****************************************************************
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        map_x = CharList(char_index).Pos.X
        map_y = CharList(char_index).Pos.Y
        Char_Pos_Get = True
    End If
End Function

Public Sub Char_Start_Anim(ByVal CharIndex As Long)
    With CharList(CharIndex)
        .arma.WeaponWalk(.heading).Started = 1
        .escudo.ShieldWalk(.heading).Started = 1
        .attaking = 255
    End With
End Sub

Public Sub Char_Start_Anim_Escudo(ByVal CharIndex As Long)
    With CharList(CharIndex)
        .escudo.ShieldWalk(.heading).Started = 1
        .attaking = 255
    End With
End Sub


Public Sub Char_Start_Anim_Arma(ByVal CharIndex As Long)
    With CharList(CharIndex)
        .arma.WeaponWalk(.heading).Started = 1
        .attaking = 255
    End With
End Sub











Public Sub Char_Render(ByVal CharIndex As Long)
'cfnc = fnc.E_Char_Render
    Dim moved As Boolean
    
    
    Dim iRender As Boolean
    Dim PixelOffsetX As Single
    Dim PixelOffsetY As Single
    Dim Yas As Integer
    Dim alphaname As Byte
    'Dim Mu As Single

    If CharIndex < 1 Then Exit Sub
    With CharList(CharIndex)
        If .active = 0 Then Exit Sub
        If .rcrc = RENDERCRC Then Exit Sub
        .rcrc = RENDERCRC
        
        'If (.Pos.x <> Map_x) Or (.Pos.y <> Map_y) Then
            'CharMap(Map_x, Map_y) = 0
            'Exit Sub
        'End If
        
        If .Velocidad.X = 0 Then
            .Velocidad.X = ScrollPixelsPerFrameX
            .Velocidad.Y = ScrollPixelsPerFrameY
        End If
        
        If .heading < 1 Or .heading > 4 Then Exit Sub
        
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + .Velocidad.X * Sgn(.scrollDirectionX) * timerTicksPerFrame
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
                .mu = Abs(.MoveOffsetX) / 32
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + .Velocidad.Y * Sgn(.scrollDirectionY) * Round(timerTicksPerFrame, 3)
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
                .mu = Abs(.MoveOffsetY) / 32
            End If
        Else
            .mu = 0
        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            
            .body.Walk(.heading).Started = 0
            .body.Walk(.heading).FrameCounter = 1
            

            If Not .attaking Then
                .arma.WeaponWalk(.invheading).Started = 0
                .arma.WeaponWalk(.invheading).FrameCounter = 1
                    
                .escudo.ShieldWalk(.invheading).Started = 0
                .escudo.ShieldWalk(.invheading).FrameCounter = 1
            End If
            
            .Moving = False
        Else
            If .body.Walk(.heading).Speed > 0 Then _
                .body.Walk(.heading).Started = 1
                
            .arma.WeaponWalk(.invheading).Started = 1
            .escudo.ShieldWalk(.invheading).Started = 1
            
            .attaking = 0
            
            If .luz Then
                DLL_Luces.Move .luz, .Pos.X, .Pos.Y, .MoveOffsetX, .MoveOffsetY
            End If
            
            
            If AlturaPie(.Pos.X, .Pos.Y) > .OffY Then
                .DirY = 1
            ElseIf AlturaPie(.Pos.X, .Pos.Y) < .OffY Then
                .DirY = -1 '(MapData(.Pos.X, .Pos.Y).h - .OffY)
            End If
        End If
        
        

        If .Offset_Altura_Final <> AlturaPie(.Pos.X, .Pos.Y) Then
            .Offset_Altura_Inicial = .OffY
            .Offset_Altura_Final = AlturaPie(.Pos.X, .Pos.Y)
        End If
        
        If .OffY <> .Offset_Altura_Final Then
            .OffY = Interp(.Offset_Altura_Final, .Offset_Altura_Inicial, mins(1, .mu))
        End If
        

        If CharIndex = UserCharIndex Then
            PixelOffsetX = user_screen_pos.X 'offset_map.X + (.Pos.X * 32 + .MoveOffsetX) '.Pos.X * 32 - offset_screen.X + offset_mapO.X
            PixelOffsetY = user_screen_pos.Y - .OffY + .vec.Y + Screen_Desnivel_Offset 'offset_map.Y + (.Pos.Y * 32 + .MoveOffsetY + -.OffY + .vec.Y)
        Else
            PixelOffsetX = (.Pos.X + minXOffset) * 32 + .MoveOffsetX + offset_map.X
            PixelOffsetY = (.Pos.Y + minYOffset) * 32 + .MoveOffsetY + offset_map.Y - .OffY + .vec.Y
        End If
        
        .MPPos.X = PixelOffsetX '- offset_map.X
        .MPPos.Y = PixelOffsetY '- offset_map.Y
        
        Yas = .Pos.Y - .OffY / 32
        'FIXME If MouseTileX - .Pos.x = 0 And (MouseTileY - Yas = 0 Or MouseTileY - Yas = -1) Then Protocol.aim_pj = CharIndex Xor 105
        
        If .Head.Head(.heading).GrhIndex Then
            iRender = True
            
            If (.flags And ePersonajeFlags.Oculto) Then
                If CharIndex = UserCharIndex Then
                    .alphacounter = 0
                    .alpha = 128
                Else
                    iRender = False
                End If
            ElseIf .flags And ePersonajeFlags.invisibleTotal Then
                If CharIndex = UserCharIndex Then
                    .alphacounter = 0
                    .alpha = 128
                Else
                    iRender = False
                End If
           ElseIf .flags And ePersonajeFlags.invisible Then
                If CharIndex = UserCharIndex Then
                    .alphacounter = 0
                    .alpha = 128
                Else
                    .alphacounter = .alphacounter + timerElapsedTime

                    If Int(.alphacounter \ 1000) = 4 Or Int(.alphacounter \ 1000) = 8 Then
                        .alpha = 128
                    Else
                        iRender = False
                    End If
                End If
            Else
                .alpha_sentido = False
                .alpha = 255
                .alphacounter = 0
            End If
            
   
            'Draw Body
            If Not .Particle_group(0) Is Nothing Then
                .Particle_group(0).SetPixelPos .Pos.X * 32 + .MoveOffsetX + 16, .Pos.Y * 32 + .MoveOffsetY + 16
                If .Particle_group(0).Render() = False Then
                    Set .Particle_group(0) = Nothing
                End If
            End If
            
            If iRender Then
                alphaname = min(.alpha, ((ResultColorArray(.Pos.X, .Pos.Y) And &HFF0000) / &H10000) + 150)

                If .muerto Then
                    If .body.Walk(.heading).GrhIndex Then _
                        Call Draw_Grh_Alpha(.body.Walk(.heading), PixelOffsetX, PixelOffsetY, 1, 100, .Pos.X, .Pos.Y, 1)
                    If .Head.Head(.heading).GrhIndex Then _
                        Call Draw_Grh_Alpha(.Head.Head(.heading), PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, 0, 100, .Pos.X, .Pos.Y, 1)

                    If Nombres Then
                        If Not .NickLabel Is Nothing Then
                            .NickLabel.SetPos PixelOffsetX + 16, PixelOffsetY + 30
                            .NickLabel.Render
                        End If
                        If Not .NickClan Is Nothing Then
                            .NickClan.SetPos PixelOffsetX + 16, PixelOffsetY + 43
                            .NickClan.Render
                        End If
                    End If
                Else
                    If .alpha = 255 Then
                        Render_Shadow_Pj CharIndex, alphaname

                        If mapdata(.Pos.X, .Pos.Y + 1).is_water Or (mapdata(.Pos.X, .Pos.Y + 1).trigger And eTriggers.Navegable) Then
                            render_reflejo CharIndex, PixelOffsetX, PixelOffsetY + .OffY, 1
                            draw_char CharIndex, PixelOffsetX, PixelOffsetY, alphaname
                            Render_Olitas_Pj CharIndex
                        Else
                            draw_char CharIndex, PixelOffsetX, PixelOffsetY, alphaname
                        End If
                        
                        If Nombres Then
                            If Not .NickLabel Is Nothing Then
                                .NickLabel.SetPos PixelOffsetX + 16, PixelOffsetY + 30
                                .NickLabel.Render
                            End If
                            If Not .NickClan Is Nothing Then
                                .NickClan.SetPos PixelOffsetX + 16, PixelOffsetY + 43
                                .NickClan.Render
                            End If
                        End If
                    Else
                        draw_char CharIndex, PixelOffsetX, PixelOffsetY, alphaname
                        If mapdata(.Pos.X, (.Pos.Y + 1) Mod ALTO_MAPA).is_water Then
                            Render_Olitas_Pj CharIndex
                        End If
                    End If
                End If
            End If
            #If esCLIENTE = 1 Then
                Call Dialogos.UpdateDialogPos(PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, CharIndex)
            #End If
        Else
            'Draw Body
        'If .velocidad.x = 0 Then
            '.velocidad.X = 4
            '.velocidad.Y = 4
        'End If
            If .body.Walk(.heading).GrhIndex Then _
                Call Draw_Grh_Interpolador(.body.Walk(.heading), PixelOffsetX, PixelOffsetY, 1, .Pos.X, .Pos.Y, 1, , , .mu)
            
            #If esCLIENTE = 1 Then
                Call Dialogos.UpdateDialogPos(PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, CharIndex)
            #End If
            'Call Text_Render_alpha("HAPPY NPC", PixelOffsetY + 30, PixelOffsetX + 16, &HFFFFFFFF, DT_CENTER)
        End If

        ''Update dialogs
        
        'Call Hits.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.X + 5, PixelOffsetY, CharIndex)

        'Draw FX
        If .FxIndex <> 0 Then
            Call Draw_Grh(.fX, PixelOffsetX + FxData(.FxIndex).offsetX, PixelOffsetY + FxData(.FxIndex).offsetY, 1, .Pos.X, .Pos.Y, 1)

            If .fX.Started = 0 Then _
                .FxIndex = 0
        End If
        
        If .hit_act = 1 Then
            .hit_off = .hit_off - timerTicksPerFrame * 3
            If .hit_off > -32 Then
                Call Text_Render_alpha(CStr(.hit), PixelOffsetY + 10 + .hit_off, PixelOffsetX + 24, .hit_color, DT_TOP Or DT_CENTER, CByte(255 - .hit_off * -4))
            Else
                .hit_act = 0
                .hit_off = 0
            End If
        End If

        If Not .Particle_group(1) Is Nothing Then
            .Particle_group(1).SetPixelPos .Pos.X * 32 + .MoveOffsetX + 16, .Pos.Y * 32 + .MoveOffsetY + 16 '.MPPos.x + 16, .MPPos.y + 16
            If .Particle_group(1).Render() = False Then
                Set .Particle_group(1) = Nothing
            End If
        End If
    End With
End Sub

Private Function min(ByVal val1 As Long, ByVal val2 As Long) As Long
'***************************************************
'Autor: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'It's faster than iif and I like it better
'***************************************************
    If val1 < val2 Then
        min = val1
    Else
        min = val2
    End If
End Function

Private Function Char_Check(ByVal char_index As Integer) As Boolean
    Char_Check = CharList(char_index).active = 1
End Function

Private Sub draw_char(ByVal id As Integer, ByVal PixelOffsetX!, ByVal PixelOffsetY!, Optional ByVal alpha As Byte = 0)
With CharList(id)
    If .body.Walk(.heading).GrhIndex Then
        Select Case .heading
            Case SOUTH
                'Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, .Pos.x, .Pos.y, 1)
                Call Draw_Grh_Interpolador(.body.Walk(.heading), PixelOffsetX, PixelOffsetY, 1, .Pos.X, .Pos.Y, 1, , , .mu, alpha)

                If .ropaInterior Then
                    
                    Call Draw_Grh_Interpolador(.body.Walk(.heading), PixelOffsetX, PixelOffsetY, 1, .Pos.X, .Pos.Y, 1, , , .mu, alpha, .ropaInterior)
                End If
                
                Call Draw_Grh(.Head.Head(.heading), PixelOffsetX + .body.HeadOffset.X + 1, PixelOffsetY + .body.HeadOffset.Y, 0, .Pos.X, .Pos.Y, 1, 0, 0, alpha)

                If .pelo Then
                    Call Draw_Barba(.pelo, PixelOffsetX + .body.HeadOffset.X + 1, PixelOffsetY + .body.HeadOffset.Y, .heading, .Pos.X, .Pos.Y, 1, 0, 0, alpha)
                End If
                
                If .barba Then
                    Call Draw_Barba(.barba, PixelOffsetX + .body.HeadOffset.X + 1, PixelOffsetY + .body.HeadOffset.Y, .heading, .Pos.X, .Pos.Y, 1, 0, 0, alpha)
                End If
                
                 If .casco.Head(.heading).GrhIndex Then _
                    Call Draw_Grh(.casco.Head(.heading), PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, 0, .Pos.X, .Pos.Y, 1, 0, 0, alpha)
                    
                #If NuevaVersion = 0 Then
                    If .arma.WeaponWalk(.invheading).GrhIndex Then _
                        Call Draw_Grh(.arma.WeaponWalk(.invheading), PixelOffsetX, PixelOffsetY, 1, .Pos.X, .Pos.Y, 1, .invh, , alpha)
                #Else
                    Render_Armas id, PixelOffsetX, PixelOffsetY
                #End If
                
                If .escudo.ShieldWalk(.invheading).GrhIndex Then _
                    Call Draw_Grh(.escudo.ShieldWalk(.invheading), PixelOffsetX, PixelOffsetY, 1, .Pos.X, .Pos.Y, 1, .invh, , alpha)
            Case NORTH
            
                #If NuevaVersion = 0 Then
                    If .arma.WeaponWalk(.invheading).GrhIndex Then _
                        Call Draw_Grh(.arma.WeaponWalk(.invheading), PixelOffsetX, PixelOffsetY, 1, .Pos.X, .Pos.Y, 1, .invh, , alpha)
                #Else
                    Render_Armas id, PixelOffsetX, PixelOffsetY
                #End If
                
                If .escudo.ShieldWalk(.invheading).GrhIndex Then _
                    Call Draw_Grh(.escudo.ShieldWalk(.invheading), PixelOffsetX, PixelOffsetY, 1, .Pos.X, .Pos.Y, 1, .invh, , alpha)
                    
                Call Draw_Grh(.Head.Head(.heading), PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, 0, .Pos.X, .Pos.Y, 1, 0, 0, alpha)
                
                If .pelo Then
                    Call Draw_Barba(.pelo, PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, .heading, .Pos.X, .Pos.Y, 1, 0, 0, alpha)
                End If
                
                If .barba Then
                    Call Draw_Barba(.barba, PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, .heading, .Pos.X, .Pos.Y, 1, 0, 0, alpha)
                End If
                
                Call Draw_Grh_Interpolador(.body.Walk(.heading), PixelOffsetX, PixelOffsetY, 1, .Pos.X, .Pos.Y, 1, , , .mu, alpha)
                
                If .ropaInterior Then
                    Call Draw_Grh_Interpolador(.body.Walk(.heading), PixelOffsetX, PixelOffsetY, 1, .Pos.X, .Pos.Y, 1, , , .mu, alpha, .ropaInterior)
                End If
                
                If .casco.Head(.heading).GrhIndex Then _
                    Call Draw_Grh(.casco.Head(.heading), PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, 0, .Pos.X, .Pos.Y, 1, 0, 0, alpha)
                    
            Case EAST
                Call Draw_Grh_Interpolador(.body.Walk(.heading), PixelOffsetX, PixelOffsetY, 1, .Pos.X, .Pos.Y, 1, , , .mu, alpha)
                
                If .ropaInterior Then
                    Call Draw_Grh_Interpolador(.body.Walk(.heading), PixelOffsetX, PixelOffsetY, 1, .Pos.X, .Pos.Y, 1, , , .mu, alpha, .ropaInterior)
                End If
                
                Call Draw_Grh(.Head.Head(.heading), PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, 0, .Pos.X, .Pos.Y, 1, 0, 0, alpha)
                
                If .barba Then
                    Call Draw_Barba(.barba, PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, .heading, .Pos.X, .Pos.Y, 1, 0, 0, alpha)
                End If
                
                If .pelo Then
                    Call Draw_Barba(.pelo, PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, .heading, .Pos.X, .Pos.Y, 1, 0, 0, alpha)
                End If
                
                If .casco.Head(.heading).GrhIndex Then _
                    Call Draw_Grh(.casco.Head(.heading), PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, 0, .Pos.X, .Pos.Y, 1, 0, 0, alpha)
                    
                If .escudo.ShieldWalk(.invheading).GrhIndex Then _
                    Call Draw_Grh(.escudo.ShieldWalk(.invheading), PixelOffsetX, PixelOffsetY, 1, .Pos.X, .Pos.Y, 1, .invh, , alpha)
                
                #If NuevaVersion = 0 Then
                    If .arma.WeaponWalk(.invheading).GrhIndex Then _
                        Call Draw_Grh(.arma.WeaponWalk(.invheading), PixelOffsetX, PixelOffsetY, 1, .Pos.X, .Pos.Y, 1, .invh, , alpha)
                #Else
                    Render_Armas id, PixelOffsetX, PixelOffsetY
                #End If
                
            Case WEST
                    
                Call Draw_Grh_Interpolador(.body.Walk(.heading), PixelOffsetX, PixelOffsetY, 1, .Pos.X, .Pos.Y, 1, , , .mu, alpha)
                
                If .ropaInterior Then
                    Call Draw_Grh_Interpolador(.body.Walk(.heading), PixelOffsetX, PixelOffsetY, 1, .Pos.X, .Pos.Y, 1, , , .mu, alpha, .ropaInterior)
                End If
                
                Call Draw_Grh(.Head.Head(.heading), PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, 0, .Pos.X, .Pos.Y, 1, 0, 0, alpha)
                  
                If .pelo Then
                    Call Draw_Barba(.pelo, PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, .heading, .Pos.X, .Pos.Y, 1, 0, 0, alpha)
                End If
                
                If .barba Then
                    Call Draw_Barba(.barba, PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, .heading, .Pos.X, .Pos.Y, 1, 0, 0, alpha)
                End If
                
                If .casco.Head(.heading).GrhIndex Then _
                    Call Draw_Grh(.casco.Head(.heading), PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, 0, .Pos.X, .Pos.Y, 1, 0, 0, alpha)
                    
                #If NuevaVersion = 0 Then
                    If .arma.WeaponWalk(.invheading).GrhIndex Then _
                        Call Draw_Grh(.arma.WeaponWalk(.invheading), PixelOffsetX, PixelOffsetY, 1, .Pos.X, .Pos.Y, 1, .invh, , alpha)
                #Else
                    Render_Armas id, PixelOffsetX, PixelOffsetY
                #End If
                
                If .escudo.ShieldWalk(.invheading).GrhIndex Then _
                    Call Draw_Grh(.escudo.ShieldWalk(.invheading), PixelOffsetX, PixelOffsetY, 1, .Pos.X, .Pos.Y, 1, .invh, , alpha)

        End Select
    End If
End With
End Sub

Private Sub render_reflejo(ByVal id As Integer, ByVal PixelOffsetX!, ByVal PixelOffsetY!, Optional ByVal shadow As Byte = 0)
If Not mapinfo.UsaAguatierra Then Exit Sub

With CharList(id)
    If .body.Walk(.heading).GrhIndex Then
        Call Draw_GrhE(.body.Walk(.heading), PixelOffsetX, PixelOffsetY + 32, 1, .Pos.X, .Pos.Y + 1, 1, , 1, .mu)
        
        Call Draw_GrhE(.Head.Head(.heading), PixelOffsetX + .body.HeadOffset.X, PixelOffsetY - .body.HeadOffset.Y, 0, .Pos.X, .Pos.Y + 2, 1, , 1, .mu)
        
        If .casco.Head(.heading).GrhIndex Then _
           Call Draw_GrhE(.casco.Head(.heading), PixelOffsetX + .body.HeadOffset.X, PixelOffsetY - .body.HeadOffset.Y + 16 + GrhData(.casco.Head(.heading).GrhIndex).offsetY * -1, 0, .Pos.X, .Pos.Y + 2, 1, , 1, .mu)
        
        If .escudo.ShieldWalk(.invheading).GrhIndex Then _
            Call Draw_GrhE(.escudo.ShieldWalk(.heading), PixelOffsetX + .body.HeadOffset.X, PixelOffsetY - .body.HeadOffset.Y, 0, .Pos.X, .Pos.Y + 2, 1, , 1, .mu)
        
        If .arma.WeaponWalk(.invheading).GrhIndex Then _
            Call Draw_GrhE(.arma.WeaponWalk(.invheading), PixelOffsetX + .body.HeadOffset.X, PixelOffsetY - .body.HeadOffset.Y, 0, .Pos.X, .Pos.Y + 2, 1, , 1, .mu)
            
    End If
End With
End Sub

Private Sub Char_Move(CharIndex As Integer, nHeading As E_Heading, offsetX As Integer, offsetY As Integer)
    
    With CharList(CharIndex)
        .MoveOffsetX = -1 * (32 * offsetX)
        .MoveOffsetY = -1 * (32 * offsetY)
        
        .Moving = 1
        .heading = nHeading
        .invheading = .heading
        
        If .invh Then
            If .heading = E_Heading.EAST Then
                .invheading = E_Heading.WEST
            ElseIf .heading = E_Heading.WEST Then
                .invheading = E_Heading.EAST
            End If
        End If
            
        .scrollDirectionX = offsetX
        .scrollDirectionY = offsetY
        
        If .Velocidad.X = 0 Then
            .Velocidad.X = ScrollPixelsPerFrameX
            .Velocidad.Y = ScrollPixelsPerFrameY
        End If
        
    End With
    
    Call ActualizarPosicion(CharList(CharIndex), CharList(CharIndex).Pos.X + offsetX, CharList(CharIndex).Pos.Y + offsetY)

    Call DoPasosFx(CharList(CharIndex))

End Sub

Public Sub ActualizarPosicion(caracter As Char, ByVal NuevaX As Integer, ByVal NuevaY As Integer)
    'Lo borramos de la posicion
    If CharMap(caracter.Pos.X, caracter.Pos.Y) = caracter.CharIndex Then
        CharMap(caracter.Pos.X, caracter.Pos.Y) = 0
    End If
     
     If caracter.CharIndex <> UserCharIndex Then
        'Si la distancia es mayor al rango que se visualiza entonces se "borra" el char. Se desactiva su display
        If Abs(CharList(UserCharIndex).Pos.X - NuevaX) > ARangoX Or Abs(CharList(UserCharIndex).Pos.Y - NuevaY) > ARangoY Then
            'Lo desactivamos
            caracter.active = 0
            'Sacamos el texto que esta diciendo
            #If esCLIENTE = 1 Then
                Call Dialogos.RemoveDialog(caracter.CharIndex)
            #End If
            'Terminamos
            Exit Sub
        End If
    End If
    
    caracter.active = 1
    caracter.Pos.X = NuevaX
    caracter.Pos.Y = NuevaY
    'Actualiamos la neuva posicion ya sea mia o del otro charindex
    CharMap(caracter.Pos.X, caracter.Pos.Y) = caracter.CharIndex
End Sub


Public Sub Char_Move_by_Head(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
    Dim addx As Integer
    Dim addy As Integer
      
    Select Case nHeading
        Case E_Heading.NORTH
            addy = -1
    
        Case E_Heading.EAST
            addx = 1
    
        Case E_Heading.SOUTH
            addy = 1
        
        Case E_Heading.WEST
            addx = -1
    End Select

    Call Char_Move(CharIndex, nHeading, addx, addy)

End Sub

Public Sub Char_Move_by_Pos(ByVal CharIndex As Integer, ByVal NX As Long, ByVal NY As Long)
''Marce On error resume next
    Dim addx As Integer
    Dim addy As Integer
    Dim nHeading As E_Heading
    
    With CharList(CharIndex)
        addx = NX - .Pos.X
        addy = NY - .Pos.Y
        
        nHeading = E_Heading.SOUTH
        
        If addx > 0 Then
            nHeading = E_Heading.EAST
        ElseIf addx < 0 Then
            nHeading = E_Heading.WEST
        ElseIf addy < 0 Then
            nHeading = E_Heading.NORTH
        ElseIf addy > 0 Then
            nHeading = E_Heading.SOUTH
        End If
        
    End With
    
    Call Char_Move(CharIndex, nHeading, addx, addy)
    
End Sub

Public Sub EraseChar(ByVal CharIndex As Integer)
    Dim X!, Y!
    
    'Remove char's dialog
    #If esCLIENTE = 1 Then
        Call Dialogos.RemoveDialog(CharIndex)
    #End If
    
    'charlist(CharIndex).hit_act = 0
    Call ResetCharInfo(CharIndex)
        
    'Update NumChars
    NumChars = NumChars - 1
End Sub


Public Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With CharList(CharIndex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With
End Function

Public Function Engine_GetAngle(ByVal centerX As Integer, ByVal centerY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single
'cfnc = fnc.E_Engine_GetAngle
    On Error GoTo errOut
    Dim opp!, adj!, ang1!

    opp = centerY - TargetY
    adj = centerX - TargetX
     
    If (centerX = TargetX) And (centerY = TargetY) Then
        Engine_GetAngle = 0
    Else
        If (adj = 0) Then
            If (opp >= 0) Then
                Engine_GetAngle = 0
            Else
                Engine_GetAngle = 180
            End If
        Else
            ang1 = (Atn(opp / adj)) * RadianToDegree
            If (centerX >= centerX) Then
                Engine_GetAngle = 90 - ang1
            Else
                Engine_GetAngle = 270 - ang1
            End If
        End If
    End If
Exit Function
errOut:
    Engine_GetAngle = 0
End Function

'public double GetAngle1(double x, double y)
'{
'    double rotation = Math.Asin(1 * y / (Math.Sqrt(x*x + y*y)));
'    rotation += Math.PI / 2;
'    if (x < 0.0)
'    {
'        rotation = 2 * Math.PI - rotation;
'    }
'    return rotation;
'}

Public Function Engine_GetAngle2(ByVal centerX As Integer, ByVal centerY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single
On Error GoTo errOut
    Dim DX!, dY!
    DX = centerY - TargetY
    dY = centerX - TargetX
     
    If DX = 0 And dY = 0 Then
        Engine_GetAngle2 = 0
    Else
        Dim rot As Single
        Engine_GetAngle2 = ASin(dY / Sqr(DX * DX + dY * dY)) + (pi / 2)
        If Engine_GetAngle2 < 0 Then
            Engine_GetAngle2 = Pi2 - Engine_GetAngle2
        End If
        Engine_GetAngle2 = Engine_GetAngle2 * RadianToDegree
    End If
Exit Function
errOut:
    Engine_GetAngle2 = 0
End Function

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
    
    ' Fix: por alg�n motivo algunos que no se deberian repetir vienen con loop = 1 o loop = 0
    If Loops = 1 Or Loops = 0 Then
        Loops = 0
    End If
    
    With CharList(CharIndex)
        If fX > 0 Then
            If FxData(fX).particula = 0 Then
                Call InitGrh(.fX, FxData(fX).Animacion)
                .fX.Loops = Loops
                .FxIndex = fX
                If Not .Particle_group(1) Is Nothing Then
                    .Particle_group(1).Matar 0
                End If
            Else
                Set .Particle_group(1) = New Engine_Particle_Group
                
                .Particle_group(1) = FxData(fX).particula
                .Particle_group(1).SetPos .Pos.X, .Pos.Y
                
                If Loops > 0 Then
                    .Particle_group(1).Matar CLng(Loops) * 800
                Else
                    .Particle_group(1).Matar 400
                End If
            End If
            
         '   Debug.Print "PARTICULA"; FxData(fX).particula
            If FxData(fX).wav <> 0 Then Call Sonido_Play(FxData(fX).wav)
        Else
            If Not .Particle_group(0) Is Nothing Then .Particle_group(0).Matar 0
            If Not .Particle_group(1) Is Nothing Then .Particle_group(1).Matar 0
            .FxIndex = fX
        End If
    End With
End Sub

Public Sub PlaySoundFX(ByVal fX As Integer)
    If fX > 0 Then
        If FxData(fX).wav <> 0 Then Call Sonido_Play(FxData(fX).wav)
    End If
End Sub

Public Sub Render_Shadow_Pj(ByVal charid As Integer, Optional alpha As Byte)
    With sombras(0)
        .Color = Alphas(alpha) Or &H10101 * alpha
        .rhw = 1
        .Tamanio = 32
        .v.Y = CharList(charid).MPPos.Y + 28
        .v.X = CharList(charid).MPPos.X + 16
    End With

    DibujarParticulas sombras, 1, TexturaSombra, 0
End Sub

Public Sub Render_Olitas_Pj(ByVal charid As Integer)
    Dim tex As Long
    
    If mapinfo.UsaAguatierra And mapinfo.agua_tileset Then
        tex = Tilesets(mapinfo.agua_tileset).Olitas
    End If
        
    If tex = 0 Then Exit Sub

    With sombras(0)
        .Color = ResultColorArrayAgua(CharList(charid).Pos.X, CharList(charid).Pos.Y)
        .rhw = 1
        .Tamanio = 32
        .v.Y = CharList(charid).MPPos.Y + 28 - mapinfo.agua_profundidad + AlturaPie(CharList(charid).Pos.X, CharList(charid).Pos.Y) + ModSuperWaterDD(CharList(charid).Pos.X, CharList(charid).Pos.Y).hs(0) / 2
        .v.X = CharList(charid).MPPos.X + 16
    End With
    
    DibujarParticulas sombras, 1, tex, 0
End Sub

Public Sub DoPasosFx(ByRef personaje As Char)
    Dim EfectoPisada As Integer
    
    If Not EfectosSonidoActivados Then Exit Sub
    
    If personaje.muerto = True Then Exit Sub
    
    If EstaPCarea(personaje.CharIndex) = False Then Exit Sub
    
    If isGameMaster(personaje) Then Exit Sub
    
    If Not UserNavegando Then
        ' Obtengo el efecto.
        EfectoPisada = mapdata(personaje.Pos.X, personaje.Pos.Y).EfectoPisada
            
        ' Sino tiene... Efecto por defecto
        If EfectoPisada = 0 Then EfectoPisada = 1
        'If EfectoPisada = 0 Then Exit Sub
            
        ' Cambio el pie
        personaje.pie = Not personaje.pie
            
        ' Ejecuto el Sonido
        If personaje.pie Then
            Call Sonido_Play(EfectosPisadas(EfectoPisada).sonido_derecha)
        Else
            Call Sonido_Play(EfectosPisadas(EfectoPisada).sonido_izquierda)
        End If
    Else 'NOT NOT...
        #If esMe = 1 Or esCLIENTE = 1 Then
            Call Sonido_Play(SND_NAVEGANDO)
        #End If
    End If
End Sub
