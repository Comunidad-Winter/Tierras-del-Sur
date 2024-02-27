Attribute VB_Name = "Engine_FX"
'ESTE ARCHIVO ESTÁ COMPARTIDO

''
' @require Engine.bas
' @require Engine_Landscape.bas
' @require Engine_Extend.bas

'                  ____________________________________________
'                 /_____/  http://www.arduz.com.ar/ao/   \_____\
'                //            ____   ____   _    _ _____      \\
'               //       /\   |  __ \|  __ \| |  | |___  /      \\
'              //       /  \  | |__) | |  | | |  | |  / /        \\
'             //       / /\ \ |  _  /| |  | | |  | | / /   II     \\
'            //       / ____ \| | \ \| |__| | |__| |/ /__          \\
'           / \_____ /_/    \_\_|  \_\_____/ \____//_____|_________/ \
'           \________________________________________________________/

Option Explicit

Public Type Projectile
    X As Single
    Y As Single
    tx As Single
    ty As Single
    v As Single
    uid As Integer
    Grh As Integer
    life As Single
    luz As Integer
End Type

Public ProjectileList() As Projectile
Public LastProjectile As Integer

Public Type hits
    txt     As String
    Color   As Long
    X       As Single
    Y       As Single
    Vida    As Single
    alpha   As Single
    active  As Byte
    offsetY As Single
End Type

Public HitList() As hits
Public LastHit As Integer


Public Sub FX_Hit_Create(ByVal cid As Integer, ByVal hit As Integer, ByVal Vida As Long, ByVal Color As Long)
    Dim Index As Integer
    Do
        Index = Index + 1
        If Index > LastHit Then
            LastHit = Index
            ReDim Preserve HitList(1 To LastHit)
            Exit Do
        End If
    Loop While HitList(Index).active = 1
    
    HitList(Index).Color = Color
    HitList(Index).txt = CStr(hit)
    HitList(Index).alpha = 255
    HitList(Index).active = 1
    HitList(Index).Vida = Vida + GetTimer
    HitList(Index).X = (CharList(cid).Pos.X)
    HitList(Index).Y = (CharList(cid).Pos.Y) * 32
    HitList(Index).offsetY = 0
End Sub

Public Sub FX_Hit_Create_Pos(ByVal X As Integer, ByVal Y As Integer, ByVal hit As Integer, ByVal Vida As Long, ByVal Color As Long)
    Dim Index As Integer
    Do
        Index = Index + 1
        If Index > LastHit Then
            LastHit = Index
            ReDim Preserve HitList(1 To LastHit)
            Exit Do
        End If
    Loop While HitList(Index).active = 1
    
    HitList(Index).Color = Color
    HitList(Index).txt = CStr(hit)
    HitList(Index).alpha = 255
    HitList(Index).active = 1
    HitList(Index).Vida = Vida + GetTimer
    HitList(Index).X = X
    HitList(Index).Y = Y
    HitList(Index).offsetY = 0
End Sub

Public Sub FX_Hit_Erase(ByVal Index As Integer)
    HitList(Index).active = 0
    If Index = LastHit Then
        Do Until HitList(Index).active = 1
            'Move down one projectile
            LastHit = LastHit - 1
            If LastHit = 0 Then Exit Do
        Loop
        If Index <> LastHit Then
            'We still have projectiles, resize the array to end at the last used slot
            If LastHit > 0 Then
                ReDim Preserve HitList(1 To LastHit)
            Else
                Erase HitList
            End If
        End If
    End If
End Sub

Public Sub FX_Hit_Erase_All()
    If LastHit > 0 Then
        LastHit = 0
        Erase HitList
    End If
End Sub

Public Sub FX_Hit_Render()
Dim X!, Y!, j%
Dim gtc&
gtc = GetTimer

    If LastHit > 0 Then
        For j = 1 To LastHit
            If HitList(j).active Then
                If HitList(j).Vida < gtc Then
                    FX_Hit_Erase j
                Else
                    HitList(j).alpha = HitList(j).alpha - timerTicksPerFrame * 12
                    If HitList(j).alpha > 0 Then
                        With HitList(j)
                            .offsetY = .offsetY + timerTicksPerFrame * 3
                            X = (.X + minXOffset) * 32 + offset_map.X + 16
                            Y = ((.Y + minYOffset) * 32 + offset_map.Y) - .offsetY
                            
                            Engine.Text_Render_alpha CStr(" " & .txt & " "), Y, X, .Color, 1, Abs(.alpha Mod 256)
                        End With
                    Else
                        FX_Hit_Erase j
                    End If
                End If
            End If
        Next j
    End If
End Sub

Public Sub Projectile_Render()
Dim angle!, angle1!, j%, X!, Y!
    If LastProjectile > 0 Then
        For j = 1 To LastProjectile
            If ProjectileList(j).Grh Then
            
                If ProjectileList(j).uid Then
                    angle = Engine_GetAngle(ProjectileList(j).X, ProjectileList(j).Y, CharList(ProjectileList(j).uid).Pos.X * 32, CharList(ProjectileList(j).uid).Pos.Y * 32)
                    ProjectileList(j).X = Interp(ProjectileList(j).X, CharList(ProjectileList(j).uid).Pos.X * 32, ProjectileList(j).life)
                    ProjectileList(j).Y = Interp(ProjectileList(j).Y, CharList(ProjectileList(j).uid).Pos.Y * 32, ProjectileList(j).life)
                Else
                    angle = Engine_GetAngle(ProjectileList(j).X, ProjectileList(j).Y, ProjectileList(j).tx, ProjectileList(j).ty)
                    ProjectileList(j).X = Interp(ProjectileList(j).X, ProjectileList(j).tx, ProjectileList(j).life)
                    ProjectileList(j).Y = Interp(ProjectileList(j).Y, ProjectileList(j).ty, ProjectileList(j).life)
                End If
                
                If ProjectileList(j).luz Then
                    Call DLL_Luces.MovePixel(ProjectileList(j).luz, ProjectileList(j).X, ProjectileList(j).Y)
                End If
                
'                If ProjectileList(j).life < 0.5 Then
'                    ProjectileList(j).Y = ProjectileList(j).Y - timerElapsedTime * ProjectileList(j).v * 0.5
'                End If
                
                ProjectileList(j).life = ProjectileList(j).life + timerElapsedTime * ProjectileList(j).v * 0.0005
'
'                angle1 = Round(180 - angle) ' * DegreeToRadian
'                ProjectileList(j).X = ProjectileList(j).X - Sin(angle1) * timerElapsedTime * ProjectileList(j).v
'                ProjectileList(j).Y = ProjectileList(j).Y + Cos(angle1) * timerElapsedTime * ProjectileList(j).v

                'Draw if within range
                X = ProjectileList(j).X + offset_map.X
                Y = ProjectileList(j).Y + offset_map.Y

                If Y >= -32 Then
                    If Y <= (MainViewHeight + 32) Then
                        If X >= -32 Then
                            If X <= (MainViewWidth + 32) Then
                                Grh_Proyectil ProjectileList(j).Grh, X, Y, , base_light, 180 - angle
                            End If
                        End If
                    End If
                End If
                
                If ProjectileList(j).life >= 1 Then
                    ProjectileList(j).life = 0
                    FX_Projectile_Erase j
                Else
                    If ProjectileList(j).uid Then
                        If Abs(ProjectileList(j).X - CharList(ProjectileList(j).uid).Pos.X * 32) < 10 Then
                            If Abs(ProjectileList(j).Y - CharList(ProjectileList(j).uid).Pos.Y * 32) < 10 Then
                                FX_Projectile_Erase j
                            End If
                        End If
                    Else
                        If Abs(ProjectileList(j).X - ProjectileList(j).tx < 10) Then
                            If Abs(ProjectileList(j).Y - ProjectileList(j).ty) < 10 Then
                                FX_Projectile_Erase j
                            End If
                        End If
                    End If
                End If
            End If
        Next j
    End If
End Sub

Public Sub FX_Projectile_Create(ByVal AttackerIndex As Integer, ByVal TargetIndex As Integer, ByVal GrhIndex As Long, Optional ByVal Velocidad As Single = 1)
Dim ProjectileIndex As Integer

    If AttackerIndex = 0 Then Exit Sub
    If TargetIndex = 0 Then Exit Sub
    If AttackerIndex > UBound(CharList) Then Exit Sub
    If TargetIndex > UBound(CharList) Then Exit Sub

    'Get the next open projectile slot
    Do
        ProjectileIndex = ProjectileIndex + 1
        
        'Update LastProjectile if we go over the size of the current array
        If ProjectileIndex > LastProjectile Then
            LastProjectile = ProjectileIndex
            ReDim Preserve ProjectileList(1 To LastProjectile)
            Exit Do
        End If
        
    Loop While ProjectileList(ProjectileIndex).Grh > 0
    
    'Figure out the initial rotation value
    'ProjectileList(ProjectileIndex).Rotate = Engine_GetAngle(charlist(AttackerIndex).pos.x, charlist(AttackerIndex).pos.y, charlist(TargetIndex).pos.x, charlist(TargetIndex).pos.y)
    
    'Fill in the values
    ProjectileList(ProjectileIndex).uid = TargetIndex
    ProjectileList(ProjectileIndex).ty = Velocidad
    ProjectileList(ProjectileIndex).v = Velocidad
    ProjectileList(ProjectileIndex).Grh = GrhIndex
    ProjectileList(ProjectileIndex).life = 0
    ProjectileList(ProjectileIndex).X = (CharList(AttackerIndex).Pos.X) * 32
    ProjectileList(ProjectileIndex).Y = (CharList(AttackerIndex).Pos.Y) * 32 - CharList(AttackerIndex).OffY
End Sub

Public Sub FX_Projectile_Create_pos(ByVal AttackerIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal GrhIndex As Long, Optional ByVal Velocidad As Single = 1)
Dim ProjectileIndex As Integer

    If AttackerIndex = 0 Then Exit Sub
    If AttackerIndex > UBound(CharList) Then Exit Sub


    'Get the next open projectile slot
    Do
        ProjectileIndex = ProjectileIndex + 1
        
        'Update LastProjectile if we go over the size of the current array
        If ProjectileIndex > LastProjectile Then
            LastProjectile = ProjectileIndex
            ReDim Preserve ProjectileList(1 To LastProjectile)
            Exit Do
        End If
        
    Loop While ProjectileList(ProjectileIndex).Grh > 0
    
    'Figure out the initial rotation value
    'ProjectileList(ProjectileIndex).Rotate = Engine_GetAngle(charlist(AttackerIndex).pos.x, charlist(AttackerIndex).pos.y, charlist(TargetIndex).pos.x, charlist(TargetIndex).pos.y)
    
    'Fill in the values
    With ProjectileList(ProjectileIndex)
    .uid = 0
    .v = Velocidad
    .Grh = GrhIndex
    .tx = X * 32
    .ty = Y * 32
    .life = 0
    .X = (CharList(AttackerIndex).Pos.X) * 32
    .Y = (CharList(AttackerIndex).Pos.Y) * 32
    .luz = DLL_Luces.crear(CharList(AttackerIndex).Pos.X, CharList(AttackerIndex).Pos.Y, 255, 255, 255, 5, 255, 8, 0, 0)
    End With
End Sub


Public Sub FX_Projectile_Erase(ByVal ProjectileIndex As Integer)
    ProjectileList(ProjectileIndex).Grh = 0
    ProjectileList(ProjectileIndex).X = 0
    ProjectileList(ProjectileIndex).Y = 0
    ProjectileList(ProjectileIndex).tx = 0
    ProjectileList(ProjectileIndex).ty = 0
    ProjectileList(ProjectileIndex).uid = 0
    ProjectileList(ProjectileIndex).v = 0
    
    If ProjectileList(ProjectileIndex).luz Then
        DLL_Luces.Quitar ProjectileList(ProjectileIndex).luz
    End If
 
    If ProjectileIndex = LastProjectile Then
        Do Until ProjectileList(ProjectileIndex).Grh > 1
            'Move down one projectile
            LastProjectile = LastProjectile - 1
            If LastProjectile = 0 Then Exit Do
        Loop
        If ProjectileIndex <> LastProjectile Then
            'We still have projectiles, resize the array to end at the last used slot
            If LastProjectile > 0 Then
                ReDim Preserve ProjectileList(1 To LastProjectile)
            Else
                Erase ProjectileList
            End If
        End If
    End If
 
End Sub

Public Sub FX_Projectile_Erase_All()
    If LastProjectile > 0 Then
        LastProjectile = 0
        Erase ProjectileList
    End If
End Sub

Public Sub Engine_CrearEfecto(ByVal CharAracante As Integer, ByVal CharAtacado As Integer, ByVal efecto As Byte)
    Dim entidad As Integer
    exit sub 
    If efecto = 2 Then
        entidad = Engine_Entidades.Entidades_Crear_Indexada(CharList(CharAracante).Pos.X, CharList(CharAracante).Pos.Y, 0, EntidadesIndexadas(1))
    Else
        entidad = Engine_Entidades.Entidades_Crear_Indexada(CharList(CharAracante).Pos.X, CharList(CharAracante).Pos.Y, 0, EntidadesIndexadas(4))
    End If
    
    Engine_Entidades.Entidades_SetCharDestino entidad, CharAtacado
End Sub
