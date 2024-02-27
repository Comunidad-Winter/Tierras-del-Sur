Attribute VB_Name = "modFacadeModos"
Option Explicit

Public Function crearModo(nombreModo As String) As iModoTorneo

    Select Case nombreModo
    
        Case "DEATHMATCH"
            
            Dim deathmatch As iModoTorneo_DeathMach
            
            Set deathmatch = New iModoTorneo_DeathMach
                        
            Set crearModo = deathmatch
            
        Case "PLAYOFF"
        
            Dim PlayOff As iModoTorneo_PlayOff
            
            Set PlayOff = New iModoTorneo_PlayOff
            
            Set crearModo = PlayOff
        
        Case "LIGA"
    
            Dim Liga As iModoTorneo_Liga
            
            Set Liga = New iModoTorneo_Liga
            
            Set crearModo = Liga
    End Select
    
End Function

