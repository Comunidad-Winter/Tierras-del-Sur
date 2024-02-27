Attribute VB_Name = "ME_Tools_Particulas"
Option Explicit

Private GrupoEditado As Engine_Particle_Group

Public VentanaSelectorParticulas As vw_Part_Select

Public Enum eHerramientasParticulas
    Insertar
    Editar
    Borrar
End Enum

Public infoParticulasSeleccion() As tParticulaSeleccionada

Public Type tParticulaSeleccionada
    particulaSeleccionada(2) As Engine_Particle_Group
End Type

Public herramientaInternaParticula As eHerramientasParticulas

Public Sub iniciarToolsParticulas()
    Dim loopCapa As Byte
    
    ReDim infoParticulasSeleccion(1 To 1, 1 To 1) As tParticulaSeleccionada
    
    For loopCapa = 0 To 2
        Set infoParticulasSeleccion(1, 1).particulaSeleccionada(loopCapa) = Nothing
    Next
    
End Sub
Public Sub EdPar_Grupo_Nuevo() ' Crea un nuevo grupo en la memoria
    Set GrupoEditado = New Engine_Particle_Group
    GrupoEditado.IniciarEdicion().id = -1
End Sub

Public Sub establecerParticula(particula() As Engine_Particle_Group)
    Dim loopCapa As Byte
    ReDim infoParticulasSeleccion(1 To 1, 1 To 1) As tParticulaSeleccionada
    
    For loopCapa = 0 To 2
        Set infoParticulasSeleccion(1, 1).particulaSeleccionada(loopCapa) = particula(loopCapa)
    Next
End Sub

Public Sub ValidarParticulasSeleccionadas()
    Dim i As Byte
    Dim X As Integer, Y As Integer
    
    For X = LBound(infoParticulasSeleccion, 1) To UBound(infoParticulasSeleccion, 1)
        For Y = LBound(infoParticulasSeleccion, 2) To UBound(infoParticulasSeleccion, 2)
        
            With infoParticulasSeleccion(X, Y)
                For i = 0 To 2
                    If Not .particulaSeleccionada(i) Is Nothing Then
                        If .particulaSeleccionada(i).PGID <> 0 Then
                            Set .particulaSeleccionada(i) = Nothing
                        End If
                    End If
                Next i
            End With
        Next Y
    Next X
End Sub
