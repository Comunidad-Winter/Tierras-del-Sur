Attribute VB_Name = "Cli_CacheMapas"
Public objMapManager            As clsMemMapManager


Public Sub iniciar()

    Set objMapManager = New clsMemMapManager
    objMapManager.Init 5

End Sub


Public Sub finalizar()

   If Not objMapManager Is Nothing Then
        Call objMapManager.BorrarTodo
        Set objMapManager = Nothing
    End If

End Sub

