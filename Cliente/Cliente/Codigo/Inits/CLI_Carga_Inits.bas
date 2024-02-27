Attribute VB_Name = "CLI_Carga_Inits"
Option Explicit

Private Const archivo_compilado_armas = "Armas.ind"
Private Const archivo_compilado_graficos = "Graficos.ind"
Private Const archivo_compilado_pisos = "Pisos.ind"
Private Const archivo_compilado_cabezas = "Cabezas.ind"
Private Const archivo_compilado_cascos = "Cascos.ind"
Private Const archivo_compilado_cuerpos = "Cuerpos.ind"
Private Const archivo_compilado_efectos = "Efectos.ind"
Private Const archivo_compilado_escudos = "Escudos.ind"

Sub Cargar_Armas()

    Dim handle As Integer
    Dim cantidad As Integer
    Dim arma As tIndiceArma
    Dim loopElemento As Integer
    Dim direccion As E_Heading
    
    handle = FreeFile()
    Open IniPath & archivo_compilado_armas For Binary Access Read As handle

    'Obtenemos cantidad
    Get #handle, , cantidad
    
     ' Cargamos los datos
    NumWeaponAnims = cantidad
    
    ' Redimencionamos
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopElemento = 1 To NumWeaponAnims
    
        Get #handle, , arma
        
        For direccion = E_Heading.NORTH To E_Heading.WEST
            InitGrh WeaponAnimData(loopElemento).WeaponWalk(direccion), arma.Walk(direccion), 0
        Next
                         
    Next loopElemento

    Close #handle
End Sub

Public Function Cargar_Graficos() As Boolean
    Dim Frame As Long
    Dim grhCount As Integer
    Dim handle As Integer
    Dim loopElemento As Integer
    Dim NumeroFrames As Integer
    
    handle = FreeFile()
    
    Open IniPath & archivo_compilado_graficos For Binary Access Read As handle
    
    'Get number of grhs
    Get handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(0 To grhCount) As GrhData
    
    For loopElemento = 1 To grhCount
    
        ' Obtenemos la cantidad de frames
        Get #handle, , NumeroFrames
        
        If NumeroFrames > 0 Then
        
            With GrhData(loopElemento)
       
                .NumFrames = NumeroFrames
                
                ReDim .frames(1 To .NumFrames)

                If .NumFrames > 1 Then
                    'Read a animation GRH set
                    For Frame = 1 To .NumFrames
                        Get handle, , .frames(Frame)
                    Next Frame
                
                    Get handle, , .Speed
                              
                    ' Ancho y alto lo tomo del primer frame
                    .pixelHeight = GrhData(.frames(1)).pixelHeight
                    .pixelWidth = GrhData(.frames(1)).pixelWidth
                Else
                    ' Datos
                    Get handle, , .filenum
                    Get handle, , .sx
                    Get handle, , .sy
                    Get handle, , .pixelWidth
                    Get handle, , .pixelHeight
                    Get handle, , .offsetX
                    Get handle, , .offsetY
                    Get handle, , .SombrasSize
                           
                    'Compute width and height
                    .frames(1) = loopElemento
                End If

            End With
         Else
            GrhData(loopElemento).NumFrames = 0
         End If
    
    Next loopElemento
End Function

Public Function Cargar_Pisos() As Boolean

    Dim tileset     As Integer
    Dim num         As Long
    Dim handle      As Integer
    
    handle = FreeFile()

    Open IniPath & archivo_compilado_pisos For Binary Access Read As handle
        
    'Get number of Tilesets
    Get handle, , Tilesets_count
    
    'Resize arrays
    ReDim Tilesets(1 To Tilesets_count)
    
    While Not EOF(handle)
        Get handle, , tileset
        If tileset Then
            With Tilesets(tileset)
                Get handle, , .stage_count
                Get handle, , .anim
                                
                Get handle, , .Olitas
                
                ReDim .stages(1 To .stage_count)
                
                For num = 1 To .stage_count
                    Get handle, , .stages(num)
                Next num
                
                .filenum = .stages(1)
            End With
        End If
    Wend
    
    Cargar_Pisos = True

End Function

Public Function Cargar_Cabezas() As Boolean

    Dim handle      As Integer
    Dim cabeza As tIndiceCabeza
    Dim NumeroCabezas As Integer
    Dim loopElemento As Integer
    Dim direccion As E_Heading
    
    handle = FreeFile
    
    Open IniPath & archivo_compilado_cabezas For Binary Access Read As #handle
            
        Get #handle, , NumeroCabezas
        
        ReDim HeadData(0 To NumeroCabezas) As HeadData
        
        For loopElemento = 1 To NumeroCabezas
            'Leemos la cabeza
            Get #handle, , cabeza
            
            For direccion = E_Heading.NORTH To E_Heading.WEST
                InitGrh HeadData(loopElemento).Head(direccion), cabeza.Head(direccion), 0
            Next
            
        Next loopElemento
            
    Close #handle
    
End Function

Public Function Cargar_Cascos() As Boolean
  
  Dim handle As Integer
  Dim casco As tIndiceCabeza
  Dim NumCascos As Integer
  Dim loopElemento As Integer
  Dim direccion As E_Heading
  
  handle = FreeFile
  
  Open IniPath & archivo_compilado_cascos For Binary Access Read As #handle
  
        'num de cabezas
        Get #handle, , NumCascos
        
        ReDim CascoAnimData(0 To NumCascos) As HeadData
        
        For loopElemento = 1 To NumCascos
            Get #handle, , casco
            
            For direccion = E_Heading.NORTH To E_Heading.WEST
                InitGrh CascoAnimData(loopElemento).Head(direccion), casco.Head(direccion), 0
            Next
        Next
        
  Close #handle
  
  Cargar_Cascos = True
End Function

Public Function Cargar_Cuerpos() As Boolean

    Dim handle As Integer
    Dim cuerpo As tIndiceCuerpo
    Dim NumCuerpos As Integer
    Dim loopElemento As Integer
    Dim direccion As E_Heading
    
    handle = FreeFile
  
    Open IniPath & archivo_compilado_cuerpos For Binary Access Read As #handle
             
        Get #handle, , NumCuerpos
        
        ReDim BodyData(0 To NumCuerpos) As BodyData
        
        For loopElemento = 1 To NumCuerpos
            
            Get #handle, , cuerpo
            For direccion = E_Heading.NORTH To E_Heading.WEST
                InitGrh BodyData(loopElemento).Walk(direccion), cuerpo.body(direccion), 0
            Next
        
            BodyData(loopElemento).HeadOffset.x = cuerpo.HeadOffsetX
            BodyData(loopElemento).HeadOffset.y = cuerpo.HeadOffsetY
        
        Next loopElemento
        
    Close #handle
    
    

End Function


Public Function Cargar_Efectos() As Boolean
    Dim NumFxs As Integer
    Dim loopElemento As Integer
    Dim handle As Integer
    
    handle = FreeFile
  
    Open IniPath & archivo_compilado_efectos For Binary Access Read As #handle
    
        'num de cabezas
        Get #handle, , NumFxs
    
        'Resize array
        ReDim FxData(1 To NumFxs) As tIndiceFx
    
        For loopElemento = 1 To NumFxs
            Get #handle, , FxData(loopElemento)
        Next
    Close #handle

End Function

Public Function Cargar_Escudos() As Boolean

  Dim handle As Integer
  Dim cantidad As Integer
  Dim loopElemento As Integer
  
  Dim escudo As tIndiceEscudo
  Dim direccion As E_Heading
  
  handle = FreeFile
  
  Open IniPath & archivo_compilado_escudos For Binary Access Read As #handle
  
        'num de cabezas
        Get #handle, , cantidad
        
        ReDim ShieldAnimData(0 To cantidad) As ShieldAnimData
        
        For loopElemento = 1 To cantidad
            Get #handle, , escudo
            
            For direccion = E_Heading.NORTH To E_Heading.WEST
                InitGrh ShieldAnimData(loopElemento).ShieldWalk(direccion), escudo.Walk(direccion), 0
            Next
        Next
        
  Close #handle
  
  Cargar_Escudos = True
  
End Function
