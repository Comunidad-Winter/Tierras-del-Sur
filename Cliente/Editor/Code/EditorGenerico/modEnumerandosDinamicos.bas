Attribute VB_Name = "modEnumerandosDinamicos"
Option Explicit

Public Type eEnumerado
    valor As Long
    nombre As String
End Type

Private enumerados() As eEnumerado

Private Function obtenerGraficos(filtros() As String) As eEnumerado()
    Dim enumerados() As eEnumerado
    Dim cantidadvalidos As String
    Dim loopElemento As Long
    Dim elemento As Long
    
    'Filtros: Animacion? Simple? Todos?
    Dim animaciones As Boolean
    Dim simple As Boolean
            
    If UBound(filtros) = 0 Then
        animaciones = True
        simple = True
    Else
        animaciones = False
        simple = False
        
        If filtros(1) = "A" Then
            animaciones = True
        ElseIf filtros(1) = "S" Then
            simple = True
        End If
    End If
            
    'Cuento
    cantidadvalidos = 0
            
    For loopElemento = LBound(GrhData) To UBound(GrhData)
        If Me_indexar_Graficos.existe(loopElemento) And Not GrhData(loopElemento).perteneceAunaAnimacion Then
            If (GrhData(loopElemento).NumFrames = 1 And simple) Or (GrhData(loopElemento).NumFrames > 0 And animaciones) Then
                cantidadvalidos = cantidadvalidos + 1
            End If
        End If
    Next
    
    cantidadvalidos = cantidadvalidos + 1
    
    'Redimensiono
    ReDim enumerados(0 To cantidadvalidos - 1) As eEnumerado
      
    enumerados(0).nombre = "Ninguno"
    enumerados(0).valor = 0

    'Agrego
    elemento = 1
    loopElemento = 0

    Do While elemento < cantidadvalidos
        If Me_indexar_Graficos.existe(loopElemento) And Not GrhData(loopElemento).perteneceAunaAnimacion Then
            If (GrhData(loopElemento).NumFrames = 1 And simple) Or (GrhData(loopElemento).NumFrames > 0 And animaciones) Then
                enumerados(elemento).valor = loopElemento
                enumerados(elemento).nombre = GrhData(loopElemento).nombreGrafico
                elemento = elemento + 1
            End If
        End If
        
        loopElemento = loopElemento + 1
    Loop
    
    obtenerGraficos = enumerados
End Function

Private Function obtenerSonidos(filtros() As String) As eEnumerado()

    Dim enumerados() As eEnumerado
    Dim cantidadvalidos As String
    Dim loopElemento As Long
    Dim elemento As Long
    Dim filtroTipo As Byte
    
    If UBound(filtros) > 0 Then
        filtroTipo = filtros(1)
    Else
        filtroTipo = 255
    End If
    'Cuento
    cantidadvalidos = 0

    For loopElemento = LBound(Sonidos) To UBound(Sonidos)
        If Me_indexar_Sonidos.existe(loopElemento) Then
            'Aplicamos el filtro
            If (Sonidos(loopElemento).tipo = filtroTipo) Or (filtroTipo = 255) Then
                cantidadvalidos = cantidadvalidos + 1
            End If
        End If
    Next
    
    cantidadvalidos = cantidadvalidos + 1 '+ 1El ninguno
        
    'Redimensiono
    ReDim enumerados(0 To cantidadvalidos - 1) As eEnumerado
                       
    enumerados(0).nombre = "Ninguno"
    enumerados(0).valor = 0
    
    'Agrego
    elemento = 1
    loopElemento = 0
            
    Do While elemento < cantidadvalidos
        If Me_indexar_Sonidos.existe(loopElemento) Then
            'Aplicamos el filtro
            If Sonidos(loopElemento).tipo = filtroTipo Or (filtroTipo = 255) Then
                enumerados(elemento).valor = loopElemento
                enumerados(elemento).nombre = Sonidos(loopElemento).nombre
                elemento = elemento + 1
            End If
        End If
                
        loopElemento = loopElemento + 1
    Loop
    
    obtenerSonidos = enumerados
End Function

Private Function obtenerObjetos(filtros() As String) As eEnumerado()
    Dim enumerados() As eEnumerado
    Dim cantidadvalidos As String
    Dim loopElemento As Long
    Dim elemento As Long
    Dim OBJType As Long
    
    'Cuento
    cantidadvalidos = 0
            
    For loopElemento = 1 To UBound(ObjData)
        OBJType = ObjData(loopElemento).OBJType
        If UBound(filtros) = 0 Or in_array(OBJType, filtros) Then cantidadvalidos = cantidadvalidos + 1
    Next
        
    If cantidadvalidos > 0 Then
        cantidadvalidos = cantidadvalidos + 1
        ReDim enumerados(0 To cantidadvalidos - 1) As eEnumerado
          
        enumerados(0).nombre = "Ninguno"
        enumerados(0).valor = 0
        
        'Agrego
        elemento = 1
        loopElemento = 1
        
        Do While elemento < cantidadvalidos
            OBJType = ObjData(loopElemento).OBJType
            If UBound(filtros) = 0 Or in_array(OBJType, filtros) Then
                enumerados(elemento).valor = loopElemento
                enumerados(elemento).nombre = ObjData(loopElemento).Name
                elemento = elemento + 1
            End If
                    
            loopElemento = loopElemento + 1
        Loop
    Else
        ReDim enumerados(0 To 0) As eEnumerado
        enumerados(elemento).valor = 0
        enumerados(elemento).nombre = "Ninguno"
    End If
    obtenerObjetos = enumerados
End Function

Private Function obtenerEfectosPisadas(filtros() As String) As eEnumerado()
    Dim loopElemento As Integer
     
    ReDim enumerados(0 To UBound(EfectosPisadas) + 1) As eEnumerado
     
    enumerados(0).nombre = "Ninguno"
    enumerados(0).valor = 0
    
     For loopElemento = 1 To UBound(EfectosPisadas)
            enumerados(loopElemento).valor = loopElemento
            enumerados(loopElemento).nombre = EfectosPisadas(loopElemento).nombre
     Next
          
     obtenerEfectosPisadas = enumerados
End Function

Private Function obtenerHechizos(filtros() As String) As eEnumerado()
    Dim loopElemento As Integer
     
     
    ReDim enumerados(0 To UBound(HechizosData) + 1) As eEnumerado
     
    enumerados(0).nombre = "Ninguno"
    enumerados(0).valor = 0
    
     For loopElemento = 1 To UBound(HechizosData)
            enumerados(loopElemento).valor = loopElemento
            enumerados(loopElemento).nombre = HechizosData(loopElemento).nombre
     Next
          
     obtenerHechizos = enumerados
End Function

Private Function obtenerArmas(filtros() As String) As eEnumerado()
    Dim enumerados() As eEnumerado
    Dim cantidadvalidos As String
    Dim loopElemento As Long
    Dim elemento As Long
     
    'Cuento
    cantidadvalidos = 0
            
    For loopElemento = LBound(WeaponAnimData) To UBound(WeaponAnimData)
        If Me_indexar_Armas.existe(loopElemento) Then cantidadvalidos = cantidadvalidos + 1
    Next
    
    cantidadvalidos = cantidadvalidos + 1 '+ 1El ninguno
        
    'Redimensiono
    ReDim enumerados(0 To cantidadvalidos - 1) As eEnumerado
           
    enumerados(0).nombre = "Ninguno"
    enumerados(0).valor = 0
    
    'Agrego
    elemento = 1
    loopElemento = 0
            
    Do While elemento < cantidadvalidos
        If Me_indexar_Armas.existe(loopElemento) Then
            enumerados(elemento).valor = loopElemento
            enumerados(elemento).nombre = WeaponAnimData(loopElemento).nombre
            elemento = elemento + 1
        End If
                
        loopElemento = loopElemento + 1
    Loop
    
    obtenerArmas = enumerados
End Function
Private Function obtenerCuerpos(filtros() As String) As eEnumerado()
    Dim enumerados() As eEnumerado
    Dim cantidadvalidos As String
    Dim loopElemento As Long
    Dim elemento As Long
    
    'Cuento
    cantidadvalidos = 0
            
    For loopElemento = LBound(BodyData) To UBound(BodyData)
        If Me_indexar_Cuerpos.existe(loopElemento) Then cantidadvalidos = cantidadvalidos + 1
    Next
    
    cantidadvalidos = cantidadvalidos + 1 '+ 1El ninguno
        
    'Redimensiono
    ReDim enumerados(0 To cantidadvalidos - 1) As eEnumerado
           
    enumerados(0).nombre = "Ninguno"
    enumerados(0).valor = 0
    
    'Agrego
    elemento = 1
    loopElemento = 0
            
    Do While elemento < cantidadvalidos
        If Me_indexar_Cuerpos.existe(loopElemento) Then
            enumerados(elemento).valor = loopElemento
            enumerados(elemento).nombre = BodyData(loopElemento).nombre
            elemento = elemento + 1
        End If
                
        loopElemento = loopElemento + 1
    Loop
    
    obtenerCuerpos = enumerados
End Function

Private Function obtenerCascos(filtros() As String) As eEnumerado()
    Dim enumerados() As eEnumerado
    Dim cantidadvalidos As String
    Dim loopElemento As Long
    Dim elemento As Long
    
    'Cuento
    cantidadvalidos = 0
            
    For loopElemento = LBound(CascoAnimData) To UBound(CascoAnimData)
        If Me_indexar_Cascos.existe(loopElemento) Then cantidadvalidos = cantidadvalidos + 1
    Next
    
    cantidadvalidos = cantidadvalidos + 1 '+ 1El ninguno
        
    'Redimensiono
    ReDim enumerados(0 To cantidadvalidos - 1) As eEnumerado
  
    enumerados(0).nombre = "Ninguno"
    enumerados(0).valor = 0
    
    'Agrego
    elemento = 1
    loopElemento = 0
            
    Do While elemento < cantidadvalidos
        If Me_indexar_Cascos.existe(loopElemento) Then
            enumerados(elemento).valor = loopElemento
            enumerados(elemento).nombre = CascoAnimData(loopElemento).nombre
            elemento = elemento + 1
        End If
                
        loopElemento = loopElemento + 1
    Loop
    
    obtenerCascos = enumerados
End Function
Private Function obtenerEscudos(filtros() As String) As eEnumerado()
    Dim enumerados() As eEnumerado
    Dim cantidadvalidos As String
    Dim loopElemento As Long
    Dim elemento As Long
    
    'Cuento
    cantidadvalidos = 0
            
    For loopElemento = LBound(ShieldAnimData) To UBound(ShieldAnimData)
        If Me_indexar_Escudos.existe(loopElemento) Then cantidadvalidos = cantidadvalidos + 1
    Next
    
    cantidadvalidos = cantidadvalidos + 1 '+ 1El ninguno
        
    'Redimensiono
    ReDim enumerados(0 To cantidadvalidos - 1) As eEnumerado
         
    enumerados(0).nombre = "Ninguno"
    enumerados(0).valor = 0
    
    'Agrego
    elemento = 1
    loopElemento = 0
            
    Do While elemento < cantidadvalidos
        If Me_indexar_Escudos.existe(loopElemento) Then
            enumerados(elemento).valor = loopElemento
            enumerados(elemento).nombre = ShieldAnimData(loopElemento).nombre
            elemento = elemento + 1
        End If
                
        loopElemento = loopElemento + 1
    Loop
    
    obtenerEscudos = enumerados
End Function

Private Function obtenerCabezas(filtros() As String) As eEnumerado()
    Dim enumerados() As eEnumerado
    Dim cantidadvalidos As String
    Dim loopElemento As Long
    Dim elemento As Long
    
    'Cuento
    cantidadvalidos = 0
            
    For loopElemento = LBound(HeadData) To UBound(HeadData)
        If Me_indexar_Cabezas.existe(loopElemento) Then cantidadvalidos = cantidadvalidos + 1
    Next
    
    cantidadvalidos = cantidadvalidos + 1 '+ 1El ninguno
        
    'Redimensiono
    ReDim enumerados(0 To cantidadvalidos - 1) As eEnumerado
          
    enumerados(0).nombre = "Ninguno"
    enumerados(0).valor = 0
    
    'Agrego
    elemento = 1
    loopElemento = 0
            
    Do While elemento < cantidadvalidos
        If Me_indexar_Cabezas.existe(loopElemento) Then
            enumerados(elemento).valor = loopElemento
            enumerados(elemento).nombre = HeadData(loopElemento).nombre
            elemento = elemento + 1
        End If
                
        loopElemento = loopElemento + 1
    Loop
    
    obtenerCabezas = enumerados
End Function

Private Function obtenerImagenes(filtros() As String) As eEnumerado()
    Dim enumerados() As eEnumerado
    Dim cantidadvalidos As Integer
    Dim loopElemento As Long
          
    ReDim enumerados(0 To pakGraficos.getCantidadElementos) As eEnumerado
    
    enumerados(0).nombre = "Ninguno"
    enumerados(0).valor = 0
    
    cantidadvalidos = 0
    For loopElemento = 1 To pakGraficos.getCantidadElementos
       If pakGraficos.Cabezal_GetFileSize(loopElemento) > 0 Then
            cantidadvalidos = cantidadvalidos + 1
            enumerados(cantidadvalidos).valor = loopElemento
            enumerados(cantidadvalidos).nombre = pakGraficos.Cabezal_GetFileNameSinComplementos(loopElemento)
       End If
    Next
    
    ReDim Preserve enumerados(0 To cantidadvalidos)
    
    obtenerImagenes = enumerados
End Function

Private Function obtenerCriaturas(filtros() As String) As eEnumerado()
    Dim enumerados() As eEnumerado
    Dim loopElemento As Long
    Dim elemento As Long
    Dim cantidadvalidos As Integer
    
    'Cuento
    cantidadvalidos = UBound(NpcData)
                        
    cantidadvalidos = cantidadvalidos + 1
    
    'Redimensiono
    ReDim enumerados(0 To cantidadvalidos - 1) As eEnumerado
          
    enumerados(0).nombre = "Ninguno"
    enumerados(0).valor = 0
    
    'Agrego
    elemento = 1
    loopElemento = 1
            
    Do While elemento + 1 < cantidadvalidos
    
            enumerados(elemento).valor = loopElemento
            enumerados(elemento).nombre = NpcData(loopElemento).Name
            elemento = elemento + 1
            
        loopElemento = loopElemento + 1
    Loop
    
    obtenerCriaturas = enumerados
End Function

Private Function obtenerEfectos(filtros() As String) As eEnumerado()
    Dim enumerados() As eEnumerado
    Dim cantidadvalidos As String
    Dim loopElemento As Long
    Dim elemento As Long
    
    'Cuento
    cantidadvalidos = 0
            
    For loopElemento = LBound(FxData) To UBound(FxData)
        If Me_indexar_Efectos.existe(loopElemento) Then cantidadvalidos = cantidadvalidos + 1
    Next
    
    cantidadvalidos = cantidadvalidos + 1
    'Redimensiono
    ReDim enumerados(0 To cantidadvalidos - 1) As eEnumerado
         
    enumerados(0).nombre = "Ninguno"
    enumerados(0).valor = 0
    
    'Agrego
    elemento = 1
    loopElemento = 0
            
    Do While elemento < cantidadvalidos
        If Me_indexar_Efectos.existe(loopElemento) Then
            enumerados(elemento).valor = loopElemento
            enumerados(elemento).nombre = FxData(loopElemento).nombre
            elemento = elemento + 1
        End If
                
        loopElemento = loopElemento + 1
    Loop
    
    obtenerEfectos = enumerados
End Function

Private Function obtenerParticulas(filtros() As String) As eEnumerado()
    Dim enumerados() As eEnumerado
    Dim cantidadvalidos As String
    Dim loopElemento As Long
    Dim elemento As Long
    
    'Cuento
    cantidadvalidos = 0

    For loopElemento = LBound(GlobalParticleGroup) To UBound(GlobalParticleGroup)
        'TODO Falta saber si la particula existe o no
        cantidadvalidos = cantidadvalidos + 1
    Next
    
    cantidadvalidos = cantidadvalidos + 1
    
    'Redimensiono
    ReDim enumerados(0 To cantidadvalidos - 1) As eEnumerado
    
    enumerados(0).nombre = "Ninguno"
    enumerados(0).valor = 0
    
    'Agrego
    elemento = 1
    loopElemento = 0
            
    Do While elemento < cantidadvalidos
        ' If Me_indexar_Sonidos.existe(loopElemento) Then
            enumerados(elemento).valor = loopElemento
            enumerados(elemento).nombre = GlobalParticleGroup(loopElemento).GetNombre
            elemento = elemento + 1
        '  End If
                
        loopElemento = loopElemento + 1
    Loop

    obtenerParticulas = enumerados
End Function


Private Function in_array(ByVal valor As String, vector() As String) As Boolean
    Dim loopElemento As Long
    
    For loopElemento = 1 To UBound(vector)
        If vector(loopElemento) = valor Then
            in_array = True
            Exit Function
        End If
    Next
    
    in_array = False

End Function

Public Function obtenerValorConstante(nombre_ As String) As Long
    Dim nombre As String
    Dim valor As Long

    Dim comienzoOperacion As Byte
    Dim comienzoSuma As Byte
    Dim comienzoResta As Byte
    Dim suma As Long
    Dim resta As Long
    
    'Operacion matematica que se le puede aplicar a lc onstante. Suma o resta
    comienzoSuma = InStr(1, nombre_, "+")
    If comienzoSuma > 0 Then suma = val(mid$(nombre_, comienzoSuma + 1))
    
    comienzoResta = InStr(1, nombre_, "-")
    If comienzoResta > 0 Then resta = val(mid$(nombre_, comienzoResta + 1))
    
    comienzoOperacion = comienzoResta + comienzoSuma
    
    nombre = mid$(nombre_, 1, IIf(comienzoOperacion = 0, Len(nombre_), comienzoOperacion - 1))
    
    'La constante
    Select Case nombre
    
        Case "MAXX"
            valor = SV_Constantes.X_MAXIMO_USABLE
            
        Case "MINX"
            valor = SV_Constantes.X_MINIMO_USABLE
    
        Case "MAXY"
            valor = SV_Constantes.Y_MAXIMO_USABLE
            
        Case "MINY"
            valor = SV_Constantes.Y_MINIMO_USABLE
    End Select
    
    valor = valor + suma - resta

    obtenerValorConstante = valor
End Function
Public Function obtenerEnumeradosDinamicos(nombreEnumerado As String) As eEnumerado()
    Dim nombre As String
    Dim filtros() As String
    
    filtros = Split(nombreEnumerado, ":")

    nombre = filtros(0)
    
    Select Case nombre
        Case "SONIDOS"
            obtenerEnumeradosDinamicos = obtenerSonidos(filtros)
        Case "PARTICULAS"
            obtenerEnumeradosDinamicos = obtenerParticulas(filtros)
        Case "GRAFICOS"
            obtenerEnumeradosDinamicos = obtenerGraficos(filtros)
        Case "HECHIZOS"
            obtenerEnumeradosDinamicos = obtenerHechizos(filtros)
        Case "OBJETOS"
            obtenerEnumeradosDinamicos = obtenerObjetos(filtros)
        Case "CUERPOS"
             obtenerEnumeradosDinamicos = obtenerCuerpos(filtros)
        Case "ARMAS"
            obtenerEnumeradosDinamicos = obtenerArmas(filtros)
        Case "ESCUDOS"
            obtenerEnumeradosDinamicos = obtenerEscudos(filtros)
        Case "CASCOS"
            obtenerEnumeradosDinamicos = obtenerCascos(filtros)
        Case "EFECTOS"
            obtenerEnumeradosDinamicos = obtenerEfectos(filtros)
        Case "CABEZAS"
            obtenerEnumeradosDinamicos = obtenerCabezas(filtros)
        Case "CRIATURAS"
            obtenerEnumeradosDinamicos = obtenerCriaturas(filtros)
        Case "IMAGENES"
            obtenerEnumeradosDinamicos = obtenerImagenes(filtros)
        Case "PISOS"
            'obtenerEnumeradosDinamicos = obtenerPisos(filtros)
        Case "EFECTOS_PISADAS"
            obtenerEnumeradosDinamicos = obtenerEfectosPisadas(filtros)
    End Select
    
End Function

