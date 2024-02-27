Attribute VB_Name = "modCrearPersonaje"
Option Explicit

Public Type tClase
    Nombre As String
    descripcion As String
    id As eClass
    grhId As Integer
End Type

Public Type tRazaPartes
    cabezas() As Integer
    cuerpos() As Integer
    barbas() As Integer
    pelos() As Integer
    ropaInterior() As Integer
End Type

Public Type tRaza
    Nombre As String
    descripcion As String
    id As eRaza
    grhId As Integer
    atributos(1 To 5) As Byte
    personalizacion(1 To 2) As tRazaPartes
End Type

Private clases() As tClase
Private razas() As tRaza

Public DatosCreacion As CrearPersonajeDTO

Public alineaciones(1 To 3) As String

Public Sub initAlineaciones()
    alineaciones(eAlineaciones.Real) = "Ejercito Índigo"
    alineaciones(eAlineaciones.Neutro) = "Rebelde"
    alineaciones(eAlineaciones.caos) = "Ejercito Escarlata"
End Sub


Public Sub initClases()
    ReDim clases(1 To 16) As tClase
    
    clases(1).Nombre = "Asesino"
    clases(1).descripcion = "Gran apuñalador que podrá dejar a su enemigo fuera de combate si logra asestar su puñal con efectividad. Tiene una importante evasión y poco poder mágico."
    clases(1).id = eClass.Assasin
    clases(1).grhId = 20278
    
    clases(2).Nombre = "Bardo"
    clases(2).descripcion = "Habilidosa clase en el combate cuerpo a cuerpo por su gran evasión, con importante ataque mágico."
    clases(2).id = eClass.Bard
    clases(2).grhId = 20277

    clases(3).Nombre = "Cazador"
    clases(3).descripcion = "Especialista en armas a distancia, se oculta entre las sombras y se mueve en ellas oculto. Nadie quiere acercarse a un cazador ya que el impacto de su flecha puede ser derribante. Posee mucha vida, fuerza y defensa física. No usa magias."
    clases(3).id = eClass.Hunter
    clases(3).grhId = 20275
    
    clases(4).Nombre = "Clerigo"
    clases(4).descripcion = "Versátil luchador tanto en artes mágicas como en lucha cuerpo a cuerpo. Especializado en combinar ataque mágico y golpe, hábil en ambos."
    clases(4).id = eClass.Cleric
    clases(4).grhId = 20274
    
    clases(5).Nombre = "Druida"
    clases(5).descripcion = "Domador nato de animales y criaturas de la naturaleza a la que ama, respeta y domina a la perfección, logrando invocar seres naturales que trabajarán a su favor. Utiliza el mimetismo para engañar a sus adversarios y transformarse."
    clases(5).id = eClass.Druid
    clases(5).grhId = 20273
        
    clases(6).Nombre = "Guerrero"
    clases(6).descripcion = "Combatiente cuerpo a cuerpo por excelencia con golpe devastador, con mucha vida y gran fuerza, importante defensa física. Hábil con el arco y la flecha. Se oculta eficientemente. No utiliza magias."
    clases(6).id = eClass.Warrior
    clases(6).grhId = 20272

    clases(7).Nombre = "Mago"
    clases(7).descripcion = "Maneja los secretos ocultos de las artes mágicas. Conjurador de hechizos antiguos y poderosos, con capacidad de concentración y maná más elevado de Tierras del Sur. Realiza ataques mágicos derribantes hacia sus enemigos. Puede conjurar los más poderosos conjuros de estas tierras."
    clases(7).id = eClass.Mage
    clases(7).grhId = 20269
    
    clases(8).Nombre = "Paladin"
    clases(8).descripcion = "Magnífica clase para la lucha cuerpo a cuerpo ya que posee mucha vida y mucha fuerza para enfrentar a sus enemigos. Poco poder mágico y maná."
    clases(8).id = eClass.Paladin
    clases(8).grhId = 20267
    
    clases(9).Nombre = "Pirata"
    clases(9).descripcion = "Rey de los mares, navega prontamente y oculta tesoros en sus barcos logrando así evadir a quienes pretendan arrebatárselos."
    clases(9).id = eClass.Pirat
    clases(9).grhId = 20265
    
    clases(10).Nombre = "Carpintero"
    clases(10).descripcion = "Trabajador creador de elementos en base a madera que servirán para armas y elementos indispensables para la lucha y la navegación."
    clases(10).id = eClass.Carpenter
    clases(10).grhId = 20276
    
    clases(11).Nombre = "Leñador"
    clases(11).descripcion = "Trabajador que conseguirá talar árboles consiguiendo leña efectivamentre para lograr con este elemento lo básico que usará luego el carpintero."
    clases(11).id = eClass.Lumberjack
    clases(11).grhId = 20270
    
    clases(12).Nombre = "Herrero"
    clases(12).descripcion = "Trabajador que fabrica armas, escudos, cascos, herramientas, anillos y armaduras con lingotes de minerales."
    clases(12).id = eClass.Blacksmith
    clases(12).grhId = 20271
    
    clases(13).Nombre = "Minero"
    clases(13).descripcion = "Trabajador que extrae minerales de los yacimientos para construir lingotes que servirán para innumerables armas y armaduras."
    clases(13).id = eClass.Miner
    clases(13).grhId = 20268
    
    clases(14).Nombre = "Pescador"
    clases(14).descripcion = "Trabajador que utiliza la caña de pescar o la red de pesca para obtener peces de los mares para vender o utilizar como alimento."
    clases(14).id = eClass.Fisher
    clases(14).grhId = 20266
    
    clases(15).Nombre = "Recolector"
    clases(15).descripcion = "Trabajador que se ocupa de extraer elementos de la naturaleza como lino y seda necesarios para la confección de vestimentas indispensables."
    clases(15).id = eClass.Recolector
    clases(15).grhId = 29033
    
    clases(16).Nombre = "Sastre"
    clases(16).descripcion = "Trabajador que confeccionará vestimentas indispensables en base a materias primas conseguidas por el recolector."
    clases(16).id = eClass.Sastre
    clases(16).grhId = 29032
End Sub

Public Sub initRazas()
    ReDim razas(1 To 5) As tRaza
    Dim aux As Integer
    
    razas(1).Nombre = "Enano"
    razas(1).descripcion = ""
    razas(1).id = eRaza.Enano
    razas(1).grhId = 20464
    razas(1).atributos(eAtributos.Agilidad) = 18
    razas(1).atributos(eAtributos.Fuerza) = 20
    razas(1).atributos(eAtributos.Constitucion) = 21
    razas(1).atributos(eAtributos.Carisma) = 17
    razas(1).atributos(eAtributos.Inteligencia) = 15
    
    ReDim razas(1).personalizacion(eGenero.Hombre).cabezas(6) As Integer
    ReDim razas(1).personalizacion(eGenero.Hombre).cuerpos(6) As Integer
    ReDim razas(1).personalizacion(eGenero.Hombre).barbas(65) As Integer
    ReDim razas(1).personalizacion(eGenero.Hombre).pelos(65) As Integer
    ReDim razas(1).personalizacion(eGenero.Hombre).ropaInterior(13) As Integer
    
    ReDim razas(1).personalizacion(eGenero.Mujer).cabezas(6) As Integer
    ReDim razas(1).personalizacion(eGenero.Mujer).cuerpos(6) As Integer
    ReDim razas(1).personalizacion(eGenero.Mujer).barbas(65) As Integer
    ReDim razas(1).personalizacion(eGenero.Mujer).pelos(65) As Integer
    ReDim razas(1).personalizacion(eGenero.Mujer).ropaInterior(13) As Integer
    
    With razas(1)
    
        For aux = 1 To 6
            .personalizacion(eGenero.Hombre).cuerpos(aux) = 540 + aux - 1
            .personalizacion(eGenero.Mujer).cuerpos(aux) = 546 + aux - 1
        Next
        
        .personalizacion(eGenero.Hombre).cabezas(1) = 541
        .personalizacion(eGenero.Mujer).cabezas(1) = 557
        .personalizacion(eGenero.Hombre).cabezas(2) = 542
        .personalizacion(eGenero.Mujer).cabezas(2) = 558
        .personalizacion(eGenero.Hombre).cabezas(3) = 543
        .personalizacion(eGenero.Mujer).cabezas(3) = 559
        .personalizacion(eGenero.Hombre).cabezas(4) = 544
        .personalizacion(eGenero.Mujer).cabezas(4) = 560
        .personalizacion(eGenero.Hombre).cabezas(5) = 550
        .personalizacion(eGenero.Mujer).cabezas(5) = 562
        .personalizacion(eGenero.Hombre).cabezas(6) = 554
        .personalizacion(eGenero.Mujer).cabezas(6) = 564
        
        For aux = 1 To 65
            .personalizacion(eGenero.Hombre).barbas(aux) = 5086 + aux - 1
            .personalizacion(eGenero.Hombre).pelos(aux) = 5151 + aux - 1
            .personalizacion(eGenero.Mujer).pelos(aux) = 5216 + aux - 1
        Next
        
        For aux = 1 To 13
            .personalizacion(eGenero.Hombre).ropaInterior(aux) = 5367 + aux - 1
            .personalizacion(eGenero.Mujer).ropaInterior(aux) = 5386 + aux - 1
        Next
    
    End With
        


    razas(2).Nombre = "Humano"
    razas(2).descripcion = ""
    razas(2).id = eRaza.Humano
    razas(2).grhId = 20462
    razas(2).atributos(eAtributos.Agilidad) = 19
    razas(2).atributos(eAtributos.Fuerza) = 19
    razas(2).atributos(eAtributos.Constitucion) = 20
    razas(2).atributos(eAtributos.Carisma) = 18
    razas(2).atributos(eAtributos.Inteligencia) = 18
    
    ReDim razas(2).personalizacion(eGenero.Hombre).cabezas(6) As Integer
    ReDim razas(2).personalizacion(eGenero.Hombre).cuerpos(6) As Integer
    ReDim razas(2).personalizacion(eGenero.Hombre).barbas(65) As Integer
    ReDim razas(2).personalizacion(eGenero.Hombre).pelos(65) As Integer
    ReDim razas(2).personalizacion(eGenero.Hombre).ropaInterior(13) As Integer
    
    ReDim razas(2).personalizacion(eGenero.Mujer).cabezas(6) As Integer
    ReDim razas(2).personalizacion(eGenero.Mujer).cuerpos(6) As Integer
    ReDim razas(2).personalizacion(eGenero.Mujer).barbas(1) As Integer
    ReDim razas(2).personalizacion(eGenero.Mujer).pelos(65) As Integer
    ReDim razas(2).personalizacion(eGenero.Mujer).ropaInterior(13) As Integer
    
    For aux = 1 To 6
        razas(2).personalizacion(eGenero.Hombre).cabezas(aux) = 516 + aux - 1
        razas(2).personalizacion(eGenero.Mujer).cabezas(aux) = 581 + aux - 1
        
        razas(2).personalizacion(eGenero.Hombre).cuerpos(aux) = 510 + aux - 1
        razas(2).personalizacion(eGenero.Mujer).cuerpos(aux) = 516 + aux - 1
    Next
    
    For aux = 1 To 65
        razas(2).personalizacion(eGenero.Hombre).barbas(aux) = 4614 + aux - 1
        razas(2).personalizacion(eGenero.Hombre).pelos(aux) = 4679 + aux - 1
    Next
    
    For aux = 1 To 9
        razas(2).personalizacion(eGenero.Mujer).pelos(aux) = 4954 + aux - 1
    Next
    
    For aux = 1 To 56
        razas(2).personalizacion(eGenero.Mujer).pelos(9 + aux) = 4965 + aux - 1
    Next
    
    For aux = 1 To 13
        razas(2).personalizacion(eGenero.Hombre).ropaInterior(aux) = 5287 + aux - 1
        razas(2).personalizacion(eGenero.Mujer).ropaInterior(aux) = 5310 + aux - 1
    Next
        
    
    razas(3).Nombre = "Gnomo"
    razas(3).descripcion = ""
    razas(3).id = eRaza.Gnomo
    razas(3).grhId = 20463
    razas(3).atributos(eAtributos.Agilidad) = 20
    razas(3).atributos(eAtributos.Fuerza) = 16
    razas(3).atributos(eAtributos.Constitucion) = 18
    razas(3).atributos(eAtributos.Carisma) = 19
    razas(3).atributos(eAtributos.Inteligencia) = 22
    
    ReDim razas(3).personalizacion(eGenero.Hombre).cabezas(6) As Integer
    ReDim razas(3).personalizacion(eGenero.Hombre).cuerpos(6) As Integer
    ReDim razas(3).personalizacion(eGenero.Hombre).barbas(1) As Integer
    ReDim razas(3).personalizacion(eGenero.Hombre).pelos(65) As Integer
    ReDim razas(3).personalizacion(eGenero.Hombre).ropaInterior(13) As Integer
    
    ReDim razas(3).personalizacion(eGenero.Mujer).cabezas(6) As Integer
    ReDim razas(3).personalizacion(eGenero.Mujer).cuerpos(6) As Integer
    ReDim razas(3).personalizacion(eGenero.Mujer).barbas(1) As Integer
    ReDim razas(3).personalizacion(eGenero.Mujer).pelos(65) As Integer
    ReDim razas(3).personalizacion(eGenero.Mujer).ropaInterior(13) As Integer
    
    With razas(3)
    
        For aux = 1 To 6
            .personalizacion(eGenero.Mujer).cabezas(aux) = 575 + aux - 1
            
            .personalizacion(eGenero.Hombre).cuerpos(aux) = 552 + aux - 1
            .personalizacion(eGenero.Mujer).cuerpos(aux) = 558 + aux - 1
        Next
        
        .personalizacion(eGenero.Hombre).cabezas(1) = 565
        .personalizacion(eGenero.Hombre).cabezas(2) = 566
        .personalizacion(eGenero.Hombre).cabezas(3) = 571
        .personalizacion(eGenero.Hombre).cabezas(4) = 572
        .personalizacion(eGenero.Hombre).cabezas(5) = 573
        .personalizacion(eGenero.Hombre).cabezas(6) = 574
        
        For aux = 1 To 65
            .personalizacion(eGenero.Hombre).pelos(aux) = 5021 + aux - 1
        Next
        
        
        For aux = 4744 To 4787
            .personalizacion(eGenero.Mujer).pelos(aux - 4744 + 1) = aux
        Next
        
        .personalizacion(eGenero.Mujer).pelos(44) = 4791
        .personalizacion(eGenero.Mujer).pelos(45) = 4792
        .personalizacion(eGenero.Mujer).pelos(46) = 4798
        .personalizacion(eGenero.Mujer).pelos(47) = 4799
        .personalizacion(eGenero.Mujer).pelos(48) = 4801
        .personalizacion(eGenero.Mujer).pelos(49) = 4802
        
        For aux = 4804 To 4818
            .personalizacion(eGenero.Mujer).pelos(aux - 4804 + 50) = aux
        Next
                
        For aux = 1 To 13
            .personalizacion(eGenero.Hombre).ropaInterior(aux) = 5329 + aux - 1
            .personalizacion(eGenero.Mujer).ropaInterior(aux) = 5348 + aux - 1
        Next
    
    End With

    
    razas(4).Nombre = "Elfo"
    razas(4).descripcion = ""
    razas(4).id = eRaza.Elfo
    razas(4).grhId = 20466
    razas(4).atributos(eAtributos.Agilidad) = 20
    razas(4).atributos(eAtributos.Fuerza) = 18
    razas(4).atributos(eAtributos.Constitucion) = 19
    razas(4).atributos(eAtributos.Carisma) = 20
    razas(4).atributos(eAtributos.Inteligencia) = 20

    ReDim razas(4).personalizacion(eGenero.Hombre).cabezas(6) As Integer
    ReDim razas(4).personalizacion(eGenero.Hombre).cuerpos(6) As Integer
    ReDim razas(4).personalizacion(eGenero.Hombre).barbas(1) As Integer
    ReDim razas(4).personalizacion(eGenero.Hombre).pelos(65) As Integer
    ReDim razas(4).personalizacion(eGenero.Hombre).ropaInterior(13) As Integer
    
    ReDim razas(4).personalizacion(eGenero.Mujer).cabezas(6) As Integer
    ReDim razas(4).personalizacion(eGenero.Mujer).cuerpos(6) As Integer
    ReDim razas(4).personalizacion(eGenero.Mujer).barbas(1) As Integer
    ReDim razas(4).personalizacion(eGenero.Mujer).pelos(65) As Integer
    ReDim razas(4).personalizacion(eGenero.Mujer).ropaInterior(13) As Integer
    
    With razas(4)
    
        For aux = 1 To 6
            .personalizacion(eGenero.Hombre).cabezas(aux) = 522 + aux - 1
            
            .personalizacion(eGenero.Hombre).cuerpos(aux) = 528 + aux - 1
            .personalizacion(eGenero.Mujer).cuerpos(aux) = 534 + aux - 1
        Next
        
        .personalizacion(eGenero.Mujer).cabezas(1) = 528
        .personalizacion(eGenero.Mujer).cabezas(2) = 529
        .personalizacion(eGenero.Mujer).cabezas(3) = 531
        .personalizacion(eGenero.Mujer).cabezas(4) = 532
        .personalizacion(eGenero.Mujer).cabezas(5) = 533
        .personalizacion(eGenero.Mujer).cabezas(6) = 534
        
        For aux = 1 To 65
            .personalizacion(eGenero.Hombre).pelos(aux) = 4819 + aux - 1
        Next
               
        For aux = 1 To 65
            .personalizacion(eGenero.Hombre).pelos(aux) = 4819 + aux - 1
        Next
                       
        For aux = 1 To 58
            .personalizacion(eGenero.Mujer).pelos(aux) = 4884 + aux - 1
        Next
        
        For aux = 59 To 65
            .personalizacion(eGenero.Mujer).pelos(aux) = 4948 - 59 + aux - 1
        Next
        
        For aux = 1 To 13
            .personalizacion(eGenero.Hombre).ropaInterior(aux) = 5437 + aux - 1
            .personalizacion(eGenero.Mujer).ropaInterior(aux) = 5456 + aux - 1
        Next
    
    End With
    
    razas(5).Nombre = "Elfo Oscuro"
    razas(5).descripcion = ""
    razas(5).id = eRaza.ElfoOscuro
    razas(5).grhId = 20465
    razas(5).atributos(eAtributos.Agilidad) = 19
    razas(5).atributos(eAtributos.Fuerza) = 20
    razas(5).atributos(eAtributos.Constitucion) = 19
    razas(5).atributos(eAtributos.Carisma) = 16
    razas(5).atributos(eAtributos.Inteligencia) = 20
    
    ReDim razas(5).personalizacion(eGenero.Hombre).cabezas(3) As Integer
    ReDim razas(5).personalizacion(eGenero.Hombre).cuerpos(3) As Integer
    ReDim razas(5).personalizacion(eGenero.Hombre).barbas(1) As Integer
    ReDim razas(5).personalizacion(eGenero.Hombre).pelos(65) As Integer
    ReDim razas(5).personalizacion(eGenero.Hombre).ropaInterior(13) As Integer
    
    ReDim razas(5).personalizacion(eGenero.Mujer).cabezas(3) As Integer
    ReDim razas(5).personalizacion(eGenero.Mujer).cuerpos(3) As Integer
    ReDim razas(5).personalizacion(eGenero.Mujer).barbas(1) As Integer
    ReDim razas(5).personalizacion(eGenero.Mujer).pelos(65) As Integer
    ReDim razas(5).personalizacion(eGenero.Mujer).ropaInterior(13) As Integer
    
    With razas(5)
    
        For aux = 1 To 3
            .personalizacion(eGenero.Hombre).cabezas(aux) = 535 + aux - 1
            .personalizacion(eGenero.Mujer).cabezas(aux) = 538 + aux - 1
            
            .personalizacion(eGenero.Hombre).cuerpos(aux) = 522 + aux - 1
            .personalizacion(eGenero.Mujer).cuerpos(aux) = 525 + aux - 1
        Next
               
        For aux = 1 To 65
            .personalizacion(eGenero.Hombre).pelos(aux) = 4819 + aux - 1
        Next
               
        For aux = 1 To 65
            .personalizacion(eGenero.Hombre).pelos(aux) = 4819 + aux - 1
        Next
        
        For aux = 4884 To 4942
            .personalizacion(eGenero.Mujer).pelos(aux - 4884 + 1) = aux
        Next
        
        For aux = 4948 To 4953
            .personalizacion(eGenero.Mujer).pelos(59 + 1 + aux - 4948) = aux
        Next
        
        For aux = 1 To 13
            .personalizacion(eGenero.Hombre).ropaInterior(aux) = 5402 + aux - 1
            .personalizacion(eGenero.Mujer).ropaInterior(aux) = 5418 + aux - 1
        Next
    
    End With
End Sub

Public Function getRazaById(id As eRaza) As tRaza
    Dim loopRaza As Byte
    
    For loopRaza = 1 To UBound(razas)
        If razas(loopRaza).id = id Then
            getRazaById = razas(loopRaza)
            Exit Function
        End If
    Next
    
End Function

Public Function getClaseNumero(id As eClass) As Byte
    Dim loopClase As Byte
    
    For loopClase = 1 To UBound(clases)
        If clases(loopClase).id = id Then
            getClaseNumero = loopClase
            Exit Function
        End If
    Next
    
End Function

Public Function getRazaNumero(id As eRaza) As Byte
    Dim loopRaza As Byte
    
    For loopRaza = 1 To UBound(razas)
        If razas(loopRaza).id = id Then
            getRazaNumero = loopRaza
            Exit Function
        End If
    Next
    
End Function

Public Function getClaseById(id As eClass) As tClase
    Dim loopClase As Byte
    
    For loopClase = 1 To UBound(clases)
        If clases(loopClase).id = id Then
            getClaseById = clases(loopClase)
            Exit Function
        End If
    Next
    
End Function
Public Function getClase(numero As Integer) As tClase
    getClase = clases(numero)
End Function

Public Function getRaza(numero As Integer) As tRaza
    getRaza = razas(numero)
End Function

Public Function getCantidadRazas() As Integer
    getCantidadRazas = UBound(razas)
End Function

Public Function getCantidadClases() As Integer
    getCantidadClases = UBound(clases)
End Function


Public Sub crearPersonaje()
    Call crearPersonaje_(DatosCreacion)
End Sub

Private Sub crearPersonaje_(ByVal datos As CrearPersonajeDTO)

Dim key As Integer '
Dim Data As String

key = RandomNumber(97, 122)

Data = ByteToString(datos.Genero) & ByteToString(datos.Clase) & ByteToString(datos.Raza) & ByteToString(datos.Alineacion) & LongToString(MiCuenta.cuenta.id) & ITS(datos.headId) & ITS(datos.bodyId) & ITS(datos.ropaInteriorId) & ITS(datos.barbaId) & ITS(datos.peloId) & datos.Nombre & "," & MD5String(datos.Contraseña)

EnviarPaquete Paquetes.CreatePj, Data

End Sub

