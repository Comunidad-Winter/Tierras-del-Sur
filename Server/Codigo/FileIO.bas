Attribute VB_Name = "ES"
Option Explicit

Public Sub CargarSpawnList()

    Dim N As Integer, loopC As Integer
    N = val(GetVar(App.Path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
    ReDim SpawnList(N) As tCriaturasEntrenador
    For loopC = 1 To N
        SpawnList(loopC).npcIndex = val(GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NI" & loopC))
        SpawnList(loopC).NpcName = GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NN" & loopC)
    Next loopC


End Sub

Public Sub CargarHechizos()

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."

Dim hechizo As Integer
Dim Leer As New clsLeerInis
Dim i As Integer

Leer.Abrir DatPath & "Hechizos.dat"

'obtiene el numero de hechizos

NumeroHechizos = val(Leer.DarValor("INIT", "NumeroHechizos"))
ReDim hechizos(1 To NumeroHechizos) As tHechizo

frmCargando.cargar.min = 0
frmCargando.cargar.max = NumeroHechizos
frmCargando.cargar.value = 0

'Llena la lista
For hechizo = 1 To NumeroHechizos

    hechizos(hechizo).id = hechizo
    hechizos(hechizo).nombre = Leer.DarValor("Hechizo" & hechizo, "Nombre")
    hechizos(hechizo).desc = Leer.DarValor("Hechizo" & hechizo, "Desc")
    hechizos(hechizo).PalabrasMagicas = Leer.DarValor("Hechizo" & hechizo, "PalabrasMagicas")
    
    hechizos(hechizo).HechizeroMsg = Leer.DarValor("Hechizo" & hechizo, "HechizeroMsg")
    hechizos(hechizo).TargetMsg = Leer.DarValor("Hechizo" & hechizo, "TargetMsg")
    hechizos(hechizo).PropioMsg = Leer.DarValor("Hechizo" & hechizo, "PropioMsg")
    
    hechizos(hechizo).tipo = val(Leer.DarValor("Hechizo" & hechizo, "Tipo"))
    hechizos(hechizo).WAV = val(Leer.DarValor("Hechizo" & hechizo, "WAV"))
    hechizos(hechizo).FXgrh = val(Leer.DarValor("Hechizo" & hechizo, "Fxgrh"))
    
    hechizos(hechizo).loops = val(Leer.DarValor("Hechizo" & hechizo, "Loops"))
    
   
    hechizos(hechizo).SubeHP = val(Leer.DarValor("Hechizo" & hechizo, "SubeHP"))
    hechizos(hechizo).minHP = val(Leer.DarValor("Hechizo" & hechizo, "MinHP"))
    hechizos(hechizo).MaxHP = val(Leer.DarValor("Hechizo" & hechizo, "MaxHP"))
    

    
    hechizos(hechizo).SubeHam = val(Leer.DarValor("Hechizo" & hechizo, "SubeHam"))
    hechizos(hechizo).minham = val(Leer.DarValor("Hechizo" & hechizo, "MinHam"))
    hechizos(hechizo).MaxHam = val(Leer.DarValor("Hechizo" & hechizo, "MaxHam"))
    
    hechizos(hechizo).SubeSed = val(Leer.DarValor("Hechizo" & hechizo, "SubeSed"))
    hechizos(hechizo).MinSed = val(Leer.DarValor("Hechizo" & hechizo, "MinSed"))
    hechizos(hechizo).MaxSed = val(Leer.DarValor("Hechizo" & hechizo, "MaxSed"))
    
    hechizos(hechizo).SubeAgilidad = val(Leer.DarValor("Hechizo" & hechizo, "SubeAG"))
    hechizos(hechizo).MinAgilidad = val(Leer.DarValor("Hechizo" & hechizo, "MinAG"))
    hechizos(hechizo).MaxAgilidad = val(Leer.DarValor("Hechizo" & hechizo, "MaxAG"))
    
    hechizos(hechizo).SubeFuerza = val(Leer.DarValor("Hechizo" & hechizo, "SubeFU"))
    hechizos(hechizo).MinFuerza = val(Leer.DarValor("Hechizo" & hechizo, "MinFU"))
    hechizos(hechizo).MaxFuerza = val(Leer.DarValor("Hechizo" & hechizo, "MaxFU"))
    
        
    
    hechizos(hechizo).Invisibilidad = val(Leer.DarValor("Hechizo" & hechizo, "Invisibilidad"))
    hechizos(hechizo).Paraliza = val(Leer.DarValor("Hechizo" & hechizo, "Paraliza"))
    hechizos(hechizo).Inmoviliza = val(Leer.DarValor("Hechizo" & hechizo, "Inmoviliza"))
    hechizos(hechizo).RemoverParalisis = val(Leer.DarValor("Hechizo" & hechizo, "RemoverParalisis"))
    hechizos(hechizo).AgiUpAndFuer = val(Leer.DarValor("Hechizo" & hechizo, "AgiUpAndFuer"))
    hechizos(hechizo).MinAgiFuer = val(Leer.DarValor("Hechizo" & hechizo, "MinAgiFuer"))
    hechizos(hechizo).MaxAgiFuer = val(Leer.DarValor("Hechizo" & hechizo, "MaxAgiFuer"))
    hechizos(hechizo).RemueveInvisibilidadParcial = val(Leer.DarValor("Hechizo" & hechizo, "RemueveInvisibilidadParcial"))
    hechizos(hechizo).CuraVeneno = val(Leer.DarValor("Hechizo" & hechizo, "CuraVeneno"))
    hechizos(hechizo).Envenena = val(Leer.DarValor("Hechizo" & hechizo, "Envenena"))
    hechizos(hechizo).Revivir = val(Leer.DarValor("Hechizo" & hechizo, "Revivir"))
        
    hechizos(hechizo).NumNpc = val(Leer.DarValor("Hechizo" & hechizo, "NumNpc"))
    hechizos(hechizo).cant = val(Leer.DarValor("Hechizo" & hechizo, "Cant"))
    hechizos(hechizo).Mimetiza = val(Leer.DarValor("hechizo" & hechizo, "Mimetiza"))
    
    
    
    hechizos(hechizo).MinSkill = val(Leer.DarValor("Hechizo" & hechizo, "MinSkill"))
    hechizos(hechizo).ManaRequerido = val(Leer.DarValor("Hechizo" & hechizo, "ManaRequerido"))
    hechizos(hechizo).ManaRequeridoPaladin = val(Leer.DarValor("Hechizo" & hechizo, "ManaRequeridoPaladin"))
    hechizos(hechizo).ManaRequeridoAsesino = val(Leer.DarValor("Hechizo" & hechizo, "ManaRequeridoAsesino"))
    hechizos(hechizo).ManaRequeridoBardo = val(Leer.DarValor("Hechizo" & hechizo, "ManaRequeridoBardo"))
    
    If hechizos(hechizo).ManaRequeridoAsesino = 0 Then
        hechizos(hechizo).ManaRequeridoAsesino = hechizos(hechizo).ManaRequerido
    End If
    
    If hechizos(hechizo).ManaRequeridoPaladin = 0 Then
        hechizos(hechizo).ManaRequeridoPaladin = hechizos(hechizo).ManaRequerido
    End If
    
    If hechizos(hechizo).ManaRequeridoBardo = 0 Then
        hechizos(hechizo).ManaRequeridoBardo = hechizos(hechizo).ManaRequerido
    End If
    
    hechizos(hechizo).StaRequerido = val(Leer.DarValor("Hechizo" & hechizo, "StaRequerido"))
    
    hechizos(hechizo).Target = val(Leer.DarValor("Hechizo" & hechizo, "Target"))
    
    hechizos(hechizo).NeedStaff = val(Leer.DarValor("Hechizo" & hechizo, "NeedStaff"))
    hechizos(hechizo).StaffAffected = CBool(val(Leer.DarValor("Hechizo" & hechizo, "StaffAffected")))
    
    For i = 1 To NUMCLASES
      hechizos(hechizo).ClaseProhibida(i) = claseToByte(Leer.DarValor("Hechizo" & hechizo, "CP" & i))
    Next
    
    hechizos(hechizo).manaPenalidad = val(Leer.DarValor("Hechizo" & hechizo, "ManaPenalidad"))
    
    frmCargando.cargar.value = frmCargando.cargar.value + 1
Next
 
End Sub

Sub LoadMotd()
Dim i As Integer

MaxLines = val(GetVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines"))
ReDim MOTD(1 To MaxLines)
For i = 1 To MaxLines
    MOTD(i).texto = GetVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & i)
    MOTD(i).formato = ""
Next i

End Sub

Public Sub DoBackUp()

Call LogBackup("Iniciando Backup")

haciendoBK = True

EnviarPaquete Paquetes.Pausa, "", 0, ToAll

Call GuardarTodosLosUsuarios(True) ' Guardamos todos los personajes

Call NPCs.eliminarTodasLasMascotas

Call WorldSave

EnviarPaquete Paquetes.Pausa, "", 0, ToAll

haciendoBK = False

Call LogBackup("Finaliznado Backup")

End Sub


Public Sub SaveMapData(ByVal N As Integer)
End Sub

Sub LoadArmasHerreria()

Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))

ReDim Preserve ArmasHerrero(1 To N) As Integer

For lc = 1 To N
    ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
Next lc


End Sub

Sub LoadArmadurasHerreria()

Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))

ReDim Preserve ArmadurasHerrero(1 To N) As Integer

For lc = 1 To N
    ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
Next lc

End Sub

Sub LoadObjCarpintero()
Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))
ReDim Preserve ObjCarpintero(1 To N) As Integer
For lc = 1 To N
    ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
Next lc
End Sub


Sub GuardarOBJData()

'*****************************************************************
'Carga la lista de objetos
'*****************************************************************
Dim ObjNumero As Integer
Dim objfile As String

Dim archivo As cIniManager

objfile = App.Path & "/" & "objetos.dat"

Set archivo = New cIniManager
Call archivo.Initialize(objfile)


For ObjNumero = 1 To NumObjDatas

    With ObjData(ObjNumero)


        '*************************************************************************
        Call archivo.ChangeValue(ObjNumero, "Name", .Name) ' ok
        Call archivo.ChangeValue(ObjNumero, "ObjType", .ObjType) ' ok
        Call archivo.ChangeValue(ObjNumero, "Subtipo", .subTipo) ' ok
        Call archivo.ChangeValue(ObjNumero, "GrhIndex", .GrhIndex) ' ok
        
        Call archivo.ChangeValue(ObjNumero, "Newbie", .Newbie) 'ok
        Call archivo.ChangeValue(ObjNumero, "Crucial", IIf(.Crucial = 0, 1, 0)) 'ok
        Call archivo.ChangeValue(ObjNumero, "Agarrable", IIf(.Agarrable = 0, 1, 0)) ' ok
        Call archivo.ChangeValue(ObjNumero, "NoSeCae", IIf(.SeCae = 1, 0, 1)) 'ok
        Call archivo.ChangeValue(ObjNumero, "Ubicable", .Ubicable) 'ok
        
        Call archivo.ChangeValue(ObjNumero, "Texto", .texto) 'ok
        Call archivo.ChangeValue(ObjNumero, "grhSecundario", .GrhSecundario) 'ok
        Call archivo.ChangeValue(ObjNumero, "Valor", .valor) 'ok
        Call archivo.ChangeValue(ObjNumero, "SkHerreria", .SkHerreria) 'ok
        Call archivo.ChangeValue(ObjNumero, "SkCarpinteria", .SkCarpinteria) 'ok
        
        'Solo valido para arboles
        Call archivo.ChangeValue(ObjNumero, "Leña", 0) 'ok
                
        'Solo valido para objetos construibles en carpinteria
        'Call archivo.ChangeValue(ObjNumero, "Madera", .Madera) 'ok
        'Call archivo.ChangeValue(ObjNumero, "MaderaT", .MaderaT) 'ok
        
        'Solo valido para objetos construibles en herreria
        'Call archivo.ChangeValue(ObjNumero, "LingH", .LingH) 'ok
        'Call archivo.ChangeValue(ObjNumero, "LingP", .LingP) 'ok
        'Call archivo.ChangeValue(ObjNumero, "LingO", .LingO) 'ok
            
        'Solo valido para armas
        Call archivo.ChangeValue(ObjNumero, "Apuñala", .Apuñala) 'ok
        Call archivo.ChangeValue(ObjNumero, "Proyectil", .proyectil) 'ok
        Call archivo.ChangeValue(ObjNumero, "Municiones", .Municion) 'ok
        Call archivo.ChangeValue(ObjNumero, "Refuerzo", .Refuerzo) 'ok
        Call archivo.ChangeValue(ObjNumero, "QuitaEnergia", .QuitaEnergia) 'ok
        Call archivo.ChangeValue(ObjNumero, "SkillCombate", .SkillCombate)
        Call archivo.ChangeValue(ObjNumero, "WeaponAnim", .WeaponAnim)
        
        'Solo valido para armas
        Call archivo.ChangeValue(ObjNumero, "WeaponAnim", .WeaponAnim)
        
        'Solo valido para armas y flechas
        Call archivo.ChangeValue(ObjNumero, "Envenena", .Envenena) 'ok
        Call archivo.ChangeValue(ObjNumero, "StaffPower", .StaffPower) 'ok
        Call archivo.ChangeValue(ObjNumero, "StaffDamageBonus", .StaffDamageBonus) 'ok
        
        'Solo valido para armas y barcos y flechas
        Call archivo.ChangeValue(ObjNumero, "MaxHIT", .MaxHIT) 'ok
        Call archivo.ChangeValue(ObjNumero, "MinHIT", .MinHIT) 'ok
        
        'Solo valido apra escudos
        Call archivo.ChangeValue(ObjNumero, "SkillDefe", .SkillDefe)
        Call archivo.ChangeValue(ObjNumero, "ShieldAnim", .ShieldAnim) 'ok
        
        'Solo valido para cascos
        Call archivo.ChangeValue(ObjNumero, "CascoAnim", .CascoAnim)
        
        'Solo valido para comida
        Call archivo.ChangeValue(ObjNumero, "MinHam", .minham)
         
        'Solo valido para bebida
        Call archivo.ChangeValue(ObjNumero, "MinSed", .MinSed)
        
        'Solo valido para pergaminos.
        Call archivo.ChangeValue(ObjNumero, "HechizoIndex", .HechizoIndex) 'ok
        
        'Solo valido para yacimientos
        Call archivo.ChangeValue(ObjNumero, "MineralIndex", .MineralIndex) 'ok
        
        'Solo valido para minerales
        Call archivo.ChangeValue(ObjNumero, "LingoteIndex", .LingoteIndex) 'ok
        
        'Solo valido para barcos y minerales
        Call archivo.ChangeValue(ObjNumero, "MinSkill", .MinSkill) 'ok
        
        'Solo valido para armaduras / ropas
        Call archivo.ChangeValue(ObjNumero, "Ropaje", .Ropaje) 'ok
        
        'Solo valido para cascos y anillos
        Call archivo.ChangeValue(ObjNumero, "DefensaMagicaMax", .DefensaMagicaMax)
        Call archivo.ChangeValue(ObjNumero, "DefensaMagicaMin", .DefensaMagicaMin)
        
        'Solo valido para elementos que necesitan de skills en MAGIA
        Call archivo.ChangeValue(ObjNumero, "SkillM", .SkillM)
        
        'Skills en mineria
        Call archivo.ChangeValue(ObjNumero, "SkillMin", .SkillMin)
        
        'Solo valido para pociones.
        Call archivo.ChangeValue(ObjNumero, "TipoPocion", .TipoPocion) 'ok
        Call archivo.ChangeValue(ObjNumero, "MaxModificador", .MaxModificador) 'ok
        Call archivo.ChangeValue(ObjNumero, "MinModificador", .MinModificador) 'ok
        Call archivo.ChangeValue(ObjNumero, "DuracionEfecto", .DuracionEfecto) 'ok
        
        'Solo valido para Ropa, Barca,  Armadura, Casco, Sombrero y Escudo.
        Call archivo.ChangeValue(ObjNumero, "MINDEF", .MinDef)
        Call archivo.ChangeValue(ObjNumero, "MAXDEF", .MaxDef)
        
        'Se usa un “Laud Mágico” cuando se le da a la “U”. Al usar un anillo. Al usar un tipo de objeto INSTRUMENTOS. Al usar un tipo de objeto BOTELLA_LLENA
        Call archivo.ChangeValue(ObjNumero, "SND1", .Snd1) 'ok
        
                
        'T es para cascos y sin T es para armaduras
        'If .SkillTacticassT > 0 Then
        '    Call archivo.ChangeValue(ObjNumero, "SkillTacticas", .SkillTacticassT)
       ' ElseIf .SkillTacticass > 0 Then
        '    Call archivo.ChangeValue(ObjNumero, "SkillTacticas", .SkillTacticassT)
        'Else
        '    Call archivo.ChangeValue(ObjNumero, "SkillTacticas", 0)
        'End If
    '*************************************************************************'
        
        ' Sistema de llaves y objetos que se transforman
        Call archivo.ChangeValue(ObjNumero, "Cerrada", .Cerrada)
        
        Call archivo.ChangeValue(ObjNumero, "Llave", .Llave)
        Call archivo.ChangeValue(ObjNumero, "Clave", .clave)
        
        'Puertas, botella vacia y botella llena
        Call archivo.ChangeValue(ObjNumero, "IndexAbierta", .IndexAbierta)
        Call archivo.ChangeValue(ObjNumero, "IndexCerrada", .IndexCerrada)
        Call archivo.ChangeValue(ObjNumero, "IndexCerradaLlave", .IndexCerradaLlave)
       
        'Puertas y llaves
        Call archivo.ChangeValue(ObjNumero, "Clave", .clave)
                
        '**************************************************************************
        '************************ TRANSFORMACIONES ********************************
        Dim i As Integer
        Dim j As Integer
        Dim tempString As String
        
        
        'Clases permitidas
        Dim prohibida As Boolean
            
        tempString = ""
        
        'Por cada clase me fjo si esta prohibida o no en el item
        For i = 1 To NUMCLASES - 1
    
            prohibida = False
    
            'La clase esta prohibida?
            For j = 1 To NUMCLASES - 1
                    
                   ' If .ClaseProhibida(j) = i Then
                   '         prohibida = True
                   '         Exit For
                  '  End If
            Next j
    
            'Sino esta prohibida la agrego
            If Not prohibida Then
                tempString = tempString & (2 ^ (claseToByte(modClases.clasesConfig(i).nombre) - 1)) & ","
            End If
            
        Next
                    
        If Not tempString = "" Then
            tempString = mid$(tempString, 1, Len(tempString) - 1)
        End If
             
             
        Call archivo.ChangeValue(ObjNumero, "Clases", tempString)
        
        'Sexos
        'If (.Hombre = 1 And .Mujer = 1) Or (.Hombre = 0 And .Mujer = 0) Then
        '    tempString = "1,2"
       ' ElseIf .Hombre = 1 Then
       '     tempString = "1"
       ' Else
       '     tempString = "2"
       ' End If
        'Call archivo.ChangeValue(ObjNumero, "Sexos", tempString)
       
        'Razas
        'If .RazaEnana = 1 Then
        '    tempString = "2,16"
        'Else
        '    tempString = "1,2,4,8,16"
        'End If
        
        Call archivo.ChangeValue(ObjNumero, "Razas", tempString)
        
        'Alineacion
        'If .Real = 1 Then
        '    Call archivo.ChangeValue(ObjNumero, "Alineacion", 2)
        'ElseIf .caos = 1 Then
        '    Call archivo.ChangeValue(ObjNumero, "Alineacion", 4)
        'Else
        '    Call archivo.ChangeValue(ObjNumero, "Alineacion", 0)
        'End If
        '*************************************************************************
    End With
Next ObjNumero

Call archivo.DumpFile(objfile)

End Sub

Sub LoadOBJData()

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."

'*****************************************************************
'Carga la lista de objetos
'*****************************************************************

Dim m_iniFile As cIniManager
Dim ultimo As Integer
Dim Object As Integer
    
Set m_iniFile = New cIniManager
    
m_iniFile.Initialize DatPath & "\objetos.dat"
    
ultimo = CInt(val(m_iniFile.getNameLastSection))
NumObjDatas = ultimo

frmCargando.cargar.min = 0
frmCargando.cargar.max = ultimo
frmCargando.cargar.value = 0

ReDim Preserve ObjData(1 To ultimo) As ObjData

'Llena la lista
For Object = 1 To ultimo

        ObjData(Object).Name = m_iniFile.getValue(Object, "Name") ' OK
        
        ' Call LogLenguaje(ObjData(Object).Name)

        ObjData(Object).ObjType = val(m_iniFile.getValue(Object, "ObjType")) ' OK
        ObjData(Object).subTipo = val(m_iniFile.getValue(Object, "Subtipo")) ' OK
        ObjData(Object).tier = val(m_iniFile.getValue(Object, "Tier")) ' OK
        ObjData(Object).GrhIndex = val(m_iniFile.getValue(Object, "GrhIndex")) ' OK
        
        ObjData(Object).Crucial = val(m_iniFile.getValue(Object, "Crucial")) ' OK
        
        ObjData(Object).Newbie = val(m_iniFile.getValue(Object, "Newbie")) ' OK
        
        ObjData(Object).Ubicable = val(m_iniFile.getValue(Object, "Ubicable")) ' OK

        ObjData(Object).Agarrable = val(m_iniFile.getValue(Object, "Agarrable")) ' OK
        
        ObjData(Object).SeCae = val(m_iniFile.getValue(Object, "NoSeCae")) ' OK
        
        ObjData(Object).HechizoIndex = val(m_iniFile.getValue(Object, "HechizoIndex")) ' OK
        ObjData(Object).MineralIndex = val(m_iniFile.getValue(Object, "MineralIndex")) ' OK
        ObjData(Object).LingoteIndex = val(m_iniFile.getValue(Object, "LingoteIndex")) ' OK
        ObjData(Object).LeñaIndex = val(m_iniFile.getValue(Object, "Leña")) ' OK
        ObjData(Object).Ropaje = val(m_iniFile.getValue(Object, "Ropaje"))  ' OK
        
        ObjData(Object).GrhSecundario = val(m_iniFile.getValue(Object, "VGrande")) ' OK
        
        ObjData(Object).minham = val(m_iniFile.getValue(Object, "MinHam")) ' OK
        ObjData(Object).MinSed = val(m_iniFile.getValue(Object, "MinSed")) ' OK
 
        ObjData(Object).MinDef = val(m_iniFile.getValue(Object, "MINDEF")) ' OK
        ObjData(Object).MaxDef = val(m_iniFile.getValue(Object, "MAXDEF")) ' OK
      
        ObjData(Object).valor = val(m_iniFile.getValue(Object, "Valor")) ' OK
         
        ObjData(Object).SkCarpinteria = val(m_iniFile.getValue(Object, "SkCarpinteria")) ' OK
        ObjData(Object).SkillM = val(m_iniFile.getValue(Object, "SkillM")) ' OK
         
        ObjData(Object).texto = m_iniFile.getValue(Object, "Texto") ' OK
         
        ObjData(Object).Cerrada = val(m_iniFile.getValue(Object, "Cerrada"))
        
        ObjData(Object).clave = val(m_iniFile.getValue(Object, "Clave"))
        
        ObjData(Object).alineacion = stringToBinaryLong(m_iniFile.getValue(Object, "Alineacion")) ' OK
                
        ObjData(Object).Envenena = val(m_iniFile.getValue(Object, "Envenena")) ' OK
        
        If ObjData(Object).Cerrada = 1 Then
            ObjData(Object).Llave = val(m_iniFile.getValue(Object, "Llave"))
            ObjData(Object).clave = val(m_iniFile.getValue(Object, "Clave")) ' OK
        End If
        
        ObjData(Object).StaffDamageBonus = val(m_iniFile.getValue(Object, "StaffDamageBonus")) ' OK
                    
        If ObjData(Object).ObjType = OBJTYPE_WEAPON Then
            ObjData(Object).Apuñala = val(m_iniFile.getValue(Object, "Apuñala")) ' OK
            ObjData(Object).proyectil = val(m_iniFile.getValue(Object, "Proyectil")) ' OK
            ObjData(Object).Refuerzo = val(m_iniFile.getValue(Object, "Refuerzo")) ' OK
            ObjData(Object).QuitaEnergia = val(m_iniFile.getValue(Object, "QuitaEnergia")) ' OK
            ObjData(Object).WeaponAnim = val(m_iniFile.getValue(Object, "WEAPONANIM")) ' OK
            ObjData(Object).SkillCombate = val(m_iniFile.getValue(Object, "SkillCombate")) ' OK
            ObjData(Object).StaffPower = val(m_iniFile.getValue(Object, "StaffPower")) ' OK

            ObjData(Object).MinHIT = val(m_iniFile.getValue(Object, "MinHIT")) ' OK
            ObjData(Object).MaxHIT = val(m_iniFile.getValue(Object, "MaxHIT")) ' OK
            ' FALTAN
            ObjData(Object).SkHerreria = val(m_iniFile.getValue(Object, "SkHerreria"))
            ObjData(Object).Municion = val(m_iniFile.getValue(Object, "Municiones"))
        End If
                
        ' OK
        If ObjData(Object).ObjType = OBJTYPE_PUERTAS Or ObjData(Object).ObjType = OBJTYPE_BOTELLAVACIA Or ObjData(Object).ObjType = OBJTYPE_BOTELLALLENA Then
            ObjData(Object).IndexAbierta = val(m_iniFile.getValue(Object, "IndexAbierta")) ' OK
            ObjData(Object).IndexCerrada = val(m_iniFile.getValue(Object, "IndexCerrada")) ' OK
            ObjData(Object).IndexCerradaLlave = val(m_iniFile.getValue(Object, "IndexCerradaLlave")) ' OK
        End If
        
        
        If ObjData(Object).subTipo = OBJTYPE_ESCUDO Then
            ObjData(Object).ShieldAnim = val(m_iniFile.getValue(Object, "SHIELDANIM")) ' OK
            ObjData(Object).SkillDefe = val(m_iniFile.getValue(Object, "SkillDefe")) ' OK
            ObjData(Object).SkHerreria = val(m_iniFile.getValue(Object, "SkHerreria")) ' OK
        End If
    
        If ObjData(Object).subTipo = OBJTYPE_CASCO Then
            ObjData(Object).CascoAnim = val(m_iniFile.getValue(Object, "CASCOANIM")) ' OK
            ObjData(Object).SkHerreria = val(m_iniFile.getValue(Object, "SkHerreria")) ' OK
            
            ObjData(Object).SkillTacticass = val(m_iniFile.getValue(Object, "SKILLTACTICAS"))
        End If
    
    
        Dim tacticas As Long
                
        tacticas = val(m_iniFile.getValue(Object, "SKILLTACTICAS"))
        
        If tacticas = 0 Then
            tacticas = val(m_iniFile.getValue(Object, "SKILLTACTICASS"))
        End If
        
        ObjData(Object).SkillTacticass = tacticas
               
        If ObjData(Object).ObjType = OBJTYPE_ARMOUR Then
            ObjData(Object).SkHerreria = val(m_iniFile.getValue(Object, "SkHerreria")) ' OK
        End If
    

        ObjData(Object).SkHerreria = val(m_iniFile.getValue(Object, "SkHerreria")) ' OK
  
 
        If ObjData(Object).ObjType = OBJTYPE_HERRAMIENTAS Then
            ObjData(Object).SkillMin = val(m_iniFile.getValue(Object, "SkillMin")) ' OK
            ObjData(Object).SkHerreria = val(m_iniFile.getValue(Object, "SkHerreria")) ' OK
        End If
    
        If ObjData(Object).ObjType = OBJTYPE_INSTRUMENTOS Then
            ObjData(Object).Snd1 = val(m_iniFile.getValue(Object, "SND1")) ' OK
        End If
    
        
    
        If ObjData(Object).ObjType = OBJTYPE_BARCOS Or ObjData(Object).ObjType = OBJTYPE_MINERALES Then
            ObjData(Object).MinSkill = val(m_iniFile.getValue(Object, "MinSkill")) ' OK
        End If
    
               
        ObjData(Object).Genero = stringToBinaryLong(m_iniFile.getValue(Object, "Sexos")) ' OK
        ObjData(Object).clasesPermitidas = stringToBinaryLong(m_iniFile.getValue(Object, "Clases")) ' OK
        ObjData(Object).razas = stringToBinaryLong(m_iniFile.getValue(Object, "Razas")) ' OK

        ObjData(Object).DefensaMagicaMax = val(m_iniFile.getValue(Object, "DefensaMagicaMax")) ' OK
        ObjData(Object).DefensaMagicaMin = val(m_iniFile.getValue(Object, "DefensaMagicaMin")) ' OK
        
        ObjData(Object).MaxHIT = val(m_iniFile.getValue(Object, "MaxHIT")) ' OK
        ObjData(Object).MinHIT = val(m_iniFile.getValue(Object, "MinHIT")) ' OK
            
        'Pociones
        If ObjData(Object).ObjType = OBJTYPE_POCIONES Then ' OK
            ObjData(Object).TipoPocion = val(m_iniFile.getValue(Object, "TipoPocion")) ' OK
            ObjData(Object).MaxModificador = val(m_iniFile.getValue(Object, "MaxModificador")) ' OK
            ObjData(Object).MinModificador = val(m_iniFile.getValue(Object, "MinModificador")) ' OK
            ObjData(Object).DuracionEfecto = val(m_iniFile.getValue(Object, "DuracionEfecto")) ' OK
        End If
        
        
        
        ReDim ObjData(Object).recursosNecesarios(1 To 1) As ObjectoNecesario
        ReDim ObjData(Object).premioReciclaje(1 To 1) As ObjectoNecesario
        
        ' Construccion
        Call addRecursoNecesario(ObjData(Object), val(m_iniFile.getValue(Object, "LingH")), 386) ' OK
        Call addRecursoNecesario(ObjData(Object), val(m_iniFile.getValue(Object, "LingP")), 387) ' OK
        Call addRecursoNecesario(ObjData(Object), val(m_iniFile.getValue(Object, "LingO")), 388) ' OK
        
        Call addRecursoNecesario(ObjData(Object), val(m_iniFile.getValue(Object, "Madera")), 58) ' OK
        Call addRecursoNecesario(ObjData(Object), val(m_iniFile.getValue(Object, "MaderaT")), 642) ' OK
        Call addRecursoNecesario(ObjData(Object), val(m_iniFile.getValue(Object, "MaderaRoble")), 380) ' OK
        
        Call addRecursoNecesario(ObjData(Object), val(m_iniFile.getValue(Object, "Lana")), 208) ' OK
        Call addRecursoNecesario(ObjData(Object), val(m_iniFile.getValue(Object, "Lino")), 257) ' OK
        Call addRecursoNecesario(ObjData(Object), val(m_iniFile.getValue(Object, "Seda")), 377) ' OK
        
        ' Reciclaje
        Call addPremioReciclaje(ObjData(Object), val(m_iniFile.getValue(Object, "ReciclabeHierro")), 386) ' OK
        Call addPremioReciclaje(ObjData(Object), val(m_iniFile.getValue(Object, "ReciclabePlata")), 387) ' OK
        Call addPremioReciclaje(ObjData(Object), val(m_iniFile.getValue(Object, "ReciclabeOro")), 388) ' OK
        
        Call addPremioReciclaje(ObjData(Object), val(m_iniFile.getValue(Object, "ReciclabeMadera")), 58) ' OK
        Call addPremioReciclaje(ObjData(Object), val(m_iniFile.getValue(Object, "ReciclabeTejo")), 642)  ' OK
        Call addPremioReciclaje(ObjData(Object), val(m_iniFile.getValue(Object, "ReciclabeRoble")), 543) ' OK
        
        Call addPremioReciclaje(ObjData(Object), val(m_iniFile.getValue(Object, "ReciclabeLana")), 574) ' OK
        Call addPremioReciclaje(ObjData(Object), val(m_iniFile.getValue(Object, "ReciclabeLino")), 577) ' OK
        Call addPremioReciclaje(ObjData(Object), val(m_iniFile.getValue(Object, "ReciclabeSeda")), 578) ' OK
         
    'Bebidas
    frmCargando.cargar.value = frmCargando.cargar.value + 1

Next Object

End Sub

Private Function addRecursoNecesario(ByRef objeto As ObjData, ByVal cantidad As Integer, ObjIndex As Integer)
    If cantidad = 0 Then
        Exit Function
    End If

    Dim cantidadObjetos As Integer
    
    cantidadObjetos = UBound(objeto.recursosNecesarios)
    
    If cantidadObjetos = 1 And objeto.recursosNecesarios(1).ObjIndex = 0 Then
        objeto.recursosNecesarios(1).cantidad = cantidad
        objeto.recursosNecesarios(1).ObjIndex = ObjIndex
    Else
        cantidadObjetos = cantidadObjetos + 1
        
        ReDim Preserve objeto.recursosNecesarios(1 To cantidadObjetos) As ObjectoNecesario
        
        objeto.recursosNecesarios(cantidadObjetos).cantidad = cantidad
        objeto.recursosNecesarios(cantidadObjetos).ObjIndex = ObjIndex
    End If
    

End Function

Private Function addPremioReciclaje(ByRef objeto As ObjData, ByVal cantidad As Integer, ObjIndex As Integer)
    If cantidad = 0 Then
        Exit Function
    End If

    Dim cantidadObjetos As Integer
    
    cantidadObjetos = UBound(objeto.premioReciclaje)
    
    If cantidadObjetos = 1 And objeto.premioReciclaje(1).ObjIndex = 0 Then
        objeto.premioReciclaje(1).cantidad = cantidad
        objeto.premioReciclaje(1).ObjIndex = ObjIndex
    Else
        cantidadObjetos = cantidadObjetos + 1
        
        ReDim Preserve objeto.premioReciclaje(1 To cantidadObjetos) As ObjectoNecesario
        
        objeto.premioReciclaje(cantidadObjetos).cantidad = cantidad
        objeto.premioReciclaje(cantidadObjetos).ObjIndex = ObjIndex
    End If
    

End Function

' Recibe un conjunto de números separado por comas y devuelve un Long que es el OR de todos los numeros presents en la cadena
Private Function stringToBinaryLong(cadena As String) As Long
       Dim i As Integer
       
       Dim cadenaParseada As Variant
    
        If Trim$(cadena) = "" Then
            stringToBinaryLong = 0
        Else
            cadenaParseada = Split(cadena, ",")
            
            stringToBinaryLong = 0
            
            For i = LBound(cadenaParseada) To UBound(cadenaParseada)
                stringToBinaryLong = (stringToBinaryLong Or Int(cadenaParseada(i)))
            Next
        End If
        
End Function
Public Function LoadUserStats(UserIndex As Integer, rs As Recordset) As Boolean

Dim loopC As Integer
    
    LoadUserStats = True

    With UserList(UserIndex).Stats
        .UserAtributos(eAtributos.Fuerza) = rs!AT1
        .UserAtributos(eAtributos.Agilidad) = rs!AT2
        .UserAtributos(eAtributos.Inteligencia) = rs!AT3
        .UserAtributos(eAtributos.Carisma) = rs!AT4
        .UserAtributos(eAtributos.constitucion) = rs!AT5
    
         For loopC = 1 To NUMATRIBUTOS
           .UserAtributosBackUP(loopC) = .UserAtributos(loopC)
         Next

'Carga de Skill
        .UserSkills(eSkills.ResistenciaMagica) = rs!SK1
        .UserSkills(2) = rs!SK2
        .UserSkills(3) = rs!SK3
        .UserSkills(4) = rs!SK4
        .UserSkills(5) = rs!SK5
        .UserSkills(6) = rs!SK6
        .UserSkills(7) = rs!SK7
        .UserSkills(8) = rs!SK8
        .UserSkills(9) = rs!SK9
        .UserSkills(10) = rs!SK10
        .UserSkills(11) = rs!SK11
        .UserSkills(12) = rs!SK12
        .UserSkills(13) = rs!SK13
        .UserSkills(14) = rs!SK14
        .UserSkills(15) = rs!SK15
        .UserSkills(16) = rs!SK16
        .UserSkills(17) = rs!SK17
        .UserSkills(18) = rs!SK18
        .UserSkills(19) = rs!SK19
        .UserSkills(20) = rs!SK20
        .UserSkills(21) = rs!SK21

'Carga los echizos!

        .UserHechizos(1) = rs!H1
        .UserHechizos(2) = rs!H2
        .UserHechizos(3) = rs!H3
        .UserHechizos(4) = rs!H4
        .UserHechizos(5) = rs!H5
        .UserHechizos(6) = rs!H6
        .UserHechizos(7) = rs!H7
        .UserHechizos(8) = rs!H8
        .UserHechizos(9) = rs!H9
        .UserHechizos(10) = rs!H10
        .UserHechizos(11) = rs!H11
        .UserHechizos(12) = rs!H12
        .UserHechizos(13) = rs!H13
        .UserHechizos(14) = rs!H14
        .UserHechizos(15) = rs!H15
        .UserHechizos(16) = rs!H16
        .UserHechizos(17) = rs!H17
        .UserHechizos(18) = rs!H18
        .UserHechizos(19) = rs!H19
        .UserHechizos(20) = rs!H20
        .UserHechizos(21) = rs!H21
        .UserHechizos(22) = rs!H22
        .UserHechizos(23) = rs!H23
        .UserHechizos(24) = rs!H24
        .UserHechizos(25) = rs!H25
        .UserHechizos(26) = rs!H26
        .UserHechizos(27) = rs!H27
        .UserHechizos(28) = rs!H28
        .UserHechizos(29) = rs!H29
        .UserHechizos(30) = rs!H30
        .UserHechizos(31) = rs!H31
        .UserHechizos(32) = rs!H32
        .UserHechizos(33) = rs!H33
        .UserHechizos(34) = rs!H34
        .UserHechizos(35) = rs!H35

        .GLD = rs!gldb
        .GldBackup = .GLD
        .Banco = rs!bancob

        .MaxHP = rs!MaxHPB
        .minHP = rs!MinHPB

        .MinSta = rs!MinSTAB
        .MaxSta = rs!MaxStaB

        .MaxMAN = rs!MaxMANb
        .MinMAN = rs!MinMANB

        .MaxHIT = rs!MaxHITB
        .MinHIT = rs!MinHITB

        .MaxAGU = rs!MaxAGUB
        .minAgu = rs!minAGUB

        .MaxHam = rs!MaxHAMB
        .minham = rs!MinHAMB

        .SkillPts = rs!SkillPtsLibresB
        
        .Exp = CCur(rs!EXPB)
        .ELV = rs!elvb
        .ELU = Constantes_Generales.obtenerExperienciaNecesaria(.ELV)
            
        .UsuariosMatados = rs!UserMuertesB
        .NPCsMuertos = rs!NpcsMuertesB
        .MaxItems = rs!MaxItems
        .OroGanado = rs!OG
        .OroPerdido = rs!OP
        .RetosGanadoS = rs!RG
        .RetosPerdidosB = rs!RP

        If .MaxItems < 20 Then .MaxItems = 20

    End With

End Function

Public Function LoadUserReputacion(UserIndex As Integer, rs As Recordset)

   LoadUserReputacion = True
    
    With UserList(UserIndex).Reputacion
        .AsesinoRep = rs!AsesinoB
        .BandidoRep = rs!BandidoB
        .BurguesRep = rs!BurguesiaB
        .LadronesRep = rs!LadronesB
        .NobleRep = rs!NoblesB
        .PlebeRep = rs!PlebeB
        .promedio = rs!promedioB
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadUserInit
' DateTime  : 17/03/2007 22:52
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
'CSEH: TDS_LINEA
Public Function LoadUserInit(UserIndex As Integer, rs As Recordset) As Boolean

    With UserList(UserIndex).faccion
        .ArmadaReal = rs!EjercitoRealB
        .FuerzasCaos = rs!EjercitoCaosB
        .CiudadanosMatados = rs!CiudMatadosB
        .CriminalesMatados = rs!CrimMatadosB
        .RecibioArmaduraCaos = rs!rArCaosB
        .RecibioArmaduraReal = rs!rArRealB
        .RecibioExpInicialCaos = rs!rExCaosB
        .RecibioExpInicialReal = rs!rExRealB
        .RecompensasCaos = rs!recCaosB
        .RecompensasReal = rs!recRealB
        .alineacion = rs!alineacion
    End With
        
    With UserList(UserIndex).flags
        .Muerto = rs!MuertoB
        .Hambre = rs!HambreB
        .Sed = rs!SedB
        .Desnudo = rs!DesnudoB

        .Envenenado = rs!EnvenenadoB
        .Paralizado = rs!ParalizadoB

        If .Paralizado = 1 Then
            UserList(UserIndex).Counters.Paralisis = modPersonaje.getIntervaloParalizado(UserList(UserIndex))
        End If

        .Navegando = rs!NavegandoB
        .PertAlCons = rs!PERTENECEB
        .PertAlConsCaos = rs!PERTENECECAOSB
        .Penasas = rs!penasasb
        
        'Privilegios
        If Not IsNull(rs!Privilegios) Then
           .Privilegios = rs!Privilegios
        Else
           .Privilegios = 0
        End If
    End With

    With UserList(UserIndex)
        .Counters.Pena = rs!penab

        .Email = rs!EmailB

        'El mysal me devuelve String, lo transformo en enums
        .Genero = Declaraciones.generoToByte(rs!generoB)
        .clase = modClases.claseToByte(rs!claseb)
        
        .ClaseNumero = modClases.claseToConfigID(rs!claseb)
        
        .Raza = Declaraciones.razaToByte(rs!razaB)
        
        .Hogar = rs!HogarB
        .Char.heading = rs!HeadingB


        .OrigChar.Head = rs!Headb
        .OrigChar.Body = rs!bodyb
        .OrigChar.WeaponAnim = rs!armab
        .OrigChar.ShieldAnim = rs!escudob
        .OrigChar.CascoAnim = rs!Cascob
        .OrigChar.heading = rs!HeadingB

        .desc = rs!DescB
'       WARNING
        .pos.map = rs!mapb
        .pos.x = rs!xb
        .pos.y = rs!yb
    End With


'[MARCHE]--------------------------------------------------------------------
'***********************************************************************************

    With UserList(UserIndex).BancoInvent
        .Object(1).ObjIndex = val(ReadField(1, rs!Bobj1, 45))
        .Object(1).Amount = val(ReadField(2, rs!Bobj1, 45))
        .Object(2).ObjIndex = val(ReadField(1, rs!Bobj2, 45))
        .Object(2).Amount = val(ReadField(2, rs!Bobj2, 45))
        .Object(3).ObjIndex = val(ReadField(1, rs!Bobj3, 45))
        .Object(3).Amount = val(ReadField(2, rs!Bobj3, 45))
        .Object(4).ObjIndex = val(ReadField(1, rs!Bobj4, 45))
        .Object(4).Amount = val(ReadField(2, rs!Bobj4, 45))
        .Object(5).ObjIndex = val(ReadField(1, rs!Bobj5, 45))
        .Object(5).Amount = val(ReadField(2, rs!Bobj5, 45))
        .Object(6).ObjIndex = val(ReadField(1, rs!Bobj6, 45))
        .Object(6).Amount = val(ReadField(2, rs!Bobj6, 45))
        .Object(7).ObjIndex = val(ReadField(1, rs!Bobj7, 45))
        .Object(7).Amount = val(ReadField(2, rs!Bobj7, 45))
        .Object(8).ObjIndex = val(ReadField(1, rs!Bobj8, 45))
        .Object(8).Amount = val(ReadField(2, rs!Bobj8, 45))
        .Object(9).ObjIndex = val(ReadField(1, rs!Bobj9, 45))
        .Object(9).Amount = val(ReadField(2, rs!Bobj9, 45))
        .Object(10).ObjIndex = val(ReadField(1, rs!Bobj10, 45))
        .Object(10).Amount = val(ReadField(2, rs!Bobj10, 45))
        .Object(11).ObjIndex = val(ReadField(1, rs!Bobj11, 45))
        .Object(11).Amount = val(ReadField(2, rs!Bobj11, 45))
        .Object(12).ObjIndex = val(ReadField(1, rs!Bobj12, 45))
        .Object(12).Amount = val(ReadField(2, rs!Bobj12, 45))
        .Object(13).ObjIndex = val(ReadField(1, rs!Bobj13, 45))
        .Object(13).Amount = val(ReadField(2, rs!Bobj13, 45))
        .Object(14).ObjIndex = val(ReadField(1, rs!Bobj14, 45))
        .Object(14).Amount = val(ReadField(2, rs!Bobj14, 45))
        .Object(15).ObjIndex = val(ReadField(1, rs!Bobj15, 45))
        .Object(15).Amount = val(ReadField(2, rs!Bobj15, 45))
        .Object(16).ObjIndex = val(ReadField(1, rs!Bobj16, 45))
        .Object(16).Amount = val(ReadField(2, rs!Bobj16, 45))
        .Object(17).ObjIndex = val(ReadField(1, rs!Bobj17, 45))
        .Object(17).Amount = val(ReadField(2, rs!Bobj17, 45))
        .Object(18).ObjIndex = val(ReadField(1, rs!Bobj18, 45))
        .Object(18).Amount = val(ReadField(2, rs!Bobj18, 45))
        .Object(19).ObjIndex = val(ReadField(1, rs!Bobj19, 45))
        .Object(19).Amount = val(ReadField(2, rs!Bobj19, 45))
        .Object(20).ObjIndex = val(ReadField(1, rs!Bobj20, 45))
        .Object(20).Amount = val(ReadField(2, rs!Bobj20, 45))
        .Object(21).ObjIndex = val(ReadField(1, rs!Bobj21, 45))
        .Object(21).Amount = val(ReadField(2, rs!Bobj21, 45))
        .Object(22).ObjIndex = val(ReadField(1, rs!Bobj22, 45))
        .Object(22).Amount = val(ReadField(2, rs!Bobj22, 45))
        .Object(23).ObjIndex = val(ReadField(1, rs!Bobj23, 45))
        .Object(23).Amount = val(ReadField(2, rs!Bobj23, 45))
        .Object(24).ObjIndex = val(ReadField(1, rs!Bobj24, 45))
        .Object(24).Amount = val(ReadField(2, rs!Bobj24, 45))
        .Object(25).ObjIndex = val(ReadField(1, rs!Bobj25, 45))
        .Object(25).Amount = val(ReadField(2, rs!Bobj25, 45))
        .Object(26).ObjIndex = val(ReadField(1, rs!Bobj26, 45))
        .Object(26).Amount = val(ReadField(2, rs!Bobj26, 45))
        .Object(27).ObjIndex = val(ReadField(1, rs!Bobj27, 45))
        .Object(27).Amount = val(ReadField(2, rs!Bobj27, 45))
        .Object(28).ObjIndex = val(ReadField(1, rs!Bobj28, 45))
        .Object(28).Amount = val(ReadField(2, rs!Bobj28, 45))
        .Object(29).ObjIndex = val(ReadField(1, rs!Bobj29, 45))
        .Object(29).Amount = val(ReadField(2, rs!Bobj29, 45))
        .Object(30).ObjIndex = val(ReadField(1, rs!Bobj30, 45))
        .Object(30).Amount = val(ReadField(2, rs!Bobj30, 45))
        .Object(31).ObjIndex = val(ReadField(1, rs!Bobj31, 45))
        .Object(31).Amount = val(ReadField(2, rs!Bobj31, 45))
        .Object(32).ObjIndex = val(ReadField(1, rs!Bobj32, 45))
        .Object(32).Amount = val(ReadField(2, rs!Bobj32, 45))
        .Object(33).ObjIndex = val(ReadField(1, rs!Bobj33, 45))
        .Object(33).Amount = val(ReadField(2, rs!Bobj33, 45))
        .Object(34).ObjIndex = val(ReadField(1, rs!Bobj34, 45))
        .Object(34).Amount = val(ReadField(2, rs!Bobj34, 45))
        .Object(35).ObjIndex = val(ReadField(1, rs!Bobj35, 45))
        .Object(35).Amount = val(ReadField(2, rs!Bobj35, 45))
        .Object(36).ObjIndex = val(ReadField(1, rs!Bobj36, 45))
        .Object(36).Amount = val(ReadField(2, rs!Bobj36, 45))
        .Object(37).ObjIndex = val(ReadField(1, rs!Bobj37, 45))
        .Object(37).Amount = val(ReadField(2, rs!Bobj37, 45))
        .Object(38).ObjIndex = val(ReadField(1, rs!Bobj38, 45))
        .Object(38).Amount = val(ReadField(2, rs!Bobj38, 45))
        .Object(39).ObjIndex = val(ReadField(1, rs!Bobj39, 45))
        .Object(39).Amount = val(ReadField(2, rs!Bobj39, 45))
        .Object(40).ObjIndex = val(ReadField(1, rs!Bobj40, 45))
        .Object(40).Amount = val(ReadField(2, rs!Bobj40, 45))
    End With
'------------------------------------------------------------------------------------
'[/MARCHE]*****************************************************************************

'Lista de objetos
    
    With UserList(UserIndex).Invent

        .Object(1).ObjIndex = val(ReadField(1, rs!iOBJ1, 45))
        .Object(1).Amount = val(ReadField(2, rs!iOBJ1, 45))
        .Object(1).Equipped = val(ReadField(3, rs!iOBJ1, 45))
        .Object(2).ObjIndex = val(ReadField(1, rs!iOBJ2, 45))
        .Object(2).Amount = val(ReadField(2, rs!iOBJ2, 45))
        .Object(2).Equipped = val(ReadField(3, rs!iOBJ2, 45))
        .Object(3).ObjIndex = val(ReadField(1, rs!iOBJ3, 45))
        .Object(3).Amount = val(ReadField(2, rs!iOBJ3, 45))
        .Object(3).Equipped = val(ReadField(3, rs!iOBJ3, 45))
        .Object(4).ObjIndex = val(ReadField(1, rs!iOBJ4, 45))
        .Object(4).Amount = val(ReadField(2, rs!iOBJ4, 45))
        .Object(4).Equipped = val(ReadField(3, rs!iOBJ4, 45))
        .Object(5).ObjIndex = val(ReadField(1, rs!iOBJ5, 45))
        .Object(5).Amount = val(ReadField(2, rs!iOBJ5, 45))
        .Object(5).Equipped = val(ReadField(3, rs!iOBJ5, 45))
        .Object(6).ObjIndex = val(ReadField(1, rs!iOBJ6, 45))
        .Object(6).Amount = val(ReadField(2, rs!iOBJ6, 45))
        .Object(6).Equipped = val(ReadField(3, rs!iOBJ6, 45))
        .Object(7).ObjIndex = val(ReadField(1, rs!iOBJ7, 45))
        .Object(7).Amount = val(ReadField(2, rs!iOBJ7, 45))
        .Object(7).Equipped = val(ReadField(3, rs!iOBJ7, 45))
        .Object(8).ObjIndex = val(ReadField(1, rs!iOBJ8, 45))
        .Object(8).Amount = val(ReadField(2, rs!iOBJ8, 45))
        .Object(8).Equipped = val(ReadField(3, rs!iOBJ8, 45))
        .Object(9).ObjIndex = val(ReadField(1, rs!iOBJ9, 45))
        .Object(9).Amount = val(ReadField(2, rs!iOBJ9, 45))
        .Object(9).Equipped = val(ReadField(3, rs!iOBJ9, 45))
        .Object(10).ObjIndex = val(ReadField(1, rs!iOBJ10, 45))
        .Object(10).Amount = val(ReadField(2, rs!iOBJ10, 45))
        .Object(10).Equipped = val(ReadField(3, rs!iOBJ10, 45))
        .Object(11).ObjIndex = val(ReadField(1, rs!iOBJ11, 45))
        .Object(11).Amount = val(ReadField(2, rs!iOBJ11, 45))
        .Object(11).Equipped = val(ReadField(3, rs!iOBJ11, 45))
        .Object(12).ObjIndex = val(ReadField(1, rs!iOBJ12, 45))
        .Object(12).Amount = val(ReadField(2, rs!iOBJ12, 45))
        .Object(12).Equipped = val(ReadField(3, rs!iOBJ12, 45))
        .Object(13).ObjIndex = val(ReadField(1, rs!iOBJ13, 45))
        .Object(13).Amount = val(ReadField(2, rs!iOBJ13, 45))
        .Object(13).Equipped = val(ReadField(3, rs!iOBJ13, 45))
        .Object(14).ObjIndex = val(ReadField(1, rs!iOBJ14, 45))
        .Object(14).Amount = val(ReadField(2, rs!iOBJ14, 45))
        .Object(14).Equipped = val(ReadField(3, rs!iOBJ14, 45))

        .Object(15).ObjIndex = val(ReadField(1, rs!iOBJ15, 45))
        .Object(15).Amount = val(ReadField(2, rs!iOBJ15, 45))
        .Object(15).Equipped = val(ReadField(3, rs!iOBJ15, 45))
        .Object(16).ObjIndex = val(ReadField(1, rs!iOBJ16, 45))
        .Object(16).Amount = val(ReadField(2, rs!iOBJ16, 45))
        .Object(16).Equipped = val(ReadField(3, rs!iOBJ16, 45))
        .Object(17).ObjIndex = val(ReadField(1, rs!iOBJ17, 45))
        .Object(17).Amount = val(ReadField(2, rs!iOBJ17, 45))
        .Object(17).Equipped = val(ReadField(3, rs!iOBJ17, 45))
        .Object(18).ObjIndex = val(ReadField(1, rs!iOBJ18, 45))
        .Object(18).Amount = val(ReadField(2, rs!iOBJ18, 45))
        .Object(18).Equipped = val(ReadField(3, rs!iOBJ18, 45))
        .Object(19).ObjIndex = val(ReadField(1, rs!iOBJ19, 45))
        .Object(19).Amount = val(ReadField(2, rs!iOBJ19, 45))
        .Object(19).Equipped = val(ReadField(3, rs!iOBJ19, 45))
        .Object(20).ObjIndex = val(ReadField(1, rs!iOBJ20, 45))
        .Object(20).Amount = val(ReadField(2, rs!iOBJ20, 45))
        .Object(20).Equipped = val(ReadField(3, rs!iOBJ20, 45))

        .Object(21).ObjIndex = val(ReadField(1, rs!iOBJ21, 45))
        .Object(21).Amount = val(ReadField(2, rs!iOBJ21, 45))
        .Object(21).Equipped = val(ReadField(3, rs!iOBJ21, 45))
        .Object(22).ObjIndex = val(ReadField(1, rs!iOBJ22, 45))
        .Object(22).Amount = val(ReadField(2, rs!iOBJ22, 45))
        .Object(22).Equipped = val(ReadField(3, rs!iOBJ22, 45))
        .Object(23).ObjIndex = val(ReadField(1, rs!iOBJ23, 45))
        .Object(23).Amount = val(ReadField(2, rs!iOBJ23, 45))
        .Object(23).Equipped = val(ReadField(3, rs!iOBJ23, 45))
        .Object(24).ObjIndex = val(ReadField(1, rs!iOBJ24, 45))
        .Object(24).Amount = val(ReadField(2, rs!iOBJ24, 45))
        .Object(24).Equipped = val(ReadField(3, rs!iOBJ24, 45))
        .Object(25).ObjIndex = val(ReadField(1, rs!iOBJ25, 45))
        .Object(25).Amount = val(ReadField(2, rs!iOBJ25, 45))
        .Object(25).Equipped = val(ReadField(3, rs!iOBJ25, 45))

        .Object(26).ObjIndex = val(ReadField(1, rs!iOBJ26, 45))
        .Object(26).Amount = val(ReadField(2, rs!iOBJ26, 45))
        .Object(26).Equipped = val(ReadField(3, rs!iOBJ26, 45))
        .Object(27).ObjIndex = val(ReadField(1, rs!iOBJ27, 45))
        .Object(27).Amount = val(ReadField(2, rs!iOBJ27, 45))
        .Object(27).Equipped = val(ReadField(3, rs!iOBJ27, 45))
        .Object(28).ObjIndex = val(ReadField(1, rs!iOBJ28, 45))
        .Object(28).Amount = val(ReadField(2, rs!iOBJ28, 45))
        .Object(28).Equipped = val(ReadField(3, rs!iOBJ28, 45))
        .Object(29).ObjIndex = val(ReadField(1, rs!iOBJ29, 45))
        .Object(29).Amount = val(ReadField(2, rs!iOBJ29, 45))
        .Object(29).Equipped = val(ReadField(3, rs!iOBJ29, 45))
        .Object(30).ObjIndex = val(ReadField(1, rs!iOBJ30, 45))
        .Object(30).Amount = val(ReadField(2, rs!iOBJ30, 45))
        .Object(30).Equipped = val(ReadField(3, rs!iOBJ30, 45))

                           End With
    
                       'Obtiene el indice-objeto del arma

    With UserList(UserIndex)
        .Invent.WeaponEqpSlot = rs!WeaponEqpSlotB
        If .Invent.WeaponEqpSlot > 0 Then .Invent.WeaponEqpObjIndex = .Invent.Object(.Invent.WeaponEqpSlot).ObjIndex

                       'Obtiene el indice-objeto del armadura
        .Invent.ArmourEqpSlot = rs!ArmourEqpSlotB
        If .Invent.ArmourEqpSlot > 0 Then
            .Invent.ArmourEqpObjIndex = .Invent.Object(.Invent.ArmourEqpSlot).ObjIndex
            .flags.Desnudo = 0
                               Else
            .flags.Desnudo = 1
                               End If

                       'Obtiene el indice-objeto del escudo
        .Invent.EscudoEqpSlot = rs!EscudoEqpSlotB
        If .Invent.EscudoEqpSlot > 0 Then .Invent.EscudoEqpObjIndex = .Invent.Object(.Invent.EscudoEqpSlot).ObjIndex


                       'Obtiene el indice-objeto del casco
        .Invent.CascoEqpSlot = rs!CascoEqpSlotB
        If .Invent.CascoEqpSlot > 0 Then .Invent.CascoEqpObjIndex = .Invent.Object(.Invent.CascoEqpSlot).ObjIndex

                       'Obtiene el indice-objeto barco
        .Invent.BarcoEqpSlot = rs!BarcoEqpSlotB
        .Invent.BarcoSlot = rs!BarcoSlotB
        If .Invent.BarcoSlot > 0 Then .Invent.BarcoObjIndex = .Invent.Object(.Invent.BarcoSlot).ObjIndex

        

                       'Obtiene el indice-objeto municion
        .Invent.MunicionEqpSlot = rs!MunicionSlotB
        If .Invent.MunicionEqpSlot > 0 Then .Invent.MunicionEqpObjIndex = .Invent.Object(.Invent.MunicionEqpSlot).ObjIndex

                       '[Alejo]
                       'Obtiene el indice-objeto herramienta
        .Invent.HerramientaEqpSlot = rs!HerramientaSlotB
        If .Invent.HerramientaEqpSlot > 0 Then .Invent.HerramientaEqpObjIndex = .Invent.Object(.Invent.HerramientaEqpSlot).ObjIndex

        .Invent.AnilloEqpSlot = rs!AnilloSlotB
        If .Invent.AnilloEqpSlot > 0 Then .Invent.AnilloEqpObjIndex = .Invent.Object(.Invent.AnilloEqpSlot).ObjIndex


                              Dim loopObjeto As Integer
                              Dim ObjIndex As Integer
        
        For loopObjeto = LBound(.Invent.Object) To UBound(.Invent.Object)
            ObjIndex = .Invent.Object(loopObjeto).ObjIndex
            If ObjIndex > 0 And .Invent.Object(loopObjeto).Equipped = 1 Then
                If ObjData(ObjIndex).ObjType = OBJTYPE_COLLAR Then
                    .Invent.CollarObjIndex = ObjIndex
                ElseIf ObjData(ObjIndex).ObjType = OBJTYPE_BRASALETE Then
                    .Invent.BrasaleteEqpObjIndex = ObjIndex
                                      End If
                                  End If
                              Next

                              '[Wizard 07/09/05] Baje esto, para poder hacerlo completo(Por el Inventario del barco)
        If .flags.Muerto = 0 Then         'Si esta vivo...
            
            If .flags.Navegando = 0 Then         'Si no navega...
                .Char = .OrigChar
             ElseIf .Invent.BarcoObjIndex > 0 Then 'Navega....
                .Char.WeaponAnim = NingunArma
                .Char.ShieldAnim = NingunEscudo
                .Char.CascoAnim = NingunCasco
                .Char.Head = 0
                .Char.Body = ObjData(.Invent.BarcoObjIndex).Ropaje
                                    Else 'BUG DE ALGUNA MANERA DEJAN PJ SIN BARCA NAVEGANDO
                .Char = .OrigChar
                .flags.Navegando = 0
                                    End If
                               Else 'Esta muerto!
            .Char.WeaponAnim = NingunArma
            .Char.ShieldAnim = NingunEscudo
            .Char.CascoAnim = NingunCasco
                If .flags.Navegando = 1 Then         'Ta navegando
                    .Char.Body = iFragataFantasmal
                    .Char.Head = 0
                ElseIf .faccion.FuerzasCaos <> 0 Then         'Es caos
                    .Char.Body = iCuerpoMuertoCrimi
                    .Char.Head = iCabezaMuertoCrimi
                                       Else 'No navega y no es caos: Casper blanquito^^
                    .Char.Body = iCuerpoMuerto
                    .Char.Head = iCabezaMuerto
                                       End If
                               End If
                       '/Wizard
                       '

                       ' Carga de mascotas HORRIBLE!!!
        .NroMacotas = 0
        
        If Int(rs!mas1) > 0 Then
            .MascotasType(1) = rs!mas1
            .NroMacotas = .NroMacotas + 1
                               Else
            .MascotasType(1) = 0
                               End If
        
        If Int(rs!mas2) > 0 Then
            .MascotasType(2) = rs!mas2
            .NroMacotas = .NroMacotas + 1
                               Else
            .MascotasType(2) = 0
                               End If

        If Int(rs!mas3) > 0 Then
            .MascotasType(3) = rs!mas3
            .NroMacotas = .NroMacotas + 1
                               Else
            .MascotasType(3) = 0
                               End If
        
                       ' Carga de mascotas guardadas
        .NroMascotasGuardadas = 0
        
        If Int(rs!mas1b) > 0 Then
            .MascotasGuardadas(1) = rs!mas1b
            .NroMascotasGuardadas = .NroMascotasGuardadas + 1
                               Else
            .MascotasGuardadas(1) = 0
                               End If
        
        If Int(rs!mas2b) > 0 Then
            .MascotasGuardadas(2) = rs!mas2b
            .NroMascotasGuardadas = .NroMascotasGuardadas + 1
                               Else
            .MascotasGuardadas(2) = 0
                               End If

        If Int(rs!mas3b) > 0 Then
            .MascotasGuardadas(3) = rs!mas3b
            .NroMascotasGuardadas = .NroMascotasGuardadas + 1
                               Else
            .MascotasGuardadas(3) = 0
                               End If

                       '****************************************************
                       '**    Carga de información sobre clan **************
        .GuildInfo.id = rs!IDClan

        If .GuildInfo.id > 0 Then
            Set .ClanRef = clanes.getClan(.GuildInfo.id)
              If Not .ClanRef Is Nothing Then
               .GuildInfo.GuildName = .ClanRef.getNombre()
                                   End If
                               Else
            Set .ClanRef = Nothing
            .GuildInfo.GuildName = ""
                               End If

        .GuildInfo.ClanFundadoID = rs!ClanFundadoID

        If .GuildInfo.ClanFundadoID > 0 Then
            .GuildInfo.FundoClan = 1
                               Else
            .GuildInfo.FundoClan = 0
                               End If

        .GuildInfo.EsGuildLeader = rs!EsGuildLeaderB
        .GuildInfo.echadas = rs!Echadasb
        .GuildInfo.Solicitudes = rs!SolicitudesB
        .GuildInfo.SolicitudesRechazadas = rs!SolicitudesRechazadasB
        .GuildInfo.VecesFueGuildLeader = rs!VecesFueGuildLeaderB
        .GuildInfo.ClanesParticipo = rs!ClanesParticipoB
        .GuildInfo.GuildPoints = rs!GuildPointsB
                           End With
    
    LoadUserInit = True

End Function





'---------------------------------------------------------------------------------------
' Procedure : GetVar
' DateTime  : 18/02/2007 19:15
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

szReturn = ""
  
sSpaces = Space(5000) ' This tells the computer how long the longest string can be
    
GetPrivateProfileString Main, Var, szReturn, sSpaces, Len(sSpaces), file
  
GetVar = RTrim(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function





Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Escribe VAR en un archivo
'*****************************************************************

writeprivateprofilestring Main, Var, value, file
    
End Sub

Sub SaveUser(ByVal UserIndex As Integer, Optional ByRef online As Boolean = False)

Dim variable As String
Dim i As Integer
Dim L As Long
Dim onlineByte As Byte

With UserList(UserIndex)

    'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
    If .clase = eClases.indefinido Or .Stats.ELV = 0 Then
        Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & .Name)
        Exit Sub
    End If

    If .flags.Muerto = 1 Then
        .Char.Head = iCabezaMuerto
    Else
        variable = ",BodyB=" & .Char.Body
    End If

    '///////////////MASCOTAS//////////////////////////////
    For i = 1 To MAXMASCOTAS
        If .MascotasIndex(i) > 0 Then
            If NpcList(.MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then
                variable = variable & ",mas" & i & "=" & .MascotasType(i)
            Else
                variable = variable & ",mas" & i & "=0"
            End If
        Else
            variable = variable & ",mas" & i & "=0"
        End If
    Next i


    ' Mascotas guardadas
    For i = 1 To MAXMASCOTAS
        If .MascotasGuardadas(i) > 0 Then
            variable = variable & ",mas" & i & "b=" & .MascotasGuardadas(i)
        Else
            variable = variable & ",mas" & i & "b=0"
        End If
    Next i


    '////////////MASCOTAS////////////////////////////////////
    variable = variable & ",BanrazB='" & mysql_real_escape_string(.flags.Banrazon) & "'"
    variable = variable & ",VecesCheat='" & .Stats.Veceshechado & "'"

    'ATRIBUTOS
    For i = 1 To 5
        variable = variable & "," & "AT" & i & "=" & val(.Stats.UserAtributosBackUP(i))
    Next

    'SKILLS
    For i = 1 To 21
        variable = variable & "," & "SK" & i & "=" & val(.Stats.UserSkills(i))
    Next
    
    'Objetos de boveda
    For i = 1 To 40
        variable = variable & "," & "Bobj" & i & "='" & .BancoInvent.Object(i).ObjIndex & "-" & .BancoInvent.Object(i).Amount & "'"
    Next

    'Objetos del inventario
    For i = 1 To 30
        variable = variable & "," & "iOBJ" & i & "='" & .Invent.Object(i).ObjIndex & "-" & .Invent.Object(i).Amount & "-" & .Invent.Object(i).Equipped & "'"
    Next

    'Hechizoss
    For i = 1 To 35
        variable = variable & "," & "H" & i & "=" & .Stats.UserHechizos(i)
    Next


    L = (-.Reputacion.AsesinoRep) + _
    (-.Reputacion.BandidoRep) + _
    .Reputacion.BurguesRep + _
    (-.Reputacion.LadronesRep) + _
    .Reputacion.NobleRep + _
    .Reputacion.PlebeRep
    
    L = L / 6

    variable = variable & ",PROMEDIOB=" & L
    variable = variable & ",LastIPB=" & .ip

    variable = variable & ",Maxitems=" & .Stats.MaxItems
    variable = variable & ",Unban='" & .flags.Unban & "'"
    variable = variable & ",AnilloSlotB=" & .Invent.AnilloEqpSlot
    variable = variable & ",EmailB='" & .Email & "'"

    If Len(.pin) > 0 Then
        variable = variable & ",PIN='" & .pin & "'"
    End If
    
    
    
    If online Then
        onlineByte = 1
    Else
        onlineByte = 0
    End If
   
    variable = variable & ",Online=" & onlineByte & ",IDCuenta=" & .IDCuenta & ", Alineacion=" & .faccion.alineacion

' Guardo el usuario
    sql = ("UPDATE " & DB_NAME_PRINCIPAL & ".usuarios SET Fecha=Now() ," & "MuertoB=" & .flags.Muerto & ",HambreB=" & .flags.Hambre & ",SedB=" & .flags.Sed & ",DesnudoB=" & .flags.Desnudo & ",banB=" & .flags.Ban & ",NavegandoB=" & .flags.Navegando & "," & _
"EnvenenadoB=" & .flags.Envenenado & ",ParalizadoB=" & .flags.Paralizado & ",PERTENECEB=" & .flags.PertAlCons & ",PERTENECECAOSB=" & .flags.PertAlConsCaos & ",banB=" & .flags.Ban & ",penab=" & .Counters.Pena & "," & _
"penasasb='" & mysql_real_escape_string(.flags.Penasas) & "',EjercitoRealB=" & .faccion.ArmadaReal & ",EjercitoCaosB=" & .faccion.FuerzasCaos & ",CiudMatadosB=" & .faccion.CiudadanosMatados & ",CrimMatadosB=" & .faccion.CriminalesMatados & ",rArCaosB=" & .faccion.RecibioArmaduraCaos & _
", rArRealB=" & .faccion.RecibioArmaduraReal & ",rExRealB=" & .faccion.RecibioExpInicialReal & ",recCaosB=" & .faccion.RecompensasCaos & ",recRealB=" & .faccion.RecompensasReal & ",EsGuildLeaderB=" & .GuildInfo.EsGuildLeader & _
", EchadasB=" & .GuildInfo.echadas & ",SolicitudesB=" & .GuildInfo.Solicitudes & ",SolicitudesRechazadasB=" & .GuildInfo.SolicitudesRechazadas & ",VecesFueGuildLeaderB=" & .GuildInfo.VecesFueGuildLeader & _
", ClanFundadoID=" & .GuildInfo.ClanFundadoID & ",IDClan=" & .GuildInfo.id & ",ClanesParticipoB=" & .GuildInfo.ClanesParticipo & ",guildPtsB=" & .GuildInfo.GuildPoints & _
", generoB='" & byteToGenero(.Genero) & "',razaB='" & byteToRaza(.Raza) & "',HogarB='" & .Hogar & "',claseb='" & byteToClase(.clase) & "'" & _
", PasswordB='" & .Password & "',DescB='" & .desc & "',HeadingB=" & .Char.heading & ",OG=" & .Stats.OroGanado & ",OP=" & .Stats.OroPerdido & _
", RG=" & .Stats.RetosGanadoS & ",RP=" & .Stats.RetosPerdidosB & ",Headb=" & .OrigChar.Head & ",armab=" & .Char.WeaponAnim & ",escudob=" & .Char.ShieldAnim & _
", Cascob=" & .Char.CascoAnim & ",mapb=" & .pos.map & ",yb=" & .pos.y & ",xb=" & .pos.x & _
", gldb=" & .Stats.GLD & ",bancob=" & .Stats.Banco & ",MaxHPB=" & .Stats.MaxHP & ",MinHPB=" & .Stats.minHP & _
",MaxStaB=" & .Stats.MaxSta & ",MinSTAB=" & .Stats.MinSta & ",MaxMANb=" & .Stats.MaxMAN & ",MinMANB=" & .Stats.MinMAN & _
", MaxHITB=" & .Stats.MaxHIT & ",MinHITB=" & .Stats.MinHIT & ",MaxAGUB=" & .Stats.MaxAGU & ",minAGUB=" & .Stats.minAgu & ",MaxHAMB=" & .Stats.MaxHam & _
", MinHAMB=" & .Stats.minham & ",SkillPtsLibresB=" & .Stats.SkillPts & ",EXPB='" & FormatNumber(.Stats.Exp, 0, vbTrue, vbFalse, vbFalse) & "' ,elvb=" & .Stats.ELV & _
", UserMuertesB=" & .Stats.UsuariosMatados & ",NpcsMuertesB=" & .Stats.NPCsMuertos & ",CantidadItemsB=" & .BancoInvent.NroItems & ",WeaponEqpSlotB=" & .Invent.WeaponEqpSlot & _
", ArmourEqpSlotB=" & .Invent.ArmourEqpSlot & ",CascoEqpSlotB=" & .Invent.CascoEqpSlot & ",EscudoEqpSlotB=" & .Invent.EscudoEqpSlot & ",BarcoSlotB=" & .Invent.BarcoSlot & ",BarcoEqpSlotB=" & .Invent.BarcoEqpSlot & ",MunicionSlotB=" & .Invent.MunicionEqpSlot & _
", HerramientaSlotB=" & .Invent.HerramientaEqpSlot & ",AsesinoB=" & .Reputacion.AsesinoRep & ",BandidoB=" & .Reputacion.BandidoRep & ",BurguesiaB=" & .Reputacion.BurguesRep & ",LadronesB=" & .Reputacion.LadronesRep & _
", NoblesB=" & .Reputacion.NobleRep & ",PlebeB=" & .Reputacion.PlebeRep & ",rExCaosB=" & .faccion.RecibioExpInicialCaos & ",MacAddress='" & .MacAddress & "'" & variable & _
" WHERE ID=" & .id)


'Ejecuto la sentencia que arme. Se ejecuto correctamente?
If Not modMySql.ejecutarSQL(sql) Then
    Call LogError("**** NO SE EJECUTO SQL ***** Error en SaveUser de " & UserList(UserIndex).Name & Err.Description & " " & sql)
End If


If Not online Then
    ' Registro de ingreso y egreso del juego
    sql = "INSERT DELAYED " & DB_NAME_PRINCIPAL & ".juego_personajes_logins(IDPERSONAJE, IP, MACADDRESS, NOMBREPC, FECHAINGRESO, FECHAEGRESO) VALUES (" & .id & ", " & .ip & ", '" & .MacAddress & "','" & .NombrePC & "','" & Format$(.FechaIngreso, "yyyy-mm-dd hh:nn:ss") & "','" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    
    If Not modMySql.ejecutarSQL(sql) Then
       Call LogError("**** NO SE EJECUTO SQL ***** Login")
    End If
End If

End With
 
End Sub

Public Function puedeAtacarFaccion(ByRef atacante As User, ByRef victima As User) As Boolean

    If atacante.faccion.alineacion = eAlineaciones.Neutro Then
        puedeAtacarFaccion = True
        Exit Function
    End If
    
    If atacante.faccion.alineacion = victima.faccion.alineacion Then
        puedeAtacarFaccion = True
        Exit Function
    End If
    
    puedeAtacarFaccion = False
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : CargaApuestas
' DateTime  : 18/02/2007 19:14
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub CargaApuestas()
    apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
    apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
    apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))
End Sub


