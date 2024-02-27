Attribute VB_Name = "Module1"
Option Explicit

Private Sub simularPremiosSubaNivel(ByRef personaje As User, ByRef aumentoHit As Integer, ByRef aumentoST As Integer, ByRef aumentoMana As Integer)


Dim loopNivel As Integer
Dim nivelActual As Integer
Dim aumentoHp As Integer

Dim aumentoHitMinimo As Integer
Dim aumentoSTParcial As Integer
Dim aumentoHitParcial As Integer
Dim aumentoManaParcial As Integer

nivelActual = personaje.Stats.ELV


Call getStatsIniciales(personaje, aumentoHp, aumentoST, aumentoMana, aumentoHitMinimo, aumentoHit)

For loopNivel = 2 To nivelActual
    personaje.Stats.ELV = loopNivel
    
    ' Obtenemos cuanto deberia ganar por subir a este nivel
    Call getPremioSubaNivel(personaje, aumentoSTParcial, aumentoHitParcial, aumentoManaParcial)
    
    ' Vamns calculando el total
    aumentoHit = aumentoHit + aumentoHitParcial
    aumentoMana = aumentoMana + aumentoManaParcial
    aumentoST = aumentoST + aumentoSTParcial
Next

End Sub

Public Sub actualizarLenadores()
    Dim sql As String
    Dim iFileNo As Integer
    Dim records As ADODB.Recordset

    iFileNo = FreeFile

 sql = "SELECT SQL_NO_CACHE usr.* , gms.Privilegio as Privilegios, " & _
        "IF (cuenta.FECHAVENCIMIENTO IS NOT NULL AND CUENTA.FECHAVENCIMIENTO > UNIX_TIMESTAMP(), 'SI', 'NO') AS ESPREMIUM, " & _
        "cuenta.ESTADO, cuenta.BLOQUEADA, cuenta.SEGUNDOS_TDSF " & _
        "FROM " & DB_NAME_PRINCIPAL & ".usuarios AS usr " & _
        "LEFT JOIN " & DB_NAME_PRINCIPAL & ".juego_gms AS gms ON usr.ID = gms.IDUsuario " & _
        "LEFT JOIN " & DB_NAME_CUENTAS & " AS cuenta ON cuenta.IDCuenta = usr.IDCuenta " & _
        "WHERE razab='ENANO' AND claseb != 'GUERRERO' and claseb !='CAZADOR'"
    
    Set records = New ADODB.Recordset
    
    records.Open sql, conn
    
    
    Open "E:\Test.txt" For Output As #iFileNo
    
    
    
    ' Enumerate Recordset
   Do While Not records.EOF
        UserList(1).Name = records!nickb

        Call LoadUserInit(1, records)
        Call LoadUserStats(1, records)
        Call LoadUserReputacion(1, records)
        
        ' Si es mayor a nivel 20. Hago una simulación de vida hasta su nivel X. Si la vida es menor, s la asigno y actualizo
        Dim aumentoHit As Integer
        Dim aumentoST As Integer
        Dim aumentoMana As Integer
        Dim aumentoHp As Integer
        Dim nuevaVida As Integer
        
        
        Dim vidaAnterior As Integer
        Dim manaAnterior As Integer
        Dim hitAnterior As Integer
        Dim stAnterior As Integer
        
        vidaAnterior = UserList(1).Stats.MaxHP
        manaAnterior = UserList(1).Stats.MaxMAN
        stAnterior = UserList(1).Stats.MaxSta
        hitAnterior = UserList(1).Stats.MaxHIT


        Call simularPremiosSubaNivel(UserList(1), aumentoHit, aumentoST, aumentoMana)
        
        Dim mensaje As String
        
        If manaAnterior < aumentoMana Then
            Dim sql2 As String

            mensaje = "Nueva mana de " & UserList(1).Name & " es " & aumentoMana & " antes era " & manaAnterior

            Print #iFileNo, mensaje
            Debug.Print mensaje

            sql2 = "UPDATE " & DB_NAME_PRINCIPAL & ".usuarios SET MaxMANB = " & aumentoMana & " WHERE Nickb='" & UserList(1).Name & "'"

            Debug.Print sql2

            Call modMySql.ejecutarSQL(sql2)
        ElseIf manaAnterior = aumentoMana Then
            mensaje = UserList(1).Name & " mantiene la misma vida."
            Print #iFileNo, mensaje
            Debug.Print mensaje
        Else
            mensaje = UserList(1).Name & " tiene más mana ahora que con el recalculo:  " & manaAnterior & " vs " & aumentoMana
            Print #iFileNo, mensaje
            Debug.Print mensaje
        End If

        
         'nuevaVida = simularVida(UserList(1))
        
       'Dim mensaje As String
'        If nuevaVida > vidaAnterior Then
'            Dim sql2 As String
'
'            mensaje = "Nueva vida de " & UserList(1).Name & " es " & nuevaVida & " antes era " & vidaAnterior
'
'            Print #iFileNo, mensaje
'            Debug.Print mensaje
'
'            sql2 = "UPDATE " & DB_NAME_PRINCIPAL & ".usuarios SET MaxHPB = " & nuevaVida & " WHERE Nickb='" & UserList(1).Name & "'"
'
'            Debug.Print sql2
'
'            Call modMySql.ejecutarSQL(sql2)
'        ElseIf nuevaVida = vidaAnterior Then
'            mensaje = UserList(1).Name & " mantiene la misma vida."
'            Print #iFileNo, mensaje
'            Debug.Print mensaje
'        Else
'            mensaje = UserList(1).Name & " tiene más vida ahora que con el recalculo:  " & vidaAnterior & " vs " & nuevaVida
'            Print #iFileNo, mensaje
'            Debug.Print mensaje
'        End If
        
        records.MoveNext
   Loop
   
   Close #iFileNo
   

End Sub

Private Function simularVida(ByRef personaje As User) As Integer

Dim x As Long
Dim aumentoHp As Byte
Dim vidaBase As Integer
Dim nivelPersonaje As Integer

Randomize Timer

vidaBase = 15 + Int(getPromedioAumentoVida(personaje) + 0.5)

simularVida = vidaBase

nivelPersonaje = personaje.Stats.ELV

For x = 2 To nivelPersonaje

    personaje.Stats.ELV = x
    personaje.Stats.MaxHP = simularVida
        
    aumentoHp = obtenerAumentoHp(personaje)
        
    simularVida = simularVida + aumentoHp
Next

End Function
Public Sub generarSimulacion()

Randomize Timer

Dim x As Long
Dim y As Long

'Dim distribucion(8 To 11) As Long
'For x = 0 To 90000000
'    y = RandomNumberByte(8, 11)
'    distribucion(y) = distribucion(y) + 1
'Next x
'For x = LBound(distribucion) To UBound(distribucion)
'    Debug.Print x & " -> " & distribucion(x)
'Next

Dim personaje As User

Dim rangoAumento As tRango
Dim promedio As Single
Dim vidaIdeal As Single

Dim suma As Single
Dim minimo As Single
Dim maximo As Single
Dim histograma() As Long
Dim histogramaVida() As Long
Dim arriba  As Long
Dim abajo As Long
Dim igual As Long
Dim maximaVida As Long
Dim minimaVida As Long

minimo = 9999999
maximo = -9999
arriba = 0
abajo = 0
igual = 0
maximaVida = -9999
minimaVida = 999999

personaje.clase = eClases.LEÑADOR
personaje.Stats.UserAtributos(constitucion) = 17
        
rangoAumento = getRangoAumentoVida(personaje)
    
ReDim histograma(rangoAumento.minimo To rangoAumento.maximo) As Long
ReDim histogramaVida(1 To 1000) As Long


Dim iFileNo As Integer

iFileNo = FreeFile


For y = 1 To 1
    
    Dim aumentoHp As Byte
    Dim vidaBase As Integer
    
    vidaBase = 15 + Int(getPromedioAumentoVida(personaje) + 0.5)
    
    personaje.Stats.MaxHP = vidaBase
        
    'open the file for writing
    Open "E:\Test.txt" For Output As #iFileNo
    
    Print #iFileNo, "Simulacion " & y
    
    Print #iFileNo, 1 & "." & vidaBase & "." & 0 & "." & vidaBase
        
    For x = 2 To 20
        personaje.Stats.ELV = x
        
        aumentoHp = obtenerAumentoHp(personaje)
        
        histograma(aumentoHp) = histograma(aumentoHp) + 1
        
        personaje.Stats.MaxHP = personaje.Stats.MaxHP + aumentoHp
        
        vidaIdeal = getVidaIdeal(personaje)
        
       Print #iFileNo, "UPDATE tds_alta.usuarios SET MaxHPB=" & personaje.Stats.MaxHP & " WHERE AT5=" & personaje.Stats.UserAtributos(constitucion) & " AND Claseb='LEÑADOR' And ELVB = " & personaje.Stats.ELV & " AND MAXHPB < " & personaje.Stats.MaxHP & ";"
    Next
    
    ' La sumatoria deberia dar cercano a 0
    rangoAumento = getRangoAumentoVida(personaje)
    promedio = (rangoAumento.minimo + rangoAumento.maximo) / 2
    vidaIdeal = vidaBase + (personaje.Stats.ELV - 1) * promedio
               
    'close the file (if you dont do this, you wont be able to open it again!)
    Close #iFileNo
        
    suma = suma + (personaje.Stats.MaxHP - vidaIdeal)
    
    If (personaje.Stats.MaxHP - vidaIdeal) < 0 Then
        abajo = abajo + 1
    End If
    
    If (personaje.Stats.MaxHP - vidaIdeal) > 0 Then
        arriba = arriba + 1
    End If
    
    If (personaje.Stats.MaxHP - vidaIdeal) = 0 Then
        igual = igual + 1
    End If
    
    If personaje.Stats.MaxHP - vidaIdeal < minimo Then
        minimo = personaje.Stats.MaxHP - vidaIdeal
    End If
    
    If personaje.Stats.MaxHP - vidaIdeal > maximo Then
        maximo = personaje.Stats.MaxHP - vidaIdeal
    End If
    
    If personaje.Stats.MaxHP > maximaVida Then
        maximaVida = personaje.Stats.MaxHP
    End If
    
    If personaje.Stats.MaxHP < minimaVida Then
        minimaVida = personaje.Stats.MaxHP
    End If
    

    
    histogramaVida(personaje.Stats.MaxHP) = histogramaVida(personaje.Stats.MaxHP) + 1
Next

Dim loopHistograma As Long

For loopHistograma = rangoAumento.minimo To rangoAumento.maximo
    Debug.Print loopHistograma & "," & histograma(loopHistograma)
Next

For loopHistograma = LBound(histogramaVida) To UBound(histogramaVida)
    Debug.Print loopHistograma & "," & histogramaVida(loopHistograma)
Next


Debug.Print " Minimo: " & minimo & ". M?ximo: " & maximo & ". Debajo del promedio: " & abajo & ". Arriba " & arriba & "Igual: " & igual & "Vida del peor personaje " & minimaVida & " . Vida del mejor personaje " & maximaVida

End

If App.PrevInstance Then
    MsgBox "Este programa ya est? corriendo.", vbInformation, "Tirras Del Sur"
    End
End If


#If testeo = 1 Then
    MsgBox "OJO. Estas en Modo testeo"
#End If


ChDir App.Path
ChDrive App.Path

IniPath = App.Path & "\"
DatPath = App.Path & "\Dat\"

If Not InputBox("Clave", "Tierras del Sur") = "tatetimarce" Then End

'Inicio el Manager

If Not API_Manager.iniciarManager Then
    MsgBox "No se pudo conectar con el Manager", vbCritical
    #If testeo = 0 Then 'Para que me permita testear sin el manager prendido
    End
    #End If
End If

' Configuraciones r?pidas
servidorAtacado = False                     ' AntiDDos
denunciarActivado = True                    ' Se puede denunciar
charlageneral = True                        ' Chat global
ProfilePaquetes = False                     ' Profile paquetes que se reciben.



frmCargando.Show

Call LoadSini                               ' Cargamos la configuracion

Call General.iniciarEstructuras                     ' Npcs, gms, etc

Call modPersonajes.iniciarEstructuras

Call modMySql.iniciarConexionBaseDeDatos    ' Me conecto a la base de datos

Call Admin.actualizarOnlinesDB(True)        ' Actualizo los online que hay en el juego

Call CryptoInit                             ' Inicia codigos de encriptacion

Call Constantes_Generales.inicializarConstantes   ' Constantes relacionadas al juego

Call modClases.inicializarClases

Call LoadMotd                               ' El mensaje de bienvenida

Call CargarRequisitos                       ' Requisitos para entrar a la armada

Call InitTimeGetTime                        ' Intervalos

Call NPCs.iniciarEstructurasNpcs

Call modEventos.iniciarEstructuraEventos    ' Lista de Eventos

Call modDescansos.iniciarZonasDescanso      ' Zonas usadas por los eventos

Call modRings.iniciarRings                  ' Carga de rings

Call modRetos.iniciar                       ' Iniciar Sistema de Retos

Call modCapturarPantalla.iniciar            ' Sistema de captura de pantalla del usuario

Call Anticheat_MemCheck.iniciarEstructuras  ' Sistema anticheat para chequear la edici?n de la memoria del usuario

Call modCarcel.iniciar                      ' Sistema de Carcel para los Usuarios

Call CargarHechizos                         ' Hechizos

Call CargaNpcsDat                           ' Criaturas

Call LoadOBJData                            ' Objecitos

Call LoadArmasHerreria

Call LoadArmadurasHerreria

Call LoadObjCarpintero

Call CargaApuestas                          ' Sistema de apuestas

Call CargarSpawnList                        ' Sistema de entrenameinto

Call modMapa.iniciar(BootDelBackUp)         ' Cargamos el Mundo

Call mdClanes.iniciar                       ' Sistema de clanes

frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision

'Bordes del mapa
MinXBorder = XMinMapSize + (XWindow \ 2)
MaxXBorder = XMaxMapSize - (XWindow \ 2)
MinYBorder = YMinMapSize + (YWindow \ 2)
MaxYBorder = YMaxMapSize - (YWindow \ 2)

DoEvents

'Resetea las conexiones de los usuarios
Dim loopC As Integer

For loopC = 1 To MaxUsers
    UserList(loopC).ConnID = INVALID_SOCKET
    UserList(loopC).InicioConexion = 0
    UserList(loopC).ConfirmacionConexion = 0
    
    UserList(loopC).PacketNumber = 1
    UserList(loopC).MinPacketNumber = 1
    UserList(loopC).CryptOffset = 0
    
    UserList(loopC).userIndex = loopC
Next loopC

'??????????????????????????????????????????????????????????
With frmMain
    .AutoSave.Enabled = True
    .GameTimer.Enabled = True
    .Auditoria.Enabled = True
    .TIMER_AI.Enabled = True
End With
'??????????????????????????????????????????????????????????

'Configuracion de los sockets
Call IniciaWsApi(frmMain.hWnd)

SockListen = ListenForConnect(Puerto, hWndMsg, "")

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

Unload frmCargando

'Log
Call Logs.LogMain("Server iniciado " & App.Major & "." & App.Minor & "." & App.Revision)

'Ocultar
If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

End Sub

