Attribute VB_Name = "mdlCOmercioConUsuario"
Public Const MAX_OBJETOS_COMERCIABLES = 29 ' A partir de 0. Son 30 total
Option Explicit


Public Type tCOmercioUsuario
    DestUsu As Integer 'El otro Usuario
    DestNick As String
    Acepto As Boolean
    ' Objetos
    objeto(0 To MAX_OBJETOS_COMERCIABLES) As Integer 'Indice del inventario a comerciar, que objeto desea dar
    ObjetoIndex(0 To MAX_OBJETOS_COMERCIABLES) As Integer '' Object index
    cant(0 To MAX_OBJETOS_COMERCIABLES) As Long 'Cuantos comerciar, cuantos objetos desea dar
End Type

'origen: origen de la transaccion, originador del comando
'destino: receptor de la transaccion
Public Sub IniciarComercioConUsuario(ByVal Origen As Integer, ByVal destino As Integer)

' ¿El evento les permite comerciar?
If Not UserList(Origen).evento Is Nothing Then
    If Not UserList(Origen).evento.puedeTirarObjeto(Origen, 0, 0, eDestinoObjeto.Usuario, destino) Then
        Exit Sub
    End If
End If

If Not UserList(destino).evento Is Nothing Then
    If Not UserList(destino).evento.puedeTirarObjeto(destino, 0, 0, eDestinoObjeto.Usuario, Origen) Then
        Exit Sub
    End If
End If


'Si ambos pusieron /comerciar entonces
If UserList(Origen).ComUsu.DestUsu = destino And _
   UserList(destino).ComUsu.DestUsu = Origen Then
    'Actualiza el inventario del usuario
    Call UpdateUserInv(True, Origen, 0)
    'Decirle al origen que abra la ventanita.
    EnviarPaquete Paquetes.InitComUsu, "", Origen
    UserList(Origen).flags.Comerciando = True
    'Actualiza el inventario del usuario
    Call UpdateUserInv(True, destino, 0)
    'Decirle al origen que abra la ventanita.
    EnviarPaquete Paquetes.InitComUsu, "", destino
    UserList(destino).flags.Comerciando = True
    'Call EnviarObjetoTransaccion(Origen)
    
    '[Wizard 03/09/05] Si el gm comercia seguro deja logs.
    If UserList(Origen).flags.Privilegios > 1 Then
        Call LogGM(UserList(Origen).id, " Ha mantenido comercio seguro con " & UserList(destino).Name)
    End If
    If UserList(destino).flags.Privilegios > 1 Then
        Call LogGM(UserList(destino).id, " Ha mantenido comercio seguro con " & UserList(Origen).Name)
    End If
    '[Wizard]
Else
    'Es el primero que comercia ?
    EnviarPaquete Paquetes.mensajeinfo, "Ahora debes esperar que " & UserList(destino).Name & " acepte el pedido de comercio.", Origen
    EnviarPaquete Paquetes.MensajeTalk, UserList(Origen).Name & " desea comerciar. Si deseas aceptar, Escribe /COMERCIAR.", destino
    UserList(destino).flags.TargetUser = Origen
End If

End Sub

'envia a AQuien el objeto del otro
Public Sub EnviarObjetoTransaccion(ByVal aQuien As Integer, ByVal nIndex As Integer)
'Dim Object As UserOBJ
Dim ObjInd As Integer
Dim ObjCant As Long
Dim index As Integer
Dim cadena As String

'Object.Amount = UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Cant
For index = 0 To nIndex - 1
ObjCant = UserList(UserList(aQuien).ComUsu.DestUsu).ComUsu.cant(index)
If UserList(UserList(aQuien).ComUsu.DestUsu).ComUsu.objeto(index) = FLAGORO Or UserList(UserList(aQuien).ComUsu.DestUsu).ComUsu.objeto(index) = 255 Then
    'Object.ObjIndex = iORO
    ObjInd = iORO
Else
    'Object.ObjIndex = UserList(UserList(AQuien).ComUsu.DestUsu).Invent.Object(UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto).ObjIndex
    ObjInd = UserList(UserList(aQuien).ComUsu.DestUsu).Invent.Object(UserList(UserList(aQuien).ComUsu.DestUsu).ComUsu.objeto(index)).ObjIndex
End If
'If ObjCant <= 0 Or ObjInd <= 0 Then Exit Sub
    If ObjInd > 0 And ObjCant > 0 Then
    cadena = cadena & ITS(ObjInd) & ITS(ObjData(ObjInd).GrhIndex) & Chr$(ObjData(ObjInd).ObjType) & ITS(ObjData(ObjInd).MaxHIT) & ITS(ObjData(ObjInd).MinHIT) & ITS(ObjData(ObjInd).MaxDef) & LongToString(ObjData(ObjInd).valor \ 3) & LongToString(ObjCant)
    End If
Next index


   EnviarPaquete Paquetes.ComUsuInv, cadena, aQuien
End Sub

Public Sub FinComerciarUsuAmbos(ByVal UserIndex As Integer)
With UserList(UserIndex)

    If .ComUsu.DestUsu > 0 Then
        If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex And _
            UserList(.ComUsu.DestUsu).ComUsu.DestNick = UserList(UserIndex).Name Then
            FinComerciarUsu (.ComUsu.DestUsu)
        End If
    End If

    FinComerciarUsu (UserIndex)
End With
End Sub

Public Sub FinComerciarUsu(ByVal UserIndex As Integer)
    ResetUserComercio (UserIndex)
    EnviarPaquete Paquetes.FinComUsuOk, "", UserIndex, ToIndex
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AceptarComercioUsu
' DateTime  : 23/04/2007 21:46
' Author    : Marce
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub AceptarComercioUsu(ByVal UserIndex As Integer)

Dim obj As obj
Dim OtroUserIndex As Integer
Dim TerminarAhora As Boolean
Dim i As Integer
Dim CantidadObjetosUserIndex As Byte
Dim CantidadObjetosOtroUserIndex As Byte
Dim log As String
Dim IndexItem As Integer
Dim idComercio As Long
Dim infoComercio As ADODB.Recordset

TerminarAhora = False

' Personaje con el que quiere comerciar
OtroUserIndex = UserList(UserIndex).ComUsu.DestUsu

' Chequeo que sea un Index Valido
If OtroUserIndex <= 0 Then
   TerminarAhora = True
   Call FinComerciarUsu(UserIndex)
   Exit Sub
End If

' Chequos para saber si el personaje esta logueado y si quiere comerciar conmigo
If UserList(OtroUserIndex).flags.UserLogged = False Or UserList(UserIndex).flags.UserLogged = False Then
    TerminarAhora = True
End If

If UserList(OtroUserIndex).ComUsu.DestUsu <> UserIndex Then
    TerminarAhora = True
End If


If UserList(OtroUserIndex).Name <> UserList(UserIndex).ComUsu.DestNick Then
    TerminarAhora = True
End If

If UserList(UserIndex).Name <> UserList(OtroUserIndex).ComUsu.DestNick Then
    TerminarAhora = True
End If

' ¿El evento les permite comerciar?
If Not UserList(UserIndex).evento Is Nothing Then
    If Not UserList(UserIndex).evento.puedeTirarObjeto(UserIndex, 0, 0, eDestinoObjeto.Usuario, OtroUserIndex) Then
        TerminarAhora = True
    End If
End If

If Not UserList(OtroUserIndex).evento Is Nothing Then
    If Not UserList(OtroUserIndex).evento.puedeTirarObjeto(OtroUserIndex, 0, 0, eDestinoObjeto.Usuario, UserIndex) Then
        TerminarAhora = True
    End If
End If

' Si hay algo que no esta ok termino los comercios
If TerminarAhora = True Then
    Call FinComerciarUsu(UserIndex)
    Call FinComerciarUsu(OtroUserIndex)
    Exit Sub
End If

' Acepto
UserList(UserIndex).ComUsu.Acepto = True

' El otro usuario acepto?
If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = False Then
    EnviarPaquete Paquetes.MensajeSimple, Chr$(209), UserIndex, ToIndex
    Exit Sub
End If

' Si ambos usuarios aceptaron chequeo que tengan todos los items
'***************** CHEQUEA SI TIENE TODOS LOS ITEMS ******************************'
For i = 0 To MAX_OBJETOS_COMERCIABLES
    If UserList(UserIndex).ComUsu.objeto(i) <> 0 Then
        If UserList(UserIndex).ComUsu.objeto(i) = FLAGORO Or UserList(UserIndex).ComUsu.objeto(i) = 255 Then
            If UserList(UserIndex).ComUsu.cant(i) > UserList(UserIndex).Stats.GLD Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(210), UserIndex, ToIndex
                TerminarAhora = True
            End If
        Else
            ' Chequeo que en el slot tenga el mismo item que ofrecio
            If UserList(UserIndex).ComUsu.ObjetoIndex(i) <> UserList(UserIndex).Invent.Object(UserList(UserIndex).ComUsu.objeto(i)).ObjIndex Then
                EnviarPaquete Paquetes.MensajeFight, "Intento de cargar en comercio seguro. Usuario: " & UserList(UserIndex).Name, UserIndex, ToAdmins
                TerminarAhora = True
            End If
            ' Chequeo que tenga la cantidad que esta ofreciendo
            If UserList(UserIndex).ComUsu.cant(i) > UserList(UserIndex).Invent.Object(UserList(UserIndex).ComUsu.objeto(i)).Amount Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(210), UserIndex, ToIndex
                TerminarAhora = True
            End If
        End If

        If TerminarAhora = True Then
            Call FinComerciarUsu(UserIndex)
            Call FinComerciarUsu(OtroUserIndex)
            Exit Sub
        End If
    Else ' No hay más items que revisar.
        Exit For
    End If
Next i

' No es i + 1 por que siempre se va hacer un ciclo mas hasta encontrar el igual a 0
CantidadObjetosUserIndex = i

' Chequeo que el otro usuario tambien tenga todo en orden
'*********************************************************************************
For i = 0 To MAX_OBJETOS_COMERCIABLES
    If UserList(OtroUserIndex).ComUsu.objeto(i) <> 0 Then
        If UserList(OtroUserIndex).ComUsu.objeto(i) = FLAGORO Or UserList(OtroUserIndex).ComUsu.objeto(i) = 255 Then
            If UserList(OtroUserIndex).ComUsu.cant(i) > UserList(OtroUserIndex).Stats.GLD Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(210), OtroUserIndex, ToIndex
                TerminarAhora = True
            End If
        Else
            ' Chequeo que en el slot tenga el mismo item que ofrecio
            If UserList(OtroUserIndex).ComUsu.ObjetoIndex(i) <> UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.objeto(i)).ObjIndex Then
                EnviarPaquete Paquetes.MensajeFight, "Intento de cargar en comercio seguro. Usuario: " & UserList(OtroUserIndex).Name, OtroUserIndex, ToAdmins
                TerminarAhora = True
            End If
            ' Chequeo que tenga la cantidad que esta ofreciendo
            If UserList(OtroUserIndex).ComUsu.cant(i) > UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.objeto(i)).Amount Then
                EnviarPaquete Paquetes.MensajeSimple, Chr$(210), OtroUserIndex, ToIndex
                TerminarAhora = True
            End If
        End If
       If TerminarAhora = True Then
            Call FinComerciarUsu(UserIndex)
            Call FinComerciarUsu(OtroUserIndex)
            Exit Sub
       End If
    Else ' No hay mas objetos que chequear
        Exit For
    End If
Next i

CantidadObjetosOtroUserIndex = i

'Creo el log del comercio
sql = "INSERT INTO " & DB_NAME_PRINCIPAL & ".juego_logs_comercio(IDusr0,IDusr1,IPusr0,IPusr1) VALUES(" & UserList(UserIndex).id & "," & UserList(OtroUserIndex).id & ", " & UserList(UserIndex).ip & ", " & UserList(OtroUserIndex).ip & ")"
Call modMySql.ejecutarSQL(sql)

'Obtengo el ID del comercio
sql = "SELECT last_insert_id() AS id FROM " & DB_NAME_PRINCIPAL & ".juego_logs_comercio LIMIT 1"
Set infoComercio = conn.Execute(sql, , adCmdText)
idComercio = CLng(infoComercio!id)

infoComercio.Close
Set infoComercio = Nothing

log = ""
'********************* LE DA LOS ITEMS *************************************
For i = 0 To CantidadObjetosUserIndex - 1
    'pone el oro directamente en la billetera
    If UserList(UserIndex).ComUsu.objeto(i) = FLAGORO Or UserList(UserIndex).ComUsu.objeto(i) = 255 Then
        'quito la cantidad de oro ofrecida
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - UserList(UserIndex).ComUsu.cant(i)
        'y se la doy al otro
        UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD + UserList(UserIndex).ComUsu.cant(i)
        
        IndexItem = 0
    Else
        'Si es un item, creo el item
        obj.Amount = UserList(UserIndex).ComUsu.cant(i)
        obj.ObjIndex = UserList(UserIndex).ComUsu.ObjetoIndex(i)
        
        IndexItem = obj.ObjIndex
        
        'Se lo da al otor usuario
        If MeterItemEnInventario(OtroUserIndex, obj) = False Then
            Call TirarItemAlPiso(UserList(OtroUserIndex).pos, obj)
        End If
        'Se lo quito
        Call quitarObjetos(obj.ObjIndex, obj.Amount, UserIndex)
    End If
    
    log = log & "(" & idComercio & "," & IndexItem & "," & UserList(UserIndex).ComUsu.cant(i) & ",0),"
     
Next

'***************
For i = 0 To CantidadObjetosOtroUserIndex - 1
    If UserList(OtroUserIndex).ComUsu.objeto(i) = FLAGORO Or UserList(OtroUserIndex).ComUsu.objeto(i) = 255 Then
        'quito la cantidad de oro ofrecida
        UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD - UserList(OtroUserIndex).ComUsu.cant(i)
        'y se la doy al otro
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(OtroUserIndex).ComUsu.cant(i)

        IndexItem = 0
    Else
        'Creo el objeto
        obj.Amount = UserList(OtroUserIndex).ComUsu.cant(i)
        obj.ObjIndex = UserList(OtroUserIndex).ComUsu.ObjetoIndex(i)
        
        IndexItem = obj.ObjIndex
        'Quita el objeto y se lo da al otro
        If MeterItemEnInventario(UserIndex, obj) = False Then
            Call TirarItemAlPiso(UserList(UserIndex).pos, obj)
        End If
        Call quitarObjetos(obj.ObjIndex, obj.Amount, OtroUserIndex)
    End If
    
    log = log & "(" & idComercio & "," & IndexItem & "," & UserList(OtroUserIndex).ComUsu.cant(i) & ",1)"
   
    If i < CantidadObjetosOtroUserIndex - 1 Then
        log = log & ","
    End If
Next i

'Agrego el log
sql = "INSERT INTO " & DB_NAME_PRINCIPAL & ".juego_logs_comercio_items VALUES " & log
Call modMySql.ejecutarSQL(sql)

'*************************************************************************************************
'Actualizo el oro
EnviarPaquete Paquetes.EnviarOro, Codify(UserList(UserIndex).Stats.GLD), UserIndex, ToIndex
EnviarPaquete Paquetes.EnviarOro, Codify(UserList(OtroUserIndex).Stats.GLD), OtroUserIndex, ToIndex

'Finalizo los comercios
Call FinComerciarUsu(UserIndex)
Call FinComerciarUsu(OtroUserIndex)

End Sub



Public Sub OfrecerItemsComercio(UserIndex As Integer, anexo As String)

Dim tempInt2 As Integer
Dim tempLong As Long
Dim tempbyte As Byte


For tempInt2 = 0 To (Len(anexo) / 5) - 1
        tempLong = StringToLong(anexo, tempInt2 * 5 + 2)
        tempbyte = Asc(mid(anexo, tempInt2 * 5 + 1, 1))
        
        If tempLong < 1 Then Exit Sub
        
        If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged = False Then
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        Else
            'Al que lo ofresco los items esta comerciando conmigo?
            If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu <> UserIndex And _
                 UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestNick <> UserList(UserIndex).Name Then
                 Call FinComerciarUsu(UserIndex)
                 Exit Sub
            End If
            
            If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.Muerto = 1 Then
                Call FinComerciarUsuAmbos(UserIndex)
                Exit Sub
            End If
                If tempbyte = FLAGORO Or tempbyte = 255 Then
                    'oro
                    If tempLong > UserList(UserIndex).Stats.GLD Then
                        EnviarPaquete Paquetes.MensajeSimple, Chr$(40), UserIndex
                        Call FinComerciarUsuAmbos(UserIndex)
                        Exit Sub
                    End If
                Else
                    'Hay algo en ese slot?
                    If UserList(UserIndex).Invent.Object(tempbyte).ObjIndex = 0 Then
                        EnviarPaquete Paquetes.mensajeinfo, "No tienes nada en ese slot del inventario.", UserIndex
                        Call FinComerciarUsuAmbos(UserIndex)
                        Exit Sub
                    End If
                    'Tiene la cantidad ?
                    If tempLong > UserList(UserIndex).Invent.Object(tempbyte).Amount Then
                        EnviarPaquete Paquetes.MensajeSimple, Chr$(40), UserIndex
                        Call FinComerciarUsuAmbos(UserIndex)
                        Exit Sub
                    End If
                    
                    ' Intenta ofrecer un objeto newbie?
                    If ObjData(UserList(UserIndex).Invent.Object(tempbyte).ObjIndex).Newbie = 1 Then
                        EnviarPaquete Paquetes.mensajeinfo, "No puedes intercambiar items newbies.", UserIndex
                        Call FinComerciarUsuAmbos(UserIndex)
                        Exit Sub
                    End If
                                    
                    ' No se puede tirar la armadura del dragon
                    If UserList(UserIndex).Invent.Object(tempbyte).ObjIndex = Objetos_Constantes.ARMADURA_DRAGON_E Or _
                        UserList(UserIndex).Invent.Object(tempbyte).ObjIndex = Objetos_Constantes.ARMADURA_DRAGON_H Or _
                        UserList(UserIndex).Invent.Object(tempbyte).ObjIndex = Objetos_Constantes.ARMADURA_DRAGON_M Then
                        EnviarPaquete Paquetes.mensajeinfo, "No puedes comerciar la Armadura del Dragón, el enfado del Dragón te desterrará de estas tierras.", UserIndex
                        Exit Sub
                    End If
                    
                    ' Intenta ofrecer una barca equipada?
                    
                    If ObjData(UserList(UserIndex).Invent.Object(tempbyte).ObjIndex).ObjType = OBJTYPE_BARCOS And UserList(UserIndex).Invent.Object(tempbyte).Equipped = 1 Then
                        EnviarPaquete Paquetes.mensajeinfo, "Debes desequipar el barco antes de ofrecerlo.", UserIndex
                        Call FinComerciarUsuAmbos(UserIndex)
                        Exit Sub
                    End If
                    
                    'Es un objeto faccionario.
                    If modObjeto.isFaccionario(ObjData(UserList(UserIndex).Invent.Object(tempbyte).ObjIndex)) Then
                        EnviarPaquete Paquetes.mensajeinfo, "No puedes intercambiar un item faccionario.", UserIndex
                        Call FinComerciarUsuAmbos(UserIndex)
                        Exit Sub
                    End If
                End If
       
                If UserList(UserIndex).ComUsu.objeto(tempInt2) > 0 Then
                   ' EnviarPaquete Paquetes.MensajeSimple, Chr$(127), UserIndex
                      Call FinComerciarUsuAmbos(UserIndex)
                    Exit Sub
                End If
                
                UserList(UserIndex).ComUsu.objeto(tempInt2) = tempbyte
                UserList(UserIndex).ComUsu.cant(tempInt2) = tempLong
                'Agrego que guarde el objindex por las dudas
                If Not (tempbyte = FLAGORO Or tempbyte = 255) Then
                    UserList(UserIndex).ComUsu.ObjetoIndex(tempInt2) = UserList(UserIndex).Invent.Object(tempbyte).ObjIndex
                End If
                
                If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = True Then
                    UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = False
                    EnviarPaquete Paquetes.MensajeTalk, UserList(UserIndex).Name & " ha cambiado su oferta.", UserList(UserIndex).ComUsu.DestUsu
                End If
        End If
        Next tempInt2
        Call EnviarObjetoTransaccion(UserList(UserIndex).ComUsu.DestUsu, tempInt2)
End Sub

