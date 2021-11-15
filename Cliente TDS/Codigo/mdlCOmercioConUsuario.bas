Attribute VB_Name = "mdlCOmercioConUsuario"
'Modulo para comerciar con otro usuario
'Por Alejo (Alejandro Santos)
'
'
'[Alejo]
Private Const MAX_ORO_LOGUEABLE = 100000
Option Explicit

Public Type tCOmercioUsuario
    DestUsu As Integer 'El otro Usuario
    DestNick As String
    Objeto As Integer 'Indice del inventario a comerciar, que objeto desea dar
    
    'El tipo de datos de Cant ahora es Long (antes Integer)
    'asi se puede comerciar con oro > 32k
    '[CORREGIDO]
    Cant As Long 'Cuantos comerciar, cuantos objetos desea dar
    '[/CORREGIDO]
    Acepto As Boolean
End Type

'origen: origen de la transaccion, originador del comando
'destino: receptor de la transaccion
Public Sub IniciarComercioConUsuario(ByVal Origen As Integer, ByVal Destino As Integer)
On Error GoTo errhandler

'Si ambos pusieron /comerciar entonces
If UserList(Origen).ComUsu.DestUsu = Destino And _
   UserList(Destino).ComUsu.DestUsu = Origen Then
    'Actualiza el inventario del usuario
    Call UpdateUserInv(True, Origen, 0)
    'Decirle al origen que abra la ventanita.
    Call SendData(ToIndex, Origen, 0, "INITCOMUSU")
    UserList(Origen).flags.Comerciando = True

    'Actualiza el inventario del usuario
    Call UpdateUserInv(True, Destino, 0)
    'Decirle al origen que abra la ventanita.
    Call SendData(ToIndex, Destino, 0, "INITCOMUSU")
    UserList(Destino).flags.Comerciando = True

    'Call EnviarObjetoTransaccion(Origen)
Else
    'Es el primero que comercia ?
    Call SendData(ToIndex, Destino, 0, "||" & UserList(Origen).Name & " desea comerciar. Si deseas aceptar, Escribe /COMERCIAR." & FONTTYPE_TALK)
    UserList(Destino).flags.TargetUser = Origen
    
End If

Exit Sub
errhandler:
    Call LogError("Error en IniciarComercioConUsuario: " & Err.Description)
End Sub

'envia a AQuien el objeto del otro
Public Sub EnviarObjetoTransaccion(ByVal AQuien As Integer)
'Dim Object As UserOBJ
Dim ObjInd As Integer
Dim ObjCant As Long

'[Alejo]: En esta funcion se centralizaba el problema
'         de no poder comerciar con mas de 32k de oro.
'         Ahora si funciona!!!

'Object.Amount = UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Cant
ObjCant = UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Cant
If UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto = FLAGORO Then
    'Object.ObjIndex = iORO
    ObjInd = iORO
Else
    'Object.ObjIndex = UserList(UserList(AQuien).ComUsu.DestUsu).Invent.Object(UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto).ObjIndex
    ObjInd = UserList(UserList(AQuien).ComUsu.DestUsu).Invent.Object(UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto).ObjIndex
End If

If ObjCant <= 0 Or ObjInd <= 0 Then Exit Sub

'If Object.ObjIndex > 0 And Object.Amount > 0 Then
'    Call SendData(ToIndex, AQuien, 0, "COMUSUINV" & 1 & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
'    & ObjData(Object.ObjIndex).ObjType & "," _
'    & ObjData(Object.ObjIndex).MaxHIT & "," _
'    & ObjData(Object.ObjIndex).MinHIT & "," _
'    & ObjData(Object.ObjIndex).MaxDef & "," _
'    & ObjData(Object.ObjIndex).Valor \ 3)
'End If
If ObjInd > 0 And ObjCant > 0 Then
    Call SendData(ToIndex, AQuien, 0, "COMUSUINV" & 1 & "," & ObjInd & "," & ObjData(ObjInd).Name & "," & ObjCant & "," & 0 & "," & ObjData(ObjInd).GrhIndex & "," _
    & ObjData(ObjInd).ObjType & "," _
    & ObjData(ObjInd).MaxHIT & "," _
    & ObjData(ObjInd).MinHIT & "," _
    & ObjData(ObjInd).MaxDef & "," _
    & ObjData(ObjInd).Valor \ 3)
End If

End Sub

Public Sub FinComerciarUsu(ByVal UserIndex As Integer)
With UserList(UserIndex)
    If .ComUsu.DestUsu > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "FINCOMUSUOK")
    End If
    
    .ComUsu.Acepto = False
    .ComUsu.Cant = 0
    .ComUsu.DestUsu = 0
    .ComUsu.Objeto = 0
    .ComUsu.DestNick = ""
    .flags.Comerciando = False
End With

End Sub

Public Sub AceptarComercioUsu(ByVal UserIndex As Integer)
Dim Obj1 As Obj, Obj2 As Obj
Dim OtroUserIndex As Integer
Dim TerminarAhora As Boolean

TerminarAhora = False

If UserList(UserIndex).ComUsu.DestUsu <= 0 Then
    TerminarAhora = True
End If

OtroUserIndex = UserList(UserIndex).ComUsu.DestUsu


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

If TerminarAhora = True Then
    Call FinComerciarUsu(UserIndex)
    Call FinComerciarUsu(OtroUserIndex)
    Exit Sub
End If

UserList(UserIndex).ComUsu.Acepto = True
TerminarAhora = False

If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = False Then
    Call SendData(ToIndex, UserIndex, 0, "Y209")
    Exit Sub
End If

'[Alejo]: Creo haber podido erradicar el bug de
'         no poder comerciar con mas de 32k de oro.
'         Las lineas comentadas en los siguientes
'         2 grandes bloques IF (4 lineas) son las
'         que originaban el problema.

If UserList(UserIndex).ComUsu.Objeto = FLAGORO Then
    'Obj1.Amount = UserList(UserIndex).ComUsu.Cant
    Obj1.ObjIndex = iORO
    'If Obj1.Amount > UserList(UserIndex).Stats.GLD Then
    If UserList(UserIndex).ComUsu.Cant > UserList(UserIndex).Stats.GLD Then
        Call SendData(ToIndex, UserIndex, 0, "Y210")
        TerminarAhora = True
    End If
Else
    Obj1.Amount = UserList(UserIndex).ComUsu.Cant
    Obj1.ObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).ComUsu.Objeto).ObjIndex
    If Obj1.Amount > UserList(UserIndex).Invent.Object(UserList(UserIndex).ComUsu.Objeto).Amount Then
        Call SendData(ToIndex, UserIndex, 0, "Y210")
        TerminarAhora = True
    End If
End If
If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
    'Obj2.Amount = UserList(OtroUserIndex).ComUsu.Cant
    Obj2.ObjIndex = iORO
    'If Obj2.Amount > UserList(OtroUserIndex).Stats.GLD Then
    If UserList(OtroUserIndex).ComUsu.Cant > UserList(OtroUserIndex).Stats.GLD Then
        Call SendData(ToIndex, OtroUserIndex, 0, "Y210")
        TerminarAhora = True
    End If
Else
    Obj2.Amount = UserList(OtroUserIndex).ComUsu.Cant
    Obj2.ObjIndex = UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).ObjIndex
    If Obj2.Amount > UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).Amount Then
        Call SendData(ToIndex, OtroUserIndex, 0, "Y210")
        TerminarAhora = True
    End If
End If

'Por si las moscas...
If TerminarAhora = True Then
    Call FinComerciarUsu(UserIndex)
    Call FinComerciarUsu(OtroUserIndex)
    Exit Sub
End If

'[CORREGIDO]
'Desde ac� correg� el bug que cuando se ofrecian mas de
'10k de oro no le llegaban al destinatario.

'pone el oro directamente en la billetera
If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
    'quito la cantidad de oro ofrecida
    UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD - UserList(OtroUserIndex).ComUsu.Cant
    If UserList(OtroUserIndex).ComUsu.Cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(Date & " " & UserList(OtroUserIndex).Name & " solto oro en comercio seguro con " & UserList(UserIndex).Name & ". Cantidad: " & UserList(OtroUserIndex).ComUsu.Cant)
    Call SendUserStatsBox(OtroUserIndex)
    'y se la doy al otro
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(OtroUserIndex).ComUsu.Cant
    If UserList(OtroUserIndex).ComUsu.Cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(Date & " " & UserList(OtroUserIndex).Name & " recibio oro en comercio seguro con " & UserList(UserIndex).Name & ". Cantidad: " & UserList(OtroUserIndex).ComUsu.Cant)
    Call SendUserStatsBox(UserIndex)
Else
    'Quita el objeto y se lo da al otro
    If MeterItemEnInventario(UserIndex, Obj2) = False Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, Obj2)
    End If
    Call QuitarObjetos(Obj2.ObjIndex, Obj2.Amount, OtroUserIndex)
End If

'pone el oro directamente en la billetera
If UserList(UserIndex).ComUsu.Objeto = FLAGORO Then
    'quito la cantidad de oro ofrecida
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - UserList(UserIndex).ComUsu.Cant
    If UserList(UserIndex).ComUsu.Cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(Date & " " & UserList(UserIndex).Name & " solto oro en comercio seguro con " & UserList(OtroUserIndex).Name & ". Cantidad: " & UserList(UserIndex).ComUsu.Cant)
    Call SendUserStatsBox(UserIndex)
    'y se la doy al otro
    UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD + UserList(UserIndex).ComUsu.Cant
    If UserList(UserIndex).ComUsu.Cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(Date & " " & UserList(OtroUserIndex).Name & " recibio oro en comercio seguro con " & UserList(UserIndex).Name & ". Cantidad: " & UserList(UserIndex).ComUsu.Cant)
    Call SendUserStatsBox(OtroUserIndex)
Else
    'Quita el objeto y se lo da al otro
    If MeterItemEnInventario(OtroUserIndex, Obj1) = False Then
        Call TirarItemAlPiso(UserList(OtroUserIndex).Pos, Obj1)
    End If
    Call QuitarObjetos(Obj1.ObjIndex, Obj1.Amount, UserIndex)
End If

'[/CORREGIDO] :p

Call UpdateUserInv(True, UserIndex, 0)
Call UpdateUserInv(True, OtroUserIndex, 0)

Call FinComerciarUsu(UserIndex)
Call FinComerciarUsu(OtroUserIndex)
 
End Sub

'[/Alejo]

