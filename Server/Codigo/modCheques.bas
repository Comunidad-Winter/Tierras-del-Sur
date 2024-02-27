Attribute VB_Name = "modCheques"
Option Explicit

Public Sub cobrarCheque(personaje As User, cheque As String)

Dim sql As String
Dim infoCheque As Recordset
Dim montoCheque As Long

' ¿Tiene seleccionado una criatura?
If personaje.flags.TargetNPC = 0 Then
    EnviarPaquete Paquetes.mensajeinfo, "Primero debes hacer clic sobre un banquero.", personaje.UserIndex, ToIndex
    Exit Sub
End If

'Se asegura que el target es un npc
If NpcList(personaje.flags.TargetNPC).NPCtype <> NPCTYPE_BANQUERO Then
    EnviarPaquete Paquetes.mensajeinfo, "Primero debes hacer clic sobre un banquero.", personaje.UserIndex, ToIndex
    Exit Sub
End If

If distancia(personaje.pos, NpcList(personaje.flags.TargetNPC).pos) > 10 Then
    EnviarPaquete Paquetes.mensajeinfo, Chr$(5), personaje.UserIndex, ToIndex
    Exit Sub
End If

'Primera validacion para ver si el cheque es posta
If Len(cheque) > 8 Or Not AsciiValidos(cheque) Then
    EnviarPaquete Paquetes.DescNpc2, ITS(NpcList(personaje.flags.TargetNPC).Char.charIndex) & "El cheque que intenta retirar no existe.", personaje.UserIndex, ToIndex
    Exit Sub
End If

'Lo buscamos en la base de datos
sql = "SELECT * FROM " & DB_NAME_PRINCIPAL & ".cheques WHERE Codigo = '" & mysql_real_escape_string(cheque) & "'"

Set infoCheque = conn.Execute(sql)
    
If infoCheque.EOF = False Then
   'Sólo pueden cobrar el cheque personajes que se encuentren adheridos en la cuenta para la cual fue creada el cheque
   ' A menos que sea por marketing, para lo cual se habilita que el cheque sea para la cuenta 0
    If infoCheque!IDCuenta = personaje.IDCuenta Or (infoCheque!IDCuenta = 0 And infoCheque!motivo = "MARKETING") Then
        montoCheque = val(infoCheque!dinero)
        'Primero elimino el cheque por las dudas
        
        'infoCheque.Delete TODO re veer cual forma es l amejor
        conn.Execute "DELETE FROM " & DB_NAME_PRINCIPAL & ".cheques WHERE Codigo= '" & mysql_real_escape_string(cheque) & "'", adExecuteNoRecords
        
        'Guardo el log
        conn.Execute "INSERT INTO " & DB_NAME_PRINCIPAL & ".juego_logs_cheques_cobrados(personajeId,cuentaId,cheque,personajeNick,monto) values(" & personaje.id & "," & personaje.IDCuenta & ",'" & cheque & "','" & personaje.Name & "'," & montoCheque & ")", , adExecuteNoRecords
        
        'Le doy el oro finalmente
        personaje.Stats.GLD = personaje.Stats.GLD + montoCheque
        EnviarPaquete Paquetes.DescNpc2, ITS(NpcList(personaje.flags.TargetNPC).Char.charIndex) & "Has cobrado el cheque por " & montoCheque & " monedas de oro.", personaje.UserIndex, ToIndex
        
        'Actualizamos la info
        Call SendUserStatsBox(personaje.UserIndex)
    Else
        EnviarPaquete Paquetes.DescNpc2, ITS(NpcList(personaje.flags.TargetNPC).Char.charIndex) & "No le puedo dar el oro. El cheque sólo puede ser cobrado por personajes que pertenezcan a la cuenta para la cual fue emitido.", personaje.UserIndex, ToIndex
    End If
Else
   EnviarPaquete Paquetes.DescNpc2, ITS(NpcList(personaje.flags.TargetNPC).Char.charIndex) & "El cheque que intenta retirar no existe.", personaje.UserIndex, ToIndex
End If

infoCheque.Close
Set infoCheque = Nothing
             
End Sub
