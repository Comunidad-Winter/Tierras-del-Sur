VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CallBackObtenerPersonajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements callBackUrl

Private Function callBackUrl_EndRequest(request As CHTTPRequest, response As CHTTPResponse) As Boolean

    If Not response.Code = 200 Then
       Call modDibujarInterface.mostrarError(0, "Error!")
       Exit Function
    End If
    
    Dim personajesData As collection
    Set personajesData = response.bodyJSON
    
    Dim personajeData As Dictionary
    
    Dim personajes As collection
    Set personajes = New collection
    
    For Each personajeData In personajesData
        Dim personaje As personaje
        Set personaje = New personaje
        
        personaje.Nick = personajeData.item("nick")
        personaje.cabeza = personajeData.item("cabeza")
        personaje.cuerpo = personajeData.item("cuerpo")
        personaje.cuerpoDesnudo = personajeData.item("cuerpoDesnudo")
        personaje.modoCandado = personajeData.item("modoCandado")
        personaje.Password = personajeData.item("password")
        personaje.online = personajeData.item("online")
        personaje.bloqueado = personajeData.item("bloqueado")
        personaje.pelo = personajeData.item("pelo")
        personaje.barba = personajeData.item("barba")
        personaje.ropaInterior = personajeData.item("ropaInterior")
        
        Call personajes.Add(personaje)
    Next
    
    Set MiCuenta.personajes = personajes
    
    ' Mostramos la cuenta
    Call modDibujarInterface.mostrarCuenta
    
End Function

Private Function callBackUrl_EndRequestError(request As CHTTPRequest, error As String, errornumber As Integer) As Boolean
    Call modDibujarInterface.mostrarError(0, "No se pudieron obtener los personajes de tu cuenta. Por favor, volve a intentarlo.")
End Function


