Attribute VB_Name = "ME_constructoresAccionEditor"
Option Explicit
'TODO poner las que correspondan

'************************************************************************************'
Public Function construirAccionTileExitComun() As cAccionTileEditor

    Dim parametro As cParamAccionTileEditor
    
    Set construirAccionTileExitComun = New cAccionTileEditor
    
    Call construirAccionTileExitComun.crear("Exit Común", "Al pasar por aquí el personaje es teletraportado a las coordenadas establecidas.", New cAccionExit)

    Call construirAccionTileExitComun.agregarParametro(construirCampoMapa)

    Call construirAccionTileExitComun.agregarParametro(construirCampoPosicionX())

    Call construirAccionTileExitComun.agregarParametro(construirCampoPosicionY)

    Set parametro = New cParamAccionTileEditor
    Call parametro.crear("Con efecto(0=NO/1=SI)", "0", eTIPO_PARAMETRO.NUMERICO, 0, 1)
    Call construirAccionTileExitComun.agregarParametro(parametro)
    
    Set parametro = New cParamAccionTileEditor
    Call parametro.crear("Radio", "3", eTIPO_PARAMETRO.NUMERICO, 0, 10)
    Call parametro.setAyuda("Si el radio es mayor a 0 se teletrasporta al usuario dentro del radio especificado a partir de la posicion establecida. Con esto se evita los campeos. ")
    Call construirAccionTileExitComun.agregarParametro(parametro)
End Function

Public Function construirAccionAbrirPuerta() As cAccionTileEditor

    Dim parametro As cParamAccionTileEditor
    
    Set construirAccionAbrirPuerta = New cAccionTileEditor
    
    Call construirAccionAbrirPuerta.crear("Abrir puerta", "Abre y/o cierra una puerta.", New cAccionExit)

    Call construirAccionAbrirPuerta.agregarParametro(construirCampoMapa)

    Call construirAccionAbrirPuerta.agregarParametro(construirCampoPosicionX())

    Call construirAccionAbrirPuerta.agregarParametro(construirCampoPosicionY())

   ' Set parametro = New cParamAccionTileEditor
   ' Call parametro.crear("Grh apertura", "0", eTIPO_PARAMETRO.NUMERICO, 0, 30000)
   ' Call construirAccionQuitarVida.agregarParametro(parametro)
    
    'Set parametro = New cParamAccionTileEditor
   ' Call parametro.crear("Radio", "3", eTIPO_PARAMETRO.NUMERICO, 0, 10)
   ' Call parametro.setAyuda("Si el radio es mayor a 0 se teletrasporta al usuario dentro del radio especificado a partir de la posicion establecida. Con esto se evita los campeos. ")
   ' Call construirAccionQuitarVida.agregarParametro(parametro)
End Function
'************************************************************************************'
Public Function construirAccionTileExitAutomaticoDerecha(Mapa As Integer) As cAccionTileEditor

    Set construirAccionTileExitAutomaticoDerecha = New cAccionTileEditor
    Dim parametro As cParamAccionTileEditor
    
    Call construirAccionTileExitAutomaticoDerecha.crear("Exit derecho", "Exit hacia la derecha del mapa", New cAccionExit)

    Call construirAccionTileExitAutomaticoDerecha.agregarParametro(construirCampoMapa)

    Set parametro = New cParamAccionTileEditor
    Call parametro.crear("Con efecto(S/N)", "S", eTIPO_PARAMETRO.NUMERICO, 0, 100)
    Call construirAccionTileExitAutomaticoDerecha.agregarParametro(parametro)
End Function

Public Function construirAccionTileEjecutarEntidadWorldPos() As cAccionTileEditor

    Set construirAccionTileEjecutarEntidadWorldPos = New cAccionTileEditor
    Dim parametro As cParamAccionTileEditor
    
    Call construirAccionTileEjecutarEntidadWorldPos.crear("Ejecutar Entidad WorldPos", "Ejecuta la entidad que se encuentra en la posicion indicada", New cAccionExit)

    Call construirAccionTileEjecutarEntidadWorldPos.agregarParametro(construirCampoPosicionX)
    Call construirAccionTileEjecutarEntidadWorldPos.agregarParametro(construirCampoPosicionY)

End Function
'************************************************************************************'
Public Function construirAccionTileBloquearPase() As cAccionTileEditor

    Set construirAccionTileBloquearPase = New cAccionTileEditor
    Dim parametro As cParamAccionTileEditor
    
    Call construirAccionTileBloquearPase.crear("Bloquear pase", "Luego de que hayan pasando N personas bloquea la posicion relativa indicada.", New cAccionExit)

    Set parametro = New cParamAccionTileEditor
    Call parametro.crear("Cantidad Personas", "1", eTIPO_PARAMETRO.NUMERICO, 1, 300)
    Call parametro.setAyuda("Cuando pasen N personas por este tile (no podrán volver para atras) se bloqueará el paso")
    Call construirAccionTileBloquearPase.agregarParametro(parametro)
    Call construirAccionTileBloquearPase.agregarParametro(construirCampoPosicionX)
    Call construirAccionTileBloquearPase.agregarParametro(construirCampoPosicionY)

End Function







'************************************************************************************'
' Constructores de campos comunes
'************************************************************************************'
Private Function construirCampoPosicionX() As cParamAccionTileEditor
    Set construirCampoPosicionX = New cParamAccionTileEditor
    Call construirCampoPosicionX.crear("X Posicion", "", eTIPO_PARAMETRO.NUMERICO, X_MINIMO_USABLE, X_MAXIMO_USABLE)
    Call construirCampoPosicionX.setAyuda("Coordenada X hacia donde será enviado. Número entre " & X_MINIMO_USABLE & " y " & X_MAXIMO_USABLE)
End Function

Private Function construirCampoPosicionY() As cParamAccionTileEditor
    Set construirCampoPosicionY = New cParamAccionTileEditor
    Call construirCampoPosicionY.crear("Y Posicion", "", eTIPO_PARAMETRO.NUMERICO, Y_MINIMO_USABLE, Y_MAXIMO_USABLE)
    Call construirCampoPosicionY.setAyuda("Coordenada Y hacia donde será enviado. Número entre " & Y_MINIMO_USABLE & " y " & Y_MAXIMO_USABLE)
End Function

Public Function construirCampoMapa() As cParamAccionTileEditor
    Set construirCampoMapa = New cParamAccionTileEditor
    Call construirCampoMapa.crear("Mapa", "", eTIPO_PARAMETRO.NUMERICO, 1, 1600)
    Call construirCampoMapa.setAyuda("Numero del mapa hacia donde será enviado.")
End Function

