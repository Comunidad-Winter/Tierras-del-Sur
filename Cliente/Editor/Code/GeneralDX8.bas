Attribute VB_Name = "ME_General"
Option Explicit

Public Const VERSION_EDITOR = "Final"

Public ModificandoOpcionesEditor As Boolean

Public Sub AplicarConfiguracion()

    SombrasHQ = ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("SOMBRAS") = "SI"
    cfgSoportaPointSprites = ME_Configuracion_Usuario.obtenerPreferenciaWorkSpace("SPRITES") = "SI"
          
    CambiarResolucion = False
    NoUsarSombras = False
    NoUsarLuces = False
    NoUsarParticulas = False
    AnimarAguatierra = True
    Optimizar_Textos = True
    UsarVSync = False
    usaBumpMapping = True

    IniPath = Clientpath & "Init\"
    DBPath = app.Path & "\Datos\DB\Raw\"

End Sub

