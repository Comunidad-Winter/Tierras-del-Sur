Attribute VB_Name = "Engine_General"
Option Explicit


Public NoUsarSombras            As Boolean 'true si no se usan sombras
Public NoUsarLuces              As Boolean 'true si no se usan luces
Public NoUsarParticulas         As Boolean 'true si no se usan partículas
Public CambiarResolucion        As Boolean 'true si se cambia la resolucion
Public AnimarAguatierra         As Boolean 'true si el aguatierra es animada
Public SombrasHQ                As Boolean 'Sombras para cada objeto vertical.
Public Optimizar_Textos         As Boolean 'Dibuja el texto en una textura, y despues dibuja la textura. POR AHORA NO SE USA
Public UsarVSync                As Boolean 'Sincronizar renders con el refresco del monitor
Public cfgSoportaPointSprites   As Boolean 'Error de las particulas. Se arregla: cfgSoportaPointSprites=false
Public usaBumpMapping           As Boolean


