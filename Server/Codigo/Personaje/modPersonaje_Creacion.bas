Attribute VB_Name = "modPersonaje_Creacion"
Option Explicit

' Hombres
Private Const MIN_CABEZA_HOMBRE_HUMANO = 1
Private Const MAX_CABEZA_HOMBRE_HUMANO = 26
Private Const MIN_CABEZA_HOMBRE_ELFO = 102
Private Const MAX_CABEZA_HOMBRE_ELFO = 111
Private Const MIN_CABEZA_HOMBRE_ELFO_OSCURO = 201
Private Const MAX_CABEZA_HOMBRE_ELFO_OSCURO = 205
Private Const MIN_CABEZA_HOMBRE_ENANO = 301
Private Const MAX_CABEZA_HOMBRE_ENANO = 305
Private Const MIN_CABEZA_HOMBRE_GNOMO = 401
Private Const MAX_CABEZA_HOMBRE_GNOMO = 406

Private Const MIN_CUERPO_HOMBRE_HUMANO = 21
Private Const MIN_CUERPO_HOMBRE_ELFO = 21
Private Const MIN_CUERPO_HOMBRE_ELFO_OSCURO = 32
Private Const MIN_CUERPO_HOMBRE_ENANO = 53
Private Const MIN_CUERPO_HOMBRE_GNOMO = 53

' Mujeres
Private Const MIN_CABEZA_MUJER_HUMANA = 71
Private Const MAX_CABEZA_MUJER_HUMANA = 79
Private Const MIN_CABEZA_MUJER_ELFA = 170
Private Const MAX_CABEZA_MUJER_ELFA = 176
Private Const MIN_CABEZA_MUJER_ELFA_OSCURA = 270
Private Const MAX_CABEZA_MUJER_ELFA_OSCURA = 279
Private Const MIN_CABEZA_MUJER_ENANA = 370
Private Const MAX_CABEZA_MUJER_ENANA = 371
Private Const MIN_CABEZA_MUJER_GNOMA = 470
Private Const MAX_CABEZA_MUJER_GNOMA = 475

Private Const MIN_CUERPO_MUJER_HUMANO = 36
Private Const MIN_CUERPO_MUJER_ELFO = 39
Private Const MIN_CUERPO_MUJER_ELFO_OSCURO = 40
Private Const MIN_CUERPO_MUJER_ENANO = 60
Private Const MIN_CUERPO_MUJER_GNOMO = 60

Public Sub GenerarCuerpoYCabeza(personaje As User)

Dim UserHead As Integer
Dim UserBody As Integer

Select Case personaje.Genero

   Case eGeneros.Hombre
   
        Select Case personaje.Raza
        
                Case eRazas.Humano
                    UserHead = HelperRandom.RandomIntNumber(MIN_CABEZA_HOMBRE_HUMANO, MAX_CABEZA_HOMBRE_HUMANO)
                    UserBody = MIN_CUERPO_HOMBRE_HUMANO
                    
                Case eRazas.Elfo
                    UserHead = HelperRandom.RandomIntNumber(MIN_CABEZA_HOMBRE_ELFO, MAX_CABEZA_HOMBRE_ELFO)
                    UserBody = MIN_CUERPO_HOMBRE_ELFO
                    
                Case eRazas.ElfoOscuro
                    UserHead = HelperRandom.RandomIntNumber(MIN_CABEZA_HOMBRE_ELFO_OSCURO, MAX_CABEZA_HOMBRE_ELFO_OSCURO)
                    UserBody = MIN_CUERPO_HOMBRE_ELFO_OSCURO
                    
                Case eRazas.Enano
                    UserHead = HelperRandom.RandomIntNumber(MIN_CABEZA_HOMBRE_ENANO, MAX_CABEZA_HOMBRE_ENANO)
                    UserBody = MIN_CUERPO_HOMBRE_ENANO
                    
                Case eRazas.Gnomo
                    UserHead = HelperRandom.RandomIntNumber(MIN_CABEZA_HOMBRE_GNOMO, MAX_CABEZA_HOMBRE_GNOMO)
                    UserBody = MIN_CUERPO_HOMBRE_GNOMO
                                        
        End Select
        
   Case eGeneros.Mujer
   
        Select Case personaje.Raza
        
                Case eRazas.Humano
                    UserHead = HelperRandom.RandomIntNumber(MIN_CABEZA_MUJER_HUMANA, MAX_CABEZA_MUJER_HUMANA)
                    UserBody = MIN_CUERPO_MUJER_HUMANO
                    
                Case eRazas.Elfo
                    UserHead = HelperRandom.RandomIntNumber(MIN_CABEZA_MUJER_ELFA, MAX_CABEZA_MUJER_ELFA)
                    UserBody = MIN_CUERPO_MUJER_ELFO
                    
                Case eRazas.ElfoOscuro
                    UserHead = HelperRandom.RandomIntNumber(MIN_CABEZA_MUJER_ELFA_OSCURA, MAX_CABEZA_MUJER_ELFA_OSCURA)
                    UserBody = MIN_CUERPO_MUJER_ELFO_OSCURO
                    
                Case eRazas.Enano
                    UserHead = HelperRandom.RandomIntNumber(MIN_CABEZA_MUJER_ENANA, MAX_CABEZA_MUJER_ENANA)
                    UserBody = MIN_CUERPO_MUJER_ENANO
                
                Case eRazas.Gnomo
                    UserHead = HelperRandom.RandomIntNumber(MIN_CABEZA_MUJER_GNOMA, MAX_CABEZA_MUJER_GNOMA)
                    UserBody = MIN_CUERPO_MUJER_GNOMO
                
        End Select
End Select

personaje.Char.Head = UserHead
personaje.Char.Body = UserBody

End Sub
