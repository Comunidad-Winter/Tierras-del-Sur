Attribute VB_Name = "modColeccion"
' Funciones para menejar colecciones (Collection)
Option Explicit


' Busca dentro de la coleccion el item cuyo valor para la clave ID sea igual a @identificador
Public Function Coleccion_Join(Bcoleccion As Collection) As String
   Dim elemento As Variant
   
    Coleccion_Join = ""
    
    For Each elemento In Bcoleccion
        Coleccion_Join = Coleccion_Join & CStr(elemento) & ", "
    Next
        
End Function

' Busca dentro de la coleccion el item cuyo valor para la clave ID sea igual a @identificador
Public Function obtenerDato(ByVal identificador As Long, coleccion As Collection) As Dictionary

    Dim elemento As Dictionary
    
    For Each elemento In coleccion
    
        If elemento.item("id") = identificador Then
            Set obtenerDato = elemento
            Exit Function
        End If
    Next
    
    Set obtenerDato = Nothing
    
End Function

' Devuelve la cantidad de items que hay en al coleccion cuyo valor en la clave @clave sea igual a @valor
Public Function contar(clave As String, Valor As String, coleccion As Collection) As Integer

    Dim elemento As Dictionary
    
    contar = 0
    
    For Each elemento In coleccion
        
        If elemento.item(clave) = Valor Then
            contar = contar + 1
        End If
        
    Next
End Function

' Devuelve todos los valores de @claveinteresada de todos los items cuya clave @clave sea igual a @valor
Public Sub buscarDonde(clave As String, Valor As String, coleccion As Collection, claveInteresada As String, destino As Collection)

    Dim elemento As Dictionary
        
    For Each elemento In coleccion
        If elemento.item(clave) = Valor Then
            If claveInteresada = vbNullString Then
                Call destino.Add(elemento)
            Else
                Call destino.Add(elemento.item(claveInteresada))
            End If
        End If
        
    Next
End Sub
