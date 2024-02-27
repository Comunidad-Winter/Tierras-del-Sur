Attribute VB_Name = "modEditorGenerico"
Option Explicit

'Tipos de ITEM, en vez de crear clases para heredar, definimos el tipo en la clase
Public Enum ItemType
    e_Numerico
    e_Cadena
    e_Enumerado
    e_EnumeradoDinamico
    e_MixedValue 'Muchos valores!
End Enum

Public Const SEPARE_ITEM_CHAR As Byte = 45  ' guion -
Public Const SEPARE_VALUE_CHAR As Byte = 44 ' coma ,
Public Const SEPARE_TAG_CHAR As Byte = 124  ' raya |
