Attribute VB_Name = "mIniManager"
Option Explicit


''
'Structure that contains a value and it's key in a INI file
'
' @param    key String containing the key associated to the value.
' @param    value String containing the value of the INI entry.
' @see      MainNode
'

Public Type ChildNode
    key As String
    value As String
End Type
