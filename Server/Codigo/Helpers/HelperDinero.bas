Attribute VB_Name = "HelperDinero"

Public Function formatearDinero(monto As Long) As String
    formatearDinero = FormatNumber(monto, 0, vbTrue, vbFalse, vbTrue)
End Function
