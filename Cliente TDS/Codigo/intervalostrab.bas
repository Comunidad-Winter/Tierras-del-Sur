Attribute VB_Name = "intervalostrab"
'////////////////////////////////////////////////
'/ Idea creada por Marche ///////////////////////
'/Todos los derechos reservados./////////////////
'/El uso de este modulo sin autorizacion del ////
'autor esta penado.//////////////////////////////
'Creado exclusivamente para Tierras del Sur//////
'/////////////////////////////////////////////////

Sub DameIntervalo()
Select Case UserLvl
Case Is <= 5
frmMain.IntervaloLaburar.Interval = 1000
Case Is < 14
frmMain.IntervaloLaburar.Interval = 900
Case Is < 24
frmMain.IntervaloLaburar.Interval = 700
Case Is >= 24
frmMain.IntervaloLaburar.Interval = 500
End Select
End Sub
