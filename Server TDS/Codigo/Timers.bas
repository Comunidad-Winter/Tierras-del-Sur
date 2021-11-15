Attribute VB_Name = "Timers"
Public Timer1 As ccrpTimer




Public Sub Timer1_Timer(ByVal Milliseconds As Long)
MsgBox "D"
End Sub


Sub timers()
Set Timer1 = New ccrpTimer

With Timer1
      .EventType = TimerPeriodic
      .Interval = 1000
      .Enabled = True
   End With

End Sub
