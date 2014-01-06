'Snake v1.0, 2014-01-02
'By Matt Carleton

'Class StopWatch - counts milliseconds between startTimer and EndTimer.

Private startTick As Long

Public Sub StartTimer()
    startTick = GetTickCount
End Sub

Public Function EndTimer() As Long
    EndTimer = (GetTickCount - startTick)
End Function
