Attribute VB_Name = "DevToolsMod"
Private TimerStart As Double

' Call this before the code segment to time
Public Sub TimerStartCount()
    TimerStart = Timer
End Sub

' Call this after the code segment to time, returns elapsed seconds
' Call this after the code segment to time, outputs elapsed time to Immediate Window
Public Sub TimerEndCount(PrintLabel As String)
    Dim duration As Double
    duration = Timer - TimerStart
    Debug.Print PrintLabel & " Elapsed time: " & Format(duration, "0.000") & " seconds."
End Sub

'call before running some heavy sheet operation code as true, and call again after as false to end.
Public Sub OptimizePerformance(enable As Boolean)
    If enable Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    Else
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
    End If
End Sub


