Attribute VB_Name = "modAPI"
Public KeyPressed As Boolean
Public TimeElapsed As Integer, AverageTime As Double, BeatCount As Integer, BPM As Double
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Const KEY_PRESSED As Integer = &H1000

Public Sub Update(Ticks As Integer)
    AverageTime = (AverageTime * (BeatCount) + Ticks) / (BeatCount + 1)
    BPM = 60000 / AverageTime
    frmBPMCalc.lblOutput.Caption = BPM
    BeatCount = BeatCount + 1
End Sub
