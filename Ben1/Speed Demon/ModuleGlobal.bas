Attribute VB_Name = "Module1"
Option Explicit
Public MaxSpeed, AccelRate, BrakeRate
Public TotalScore, Rank
Public Track, Name(1 To 10) As String, SavedScore(1 To 10)
Type ScoreType
Name As String * 20
Score As Integer
End Type
Global Players As ScoreType
Global ScoreNumber As Integer


