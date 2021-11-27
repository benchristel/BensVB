Attribute VB_Name = "Module1"
Type Star
    XCoord As Double
    YCoord As Double
    XVector As Double
    YVector As Double
    LastX As Double
    LastY As Double
    Mass As Integer
    Color As Long
End Type
Type Slope
    Rise As Double
    Run As Double
End Type
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Star() As Star, starcount As Integer
Public StarDC As Long, trails As Boolean, gravity As Integer

Public Function FindLine(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Length) As Slope
Dim proportion, pointdistance
pointdistance = Distance(X1, Y1, X2, Y2)
If pointdistance = 0 Then
FindLine.Rise = 0
FindLine.Run = 0
Exit Function
End If
proportion = Length / pointdistance
FindLine.Run = proportion * (X2 - X1)
FindLine.Rise = proportion * (Y2 - Y1)
End Function

Public Function Distance(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Double
Distance = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
End Function
