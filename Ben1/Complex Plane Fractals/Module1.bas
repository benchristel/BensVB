Attribute VB_Name = "Module1"
Public Coeff(0 To 9) As Single, Tolerance As Double
Public Const PI = 3.14159265898
Type Complex
    real As Double
    imag As Double
End Type


Public Function ComplexPwr(z As Complex, n As Integer) As Complex
Dim r As Integer 'index variable
Dim term As Double
For r = 0 To n
    term = Factorial(n) / (Factorial(r) * Factorial(n - r)) * z.real ^ (n - r) * z.imag ^ r
    If r Mod 4 = 2 Or r Mod 4 = 3 Then
        term = -term
    End If
    If r Mod 4 = 1 Or r Mod 4 = 3 Then
        ComplexPwr.imag = ComplexPwr.imag + term
    Else
        ComplexPwr.real = ComplexPwr.real + term
    End If
Next r
End Function

Public Function Factorial(x As Integer) As Long
Dim j As Integer
Factorial = 1
On Error GoTo 1
For j = 2 To x
    Factorial = Factorial * j
Next j
1:
End Function

Public Function ComplexDiv(a As Complex, b As Complex) As Complex
ComplexDiv.real = (a.real * b.real + a.imag * b.imag) / (b.real ^ 2 + b.imag ^ 2)
ComplexDiv.imag = (a.imag * b.real - a.real * b.imag) / (b.real ^ 2 + b.imag ^ 2)
End Function

Public Function FVal(z As Complex) As Complex
FVal.real = Coeff(0) + Coeff(1) * z.real + Coeff(2) * ComplexPwr(z, 2).real _
    + Coeff(3) * ComplexPwr(z, 3).real + Coeff(4) * ComplexPwr(z, 4).real _
    + Coeff(5) * ComplexPwr(z, 5).real + Coeff(6) * ComplexPwr(z, 6).real _
    + Coeff(7) * ComplexPwr(z, 7).real + Coeff(8) * ComplexPwr(z, 8).real _
    + Coeff(9) * ComplexPwr(z, 9).real
FVal.imag = Coeff(0) + Coeff(1) * z.imag + Coeff(2) * ComplexPwr(z, 2).imag _
    + Coeff(3) * ComplexPwr(z, 3).imag + Coeff(4) * ComplexPwr(z, 4).imag _
    + Coeff(5) * ComplexPwr(z, 5).imag + Coeff(6) * ComplexPwr(z, 6).imag _
    + Coeff(7) * ComplexPwr(z, 7).imag + Coeff(8) * ComplexPwr(z, 8).imag _
    + Coeff(9) * ComplexPwr(z, 9).imag
End Function

Public Function FDeriv(z As Complex) As Complex
FDeriv.real = Coeff(1) + 2 * Coeff(2) * z.real _
    + 3 * Coeff(3) * ComplexPwr(z, 2).real + 4 * Coeff(4) * ComplexPwr(z, 3).real _
    + 5 * Coeff(5) * ComplexPwr(z, 4).real + 6 * Coeff(6) * ComplexPwr(z, 5).real _
    + 7 * Coeff(7) * ComplexPwr(z, 6).real + 8 * Coeff(8) * ComplexPwr(z, 7).real _
    + 9 * Coeff(9) * ComplexPwr(z, 8).real
FDeriv.imag = Coeff(1) + 2 * Coeff(2) * z.imag _
    + 3 * Coeff(3) * ComplexPwr(z, 2).imag + 4 * Coeff(4) * ComplexPwr(z, 3).imag _
    + 5 * Coeff(5) * ComplexPwr(z, 4).imag + 6 * Coeff(6) * ComplexPwr(z, 5).imag _
    + 7 * Coeff(7) * ComplexPwr(z, 6).imag + 8 * Coeff(8) * ComplexPwr(z, 7).imag _
    + 9 * Coeff(9) * ComplexPwr(z, 8).imag
End Function

Public Function complexdist(a As Complex, b As Complex) As Double
complexdist = ((a.imag - b.imag) ^ 2 + (a.real - b.real) ^ 2) ^ 0.5
End Function
