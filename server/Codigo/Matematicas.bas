Attribute VB_Name = "Matematicas"
Option Explicit

Public Function Porcentaje(ByVal Total As Long, ByVal Porc As Long) As Long
    Porcentaje = (Total * Porc) / 100
End Function

Public Function SD(ByVal N As Integer) As Integer
'Call LogTarea("Function SD n:" & n)
'Suma digitos

Do
    SD = SD + (N Mod 10)
    N = N \ 10
Loop While (N > 0)

End Function

Public Function SDM(ByVal N As Integer) As Integer
'Call LogTarea("Function SDM n:" & n)
'Suma digitos cada digito menos dos

Do
    SDM = SDM + (N Mod 10) - 1
    N = N \ 10
Loop While (N > 0)

End Function

Public Function Complex(ByVal N As Integer) As Integer
'Call LogTarea("Complex")

If N Mod 2 <> 0 Then
    Complex = N * SD(N)
Else
    Complex = N * SDM(N)
End If

End Function

Function Distancia(ByRef wp1 As WorldPos, ByRef wp2 As WorldPos) As Long
    'Encuentra la distancia entre dos WorldPos
    Distancia = Abs(wp1.X - wp2.X) + Abs(wp1.Y - wp2.Y) + (Abs(wp1.Map - wp2.Map) * 100)
End Function

Function Distance(X1 As Variant, Y1 As Variant, X2 As Variant, Y2 As Variant) As Double

'Encuentra la distancia entre dos puntos

Distance = Sqr(((Y1 - Y2) ^ 2 + (X1 - X2) ^ 2))

End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Generates a random number in the range given - recoded to use longs and work properly with ranges
'**************************************************************
    Randomize Timer
    
    RandomNumber = Fix(Rnd * (UpperBound - LowerBound + 1)) + LowerBound
End Function
