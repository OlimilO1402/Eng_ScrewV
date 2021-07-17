Attribute VB_Name = "MMath"
Option Explicit

Private Type Point2
    X As Double
    Y As Double
End Type

Private Type Triangle
    P1 As Point2
    P2 As Point2
    P3 As Point2
End Type

Public Enum ERound
    eRoundDown = -1 'Abrunden
    eRoundMath = 0  'kaufmännisch/mathematisch runden 0-4 =Ab, 5-9=Auf
    eRoundUp = 1    'Aufrunden
    'eRoundBin = 2   'binär Runden wie in VBA.Math.Round 'No, not implemented
End Enum

Public Function CInch(v)
    CInch = v * 25.4
End Function

Public Function RoundUp(ByVal v As Double, Optional ByVal digits As Long = 0) As Double
    RoundUp = Round(v, digits, eRoundUp)
End Function
Public Function RoundDown(ByVal v As Double, Optional ByVal digits As Long = 0) As Double
    RoundDown = Round(v, digits, eRoundDown)
End Function
Public Function Round(ByVal v As Double, Optional ByVal digits As Long = 0, Optional ByVal er As ERound = eRoundMath) As Double
    'rundet auf die angegebene Stellenanzahl
    'negative Stellen = Stellen vor dem Komma
    'positive Stellen = Stellen nach dem Komma
    Round = v
    If digits <> 0 Then Round = Round * 10 ^ digits
    Select Case er
    'Case eRoundDown
        'r = Int(v * 10 ^ digits) / 10 ^ digits
        'r = Int(r)
    Case eRoundMath
        'r = Int(v * 10 ^ digits + 0.5) / 10 ^ digits
        Round = Round + 0.5
    Case eRoundUp
        'r = Int(v * 10 ^ digits + 1) / 10 ^ digits
        If Round <> 0 Then If Abs(Round / Int(Round)) <> 1 Then Round = Round + 1
    'Case eRoundBin
    '    r = VBA.Math.Round(v, Abs(digits))
    End Select
    Round = Int(Round)
    If digits <> 0 Then Round = Round / 10 ^ digits
End Function

Public Function Min(v1, v2)
    If v1 < v2 Then Min = v1 Else Min = v2
End Function
Private Function Triangle_Area(t As Triangle) As Double
    'berechnet/liefert die Fläche eines Dreiecks das durch 3 Punkte im 2D gegeben ist.
    With t
        '.P2.X -.P1.X 'AB_x
        '.P2.Y -.P1.Y 'AB_y
        '.P3.X -.P1.X 'AC_x
        '.P3.Y -.P1.Y 'AC_y
        'AC x AB
        'ABS(AC_x*AB_y-AC_y*AB_x)
        Triangle_Area = 0.5 * Abs((.P3.X - .P1.X) * (.P2.Y - .P1.Y) - (.P3.Y - .P1.Y) * (.P2.X - .P1.X))
    End With
End Function
Private Function Triangle_AreaXY(P1X As Double, P1Y As Double, P2X As Double, P2Y As Double, P3X As Double, P3Y As Double) As Double
    'mit dem Vektorprodukt aka Kreuzprodukt errechnet man die Fläche des Parallelogramms, das durch zwei Vektoren aufgespannt wird.
    Triangle_AreaXY = 0.5 * Abs((P3X - P1X) * (P2Y - P1Y) - (P3Y - P1Y) * (P2X - P1X))
End Function


