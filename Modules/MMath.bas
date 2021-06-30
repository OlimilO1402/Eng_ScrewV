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
Public posINF As Double
Public negINF As Double
Public NaN    As Double
Public Declare Sub RtlMoveMemory Lib "kernel32" ( _
    ByRef pDst As Any, ByRef pSrc As Any, ByVal bLength As Long)

Public Sub Init()
    posINF = GetINF
    negINF = GetINF(-1)
    Call GetNaN(NaN)
End Sub

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
    If IsNaN(v) Then Exit Function
    If IsPosINF(v) Then Exit Function
    If IsNegINF(v) Then Exit Function
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
Public Function Round5(ByVal v As Double, Optional ByVal digits As Long = 0, Optional ByVal er As ERound = eRoundMath) As Double
    'rundet auf die angegebene Stellenanzahl
    'negative Stellen = Stellen vor dem Komma
    'positive Stellen = Stellen nach dem Komma
    Round5 = Round(v / 5, digits) * 5
End Function
Public Function RoundUp5(ByVal v As Double, ByVal digits As Long) As Double
    RoundUp5 = RoundUp(v / 5, digits) * 5
End Function

Public Function Double_TryParse(ByVal s As String, ByRef dblOut As Double) As Boolean
Try: On Error GoTo Catch
    s = Trim$(s)
    If Len(s) > 0 Then
        s = Replace$(s, ",", ".")
        If StrComp(s, "1.#QNAN") = 0 Then
            Call GetNaN(dblOut)
        ElseIf StrComp(s, "1.#INF") = 0 Then
            dblOut = GetINF
        ElseIf StrComp(s, "-1.#INF") = 0 Then
            dblOut = GetINF(-1)
        Else
            dblOut = Val(s)
        End If
        Double_TryParse = True
    End If
    Exit Function
Catch:
End Function
'Public Function Double_ToStr(ByVal Value As Double) As String
'    Double_ToStr = CStr(Value)
'End Function

Public Function Min(v1, v2)
    If v1 < v2 Then Min = v1 Else Min = v2
End Function
Public Function Max(v1, v2)
    If v1 > v2 Then Max = v1 Else Max = v2
End Function
Public Function IsPositive(ByVal v As Double) As Boolean
    If IsPosINF(v) Then IsPositive = True: Exit Function
    If IsNegINF(v) Then IsPositive = False: Exit Function
    
    IsPositive = Sgn(v) > 0
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

Public Function GetINF(Optional ByVal sign As Long = 1) As Double
    Dim L(1 To 2) As Long
    If Sgn(sign) > 0 Then
        L(2) = &H7FF00000
    ElseIf Sgn(sign) < 0 Then
        L(2) = &HFFF00000
    End If
    Call RtlMoveMemory(GetINF, L(1), 8)
End Function

Public Sub GetNaN(ByRef DblVal As Double)
    Dim L(1 To 2) As Long
    L(1) = 1
    L(2) = &H7FF00000
    Call RtlMoveMemory(DblVal, L(1), 8)
End Sub

Public Function IsNaN(ByRef DblVal As Double) As Boolean
    Dim b(0 To 7) As Byte
    Dim i As Long
    
    Call RtlMoveMemory(b(0), DblVal, 8)
    
    If (b(7) = &H7F) Or (b(7) = &HFF) Then
        If (b(6) >= &HF0) Then
            For i = 0 To 5
                If b(i) <> 0 Then
                    IsNaN = True
                    Exit Function
                End If
            Next
        End If
    End If
End Function

Public Function IsPosINF(ByVal DblVal As Double) As Boolean
    IsPosINF = (DblVal = posINF)
End Function

Public Function IsNegINF(ByVal DblVal As Double) As Boolean
    IsNegINF = (DblVal = negINF)
End Function

Public Function NaNToStr() As String
    On Error Resume Next
    NaNToStr = CStr(NaN)
    On Error GoTo 0
End Function

Public Function PosINFToStr() As String
    PosINFToStr = CStr(posINF)
End Function

Public Function NegINFToStr() As String
    NegINFToStr = CStr(negINF)
End Function

