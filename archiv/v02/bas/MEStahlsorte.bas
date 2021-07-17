Attribute VB_Name = "MEStahlsorte"
Option Explicit

Public Enum EStahlsorte
    S235 = 0 'aka ST37-2
    S275 = 1 'aka ST42
    S355 = 2 'aka ST52-3
    S420 = 3
    S460 = 4
End Enum

Public Function EStahlsorte_ToStr(ByVal e As EStahlsorte, ByVal n As ENorm) As String
    Dim s As String
    If n = Norm_DIN18800 Then
        Select Case e
        Case S235: s = "ST37"
        Case S275: s = "ST42"
        Case S355: s = "ST52"
        End Select
    Else
        Select Case e
        Case S235: s = "S235"
        Case S275: s = "S275"
        Case S355: s = "S355"
        Case S420: s = "S420"
        Case S460: s = "S460"
        End Select
    End If
    EStahlsorte_ToStr = s
End Function

Public Sub EStahlsorte_FillComboBox(aCB As ComboBox, ByVal n As ENorm)
    With aCB
        .Clear
        .AddItem EStahlsorte_ToStr(S235, n)
        .AddItem EStahlsorte_ToStr(S275, n)
        .AddItem EStahlsorte_ToStr(S355, n)
    End With
End Sub

Public Function EStahlsorte_Parse(s As String) As EStahlsorte
    Dim p As Integer: p = CInt(Left(Right(s, 2), 1))
    Select Case p
    Case 3:    EStahlsorte_Parse = S235
    Case 4, 7: EStahlsorte_Parse = S275
    Case 5:    EStahlsorte_Parse = S355
    End Select
End Function
