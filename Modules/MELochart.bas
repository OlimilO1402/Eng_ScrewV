Attribute VB_Name = "MELochart"
Option Explicit

Public Enum ELochart
    Normal
    ‹bergroﬂ
    LanglochKurz
    LanglochLang
End Enum

Public Function ELochart_ToStr(ByVal e As ELochart) As String
    Dim s As String
    Select Case e
    Case Normal:       s = "Normal"
    Case ‹bergroﬂ:     s = "‹bergroﬂ"
    Case LanglochKurz: s = "Lang-Kurz"
    Case LanglochLang: s = "Lang-Lang"
    End Select
    ELochart_ToStr = s
End Function
Public Sub ELochart_FillComboBox(aCB As ComboBox)
    With aCB
        .Clear
        .AddItem ELochart_ToStr(ELochart.Normal)
        .AddItem ELochart_ToStr(ELochart.‹bergroﬂ)
        .AddItem ELochart_ToStr(ELochart.LanglochKurz)
        .AddItem ELochart_ToStr(ELochart.LanglochLang)
    End With
End Sub
Public Function ELochart_Parse(s As String) As ELochart
    Select Case Left(s, 1)
    Case "N":  ELochart_Parse = ELochart.Normal
    Case "‹":  ELochart_Parse = ELochart.‹bergroﬂ
    Case Else: ELochart_Parse = IIf(Right(s, 1) = "z", ELochart.LanglochKurz, ELochart.LanglochLang)
    End Select
End Function

