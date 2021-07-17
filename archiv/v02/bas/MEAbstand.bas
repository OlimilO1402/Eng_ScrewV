Attribute VB_Name = "MEAbstand"
Option Explicit

Public Enum EAbstand
    AbstandMinRed
    AbstandMinMit
    AbstandMinVol
    AbstandMaximal
End Enum

Public Function EAbstand_ToStr(ByVal e As EAbstand) As String
    Select Case e
    Case AbstandMinRed:  EAbstand_ToStr = "min red."
    Case AbstandMinMit:  EAbstand_ToStr = "min mitt"
    Case AbstandMinVol:  EAbstand_ToStr = "min voll"
    Case AbstandMaximal: EAbstand_ToStr = "maximal"
    End Select
End Function

Public Sub EAbstand_FillComboBox(aCB As ComboBox)
    With aCB
        .Clear
        .AddItem EAbstand_ToStr(EAbstand.AbstandMinRed)
        .AddItem EAbstand_ToStr(EAbstand.AbstandMinMit)
        .AddItem EAbstand_ToStr(EAbstand.AbstandMinVol)
        .AddItem EAbstand_ToStr(EAbstand.AbstandMaximal)
    End With
End Sub

Public Function EAbstand_Parse(ByVal s As String) As EAbstand
    Select Case s
    Case "min red.": EAbstand_Parse = EAbstand.AbstandMinRed
    Case "min mitt": EAbstand_Parse = EAbstand.AbstandMinMit
    Case "min voll": EAbstand_Parse = EAbstand.AbstandMinVol
    Case "maximal":  EAbstand_Parse = EAbstand.AbstandMaximal
    End Select
End Function

