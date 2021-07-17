Attribute VB_Name = "MENorm"
Option Explicit

Public Enum ENorm
    Norm_DIN18800
    Norm_EuroCode3
End Enum

Public Function ENorm_ToStr(e As ENorm) As String
    Dim s As String
    Select Case e
    Case Norm_DIN18800:  s = "DIN 18800"
    Case Norm_EuroCode3: s = "EuroCode3"
    End Select
    ENorm_ToStr = s
End Function
Public Sub ENorm_FillComboBox(aCB As ComboBox)
    With aCB
        .Clear
        .AddItem ENorm_ToStr(Norm_DIN18800)
        .AddItem ENorm_ToStr(Norm_EuroCode3)
    End With
End Sub
Public Function ENorm_Parse(s As String) As ENorm
    ENorm_Parse = IIf(Left(s, 1) = "D", Norm_DIN18800, Norm_EuroCode3)
End Function

