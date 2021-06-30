Attribute VB_Name = "MESFK"
Option Explicit

Public Enum ESFK 'Schraubenfestigkeitsklasse
    SFK_36 = 1   ' 36 ' 3.6
    SFK_46 = 2   ' 46 ' 4.6
    SFK_48 = 3   ' 48 ' 4.8
    SFK_56 = 4   ' 56 ' 5.6
    SFK_58 = 5   ' 58 ' 5.8
    SFK_68 = 6   ' 68 ' 6.8
    SFK_88 = 7   ' 88 ' 8.8
    SFK_98 = 8   ' 98 ' 9.8
    SFK_109 = 9  '109 '10.9
    SFK_129 = 10 '129 '12.9
End Enum

Public Enum EMFK 'Mutternfestigkeitsklasse
    MFK_4
    MFK_5
    MFK_8
    MFK_10
End Enum
Private Function SG_to_ESFK(ByVal sg As Single) As ESFK
    Dim s1 As Integer: s1 = Int(sg)        ' vor dem Komma
    Dim s2 As Integer: s2 = (sg - s1) * 10 ' nach dem Komma
    Dim e As ESFK
    Select Case s1
    Case 3: e = SFK_36
    Case 4: e = IIf(s2 < 8, ESFK.SFK_46, ESFK.SFK_48)
    Case 5: e = IIf(s2 < 8, ESFK.SFK_56, ESFK.SFK_58)
    Case 6: e = SFK_68
    Case 8: e = SFK_88
    Case 9: e = SFK_98
    Case 10: e = SFK_109
    Case 12: e = SFK_129
    End Select
    SG_to_ESFK = e
End Function
'SG=Schraubengüte und ESFK=Schraubenfestigkeitsklasse müssen das gleiche Resultat liefern
Public Function EFSK_ToStr(e As ESFK) As String
    Dim s As String
    Select Case e
    Case SFK_36:  s = "3.6"
    Case SFK_46:  s = "4.6"
    Case SFK_48:  s = "4.8"
    Case SFK_56:  s = "5.6"
    Case SFK_58:  s = "5.8"
    Case SFK_68:  s = "6.8"
    Case SFK_88:  s = "8.8"
    Case SFK_98:  s = "9.8"
    Case SFK_109: s = "10.9"
    Case SFK_129: s = "12.9"
    End Select
    EFSK_ToStr = s
End Function
Public Sub ESFK_FillComboBox(aCB As ComboBox, Optional OnlyInD As Boolean = False, Optional ByVal Only_HV As Boolean = False)
    aCB.Clear
    If Not Only_HV Then
        If Not OnlyInD Then aCB.AddItem "3.6"  'in D nicht unterstützt
                            aCB.AddItem "4.6"  'in D unterstützt, nicht für vorgespannt
        If Not OnlyInD Then aCB.AddItem "4.8"  'in D nicht unterstützt
                            aCB.AddItem "5.6"  'in D unterstützt, nicht für vorgespannt
        If Not OnlyInD Then aCB.AddItem "5.8"  'in D nicht unterstützt
        If Not OnlyInD Then aCB.AddItem "6.8"  'in D nicht unterstützt
    End If
                            aCB.AddItem "8.8"  'in D unterstützt, für vorgespannt
        If Not OnlyInD Then aCB.AddItem "9.8"  'in D nicht unterstützt
                            aCB.AddItem "10.9" 'in D unterstützt, für vorgespannt
        If Not OnlyInD Then aCB.AddItem "12.9" 'in D nicht unterstützt
End Sub
Public Function ESFK_Parse(ByVal s As String) As ESFK
    If Len(Trim(s)) > 4 Then Exit Function
    Select Case Int(Left(s, 1))
    Case 3: ESFK_Parse = SFK_36
    Case 4: ESFK_Parse = IIf(Int(Right(s, 1)) < 8, ESFK.SFK_46, ESFK.SFK_48)
    Case 5: ESFK_Parse = IIf(Int(Right(s, 1)) < 8, ESFK.SFK_56, ESFK.SFK_58)
    Case 6: ESFK_Parse = SFK_68
    Case 8: ESFK_Parse = SFK_88
    Case 9: ESFK_Parse = SFK_98
    Case 1: ESFK_Parse = IIf(Int(Mid(s, 2, 1)) = 0, ESFK.SFK_109, ESFK.SFK_129)
    End Select
End Function
