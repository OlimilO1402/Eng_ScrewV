Attribute VB_Name = "MCtors"
Option Explicit

Public Enum Stahlsorte
    S235 'aka ST37-2
    S275 'aka ST42
    S355 'aka ST52-3
End Enum

Public Function New_Schraube(ByVal aCalcNorm As Norm, _
                             ByVal Schraubendurchmesser As Double, _
                             Optional ByVal Schraubengüte As Double = 4.6, _
                             Optional ByVal isSenk As Boolean = False, _
                             Optional ByVal isPass As Boolean = False, _
                             Optional ByVal isSFS As Boolean = False, _
                             Optional ByVal isGlf As Boolean = False, _
                             Optional ByVal isZug As Boolean = False, _
                             Optional ByVal isVor As Boolean = False) As Schraube
    Set New_Schraube = New Schraube
    Call New_Schraube.New_(aCalcNorm, Schraubendurchmesser, Schraubengüte, isSenk, isPass, isSFS, isGlf, isZug, isVor)
End Function

Public Function New_Schraubenloch(ByVal s As Schraube, ByVal la As Lochart, _
                                  Optional ByVal isVert As Boolean = True) As Schraubenloch
    Set New_Schraubenloch = New Schraubenloch
    Call New_Schraubenloch.New_(s, la, isVert)
End Function

Public Function New_Schraubengruppe(ByVal SL As Schraubenloch, _
                                    ByVal nSchraubenX As Byte, ByVal nSchraubenZ As Byte, _
                                    ByVal RandX As Double, ByVal RandZ As Double, _
                                    ByVal LochX As Double, ByVal LochZ As Double) As Schraubengruppe
    Set New_Schraubengruppe = New Schraubengruppe
    Call New_Schraubengruppe.New_(SL, nSchraubenX, nSchraubenZ, RandX, RandZ, LochX, LochZ)
End Function

Public Function New_Blech(ByVal aNorm As Norm, _
                          ByVal SG As Schraubengruppe, _
                          ByVal sso As Stahlsorte, _
                          ByVal t As Double, _
                          ByVal l_x As Double, _
                          ByVal h_z As Double) As Blech
    Set New_Blech = New Blech
    Call New_Blech.New_(aNorm, SG, t, sso, l_x, h_z)
End Function
Public Function New_Norm(aNorm As ENorm) As Norm
    Set New_Norm = New Norm
    Call New_Norm.New_(aNorm)
End Function

