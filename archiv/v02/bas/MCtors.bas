Attribute VB_Name = "MNew"
Option Explicit
Private m_Norm As Norm

Public Function Norm(Optional aNorm As ENorm) As Norm
'Norm soll ein Singleton/Monoton sein
    If m_Norm Is Nothing Then
        Set m_Norm = New Norm: m_Norm.New_ aNorm
        Set Norm = m_Norm
    Else
        If m_Norm.Norm = aNorm Then
            Set Norm = m_Norm
        Else
            'darf nicht vorkommen
            
        End If
    End If
End Function

Public Function Schraube(ByVal aCalcNorm As Norm, _
                         ByVal Schraubendurchmesser As Double, _
                         Optional ByVal Schraubengüte As Double = 4.6, _
                         Optional ByVal isSenk As Boolean = False, _
                         Optional ByVal isPass As Boolean = False, _
                         Optional ByVal isSFS As Boolean = False, _
                         Optional ByVal isGlf As Boolean = False, _
                         Optional ByVal isZug As Boolean = False, _
                         Optional ByVal isVor As Boolean = False) As Schraube
    Set Schraube = New Schraube: Schraube.New_ aCalcNorm, Schraubendurchmesser, Schraubengüte, isSenk, isPass, isSFS, isGlf, isZug, isVor
End Function

Public Function Schraubenloch(ByVal s As Schraube, ByVal la As ELochart, _
                              Optional ByVal isVert As Boolean = False) As Schraubenloch
    Set Schraubenloch = New Schraubenloch: Schraubenloch.New_ s, la, isVert
End Function

Public Function Schraubengruppe(ByVal sl As Schraubenloch, _
                                ByVal nSchraubenX As Long, _
                                ByVal nSchraubenZ As Long, _
                                ByVal AbstandSelected As AbstandLR, _
                                ByVal eaL As EAbstand, ByVal eaR As EAbstand, _
                                Ewk As EinwirkungsKombi, _
                                bl As Blech, br As Blech) As Schraubengruppe
    Set Schraubengruppe = New Schraubengruppe: Schraubengruppe.New_ sl, nSchraubenX, nSchraubenZ, AbstandSelected, eaL, eaR, Ewk, bl, br
End Function

Public Function VectorXZ(ByVal aX As Double, ByVal aZ As Double) As VectorXZ
    Set VectorXZ = New VectorXZ: VectorXZ.New_ aX, aZ
End Function
Public Function VectorXZCopy(ByVal other As VectorXZ) As VectorXZ
    Set VectorXZCopy = New VectorXZ: VectorXZCopy.NewC other
End Function

Public Function AbstandLR(ByVal LochXZ As VectorXZ, ByVal RandXZ As VectorXZ) As AbstandLR
    Set AbstandLR = New AbstandLR: Call AbstandLR.New_(LochXZ, RandXZ)
End Function
Public Function AbstandLRCopy(ByVal other As AbstandLR) As AbstandLR
    Set AbstandLRCopy = New AbstandLR: AbstandLRCopy.NewC other
End Function

Public Function Blech(ByVal aNorm As Norm, _
                      ByVal sso As EStahlsorte, _
                      ByVal t As Double, _
                      ByVal l_x As Double, _
                      ByVal h_z As Double, _
                      ByVal bIsLinks As Boolean, _
                      ByVal bIsZange As Boolean, _
                      ByVal bIsMehrschnittig As Boolean) As Blech
    Set Blech = New Blech: Blech.New_ aNorm, sso, t, l_x, h_z, bIsLinks, bIsZange, bIsMehrschnittig
'sg,                       ByVal sg As Schraubengruppe,
End Function

Public Function CairoPicBox(aPb As PictureBox, c As cCairo)
    Set CairoPicBox = New CairoPicBox: CairoPicBox.New_ aPb, c
End Function
Public Function CairoPdfDoc(c As cCairo, ByVal epo As EPageOrientation, ByVal epf As EPageFormat, ByVal zoom As Double) As CairoPdfDoc
    Set CairoPdfDoc = New CairoPdfDoc: CairoPdfDoc.New_ c, epo, epf, zoom
End Function

Public Function EinwirkungsKombi(ByVal MEd As Double, ByVal VEd As Double, ByVal NEd As Double, _
                                 ByVal OffX As Double, ByVal OffZ As Double) As EinwirkungsKombi
    Set EinwirkungsKombi = New EinwirkungsKombi: EinwirkungsKombi.New_ MEd, VEd, NEd, OffX, OffZ
End Function

Public Function SchraubenNachweis(ByVal aCalcNorm As Norm, s As Schraube, sg As Schraubengruppe) As SchraubenNachweis
    Set SchraubenNachweis = New SchraubenNachweis: SchraubenNachweis.New_ aCalcNorm, s, sg
End Function

