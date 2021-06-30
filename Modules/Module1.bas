Attribute VB_Name = "MBsps"
Option Explicit

Public Function Bsp1(ByRef isupd_inout As Boolean) As SchrVerbFL 'n As Norm, s As Schraube, sl As Schraubenloch, ekd As EinwirkungsKombi, bl As Blech, br As Blech, sg As Schraubengruppe, nw As SchraubenNachweis)
    'Beispiel aus dem Buch "Verbindungen im Stahl und Verbundbau" von Rolf Kindmann, Michael Stracke
    'Kap. 3.4.4 Zugstoﬂ eines Stabes aus Flachst‰hlen
    'siehe Bild 3.44 Geschraubter Zugstoﬂ von Flachst‰hlen
    isupd_inout = True
    
    Dim n As Norm:                Set n = MNew.Norm(MNew.NormFromENorm(ENorm.Norm_EuroCode3))
    Dim s As Schraube:            Set s = MNew.Schraube(n, 20#, 5.6, , , True)
    Dim sl As Schraubenloch:      Set sl = MNew.Schraubenloch(s, Normal)
    Dim ekd As EinwirkungsKombi:  Set ekd = MNew.EinwirkungsKombi(0, 0, 500, 0, 0)
    Dim bl  As Blech:             Set bl = MNew.Blech(n, S235, 10, 350, 170, True, True, True)
    Dim br  As Blech:             Set br = MNew.Blech(n, S235, 18, 1000, 170, False, False, True)
    Dim sg  As Schraubengruppe:   Set sg = MNew.Schraubengruppe(sl, 2, 2, MNew.AbstandLR(MNew.VectorXZ(65, 70), _
                                                                                         MNew.VectorXZ(55, 50)), _
                                                                          EAbstand.AbstandMinVol, _
                                                                          EAbstand.AbstandMinVol, ekd, bl, br)
    Dim nw  As SchraubenNachweis: Set nw = MNew.SchraubenNachweis(n, s, sg)
    
    Set Bsp1 = MNew.SchrVerbFL(n, s, sl, sg, bl, br, nw, ekd)
    
    frmSchrauben.CkDrawHole.Value = vbChecked
    frmSchrauben.CkDrawMutter.Value = vbChecked
    frmSchrauben.CkDrawUScheibe.Value = vbChecked
    
    frmSchrauben.CkBBoltGroup.Value = vbChecked
    frmSchrauben.CkBBeamLeft.Value = vbChecked
    frmSchrauben.CkBBeamRight.Value = vbChecked
    frmSchrauben.CkRoundUp5.Value = vbChecked
    'frmSchrauben.CkUpdateAbstand.Value = vbChecked
    frmSchrauben.CmBAbstRand.ListIndex = 2
    frmSchrauben.CmBAbstLoch.ListIndex = 2
    'isUpdating = True
    'frmSchrauben.CkUpdateAbstand.Value = vbUnchecked
    
    'frmSchrauben.CkZange.Value = vbChecked
    frmSchrauben.CkPdfQuer.Value = vbChecked
    'bis hierher, ab jetzt den View Updaten
    isupd_inout = False
    frmSchrauben.CbZoom.ListIndex = 16
    'frmSchrauben.CkZange.Value = vbChecked
    'frmSchrauben.UpdateView
End Function


Public Function Bsp2(ByRef isupd_inout As Boolean) As SchrVerbFL 'n As Norm, s As Schraube, sl As Schraubenloch, ekd As EinwirkungsKombi, ) 'bl As Blech, br As Blech, sg As Schraubengruppe, nw As SchraubenNachweis)
    'Beispiel aus dem Buch "Verbindungen im Stahl und Verbundbau" von Rolf Kindmann, Michael Stracke
    'Kap. 3.5.3 Stoﬂ mit Steglaschen
    'Gelenkiger Tr‰gerstoﬂ eines HEB 400 mit Steglaschen
    isupd_inout = True
    
    Dim n As Norm:               Set n = MNew.Norm(MNew.NormFromENorm(ENorm.Norm_EuroCode3))
    Dim s As Schraube:           Set s = MNew.Schraube(n, 24#, 5.6, , , True)
    Dim sl As Schraubenloch:     Set sl = MNew.Schraubenloch(s, Normal)
    Dim ekd As EinwirkungsKombi: Set ekd = MNew.EinwirkungsKombi(0, -432, 0, 47.5, 0)
    Dim bl As Blech:             Set bl = MNew.Blech(n, S355, 10, 185, 290, True, True, True)
    Dim br As Blech:             Set br = MNew.Blech(n, S355, 13.5, 1000, 400, False, False, True)
    Dim sg As Schraubengruppe:   Set sg = MNew.Schraubengruppe(sl, 1, 3, MNew.AbstandLR(MNew.VectorXZ(0, 100), _
                                                                                        MNew.VectorXZ(45, 45)), _
                                                                         EAbstand.AbstandMinVol, _
                                                                         EAbstand.AbstandMinVol, ekd, bl, br)
    Dim nw As SchraubenNachweis: Set nw = MNew.SchraubenNachweis(n, s, sg)
    
    Set Bsp2 = MNew.SchrVerbFL(n, s, sl, sg, bl, br, nw, ekd)
    
    frmSchrauben.CkDrawHole.Value = vbChecked
    frmSchrauben.CkDrawMutter.Value = vbChecked
    frmSchrauben.CkDrawUScheibe.Value = vbChecked
    
    frmSchrauben.CkBBoltGroup.Value = vbChecked
    frmSchrauben.CkBBeamLeft.Value = vbChecked
    frmSchrauben.CkBBeamRight.Value = vbChecked
    frmSchrauben.CkRoundUp5.Value = vbChecked
    'frmSchrauben.CkUpdateAbstand.Value = vbChecked
    frmSchrauben.CmBAbstRand.ListIndex = 2
    frmSchrauben.CmBAbstLoch.ListIndex = 2
    'isUpdating = True
    'frmSchrauben.CkUpdateAbstand.Value = vbUnchecked
    
    'frmSchrauben.CkZange.Value = vbChecked
    frmSchrauben.CkPdfQuer.Value = vbChecked
    'bis hierher, ab jetzt den View Updaten
    isupd_inout = False
    frmSchrauben.CbZoom.ListIndex = 16
    'frmSchrauben.CkZange.Value = vbChecked
    'frmSchrauben.UpdateView
End Function


Public Function Bsp3(ByRef isupd_inout As Boolean) As SchrVerbFL ') ' n As Norm, s As Schraube, sl As Schraubenloch, ekd As EinwirkungsKombi, bl As Blech, br As Blech, sg As Schraubengruppe, nw As SchraubenNachweis)
    isupd_inout = True
    Dim n   As Norm:              Set n = MNew.Norm(MNew.NormFromENorm(ENorm.Norm_EuroCode3))
    Dim s   As Schraube:          Set s = MNew.Schraube(n, 24#, 10.9, , , True)
    Dim sl  As Schraubenloch:     Set sl = MNew.Schraubenloch(s, Normal)
    Dim ekd As EinwirkungsKombi:  Set ekd = MNew.EinwirkungsKombi(33.71, 90#, 36.3, 85, 0)
    Dim bl  As Blech:             Set bl = MNew.Blech(n, S235, 10, 1000, 380, True, True, True)
    Dim br  As Blech:             Set br = MNew.Blech(n, S235, 12, 1000, 380, False, False, True)
    Dim sg  As Schraubengruppe:   Set sg = MNew.Schraubengruppe(sl, 1, 3, MNew.AbstandLR(MNew.VectorXZ(0, 115), _
                                                                                         MNew.VectorXZ(80, 75)), _
                                                                          EAbstand.AbstandMinVol, _
                                                                          EAbstand.AbstandMinVol, ekd, bl, br)
    Dim nw  As SchraubenNachweis: Set nw = MNew.SchraubenNachweis(n, s, sg)

    Set Bsp3 = MNew.SchrVerbFL(n, s, sl, sg, bl, br, nw, ekd)
   
    frmSchrauben.CkDrawHole.Value = vbChecked
    frmSchrauben.CkDrawMutter.Value = vbChecked
    frmSchrauben.CkDrawUScheibe.Value = vbChecked
    
    frmSchrauben.CkBBoltGroup.Value = vbChecked
    frmSchrauben.CkBBeamLeft.Value = vbChecked
    frmSchrauben.CkBBeamRight.Value = vbChecked
    frmSchrauben.CkRoundUp5.Value = vbChecked
    'frmSchrauben.CkUpdateAbstand.Value = vbChecked
    frmSchrauben.CmBAbstRand.ListIndex = 2
    frmSchrauben.CmBAbstLoch.ListIndex = 2
    'isUpdating = True
    'frmSchrauben.CkUpdateAbstand.Value = vbUnchecked
    
    'frmSchrauben.CkZange.Value = vbChecked
    frmSchrauben.CkPdfQuer.Value = vbChecked
    'bis hierher, ab jetzt den View Updaten
    isupd_inout = False
    frmSchrauben.CbZoom.ListIndex = 16
    'frmSchrauben.CkZange.Value = vbChecked
    'frmSchrauben.UpdateView
End Function

Public Function Bsp4(ByRef isupd_inout As Boolean) As SchrVerbFL ', ByRef aSVFL_out As SchrVerbFL) '  n As Norm, s As Schraube, sl As Schraubenloch, ekd As EinwirkungsKombi, bl As Blech, br As Blech, sg As Schraubengruppe, nw As SchraubenNachweis)
    isupd_inout = True
    
    Dim n   As Norm:              Set n = MNew.Norm(MNew.NormFromENorm(ENorm.Norm_EuroCode3))
    Dim s   As Schraube:          Set s = MNew.Schraube(n, 12, 5.6, , , True)
    Dim sl  As Schraubenloch:     Set sl = MNew.Schraubenloch(s, Normal)
    Dim ekd As EinwirkungsKombi:  Set ekd = MNew.EinwirkungsKombi(0.8, 1.6, 0, 0, 0)
    Dim bl  As Blech:             Set bl = MNew.Blech(n, S235, 10, 1000, 100, True, False, False) 'sg,
    Dim br  As Blech:             Set br = MNew.Blech(n, S235, 10, 1000, 100, False, False, False) 'sg,
    Dim sg  As Schraubengruppe:   Set sg = MNew.Schraubengruppe(sl, 2, 1, MNew.AbstandLR(MNew.VectorXZ(60, 0), _
                                                                                         MNew.VectorXZ(30, 50)), _
                                                                          EAbstand.AbstandMinVol, _
                                                                          EAbstand.AbstandMinVol, ekd, bl, br)
    Dim nw As SchraubenNachweis
    Set nw = MNew.SchraubenNachweis(n, s, sg)
    
    Set Bsp4 = MNew.SchrVerbFL(n, s, sl, sg, bl, br, nw, ekd)
    
    frmSchrauben.CkDrawHole.Value = vbChecked
    frmSchrauben.CkDrawMutter.Value = vbChecked
    frmSchrauben.CkDrawUScheibe.Value = vbChecked
    
    frmSchrauben.CkBBoltGroup.Value = vbChecked
    frmSchrauben.CkBBeamLeft.Value = vbChecked
    frmSchrauben.CkBBeamRight.Value = vbChecked
    frmSchrauben.CkZange.Value = vbChecked
    frmSchrauben.CkRoundUp5.Value = vbChecked
    'frmSchrauben.CkUpdateAbstand.Value = vbChecked
    frmSchrauben.CmBAbstRand.ListIndex = 2
    frmSchrauben.CmBAbstLoch.ListIndex = 2
    'isUpdating = True
    'frmSchrauben.CkUpdateAbstand.Value = vbUnchecked
    
    'frmSchrauben.CkZange.Value = vbChecked
    frmSchrauben.CkPdfQuer.Value = vbChecked
    isupd_inout = False
    frmSchrauben.CbZoom.ListIndex = 21
    'frmSchrauben.CbZoom.ListIndex = 16
    'frmSchrauben.UpdateView
End Function

