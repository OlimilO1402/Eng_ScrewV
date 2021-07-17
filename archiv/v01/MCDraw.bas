Attribute VB_Name = "MCDraw"
Option Explicit

 
'Sub ZeichneSchraube(sch As Schraube)
'    Dim d   As Double:   d = sch.Durchmesser
'    Dim d_L As Double: d_L = sch.Durchmesser
'    Dim s   As Double:   s = sch.Schlüsselweite
'    Dim e   As Double:   e = sch.Eckenmass
'    Dim SbD As Double: SbD = sch.Scheibendurchmesser
'    Dim SL  As Double:  SL = sch.MinSchraubenlänge
'
'    Dim r_in   As Double:   r_in = CInch(d / 2)
'    Dim r_L_in As Double: r_L_in = CInch(d_L / 2)
'    Dim r_s_in As Double: r_s_in = CInch(s / 2)
'    Dim r_e_in As Double: r_e_in = CInch(e / 2)
'    Dim r_sbIn As Double: r_sbIn = CInch(SbD / 2)
'    'Dim r_x_in As Double: r_x_in = r_x 'Inch(r_x)
'
'    'Schraube in Ansicht
'    'Kreis für Gewinde
'    Dim s1 As Shape: Set s1 = ActiveLayer.CreateEllipse2(0#, 0#, r_in, r_in)
'    s1.Fill.ApplyNoFill
'    s1.Outline.SetProperties 0.007874, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
'
'    'Kreis für Loch
'
'    Dim s2 As Shape: Set s2 = ActiveLayer.CreateEllipse2(0#, 0#, r_L_in, r_L_in)
'    s2.Fill.ApplyNoFill
'    s2.Outline.SetProperties 0.007874, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
'
'    'Kreis für Schraubenkopf
'    Dim s3 As Shape: Set s3 = ActiveLayer.CreateEllipse2(0#, 0#, r_s_in, r_s_in)
'    s3.Fill.ApplyNoFill
'    s3.Outline.SetProperties 0.007874, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
'
'    '6-Eck Polygon für Schraubenkopf
'    Dim s4 As Shape: Set s4 = ActiveLayer.CreatePolygon2(0#, 0#, r_e_in, 6)
'    s4.Fill.ApplyNoFill
'    s4.Outline.SetProperties 0.007874, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
'
'    'Kreis für Scheibe
'    Dim s5 As Shape: Set s5 = ActiveLayer.CreateEllipse2(0#, 0#, r_sbIn, r_sbIn)
'    s5.Fill.ApplyNoFill
'    s5.Outline.SetProperties 0.007874, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
'
'    'ActiveDocument.ReferencePoint = cdrCenter
'
'    'Alle Schraubenshapes auswählen
'
'    s1.Selected = True
'    s3.Selected = True
'    s4.Selected = True
'
'    'die Schraubenshapes gruppieren
'    Dim sGrp1 As Shape: Set sGrp1 = ActiveSelection.Group
'    sGrp1.Selected = False
'
'    'das Schraubenloch dazugruppieren
'    s2.Selected = True
'    s5.Selected = True
'    Dim sGrp2 As Shape: Set sGrp2 = ActiveSelection.Group
'
'    sGrp1.Selected = True
'    sGrp2.Selected = True
'
'    Dim sGrp3 As Shape: Set sGrp3 = ActiveSelection.Group
'
'    Dim s6 As Shape: Set s6 = ActiveLayer.CreateRectangle(-r_in, CInch(40), r_in, -CInch(40))
'    s6.Rectangle.CornerType = cdrCornerTypeRound
'    s6.Rectangle.RelativeCornerScaling = True
'    s6.Fill.ApplyNoFill
'    s6.Outline.SetProperties 0.007874, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
'
'    Dim s7 As Shape: Set s7 = ActiveLayer.CreateRectangle(-0.555665, -3.240358, 0.551972, -3.542441)
'    s7.Rectangle.CornerType = cdrCornerTypeRound
'    s7.Rectangle.RelativeCornerScaling = True
'    s7.Fill.ApplyNoFill
'    s7.Outline.SetProperties 0.007874, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
'
'    Dim s8 As Shape: Set s8 = ActiveLayer.CreateRectangle(-0.4515, -1.698693, 0.440862, -2.118831)
'    s8.Rectangle.CornerType = cdrCornerTypeRound
'    s8.Rectangle.RelativeCornerScaling = True
'    s8.Fill.ApplyNoFill
'    s8.Outline.SetProperties 0.007874, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
'
'End Sub
'
'
