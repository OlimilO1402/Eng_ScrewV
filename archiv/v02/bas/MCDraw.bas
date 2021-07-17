Attribute VB_Name = "MCDraw"
Option Explicit

'Public Function CRad(ByVal angle_Degree)
'    'wandelt Grad in Radians um
'    CRad = angle_Degree * VBA.Math.Atn(1) / 45
'End Function
'Public Function CInch(ByVal aMM)
'    'wandelt mm in Inch um
'    CInch = aMM / 25.4
'End Function
Public Sub DrawSystem(c As cCairoContext, ByVal X_m As Double, ByVal Y_m As Double, s As Schraube, sl As Schraubenloch, sg As Schraubengruppe, bl As Blech, br As Blech)
    'Dim c As cCairoContext: Set c = aCPB.Canvas
    'c.AntiAlias = CAIRO_ANTIALIAS_NONE
    'alles löschen
    'c.Save
    c.SetSourceColor vbWhite
    c.Paint
    c.SetSourceColor vbBlack
    c.SetLineWidth 0.2, RoundToFullPixels:=True
    c.SetLineCap CAIRO_LINE_CAP_SQUARE
    If Not s Is Nothing Then
        If sg Is Nothing Then
            DrawSchrGarnitur c, s, sl, X_m, Y_m
        Else
            DrawSchraubengruppe c, sg, X_m, Y_m
            If Not bl Is Nothing Then
                DrawBlechL c, sg, bl, X_m, Y_m
            End If
            If Not br Is Nothing Then
                DrawBlechR c, sg, br, X_m, Y_m, 2 * X_m
            End If
        End If
    End If
End Sub
Public Sub DrawSchraubengruppe(c As cCairoContext, sg As Schraubengruppe, ByVal X_m As Double, ByVal Y_m As Double)
'Try: On Error GoTo Catch
    Dim X As Double:   X = X_m - sg.BRectSelInnW / 2
    Dim Y As Double:   Y = Y_m - sg.BRectSelInnH / 2
    Dim x0 As Double: x0 = X
    Dim y0 As Double: y0 = Y
    Dim dx As Double: dx = sg.AbstandSel.Loch.X
    Dim dz As Double: dz = sg.AbstandSel.Loch.Z
    Dim s As Schraube, sl As Schraubenloch, V As VectorXZ
    Dim i As Long, j As Long
    For j = 0 To sg.AnzahlZ - 1
        For i = 0 To sg.AnzahlX - 1
            Set V = sg.V(i, j)
            Set s = sg.Schraube(i, j)
            Set sl = sg.Schraubenloch(i, j)
            DrawSchrGarnitur c, s, sl, X, Y
            DrawKräfte c, V, X, Y
            X = X + dx
        Next
        Y = Y + dz
        X = x0
    Next
    DrawAbstände c, sg, X_m, Y_m
'    Exit Sub
'Catch:
'    If Err.Number = 6 Then Resume Next 'Überlauf wg unendlich
End Sub
Public Sub DrawKräfte(c As cCairoContext, V As VectorXZ, ByVal X_m As Double, ByVal Y_m As Double)
    'X=>X; Y=>Z
    c.Save
    c.SetSourceColor vbRed
    c.SetLineWidth 0.5, True
    c.DrawLine X_m, Y_m, X_m - V.X, Y_m - V.Z
    c.Stroke
    'c.DrawLine X_m, Y_m, V.X, Y_m
    c.Restore
End Sub
Public Sub DrawAbstände(c As cCairoContext, sg As Schraubengruppe, ByVal X_m As Double, ByVal Y_m As Double)
    Dim X As Double, dx As Double
    Dim Z As Double, dz As Double
    c.Save
    
    'Rechteck der rechnerischen Abstände zeichnen
    c.SetSourceColor vbBlue
    c.SetDashes 0, 2, 2, 2, 2 ', 5, 5, 5
    'das innere Rechteck
    X = X_m - sg.BRectOptInnW / 2 '* PixProMM * ZoomFact
    Z = Y_m - sg.BRectOptInnH / 2 '* PixProMM * ZoomFact
    dx = sg.BRectOptInnW '* PixProMM * ZoomFact
    dz = sg.BRectOptInnH '* PixProMM * ZoomFact
    Call c.Rectangle(X, Z, dx, dz)
    Call c.Stroke
    'das äußere Rechteck
    X = X_m - sg.BRectOptOutW / 2 '* PixProMM * ZoomFact
    Z = Y_m - sg.BRectOptOutH / 2 '* PixProMM * ZoomFact
    dx = sg.BRectOptOutW '* PixProMM * ZoomFact
    dz = sg.BRectOptOutH '* PixProMM * ZoomFact
    Call c.Rectangle(X, Z, dx, dz)
    Call c.Stroke
    
    'Rechteck der gewählten Abstände zeichnen
    c.SetSourceColor &H9900 'vbGreen'vbRed
    'gestrichelt
    c.SetDashes 0, 0.2, 2, 3, 2
    'das innere Rechteck
    X = X_m - sg.BRectSelInnW / 2
    Z = Y_m - sg.BRectSelInnH / 2
    dx = sg.BRectSelInnW
    dz = sg.BRectSelInnH
    Call c.Rectangle(X, Z, dx, dz)
    Call c.Stroke
    'das äußere Rechteck
    X = X_m - sg.BRectSelOutW / 2 '* PixProMM * ZoomFact
    Z = Y_m - sg.BRectSelOutH / 2 '* PixProMM * ZoomFact
    dx = sg.BRectSelOutW '* PixProMM * ZoomFact
    dz = sg.BRectSelOutH '* PixProMM * ZoomFact
    Call c.Rectangle(X, Z, dx, dz)
    Call c.Stroke
    
    c.Restore
End Sub
Public Sub DrawSchrGarnitur(c As cCairoContext, s As Schraube, sl As Schraubenloch, ByVal X_m As Double, ByVal Y_m As Double)
    DrawSchraube c, s, X_m, Y_m
    If frmSchrauben.CkDrawUScheibe.Value = vbChecked Then
        DrawUScheibe c, s, X_m, Y_m
    End If
    If frmSchrauben.CkDrawHole.Value = vbChecked Then
        If Not sl Is Nothing Then
            DrawSchrLoch c, sl, X_m, Y_m
        End If
    End If
    If frmSchrauben.CkDrawMutter.Value = vbChecked Then
        DrawMutter c, s, X_m, Y_m
    End If
End Sub
Public Sub DrawSchraube(c As cCairoContext, s As Schraube, ByVal X_m As Double, ByVal Y_m As Double)
    If s Is Nothing Then Exit Sub
    Dim m  As Double:  m = s.Durchmesser '* PixProMM * ZoomFact
    Dim e  As Double:  e = s.Eckenmass '* PixProMM * ZoomFact
    Dim sw As Double: sw = IIf(s.IsSenkschraube, s.Kopfdurchmesser, s.Schlüsselweite) '* PixProMM * ZoomFact
    Dim sd As Double: sd = s.Schaftdurchmesser '* PixProMM * ZoomFact
    'Schraubengewinde
    'c.Save
    Call c.Ellipse(X_m, Y_m, m, m)
    Call c.Stroke
    'Schraubenschaft bei Passschraube
    If s.IsPassschraube Then
        Call c.Ellipse(X_m, Y_m, sd, sd)
        Call c.Stroke
    End If
    'Schraubenkopf
    Call c.DrawRegularPolygon(X_m, Y_m, e / 2, 6)
    Call c.Stroke
    Call c.Ellipse(X_m, Y_m, sw, sw)
    Call c.Stroke
    'ein kleines Kreuz in der Mitte zeichnen
    Dim k As Double: k = s.Durchmesser / 6 ' 2
    c.DrawLine X_m, Y_m, X_m, Y_m - k
    c.DrawLine X_m, Y_m, X_m + k, Y_m
    c.DrawLine X_m, Y_m, X_m, Y_m + k
    c.DrawLine X_m, Y_m, X_m - k, Y_m
    Call c.Stroke
  'c.Restore
End Sub
Public Sub DrawUScheibe(c As cCairoContext, s As Schraube, ByVal X_m As Double, ByVal Y_m As Double)
    If s Is Nothing Then Exit Sub
    Dim u As Double: u = s.Scheibendurchmesser '* PixProMM * ZoomFact
    c.SetSourceColor vbBlack
    Call c.Ellipse(X_m, Y_m, u, u)
    Call c.Stroke
End Sub
Public Sub DrawBlechL(c As cCairoContext, sg As Schraubengruppe, b As Blech, ByVal X_m As Double, ByVal Y_m As Double)
    If b Is Nothing Then Exit Sub
    Dim h As Double: h = b.Höhe '* PixProMM * ZoomFact
    Dim w As Double: w = sg.BRectSelOutW '* PixProMM * ZoomFact ' die Koordinate der rechtesten Schraube + AbstandRandX
    c.SetSourceColor vbBlack
    Call c.Rectangle(0, Y_m - h / 2, X_m + w / 2, h)
    Call c.Stroke
End Sub
Public Sub DrawBlechR(c As cCairoContext, sg As Schraubengruppe, b As Blech, ByVal X_m As Double, ByVal Y_m As Double, ww As Double)
    If b Is Nothing Then Exit Sub
    Dim h As Double: h = b.Höhe '* PixProMM * ZoomFact
    Dim w As Double: w = sg.BRectSelOutW '* PixProMM * ZoomFact ' die Koordinate der rechtesten Schraube + AbstandRandX
    c.SetSourceColor vbBlack
    Call c.Rectangle(X_m - w / 2, Y_m - h / 2, ww, h)
    Call c.Stroke
End Sub

Public Sub DrawSchrLoch(c As cCairoContext, sl As Schraubenloch, ByVal X_m As Double, ByVal Y_m As Double)
    If sl Is Nothing Then Exit Sub
    Dim d As Double: d = sl.Durchmesser '* PixProMM * ZoomFact
    c.SetSourceColor vbBlack
    Select Case sl.Lochart
    Case Normal, Übergroß
        Call c.Ellipse(X_m, Y_m, d, d)
        Call c.Stroke
    Case LanglochKurz, LanglochLang
        'Rechteck mit gerundeten Ecken zeichnen
        Dim w As Double: w = sl.Durchmesser '* PixProMM * ZoomFact
        Dim h As Double: h = sl.DurchmesserSenkr '* PixProMM * ZoomFact
        If sl.IsVertikal Then
            'fliparound 'swaparound
            Dim t As Double: t = w: w = h: h = t
        End If
        Dim X As Double: X = X_m - w / 2
        Dim r As Double: r = Min(w, h)
        Dim Y As Double: Y = Y_m - h / 2
        Call c.RoundedRect(X, Y, w, h, r)
        Call c.Stroke
    End Select
  'c.Restore
End Sub
Public Sub DrawMutter(c As cCairoContext, s As Schraube, ByVal X_m As Double, ByVal Y_m As Double)
    If s Is Nothing Then Exit Sub
    Dim e  As Double:  e = s.MutterEckenmass '* PixProMM * ZoomFact
    Dim sw As Double: sw = s.MutterSchlüsselweite '* PixProMM * ZoomFact
    c.SetSourceColor vbBlack
    Call c.DrawRegularPolygon(X_m, Y_m, e / 2, 6)
    Call c.Stroke
    Call c.Ellipse(X_m, Y_m, sw, sw)
    Call c.Stroke
  'c.Restore
End Sub

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
