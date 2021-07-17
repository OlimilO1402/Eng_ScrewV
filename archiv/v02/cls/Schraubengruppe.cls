VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Schraubengruppe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Die Klasse Schraubengruppe definiert eine 2D-Gruppe von Anzahl-X (in x-Richtung) mal Anzahl-Z (in z-Richtung) Schrauben
'man könnte auch andere Schraubenanordnungen zulassen, z.B. Dreiecks/Waben-Raster, oder quadrat. um 45° gedreht
'von gleichen Schrauben mit einem bestimmten Abstand zum Rand des Blechs und der Schrauben untereinander
'Private M_S As Schraube 'in Schraubenloch!
Private m_NX   As Long 'Anzahl der Schrauben in X-Richtung
Private m_NZ   As Long 'Anzahl der Schrauben in Z-Richtung
Private m_eaL  As EAbstand 'Lochabstand 'Siehe Modul MEAbstand
Private m_eaR  As EAbstand 'Randabstand 'Siehe Modul MEAbstand
Private WithEvents mASel As AbstandLR 'die gewählten Abstände
Attribute mASel.VB_VarHelpID = -1
Private WithEvents mAOpt As AbstandLR 'die Min-Max-Abstände
Attribute mAOpt.VB_VarHelpID = -1
Private Type TScrew
    Pos  As VectorXZ 'die Position
    v    As VectorXZ 'die Kräfte in x- und z-Ri
    sl As Schraubenloch 'das Schrabuenloch mit Schraube
End Type
Private m_Screws() As TScrew ' Schraubenloch
Private m_Ip   As Double
Private m_Ekd  As EinwirkungsKombi
Private WithEvents mBl As Blech 'Blech von Links
Attribute mBl.VB_VarHelpID = -1
Private WithEvents mBr As Blech 'Blech von Rechts
Attribute mBr.VB_VarHelpID = -1
Private Enum ETrMainLR
    None = 0
    Links = 1
    Rechts = 2
    LinksURechts = 3
End Enum
Private m_gllkMainLA As ETrMainLR 'das Maßgebende Blech in der Berechnung Grenzlochleibung Lochabstand
Private m_gllkMainRA As ETrMainLR 'das Maßgebende Blech in der Berechnung Grenzlochleibung Randabstand
Private m_bvnzMain   As ETrMainLR 'das Maßgebende Blech in der Berechnung BlockversagenNzentr
Private m_bvveMain   As ETrMainLR 'das Maßgebende Blech in der Berechnung BlockversagenVexztr
Private m_gdstMain   As ETrMainLR 'das Maßgebende Blech in der Berechnung Grenzdurchstanzkraft
Private m_N As Norm
Private m_alpha As Double 'alpha-Wert für die Grenzlochleibungskraft, wird hier nur zwischengespeichert für die Ausgabe
Private m_k1    As Double '  k_1-Wert für die Grenzlochleibungskraft, wird hier nur zwischengespeichert für die Ausgabe

Public Sub New_(sl As Schraubenloch, ByVal nSchraubenX As Long, ByVal nSchraubenZ As Long, _
                AbstSel As AbstandLR, ByVal eaL As EAbstand, ByVal eaR As EAbstand, Ewk As EinwirkungsKombi, _
                bl As Blech, br As Blech)
       m_NX = nSchraubenX:    m_NZ = nSchraubenZ
    Set mASel = AbstSel
    Set m_Ekd = Ewk
    m_eaL = eaL: m_eaR = eaR
    
    Set mAOpt = New AbstandLR
    Resize sl
    SetAbstandRandOpt m_eaR
    SetAbstandLochOpt m_eaL
    Set m_N = sl.Schraube.Norm
    Set mBl = bl
    Set mBr = br
    CalcSchraubenlänge
End Sub
Private Function ETrMainLR_ToStr(e As ETrMainLR) As String
    Select Case e
    Case ETrMainLR.Links:        ETrMainLR_ToStr = "Links"
    Case ETrMainLR.Rechts:       ETrMainLR_ToStr = "Rechts"
    Case ETrMainLR.LinksURechts: ETrMainLR_ToStr = "Links/Rechts"
    End Select
End Function
Public Property Get Norm() As Norm
    Set Norm = Me.Schraube(0, 0).Norm
End Property
Public Property Get TrägerLinks() As Blech
    Set TrägerLinks = mBl
End Property
Public Property Set TrägerLinks(ByVal Value As Blech)
    Set mBl = Value
    CalcSchraubenlänge
End Property
Public Property Get TrägerRechts() As Blech
    Set TrägerRechts = mBr
End Property
Public Property Set TrägerRechts(ByVal Value As Blech)
    Set mBr = Value
    CalcSchraubenlänge
End Property

Private Sub mAOpt_Changed(sender As AbstandLR)
    Resize
End Sub
Private Sub mASel_Changed(sender As AbstandLR)
    Resize
End Sub
Private Sub mBl_Changed(sender As Blech)
    CalcSchraubenlänge 'die Schraubenlänge neu berechnen
End Sub
Private Sub mBr_Changed(sender As Blech)
    CalcSchraubenlänge 'die Schraubenlänge/Klemmlänge neu berechnen
End Sub
Private Sub CalcSchraubenlänge()
    Dim kl As Double
    If Not mBl Is Nothing Then kl = kl + mBl.GesamtT '+ IIf(mBl.IsZange, mBl.Blechdicke, 0)
    If Not mBr Is Nothing Then kl = kl + mBr.GesamtT
    'und in alle eintragen
    'm_S.Klemmlänge = kl
    'Dim ts 'As TScrew
    Dim i_x As Long, j_z As Long
    For i_x = 0 To m_NX - 1 'Each ts In m_Screws()
        For j_z = 0 To m_NZ - 1
            With m_Screws(i_x, j_z)
                If Not .sl Is Nothing Then
                    .sl.Schraube.Klemmlänge = kl
                End If
            End With
        Next
    Next
End Sub

Private Sub Resize(Optional sl As Schraubenloch)
    'ReDim Preserve m_SLs(0 To Anzahl - 1)
    'If (UBound(m_Screws, 1) <> m_NX - 1) Or (UBound(m_Screws, 2) <> m_NZ - 1) Then
        'Erase m_Screws
    If sl Is Nothing Then Set sl = m_Screws(0, 0).sl
    ReDim m_Screws(0 To m_NX - 1, 0 To m_NZ - 1) 'As TScrews
    'End If
    Dim xs As Double: xs = BRectSelInnW / 2
    Dim zs As Double: zs = BRectSelInnH / 2
    Dim ex As Double: ex = mASel.Loch.X
    Dim ez As Double: ez = mASel.Loch.Z
    Dim Ipl As Double
    Dim X As Double, Z As Double
    Dim i_x As Long, j_z As Long
    For i_x = 0 To m_NX - 1
        For j_z = 0 To m_NZ - 1
            With m_Screws(i_x, j_z)
                X = i_x * ex - xs
                Z = j_z * ez - zs
                Set .Pos = MNew.VectorXZ(X, Z)
                If .sl Is Nothing Then Set .sl = sl
                Ipl = Ipl + ((X / 1000) ^ 2 + (Z / 1000) ^ 2)
            End With
        Next
    Next
    m_Ip = Ipl
End Sub
Private Sub SetVxVz() 'ed As EinwirkungsKombi)
    Dim Ms As Double: Ms = m_Ekd.MomentS
    Dim Vs As Double: Vs = m_Ekd.QuerkraftS
    Dim Ns As Double: Ns = m_Ekd.NormalkraftS
    Dim n As Long: n = Me.Anzahl
    Dim Vx As Double, Vz As Double
    Dim i_x As Long, j_z As Long
    For i_x = 0 To m_NX - 1
        For j_z = 0 To m_NZ - 1
            With m_Screws(i_x, j_z)
                Vx = Ns / n
                If m_Ip <> 0 Then Vx = Vx + Ms / m_Ip * (.Pos.Z / 1000) 'geht nicht mit IIf
                'Vx = IIf(m_Ip <> 0, Ns / n + Ms / m_Ip * (.Pos.Z / 1000), Ns)
                'Vz = IIf(m_Ip <> 0, Vs / n - Ms / m_Ip * (.Pos.X / 1000), Vs)
                Vz = Vs / n
                If m_Ip <> 0 Then Vz = Vz - Ms / m_Ip * (.Pos.X / 1000) 'geht nicht mit IIf
                Set .v = MNew.VectorXZ(Vx, Vz)
            End With
        Next
    Next
End Sub
Public Property Get Pos(ByVal i_x As Long, ByVal j_z As Long) As VectorXZ 'die Position
    Set Pos = m_Screws(i_x, j_z).Pos
End Property
Public Property Get v(ByVal i_x As Long, ByVal j_z As Long) As VectorXZ 'die Kraft in x und z-ri
    Set v = m_Screws(i_x, j_z).v
End Property
Public Property Get Vmax() As VectorXZ
    Dim i_x As Long, j_z As Long
    Set Vmax = m_Screws(0, 0).v
    For i_x = 0 To m_NX - 1
        For j_z = 0 To m_NZ - 1
            With m_Screws(i_x, j_z)
                Set Vmax = IIf(Vmax.Length >= .v.Length, Vmax, .v)
            End With
        Next
    Next
End Property
Public Property Get Rmax() As Double
    Rmax = Vmax.Length
End Property

Public Property Get Ip() As Double
    'liefert das Polare Trägheitsmoment der Schrauben in m4
    Ip = m_Ip
'    Dim i_x As Long, j_z As Long
'    For i_x = 0 To m_NX - 1
'        For j_z = 0 To m_NZ - 1
'            With m_Screws(i_x, j_z)
'            End With
'        Next
'    Next
End Property
Public Property Get Schraube(ByVal i_x As Long, ByVal j_z As Long) As Schraube
    Set Schraube = m_Screws(i_x, j_z).sl.Schraube
'    Set Schraube = m_SLs(ix * m_N).Schraube
'    Dim n As Long: n = m_NX * m_NZ
'    If Index > n - 1 Then Index = n - 1
'    Set Schraube = m_SLs(Index).Schraube
'    If Not Schraube Is Nothing Then Exit Property
'    For i = Index To 0 Step -1
'        Set Schraube = m_SLs(i).Schraube
'        If Not Schraube Is Nothing Then Exit For
'    Next
End Property

Public Property Get MinGrenzabscherkraft() As Double
    Dim i_x As Long, j_z As Long
    Dim m As Double: m = Schraube(0, 0).Grenzabscherkraft
    For j_z = 0 To m_NZ - 1
        For i_x = 0 To m_NX - 1
            m = Min(m, Schraube(i_x, j_z).Grenzabscherkraft)
        Next
    Next
    MinGrenzabscherkraft = m
End Property
Public Property Get AGes() As Double
    Dim i_x As Long, j_z As Long
    For i_x = 0 To m_NX - 1
        For j_z = 0 To m_NZ - 1
            With m_Screws(i_x, j_z)
                AGes = AGes + .sl.Schraube.A
            End With
        Next
    Next
'    Dim i As Long, A As Double
'    Dim s As Schraube
'    Set s = m_SLs(0).Schraube
'    A = s.A
'    For i = 0 To Me.Anzahl - 1
'        Set s = m_SLs(i).Schraube
'        If s Is Nothing Then AGes = AGes + A Else AGes = AGes + s.A
'    Next
End Property

'Public Property Get Xi(ByVal i_x As Long, j_z As Long) As Double
'    '(I)ij
'    '(1)11 (2)12 (3)13 nx
'    '(4)21 (5)22 (6)23
'    '(7)31 (8)32 (9)33
'    '   nz
'    'i und j berechnen
'
'End Property
'Public Property Get Zi(ByVal Index As Long) As Double
'
'End Property

Public Property Get Schraubenloch(ByVal i_x As Long, ByVal j_z As Long) As Schraubenloch
    Set Schraubenloch = m_Screws(i_x, j_z).sl
'    If Index >= Anzahl Then Index = Anzahl - 1
'    If Index < 0 Then Index = 0
'    Set Schraubenloch = m_SLs(Index)
'    If Schraubenloch Is Nothing Then Set Schraubenloch = m_SLs(0)
End Property

Public Property Get Anzahl() As Long
    Anzahl = m_NX * m_NZ
End Property

Public Property Get AnzahlX() As Long
    AnzahlX = m_NX
End Property
Public Property Let AnzahlX(ByVal Value As Long)
    If Value < 1 Then Exit Property 'Value = 1
    m_NX = Value
    Dim sl As Schraubenloch: Set sl = m_Screws(0, 0).sl
    Resize sl
End Property
Public Property Get AnzahlZ() As Long
    AnzahlZ = m_NZ
End Property
Public Property Let AnzahlZ(ByVal Value As Long)
    If Value < 1 Then Exit Property 'Value = 1
    m_NZ = Value
    Dim sl As Schraubenloch: Set sl = m_Screws(0, 0).sl
    Resize sl
End Property

Public Property Get EAbstLoch() As EAbstand
    EAbstLoch = m_eaL
End Property
'Public Property Let EAbstLoch(bl As Blech, ByVal Value As EAbstand)
'    m_eaL = Value
'    SetAbstandLochOpt Value, bl
'End Property

Public Property Get EAbstRand() As EAbstand
    EAbstRand = m_eaR
End Property
'Public Property Let EAbstRand(ByVal Value As EAbstand)
'    m_eaR = Value
'    SetAbstandRandOpt bl
'End Property

Public Property Get AbstandSel() As AbstandLR
    Set AbstandSel = mASel 'die gewählten Abstände
End Property
Public Property Get AbstandOpt() As AbstandLR
    Set AbstandOpt = mAOpt 'die Min-Max-Abstände
End Property

'Public Property Get FbrdMin() As Double
'    'minmale Grenzlochleibung
'    If Not m_bl Is Nothing Then
'        '
'    End If
'End Property
'Beiwerte zur Berechnung der maximalen Schraubenkraft
'obsolete weil nicht funktioniert
'Public Property Get alpha() As Double
'    alpha = 6 / (m_NX * m_NZ) * (m_NZ - 1) / (m_NZ + 1)
'End Property
'Public Property Get beta() As Double
'    If m_NX = 1 Then Exit Property
'    beta = (m_NX + 1) / (m_NX - 1) * (m_NZ - 1) / (m_NZ + 1)
'End Property
'Public Property Get AbstandRandX() As Double
'    AbstandRandX = m_RandX 'D: e_1 ' E: e_1 'in Kraftrichtung
'End Property
'Public Property Let AbstandRandX(ByVal Value As Double)
'    m_RandX = Value        'D: e_1 ' E: e_1 'in Kraftrichtung
'End Property
'Public Property Get AbstandRandZ() As Double
'    AbstandRandZ = m_RandZ 'D: e_2 ' E: e_2 'senkr zu Kraftrichtung
'End Property
'Public Property Let AbstandRandZ(ByVal Value As Double)
'    m_RandZ = Value        'D: e_2 ' E: e_2 'senkr zu Kraftrichtung
'End Property
'
'Public Property Get AbstandLochX() As Double
'    'soll man bei nx=1 hier 0 zurückliefern? oder vlt unendl?
'    'AbstandLochX = m_LochX 'D: e_0 ' E: p_1 'in Kraftrichtung
'    AbstandLochX = m_LochX 'IIf(m_NX = 1, MMath.posINF, )
'End Property
'Public Property Let AbstandLochX(ByVal Value As Double)
'    m_LochX = Value        'D: e_0 ' E: p_1 'in Kraftrichtung
'End Property
'Public Property Get AbstandLochZ() As Double
'    'soll man bei nz=1 hier 0 zurückliefern? oder vlt unendl?
'    'AbstandLochZ = m_LochZ 'D: e_3 ' E: p_2 'senkr zu Kraftrichtung
'    AbstandLochZ = m_LochZ 'IIf(m_NZ = 1, MMath.posINF, )
'End Property
'Public Property Let AbstandLochZ(ByVal Value As Double)
'    m_LochZ = Value        'D: e_3 ' E: p_2 'senkr zu Kraftrichtung
'End Property

'
'Public Property Get EAbstandRand() As EAbstand
'    EAbstandRand = m_eaR
'End Property
'Public Property Let EAbstandRand(ByVal Value As EAbstand)
'    m_eaR = Value
'End Property
'
'Public Property Get EAbstandLoch() As EAbstand
'    EAbstandLoch = m_eaL
'End Property
'Public Property Let EAbstandLoch(ByVal Value As EAbstand)
'    m_eaL = Value
'End Property

'Public Property Get BoundingRectL() As Double
'    BoundingRectL = m_N
'End Property
'Public Property Get BoundingRectT() As Double
'
'End Property

'Rechteck der gewählten Abstände
'Innen nur Lochabstände
Public Property Get BRectSelInnW() As Double
    BRectSelInnW = mASel.RectInnWidth(m_NX)
End Property
Public Property Get BRectSelInnH() As Double
    BRectSelInnH = mASel.RectInnHeight(m_NZ)
End Property

'Außen Lochabstände + Randabstände
Public Property Get BRectSelOutW() As Double
    BRectSelOutW = mASel.RectOutWidth(m_NX)
End Property
Public Property Get BRectSelOutH() As Double
    BRectSelOutH = mASel.RectOutHeight(m_NZ)
End Property

'Rechteck der rechnerischen Abstände
'Innen nur die Lochabstände
Public Property Get BRectOptInnW() As Double
    BRectOptInnW = mAOpt.RectInnWidth(m_NX)
End Property
Public Property Get BRectOptInnH() As Double
    BRectOptInnH = mAOpt.RectInnHeight(m_NZ)
End Property

'Außen Lochabstände plus Randabstand
Public Property Get BRectOptOutW() As Double
    BRectOptOutW = mAOpt.RectOutWidth(m_NX)
End Property
Public Property Get BRectOptOutH() As Double
    BRectOptOutH = mAOpt.RectOutHeight(m_NZ)
End Property

Public Function GrenzlochleibungskraftRA() As Double 'RA=RandAbstand
    Dim gllkl As Double
    If Not mBl Is Nothing Then _
        gllkl = pGrenzlochleibungskraftRA(mBl)
    Dim gllkr As Double
    If Not mBr Is Nothing Then _
        gllkr = pGrenzlochleibungskraftRA(mBr)
    m_gllkMainRA = IIf(gllkl = gllkr, ETrMainLR.LinksURechts, IIf(gllkl < gllkr, ETrMainLR.Links, ETrMainLR.Rechts))
    'Set m_gllkMainRA = IIf(gllkl < gllkr, mBl, mBr)
    GrenzlochleibungskraftRA = Min(gllkl, gllkr)
End Function
'umpf eigentilch müßte es ein Interface NormAlgos geben
'mit einer Klasse NormAlgosDIN18800
'und einer klasse NormAlgosEurocode3
Public Function Alpha() As Double
    'liefert den Alpha-Wert für die Grenzlochleibungskraft
    'nö nicht so, anders, soll bei der Berechnung zwischengespeichert werden
'    Dim d_L As Double:     d_L = Me.Schraubenloch(0, 0).Durchmesser
'    Dim e_1 As Double:     e_1 = Me.AbstandSel.Rand.X
'    If m_N.Norm = Norm_DIN18800 Then
'        Dim e_2 As Double:     e_2 = Me.AbstandSel.Rand.Z
'        Dim e_3 As Double:     e_3 = IIf(m_NZ > 1, sg.AbstandSel.Loch.Z, MMath.posINF)
'        If (e_2 >= 1.5 * d_L) And (e_3 >= 3# * d_L) Then
'            Alpha = 1.1 * Min(e_1, 3 * d_L) / d_L - 0.3
'        Else
'            Alpha = 0.73 * Min(e_1, 3 * d_L) / d_L - 0.2
'        End If
'    Else
'        Alpha = Min(e_1 / (3 * d_L), Min(fub / b.fu, 1#))
'    End If
    Alpha = m_alpha
End Function
Public Function k1() As Double
    'wird nur für Ausgabe für Nachweis Grenzlochleibung nach Eurocode3 gebraucht
    k1 = m_k1
End Function
Private Function pGrenzlochleibungskraftRA(b As Blech) As Double 'RA=RandAbstand
    'liefert die Grenzlochleibungskraft in kN
    Dim sg As Schraubengruppe: Set sg = Me
    Dim d_S As Double:     d_S = sg.Schraube(0, 0).Durchmesser
    Dim d_L As Double:     d_L = sg.Schraubenloch(0, 0).Durchmesser
    Dim e_0 As Double:     e_0 = IIf(m_NX > 1, sg.AbstandSel.Loch.X, MMath.posINF)
    Dim e_1 As Double:     e_1 = sg.AbstandSel.Rand.X
    Dim e_2 As Double:     e_2 = sg.AbstandSel.Rand.Z
    Dim e_3 As Double:     e_3 = IIf(m_NZ > 1, sg.AbstandSel.Loch.Z, MMath.posINF)
    Dim p_1 As Double: p_1 = e_0 '= m_SG.AbstandLochX
    Dim p_2 As Double: p_2 = e_3 '= m_SG.AbstandLochZ
    If m_N.Norm = Norm_DIN18800 Then
        'V_l_Rd = t * d_Sch * alpha_l * f_yk / gam_M
        Dim f_yk As Double:   f_yk = b.fyk
        Dim gam_M As Double: gam_M = m_N.gam_M
        
        'Randabstand in Kraftrichtung
        Dim alp_l As Double
        If (e_2 >= 1.5 * d_L) And (e_3 >= 3# * d_L) Then
            alp_l = 1.1 * Min(e_1, 3 * d_L) / d_L - 0.3
        Else
            alp_l = 0.73 * Min(e_1, 3 * d_L) / d_L - 0.2
        End If
        m_alpha = alp_l 'zwischenspeichern für Ausgabe
        pGrenzlochleibungskraftRA = b.GesamtT * d_S * alp_l * f_yk / gam_M / 1000
        If Not b.IsZange Or sg.AnzahlX = 1 Then
            pGrenzlochleibungskraftRA = pGrenzlochleibungskraftRA / 1.2
        End If
    Else
        Dim d As Double:         d = sg.Schraube(0, 0).Schaftdurchmesser
        Dim fub As Double:     fub = sg.Schraube(0, 0).fubk
        Dim alp_b As Double: alp_b = Min(e_1 / (3 * d_L), Min(fub / b.fu, 1#))
        Dim k_1 As Double:     k_1 = Min(2.8 * e_2 / d_L - 1.7, Min(1.4 * p_2 / d_L - 1.7, 2.5))
        m_alpha = alp_b 'zwischenspeichern für Ausgabe
        m_k1 = k_1      'zwischenspeichern für Ausgabe
         pGrenzlochleibungskraftRA = k_1 * alp_b * b.fu * d * b.GesamtT / m_N.gam_M2 / 1000
        If Not b.IsMehrschnittig Then
            pGrenzlochleibungskraftRA = 1.5 * b.fu * d * b.GesamtT / m_N.gam_M2 / 1000
        End If
    End If
End Function
Public Function GrenzlochleibungskraftLA() As Double
    Dim gllkl As Double
    If Not mBl Is Nothing Then _
        gllkl = pGrenzlochleibungskraftLA(mBl)
    Dim gllkr As Double
    If Not mBr Is Nothing Then _
        gllkr = pGrenzlochleibungskraftLA(mBr)
    m_gllkMainLA = IIf(gllkl = gllkr, ETrMainLR.LinksURechts, IIf(gllkl < gllkr, ETrMainLR.Links, ETrMainLR.Rechts))
    GrenzlochleibungskraftLA = Min(gllkl, gllkr)
End Function
Private Function pGrenzlochleibungskraftLA(b As Blech) As Double 'LA=LochAbstand
    'liefert die Grenzlochleibungskraft in kN
    'V_l_Rd = t * d_Sch * alpha_l * f_yk / gam_M
    Dim sg As Schraubengruppe: Set sg = Me
    Dim d_S As Double:     d_S = sg.Schraube(0, 0).Durchmesser
    Dim d_L As Double:     d_L = sg.Schraubenloch(0, 0).Durchmesser
    Dim e_0 As Double:     e_0 = IIf(m_NX > 1, sg.AbstandSel.Loch.X, MMath.posINF)
    Dim e_1 As Double:     e_1 = sg.AbstandSel.Rand.X
    Dim e_2 As Double:     e_2 = sg.AbstandSel.Rand.Z
    Dim e_3 As Double:     e_3 = IIf(m_NZ > 1, sg.AbstandSel.Loch.Z, MMath.posINF)
    Dim p_1 As Double: p_1 = e_0 '= m_SG.AbstandLochX
    Dim p_2 As Double: p_2 = e_3 '= m_SG.AbstandLochZ
    If m_N.Norm = Norm_DIN18800 Then
        'mehr-schnittig
        Dim f_yk As Double:   f_yk = b.fyk
        Dim gam_M As Double: gam_M = m_N.gam_M
                
        'Lochabstand in Kraftrichtung
        Dim alp_l As Double
        If (e_2 >= 1.5 * d_L) And (e_3 >= 3# * d_L) Then
            alp_l = 1.08 * Min(e_0, 3.5 * d_L) / d_L - 0.77
        Else
            alp_l = 0.72 * Min(e_0, 3.5 * d_L) / d_L - 0.51
        End If
        m_alpha = alp_l 'zwischenspeichern für Ausgabe
        pGrenzlochleibungskraftLA = b.GesamtT * d_S * alp_l * f_yk / gam_M / 1000
        If sg.AnzahlX = 1 Then
            pGrenzlochleibungskraftLA = b.GesamtT * d_S * alp_l * f_yk / gam_M / 1000
        End If
        If Not b.IsZange Or sg.AnzahlX = 1 Then
            'ein-schnittig
            'einreihig
            pGrenzlochleibungskraftLA = pGrenzlochleibungskraftLA / 1.2
        End If
    Else
        'mehr-schnittig
        Dim d As Double:         d = sg.Schraube(0, 0).Schaftdurchmesser
        Dim fub As Double:     fub = sg.Schraube(0, 0).fubk
        Dim alp_b As Double: alp_b = Min(p_1 / (3 * d_L) - 1 / 4, Min(fub / b.fu, 1#))
        'If alp_b < 0 Then Debug.Print p_1
        Dim k_1 As Double:     k_1 = Min(1.4 * p_2 / d_L - 1.7, 2.5)
        
          pGrenzlochleibungskraftLA = k_1 * alp_b * b.fu * d * b.GesamtT / m_N.gam_M2 / 1000
        'If Not m_isZange Or m_sg.AnzahlX = 1 Then
        If Not b.IsMehrschnittig Then
            'ein-schnittig
            'einreihig
            pGrenzlochleibungskraftLA = 1.5 * b.fu * d * b.GesamtT / m_N.gam_M2 / 1000
        End If
    End If
End Function

Public Property Get MinGrenzlochleibungskraft() As Double
    MinGrenzlochleibungskraft = Min(GrenzlochleibungskraftLA, GrenzlochleibungskraftRA)
End Property

'        If e_0 < 2.2 * d_L Then e_0 = RoundUp(2.2 * d_L)
'        If e_0 > 3.5 * d_L Then e_0 = RoundUp(3.5 * d_L)
'        If e_0 > 12# * m_t Then e_0 = RoundUp(12# * m_t)
'
'        If e_1 < 1.2 * d_L Then e_1 = RoundUp(1.2 * d_L)
'        If e_1 > 3# * d_L Then e_1 = RoundUp(3# * d_L)
'        If e_1 > 6# * m_t Then e_1 = RoundUp(6 * m_t)
'
'        If e_2 < 1.2 * d_L Then e_2 = RoundUp(1.2 * d_L)
'        If e_2 > 1.5 * d_L Then e_2 = RoundUp(1.5 * d_L)
'        If e_2 > 6# * m_t Then e_2 = RoundUp(6 * m_t)
'
'        If e_3 < 2.4 * d_L Then e_3 = RoundUp(2.4 * d_L)
'        If e_3 > 3# * d_L Then e_3 = RoundUp(3# * d_L)
'        If e_3 > 12# * m_t Then e_3 = RoundUp(12# * m_t)
        
Public Function Grenzdurchstanzkraft() As Double
    Dim gdstl As Double
    If Not mBl Is Nothing Then _
        gdstl = pGrenzdurchstanzkraft(mBl)
    Dim gdstr As Double
    If Not mBr Is Nothing Then _
        gdstr = pGrenzdurchstanzkraft(mBr)
    m_gdstMain = IIf(gdstl = gdstr, ETrMainLR.LinksURechts, IIf(gdstl < gdstr, ETrMainLR.Links, ETrMainLR.Rechts))
    'Set m_gdstMain = IIf(gdstl < gdstr, mBl, mBr)
    Grenzdurchstanzkraft = Min(gdstl, gdstr)
End Function
Private Function pGrenzdurchstanzkraft(b As Blech) As Double
    'liefert die Grenzdurchstanzkraft B_pRd in kN je Schraube bezogen auf 10mm Blechdicke
    'nur für vorgespannte Schrauben
    'berechnet die
    Dim sg As Schraubengruppe: Set sg = Me
    Dim d_m As Double:       d_m = sg.Schraube(0, 0).Kopfdurchmesser
    Dim gam_M2 As Double: gam_M2 = m_N.gam_M2 '1.25
    Dim B_pRd As Double:   B_pRd = 0.6 * 4 * Atn(1) * d_m * b.GesamtT * b.fu / gam_M2 / 1000
    pGrenzdurchstanzkraft = B_pRd
End Function
Public Function BlockversagenNzentr() As Double
    Dim bvnzl As Double
    If Not mBl Is Nothing Then _
        bvnzl = pBlockversagenNzentr(mBl)
    Dim bvnzr As Double
    If Not mBr Is Nothing Then _
        bvnzr = pBlockversagenNzentr(mBr)
    m_bvnzMain = IIf(bvnzl = bvnzr, ETrMainLR.LinksURechts, IIf(bvnzl < bvnzr, ETrMainLR.Links, ETrMainLR.Rechts))
    BlockversagenNzentr = Min(bvnzl, bvnzr)
End Function
Private Function pBlockversagenNzentr(b As Blech) As Double
    'Berechnet die Kraft V_eff,1,Rd in kN
Try: On Error GoTo Catch
    Dim sg As Schraubengruppe: Set sg = Me
    Dim nh As Long:               nh = sg.AnzahlZ 'ja horizontal ist hier Z, weil Formeln im Schneider umgedreht
    Dim nv As Long:               nv = sg.AnzahlX 'ja vertikal   ist hier X, weil Formeln im Schneider umgedreht
    Dim Lh As Double: Lh = IIf(nh > 1, (nh - 1) * sg.AbstandSel.Loch.Z, 0) 'hier so lassen!!! nicht BRectInnW oder BRectInnH verwenden
    'If nh > 1 Then Lh = (nh - 1) * m_sg.AbstandLochZ Else Lh = 0
    Dim Lv As Double: Lv = IIf(nv > 1, (nv - 1) * sg.AbstandSel.Loch.X, 0) 'hier so lassen!!! nicht BRectInnW oder BRectInnH verwenden
    'If nv > 1 Then
        
    'Else
    '    Lv = 0
    'End If
    Dim d_0 As Double:           d_0 = sg.Schraubenloch(0, 0).Durchmesser
    Dim Ant As Double:           Ant = (Lh - (nv - 1#) * d_0) * b.GesamtT
    Dim Anv As Double:           Anv = 2 * (Lv - (nh - 0.5) * d_0) * b.GesamtT
    Dim n As Norm:             Set n = sg.Schraube(0, 0).Norm
    Dim V_eff1Rd As Double: V_eff1Rd = b.fu * Ant / n.gam_M2 + b.fy * Anv / (Sqr(3) * n.gam_M0)
    pBlockversagenNzentr = V_eff1Rd / 1000
    Exit Function
Catch:
    If Err.Number = 6 Then Resume Next 'Überlauf wg unendlich
End Function
Public Function BlockversagenVexztr() As Double
    Dim bvvel As Double
    If Not mBl Is Nothing Then _
        bvvel = pBlockversagenVexztr(mBl)
    Dim bvver As Double
    If Not mBr Is Nothing Then _
        bvver = pBlockversagenVexztr(mBr)
    m_bvveMain = IIf(bvvel = bvver, ETrMainLR.LinksURechts, IIf(bvvel < bvver, ETrMainLR.Links, ETrMainLR.Rechts))
    BlockversagenVexztr = Min(bvvel, bvver)
End Function
Private Function pBlockversagenVexztr(b As Blech) As Double
    'Berechnet die Kraft V_eff,2,Rd in kN
Try: On Error GoTo Catch
    Dim sg As Schraubengruppe: Set sg = Me
    Dim nh As Long:               nh = sg.AnzahlX
    Dim nv As Long:               nv = sg.AnzahlZ
    Dim a1 As Double:             a1 = sg.AbstandSel.Rand.Z
    Dim a2 As Double:             a2 = sg.AbstandSel.Rand.X
    Dim a3 As Double:             a3 = sg.AbstandSel.Rand.Z
    Dim Lv As Double: Lv = IIf(nv > 1, (nv - 1) * sg.AbstandSel.Loch.Z, 0) 'hier so lassen!!! nicht BRectInnW oder BRectInnH verwenden
    Dim d_0 As Double:           d_0 = sg.Schraubenloch(0, 0).Durchmesser
    Dim Ant As Double:           Ant = (a2 - (nh - 0.5) * d_0) * b.GesamtT
    Dim Anv As Double:           Anv = (a1 + Lv - (nv - 0.5) * d_0) * b.GesamtT
    Dim n As Norm:             Set n = sg.Schraube(0, 0).Norm
    Dim V_eff2Rd As Double: V_eff2Rd = 0.5 * b.fu * Ant / n.gam_M2 + b.fy * Anv / (Sqr(3) * n.gam_M0)
    pBlockversagenVexztr = V_eff2Rd / 1000
    Exit Function
Catch:
    If Err.Number = 6 Then
        'err.
        Err.Clear
        'Resume Next 'Überlauf wg unendlich
    End If
End Function

Public Sub ToListBox(aLB As ListBox)
    'Dim isEC As Boolean: isEC = m_N.Norm = Norm_EuroCode3
    With aLB
        .AddItem "Schraubengruppe"
        .AddItem "  Anzahl-X : " & AnzahlX & ";  Anzahl-Z : " & AnzahlZ
        .AddItem "  Gewählte Abstände "
        mASel.ToListBox aLB
        .AddItem "  Rechnerische Abstände "
        mAOpt.ToListBox aLB
        .AddItem "  Gesamthöhe: " & BRectSelOutH & "; Gesamtbreite: " & BRectSelOutW
        'hmm ja man muss Grenzlochleibungskraft für beide Bleche durchführen
        Dim n As String
        Dim v As Double
        
        n = m_N.AbkName_Grenzlochleibungskraft
        v = GrenzlochleibungskraftLA
        If v > 0 Then
            .AddItem "  Grenzlochleibg.LA " & n & ": " & Round(v, 2) & "[kN]"
            If m_gllkMainLA > 0 Then _
                .AddItem "    Maßgebender Träger " & ETrMainLR_ToStr(m_gllkMainLA)
        End If
        v = GrenzlochleibungskraftRA
        If v > 0 Then
            .AddItem "  Grenzlochleibg.RA " & n & ": " & Round(v, 2) & "[kN]"
            If m_gllkMainRA > 0 Then _
                .AddItem "    Maßgebender Träger " & ETrMainLR_ToStr(m_gllkMainRA)
        End If
        n = m_N.AbkName_Grenzdurchstanzkraft
        v = Grenzdurchstanzkraft
        If Me.Schraube(0, 0).IsVorgespannt Then
            .AddItem "  Grenzdurchstanzkraft " & n & ": " & Round(v, 1) & "[kN]"
            If m_gllkMainRA > 0 Then _
                .AddItem "    Maßgebender Träger " & ETrMainLR_ToStr(m_gdstMain)
        End If
        
        n = m_N.AbkName_BlockversagenskraftNzentr
        v = BlockversagenNzentr
        If IsPositive(v) Then
            .AddItem "  Blockversagen " & n & ": " & Round(v, 1) & "[kN]"
            If m_bvnzMain > 0 Then _
                .AddItem "    Maßgebender Träger " & ETrMainLR_ToStr(m_bvnzMain)
        End If
        n = m_N.AbkName_BlockversagenskraftVexztr
        v = BlockversagenVexztr
        If IsPositive(v) Then
            .AddItem "  Blockversagen " & n & ": " & Round(v, 1) & "[kN]"
            If m_bvveMain > 0 Then _
                .AddItem "    Maßgebender Träger " & ETrMainLR_ToStr(m_bvveMain)
        End If
        .AddItem "  Polares Trägheitsmoment Ip: " & Round(m_Ip, 6)
        If Not m_Ekd Is Nothing Then
            SetVxVz
            .AddItem "Max. Schraubenbelastung: " & Round(Me.Rmax, 3) & " kN"
        End If
    End With
End Sub


'Minimale Abstände red. Tragkraft
Private Function MinRedE1() As Double 'AbstandRandX
    Dim d_L As Double: d_L = m_Screws(0, 0).sl.Durchmesser
    MinRedE1 = RoundUp(1.2 * d_L)
End Function
Private Function MinRedE2() As Double 'AbstandRandZ
    Dim d_L As Double: d_L = m_Screws(0, 0).sl.Durchmesser
    MinRedE2 = RoundUp(1.2 * d_L)
End Function
Private Function MinRedP1() As Double 'AbstandLochX aka e_0
    Dim d_L As Double: d_L = m_Screws(0, 0).sl.Durchmesser
    MinRedP1 = RoundUp(2.2 * d_L)
End Function
Private Function MinRedP2() As Double 'AbstandLochZ aka e_3
    Dim d_L As Double: d_L = m_Screws(0, 0).sl.Durchmesser
    MinRedP2 = RoundUp(2.4 * d_L)
End Function
Private Function MinRedE3() As Double 'AbstandRandX
    Dim d_L As Double: d_L = m_Screws(0, 0).sl.Durchmesser
    MinRedE3 = RoundUp(1.5 * d_L)
End Function
Private Function MinRedE4() As Double 'AbstandRandZ
    Dim d_L As Double: d_L = m_Screws(0, 0).sl.Durchmesser
    MinRedE4 = RoundUp(1.5 * d_L)
End Function

'Minimale Abstände volle Tragkraft
Private Function MinVolE1() As Double 'AbstandRandX
    Dim d_L As Double: d_L = m_Screws(0, 0).sl.Durchmesser
    MinVolE1 = RoundUp(3# * d_L) 'auf Ganze runden
End Function
Private Function MinVolE2() As Double 'AbstandRandZ
    Dim d_L As Double: d_L = m_Screws(0, 0).sl.Durchmesser
    MinVolE2 = RoundUp(1.5 * d_L) 'auf Ganze runden
End Function
Private Function MinVolP1() As Double 'AbstandLochX aka e_0
    Dim d_L As Double: d_L = m_Screws(0, 0).sl.Durchmesser
    MinVolP1 = RoundUp(3.75 * d_L) 'auf Ganze runden
End Function
Private Function MinVolP2() As Double 'AbstandLochZ aka e_3
    Dim d_L As Double: d_L = m_Screws(0, 0).sl.Durchmesser
    MinVolP2 = RoundUp(3# * d_L) 'auf Ganze runden
End Function

Public Sub SetAbstandRandOpt(ByVal A As EAbstand, Optional bl As Blech) ', Optional ByVal br5 As Boolean = False)
    'setzt Schraubenabstand in Abhängigkeit vom Schraubendurchmesser
    m_eaR = A
    Select Case m_eaR
    Case AbstandMinRed: Set mAOpt.Rand = MNew.VectorXZ(MinRedE1, MinRedE2)
    Case AbstandMinMit: Set mAOpt.Rand = MNew.VectorXZ((MinRedE1 + MinVolE1) / 2, (MinRedE2 + MinVolE2) / 2)
    Case AbstandMinVol: Set mAOpt.Rand = MNew.VectorXZ(MinVolE1, MinVolE2)
    Case AbstandMaximal
        If Not bl Is Nothing Then Set mAOpt.Rand = MNew.VectorXZ(bl.MaxE1, bl.MaxE2)
    End Select
End Sub
Public Sub SetAbstandLochOpt(ByVal A As EAbstand, Optional bl As Blech) ', Optional ByVal br5 As Boolean = False)
    'setzt Schraubenabstand in Abhängigkeit vom Schraubendurchmesser
    m_eaL = A
    Select Case m_eaL
    Case AbstandMinRed: Set mAOpt.Loch = MNew.VectorXZ(MinRedP1, MinRedP2)
    Case AbstandMinMit: Set mAOpt.Loch = MNew.VectorXZ((MinRedP1 + MinVolP1) / 2, (MinRedP2 + MinVolP2) / 2)
    Case AbstandMinVol: Set mAOpt.Loch = MNew.VectorXZ(MinVolP1, MinVolP2)
    Case AbstandMaximal
        If Not bl Is Nothing Then Set mAOpt.Loch = MNew.VectorXZ(bl.MaxP1, bl.MaxP2)
    End Select
End Sub

Public Sub SyncAbstandLochSel(Optional ByVal br5 As Boolean = False)
    Set mASel.Loch = IIf(br5, MNew.VectorXZ(RoundUp5(mAOpt.Loch.X, 0), RoundUp5(mAOpt.Loch.Z, 0)), MNew.VectorXZCopy(mAOpt.Loch))
End Sub
Public Sub SyncAbstandRandSel(Optional ByVal br5 As Boolean = False)
    Set mASel.Rand = IIf(br5, MNew.VectorXZ(RoundUp5(mAOpt.Rand.X, 0), RoundUp5(mAOpt.Rand.Z, 0)), MNew.VectorXZCopy(mAOpt.Rand))
End Sub

