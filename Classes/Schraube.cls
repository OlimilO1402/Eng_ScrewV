VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Schraube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Schraubenklasse für Schrauben nach DIN 7968, DIN 7990, DIN7991 . . .
'HV (vorgespannt) mit großen Schlüsselweiten
'  alt: DIN 6914, DIN 7999
'  neu: DIN EN 14399-4, DIN EN 14399-8
Private m_N As Norm 'DIN18800 oder EuroCode 3
Private m_d As Double   ' Schraubendurchmesser in [mm] normalerweise nur ganze Millimeter
                        ' => Byte müßte reichen
Private m_l As Double   'die Klemmlänge, daraus ergibt sich die erforderliche Schraubenlänge
Private m_g As Double   ' Schraubengüte 4.6, 5.6, 8.8, 10.9, was gibbet no?
                        ' 4.8, 5.8, 6.8: in D nicht zugelassen
                        ' aka SFK = SchraubenFestigkeitsKlasse
Private m_isSenk As Boolean ' True: Schraube hat Senkkopf DIN7991
Private m_isPass As Boolean ' True: Schraube ist Passschraube DIN 7968
Private m_isSFSchaft As Boolean 'Scherfuge True: im Schaft, False: im Gewinde
Private m_isGleitf   As Boolean 'True: Gleitfeste Verbindung, False: Scher-Lochleibungsverb.
Private m_isGlImGZT  As Boolean 'True: Gleiten im GZT, False: Gleiten im GZG
Private m_isZugvbdg  As Boolean 'True: Zugverbindung, Schrauben werden auf Zug belastet
Private m_isVorgesp  As Boolean 'True: Vorgespannte HV-Schrauben, False: SL
Private Enum VerbArt
    sl = 0       '00000 '(Scher) Kategorie A Scher-Lochleibungsverb.
    pas = 1      '00001
    glf = 2      '00010
    gzt = 4      '00100
    zug = 8      '01000
    vor = 16     '10000
    SLP = 1      '00001 (Scher) Kategorie A Scher-Lochleibungsverb. mit Passschraube
    GV_GZG = 18  '10010 (Scher) Kategorie B Gleitfeste vorgespannte Verb. im GZG
    GV_GZT = 22  '10110 (Scher) Kategorie C Gleitfeste vorgespannte Verb. im GZT
    GVP_GZG = 19 '10011 (Scher) Kategorie B Gleitfeste vorgespannte Verb. im GZG mit Passschraube
    GVP_GZT = 23 '10111 (Scher) Kategorie C Gleitfeste vorgespannte Verb. im GZT mit Passschraube
    'zug = 8     '01000 ( Zug ) Kategorie D Zugverbindung nicht vorgespannt
    zugP = 9     '01001 ( Zug ) Kategorie D Zugverbindung nicht vorgespannt mit Passschraube
    zugV = 24    '11000 ( Zug ) Kategorie D Zugverbindung vorgespannt
    zugVP = 25   '11001 ( Zug ) Kategorie D Zugverbindung vorgespannt mit Passschraube
    SLV = 16     '10000 ( Zug ) Kategorie E Scher-Lochleibungsverb. vorgespannt
    SLVP = 17    '10001 ( Zug ) Kategorie E Scher-Lochleibungsverb. vorgespannt mit Passschraube
End Enum

Friend Sub New_(ByVal aCalcNorm As Norm, _
                ByVal Schraubendurchmesser As Double, _
                Optional ByVal Schraubengüte As Double = 4.6, _
                Optional ByVal isSenk As Boolean = False, _
                Optional ByVal isPass As Boolean = False, _
                Optional ByVal isSFS As Boolean = False, _
                Optional ByVal isGlf As Boolean = False, _
                Optional ByVal isZug As Boolean = False, _
                Optional ByVal isVor As Boolean = False)
    Set m_N = aCalcNorm
    m_d = Schraubendurchmesser
    m_g = Schraubengüte
    m_isSenk = isSenk
    m_isPass = isPass
    m_isSFSchaft = isSFS
    m_isGleitf = isGlf
    m_isZugvbdg = isZug
    m_isVorgesp = isVor
End Sub

' ######################################## ' v   ISerializable   v ' ######################################## '
Public Property Get ID() As Long
    ID = ObjPtr(Me)
End Property
Public Property Let ID(ByVal Value As Long)
    'nix
End Property

Public Property Get Norm() As Norm
    Set Norm = m_N
End Property
Public Property Set Norm(ByVal Value As Norm)
    Set m_N = Value
End Property

Public Property Get Durchmesser() As Double
    Durchmesser = m_d
End Property
Public Property Let Durchmesser(ByVal Value As Double)
    m_d = Value
End Property

Public Property Get Klemmlänge() As Double
    Klemmlänge = m_l
End Property
Public Property Let Klemmlänge(ByVal Value As Double)
    m_l = Value
End Property

Public Property Get Schraubengüte() As Double
    Schraubengüte = m_g
End Property
Public Property Let Schraubengüte(ByVal Value As Double)
    m_g = Value
End Property

Public Property Get IsSenkschraube() As Boolean 'DIN 7991
    IsSenkschraube = m_isSenk
End Property
Public Property Let IsSenkschraube(ByVal Value As Boolean)
    m_isSenk = Value
End Property

Public Property Get IsPassschraube() As Boolean 'DIN 7968
    IsPassschraube = m_isPass
End Property
Public Property Let IsPassschraube(ByVal Value As Boolean)
    m_isPass = Value
    'OK hier sollte man es nicht anpassen, weil dann die User-Einstellung verloren geht
    'einfach den View disablen ist besser
    'm_isSFSchaft = Not Value ' True
    'OK wenn Passschraube dann muss doch auch das Schraubenloch angepasst werden.
End Property

Public Property Get IsScherfugeSchaft() As Boolean
    IsScherfugeSchaft = m_isSFSchaft
End Property
Public Property Let IsScherfugeSchaft(ByVal Value As Boolean)
    m_isSFSchaft = Value
End Property

Public Property Get IsGleitfest() As Boolean
    IsGleitfest = m_isGleitf
End Property
Public Property Let IsGleitfest(ByVal Value As Boolean)
    m_isGleitf = Value  'wenn Gleitfest dann immer auch vorgespannt!!!
    'OK hier sollte man es nicht anpassen, weil dann die User-Einstellung verloren geht
    'einfach den View disablen ist besser
    'm_isVorgesp = Value
    'm_isZugvbdg = False 'Not Value
    If Not m_isGleitf Then m_isGlImGZT = False
End Property

Public Property Get IsGleitfestImGZT() As Boolean
    IsGleitfestImGZT = m_isGlImGZT 'False: im GZG; True: im GZT
End Property
Public Property Let IsGleitfestImGZT(ByVal Value As Boolean)
    m_isGlImGZT = Value            'False: im GZG; True: im GZT
End Property

Public Property Get IsZugverbindung() As Boolean
    IsZugverbindung = m_isZugvbdg
End Property
Public Property Let IsZugverbindung(ByVal Value As Boolean)
    m_isZugvbdg = Value
End Property

Public Property Get IsVorgespannt() As Boolean
    IsVorgespannt = m_isVorgesp
End Property
Public Property Let IsVorgespannt(ByVal Value As Boolean)
    m_isVorgesp = Value
End Property
' ######################################## ' ^   ISerializable   ^ ' ######################################## '

Public Property Get Länge() As Double
    Länge = RoundUp5(m_l + Me.Mutterhöhe + 2 * Me.Scheibendicke, 0)
End Property

Public Property Get fybk() As Double
    'fyb = fub / 10 * (m_g - fub / 100) * 10
    fybk = Round(Int(m_g) * 100# * (m_g - Int(m_g)))
End Property
Public Property Get fubk() As Double
    fubk = Int(m_g) * 100
End Property
 
'Public Property Get Lochdurchmesser() As Double
'    'liefert den Lochdurchmesser für Normalloch d_0 in [mm] für Schraubendurchmesser
'    'halt Stop eigentlich in Abhängigkeit von der Lochart
'    ' +1 , +2, +3 siehe Nennlochspiel in Klasse Schraubenloch
'    Dim dL As Double
'    Select Case m_d
'    Case 12: dL = m_d + 1
'    Case 16: dL = m_d + 2
'    Case 20: dL = m_d + 2
'    Case 22: dL = m_d + 2
'    Case 24: dL = m_d + 2
'    Case 27: dL = m_d + 3
'    Case 30: dL = m_d + 3
'    Case 36: dL = m_d + 3
'    Case Else
'        Select Case True
'        Case m_d <= 12: dL = m_d + 1
'        Case m_d <= 24: dL = m_d + 2
'        Case Else:      dL = m_d + 3
'        End Select
'    End Select
'    Lochdurchmesser = dL
'End Property

'OK, das sollte eigentlich in Norm rein:
Public Property Get SchraubenNorm() As String
    Dim s As String: s = "DIN "
    If m_N.ENorm = Norm_DIN18800 Then
        Select Case m_g
        Case 4.6:  s = s & IIf(IsPassschraube, "7969", "7990")
        Case 5.6:  s = s & IIf(IsPassschraube, "7968", "7990")
        Case 8.8:  s = s & "EN ISO " & IIf(IsPassschraube, "4017", "4014")
        Case 10.9: s = s & IIf(IsPassschraube, "7999", "6914")
        End Select
    Else
        Select Case m_g
        Case 4.6:  s = s & IIf(IsPassschraube, "7969", "7990")
        Case 5.6:  s = s & IIf(IsPassschraube, "7968", "7990")
        Case 8.8:  s = s & "EN ISO " & IIf(IsPassschraube, "4017", "4014")
        Case 10.9: s = s & "EN 14399-" & IIf(IsPassschraube, "-8", "-4")
        End Select
    End If
    SchraubenNorm = s
End Property
Public Property Get Kategorie() As String
    Dim s As String
    If IsVorgespannt And Not IsGleitfest Then
        s = "E"
    ElseIf IsZugverbindung Then
        s = "D"
    ElseIf IsGleitfest Then
        If IsGleitfestImGZT Then
            s = "C"
        Else
            s = "B"
        End If
    Else
        s = "A"
    End If
    Kategorie = s
End Property
Private Property Get VerbArt() As VerbArt
    Dim va As VerbArt
    If IsPassschraube Then _
        va = va Or pas
    If IsGleitfest Then _
        va = va Or glf: va = va Or vor
    If IsGleitfestImGZT Then _
        va = va Or GV_GZT
    If IsZugverbindung Then _
        va = va Or zug
    If IsVorgespannt Then _
        va = va Or vor
    VerbArt = va
'    SL = 0       '00000 '(Scher) Kategorie A Scher-Lochleibungsverb.
'    pas = 1      '00001
'    glf = 2      '00010
'    gzt = 4      '00100
'    zug = 8      '01000
'    vor = 16     '10000
'    SLP = 1      '00001 (Scher) Kategorie A (Scher) Scher-Lochleibungsverb. mit Passschraube
'    GV_GZG = 18  '10010 (Scher) Kategorie B Gleitfeste vorgespannte Verb. im GZG
'    GV_GZT = 22  '10110 (Scher) Kategorie C Gleitfeste vorgespannte Verb. im GZT
'    GVP_GZG = 19 '10011 (Scher) Kategorie B Gleitfeste vorgespannte Verb. im GZG mit Passschraube
'    GVP_GZT = 23 '10111 (Scher) Kategorie C Gleitfeste vorgespannte Verb. im GZT mit Passschraube
'    zug = 8      '01000 ( Zug ) Kategorie D Zugverbindung nicht vorgespannt
'    SLV = 16     '10000 ( Zug ) Kategorie E Scher-Lochleibungsverb. mit Vorspannung
'    SLVP = 17    '10001 ( Zug ) Kategorie E Scher-Lochleibungsverb. mit Passschraube und Vorspannung
End Property
Private Function VerbArtToStr() As String
    Dim va As VerbArt: va = VerbArt
    Dim s As String
    Select Case va
    Case sl:      s = "SL"
    Case SLP:     s = "SLP"
    Case SLV:     s = "SLV"
    Case SLVP:    s = "SLVP"
    Case GV_GZG:  s = "GV_GZG"
    Case GV_GZT:  s = "GV_GZT"
    Case GVP_GZG: s = "GVP_GZG"
    Case GVP_GZT: s = "GVP_GZT"
    Case zug:     s = "Zug"
    Case zugP:    s = "ZugP"
    Case zugV:    s = "ZugV"    '11000 ( Zug ) Kategorie D Zugverbindung vorgespannt
    Case zugVP:   s = "ZugVP"   '11001 ( Zug ) Kategorie D Zugverbindung vorgespannt mit Passschraube
    End Select
    VerbArtToStr = s
End Function
Public Property Get Ausführungsform() As String
    Ausführungsform = VerbArtToStr
End Property
Public Property Get Schaftdurchmesser() As Double
    'liefert den Schaftdurchmesser für Schraubendurchmesser
    'bei Normalschrauben ist der Schaftdurchmesser = Gew.-Durchmesser
    'bei Passschrauben ist der Schaftdurchmesser = Gew.-Durchmesser + 1
    Schaftdurchmesser = m_d + IIf(m_isPass, 1, 0)
    'wie ist das, sollte bei Passschrauben der Schaftdurchmesser nicht gleich dem Gewinde-Durchmesser + Nennlochspiel sein?
    'oder ist bei Passschrauben das Nennlochspiel immer = 0
    'gibt es bei Passschrauben eigentlich ein Langloch?
End Property

Public Property Get Kopfdurchmesser() As Double 'hauptsächlich bei Senkkopf relevant
    Dim k As Double
    If IsSenkschraube Then
        Select Case m_d
        Case 8:  k = 16
        Case 10: k = 20
        Case 12: k = 24
        Case 16: k = 30
        Case 20: k = 36
        Case 22: k = 38
        Case 24: k = 39
        Case 27: k = 42 'extrapoliert/geschätzt noch nicht verifiziert: om 10.07.2015
        Case 30: k = 46 'extrapoliert/geschätzt noch nicht verifiziert: om 10.07.2015
        Case 36: k = 52 'extrapoliert/geschätzt noch nicht verifiziert: om 10.07.2015
        Case Else: k = RoundDown(m_d / 3) * 2
        End Select
    Else
        'wird hier zweckentfremdet, wird gebraucht bei -> siehe Blech Grenzdurchstanzkraft
        k = (Schlüsselweite + EckenmassMin) / 2
    End If
    Kopfdurchmesser = k
End Property
Public Property Get Schlüsselweite() As Double
    'liefert die Schlüsselweite für Schraubendurchmesser
    Dim s As Double
    If IsSenkschraube Then 'Innensechskant (Inbus)
        Select Case m_d
        Case 8:  s = 5
        Case 10: s = 6
        Case 12: s = 8
        Case 16: s = 10
        Case 20: s = 12
        Case 22: s = 13
        Case 24: s = 14
        Case 27: s = 16 'noch nicht verifiziert: om 10.07.2015
        Case 30: s = 18 'noch nicht verifiziert: om 10.07.2015
        Case 36: s = 22 'noch nicht verifiziert: om 10.07.2015
        Case Else: s = RoundDown(m_d / 3) * 2
        End Select
    Else
        s = MutterSchlüsselweite
    End If
    Schlüsselweite = s
End Property
 
Public Property Get Eckenmass() As Double
    'liefert das Eckenmass zum Zeichnen des Sechsecks(=Kopfdurchmesser an breitester Stelle) für Schlüsselweite
    'Eckenmass = VBA.Math.Sqr(s ^ 2 + (VBA.Math.Tan(CRad(30)) * s) ^ 2)
    Eckenmass = Schlüsselweite * 2 / Sqr(3)
End Property
Public Property Get EckenmassMin() As Double
    'liefert das Eckenmass e_min
    Dim e As Double
    If IsVorgespannt Then
        Select Case m_d
        Case 8:  e = 14.83
        Case 10: e = 18.53
        Case 12: e = 23.91
        Case 16: e = 29.56
        Case 20: e = 35.03
        Case 22: e = 39.55
        Case 24: e = 45.2
        Case 27: e = 50.85
        Case 30: e = 55.37
        Case 36: e = 66.44
        Case Else: e = m_d + m_d / 2
        End Select
    Else
        Select Case m_d
        Case 8:  e = 11.82
        Case 10: e = 14.55
        Case 12: e = 19.85
        Case 16: e = 26.17
        Case 20: e = 32.95
        Case 22: e = 37.29
        Case 24: e = 39.55
        Case 27: e = 45.2
        Case 30: e = 50.85
        Case 36: e = 60.79
        Case Else: e = m_d + m_d / 2
        End Select
    End If
    EckenmassMin = e
End Property
Public Property Get Kopfhöhe() As Double
    'liefert die Kopfhöhe für Schraubendurchmesser
    Dim k As Double
    Select Case m_d
    Case 8:  k = 6  '= 2 / 3 * d
    Case 10: k = 7 '= 5 / 8 * d
    Case 12: k = 8  '= 2 / 3 * d
    Case 16: k = 10 '= 5 / 8 * d
    Case 20: k = 13 '= 2 / 3 * d
    Case 22: k = 14 '= 2 / 3 * d
    Case 24: k = 15 '= 2 / 3 * d
    Case 27: k = 17 '
    Case 30: k = 19
    Case 36: k = 23
    Case Else: k = 2 / 3 * m_d
    End Select
    Kopfhöhe = k
End Property

Public Property Get MutterEckenmass() As Double
    MutterEckenmass = MutterSchlüsselweite * 2 / Sqr(3)
End Property
Public Property Get MutterSchlüsselweite() As Double
    Dim s As Double
    If IsVorgespannt Then
        Select Case m_d
        Case 8:  s = 14
        Case 10: s = 17
        Case 12: s = 22
        Case 16: s = 27
        Case 20: s = 32
        Case 22: s = 36
        Case 24: s = 41
        Case 27: s = 46
        Case 30: s = 50
        Case 36: s = 60
        Case Else: s = m_d / 0.6
        End Select
    Else
        Select Case m_d
        Case 8:  s = 13 ' vormals: 14
        Case 10: s = 16 ' vormals: 17
        Case 12: s = m_d + m_d / 2 '=18 nach ISO; =19 nach alte DIN
        Case 16: s = m_d + m_d / 2
        Case 20: s = m_d + m_d / 2
        Case 22: s = 34 'm_d + m_d / 2
        Case 24: s = m_d + m_d / 2
        Case 27: s = 41 'm_d + m_d / 2
        Case 30: s = 46 'm_d + m_d / 2
        Case 36: s = 55 'm_d + m_d / 2
        Case Else: s = m_d + m_d / 2
        End Select
    End If
    MutterSchlüsselweite = s
End Property

Public Property Get Mutterhöhe() As Double
    'liefert die Mutterhöhe (nach DIN EN ISO 4034) für Schraubendurchmesser
    Dim m As Double
    Select Case m_d
    Case 8: m = 7.9
    Case 10: m = 9.5
    Case 12: m = 12.2
    Case 16: m = 15.9
    Case 20: m = 19#
    Case 22: m = 20.2
    Case 24: m = 22.3
    Case 27: m = 24.7
    Case 30: m = 26.4
    Case 36: m = 31.5
    Case Else: m = 0.95 * m_d
    End Select
    Mutterhöhe = m
End Property
 
Public Property Get Scheibendurchmesser() As Double
    'liefert den (Unterleg-/Beilag-) Scheibendurchmesser für Schraubendurchmesser
    'nach DIN 7989-1 für Sechskantschraube
    'nach DIN 7989-2 für Sechskant-Passschraube
    Dim SchD As Double
    Select Case m_d
    Case 8:  SchD = 16
    Case 10: SchD = 21
    Case 12: SchD = 24
    Case 16: SchD = 30
    Case 20: SchD = 37
    Case 22: SchD = 39
    Case 24: SchD = 44
    Case 27: SchD = 50
    Case 30: SchD = 56
    Case 36: SchD = IIf(IsVorgespannt, 66, 66)
    Case Else: SchD = RoundUp(1.9 * m_d)
    End Select
    Scheibendurchmesser = SchD
End Property
Public Property Get Scheibendicke() As Double
    Dim d As Double
    If IsVorgespannt Then
        Select Case m_d
        Case 8:  d = 3
        Case 10: d = 3
        Case 12: d = 3
        Case 16 To 24: d = 4
        Case 27 To 30: d = 5
        Case 36:     d = 6
        Case Else:   d = 8
        End Select
    Else
        d = 8
    End If
    Scheibendicke = d
End Property
Public Property Get FiktiveBlechstärke() As Double
    FiktiveBlechstärke = Round((2 * m_d) / 10) * 10 / 2
End Property
 
Public Function MinSchraubenlänge() As Double
    'liefert die minimale Schraubenlänge (ohne Kopf)
    'für fiktive Blechstärke abh. v. Schraubendruchmesser
    'bei 12 == 20 '= * 2 Abrunden auf 10
    'bei 16 == 30 ' . . .
    'MinSchraubenlänge = (Round(FiktiveBlechstärke + Mutterhöhe / 10) + 0.5) * 10
    MinSchraubenlänge = Round(((FiktiveBlechstärke + Mutterhöhe) / 10) - 0.5) * 10 + 5
End Function
Public Property Get Schaftquerschnitt() As Double
    'liefert den Schaftquerscnitt A in cm2 in Abhängigkeit ob Passschrauben oder nicht
    Schaftquerschnitt = 4 * Atn(1) * ((m_d + IIf(m_isPass, 1, 0)) / 2 / 10) ^ 2
End Property
Public Property Get Spannungsquerschnitt() As Double
    'liefert den Spannungsquerschnitt A_s in cm2
    Dim sq As Double
    Select Case m_d
    Case 12: sq = 0.843
    Case 16: sq = 1.57
    Case 20: sq = 2.45
    Case 22: sq = 3.03
    Case 24: sq = 3.53
    Case 27: sq = 4.59
    Case 30: sq = 5.61
    Case 36: sq = 8.17
    Case Else: sq = Schaftquerschnitt / 1.25
    End Select
    Spannungsquerschnitt = sq
End Property
Public Property Get A() As Double
    A = Schaftquerschnitt
End Property
Public Property Get A_s() As Double
    A_s = Spannungsquerschnitt
End Property
 
'OK, das sollte eigentlich in Norm rein:
Public Function Grenzabscherkraft() As Double
    'liefert die Grenzabscherkraft in kN je Scherfuge
    Dim AA As Double
    If m_N.ENorm = Norm_DIN18800 Then
        'berechnet nach DIN 18800 die Grenzabscherkraft V_aRd
        'für einschnittige gestützte und mehrschnittige Verbindungen
        'getestet und alle Werte OK
        AA = IIf(IsPassschraube Or IsScherfugeSchaft, A, A_s)
        Dim gam_M As Double: gam_M = m_N.gamM
        Dim alp_a As Double: alp_a = IIf(m_g = 10.9, IIf(Not IsScherfugeSchaft, 0.44, 0.55), 0.6)
        Dim V_aRd As Double: V_aRd = AA * alp_a * fubk / gam_M / 10
        Grenzabscherkraft = V_aRd
    Else
        'berechnet nach Eurocode3 die Grenzabscherkraft F_v_Rd
        'getestet und alle Werte OK
        AA = IIf(IsPassschraube Or IsScherfugeSchaft, A, A_s)
        Dim gam_M2 As Double: gam_M2 = m_N.gamM2
        'bei Passschrauben muss i.d.Regel die Scherfuge im Schaft liegen
        Dim alp_v As Double: alp_v = IIf(IsPassschraube Or IsScherfugeSchaft Or Schraubengüte < 10.9, 0.6, 0.5)
        Dim F_vRd As Double: F_vRd = alp_v * fubk * AA / gam_M2 / 10
        Grenzabscherkraft = F_vRd
    End If
End Function

'OK, das sollte eigentlich in Norm rein:
Public Function Grenzzugkraft() As Double
    'liefert die Grenzzugkraft in kN je Schraube
    If m_N.ENorm = Norm_DIN18800 Then
        'berechnet nach DIN 18800 die Grenzzugkraft N_Rd
        'getestet und alle Werte OK
        Dim gam_M As Double: gam_M = 1.1
        Dim sig_1Rd As Double: sig_1Rd = fybk / (1.1 * gam_M) / 10
        Dim sig_2Rd As Double: sig_2Rd = fubk / (1.25 * gam_M) / 10
        Dim N_Rd As Double: N_Rd = Min(A * sig_1Rd, A_s * sig_2Rd)
        Grenzzugkraft = N_Rd
    Else
        'berechnet nach Eurocode3 die Grenzzugkraft F_t_Rd
        'getestet und alle Werte OK
        Dim k_2 As Double: k_2 = IIf(m_isSenk, 0.63, 0.9)
        Dim gam_M2 As Double: gam_M2 = 1.25
        Dim F_t_Rd As Double: F_t_Rd = k_2 * fubk * A_s / gam_M2 / 10
        Grenzzugkraft = F_t_Rd
    End If
End Function

'hmm ist das nicht Norm-Abhängig?, OK vlt in Din und EC gleich
Public Function Grenzvorspannkraft() As Double
    'liefert die Grenzvorspannkraft F_p,C in kN je Schraube
    Grenzvorspannkraft = 0.7 * fubk * A_s / 10
End Function

'OM 2016_10_24: durch ToStr ersetzt
'Public Sub ToListBox(aLB As ListBox)
'    'Dim isEC As Boolean: isEC = m_N.Norm = Norm_EuroCode3
'    With aLB
'        '.Clear
'        .AddItem "Schraube            : " & " M" & Durchmesser & IIf(m_l > 0, " x " & Länge, "") & " " & Trim(Str(Schraubengüte))
'        .AddItem "  Schraubennorm     : " & SchraubenNorm
'        '.AddItem "Lochdurchmesser     : " & Lochdurchmesser & "[mm]"
'        '.AddItem "Schraubengüte       : " & Str(Schraubengüte)
'        .AddItem "  fyb: " & fybk & "[N/mm2]; fub: " & fubk & "[N/mm2]"
'        .AddItem "  Kategorie " & Kategorie & ", Ausf.form: " & Ausführungsform
'        If m_isPass Then
'            .AddItem "  Passschraube Schaftdurchmesser: " & Schaftdurchmesser & "[mm]"
'        Else
'            .AddItem "  Lage der Scherfuge: " & IIf(IsScherfugeSchaft, "im Schaft", "im Gewinde")
'        End If
'        If IsVorgespannt Then
'            .AddItem "  vorgespannte HV-Schraube " '& Schlüsselweite & "[mm]"
'        End If
'        '.AddItem "Schlüsselweite      : " & Schlüsselweite & "[mm]"
'        '.AddItem "Eckenmass           : " & Round(Eckenmass, 2) & "[mm]"
'        '.AddItem "Eckenmass e_min     : " & EckenmassMin & "[mm]"
'        'If m_isSenk Then
'        '    .AddItem "Senkkopfhöhe        : " & Round(Kopfhöhe, 3) & "[mm]"
'        'Else
'        '    .AddItem "Kopfhöhe            : " & Round(Kopfhöhe, 3) & "[mm]"
'        'End If
'        '.AddItem "  Mutternhöhe       : " & Mutterhöhe & "[mm]"
'        '.AddItem "Scheibendurchmesser : " & Scheibendurchmesser & "[mm]"
'        '.AddItem "  Scheibendicke     : " & Scheibendicke & "[mm]"
'        If Klemmlänge > 0 Then _
'            .AddItem "  Klemmlänge        : " & Klemmlänge & "[mm]"
'        '.AddItem "Fiktive Blechstärke : " & FiktiveBlechstärke & "[mm]"
'        '.AddItem "Min: Schraubenlänge : " & MinSchraubenlänge & "[mm]" 'könnte man auch als Klemmlänge ausgeben
'        '.AddItem "Schaftquerschnitt    A  : " & Round(Schaftquerschnitt, 3) & "[cm2]"
'        '.AddItem "Spannungsquerschnitt A_s: " & Round(Spannungsquerschnitt, 3) & "[cm2]"
'        .AddItem "  Grenzabscherkraft " & m_N.AbkName_Grenzabscherkraft & " = " & _
'            Round(Grenzabscherkraft, 1) & "[kN]"
'        .AddItem "  Grenzzugkraft     " & m_N.AbkName_Grenzzugkraft & " = " & _
'            Round(Grenzzugkraft, 1) & "[kN]"
'        If m_isVorgesp Then
'            .AddItem "Grenzvorspannkraft " & m_N.AbkName_Grenzvorspannkraft & " = " & _
'                Round(Grenzvorspannkraft, 1) & "[kN]"
'        End If
'    End With
'End Sub
'
Friend Function ToStr() As String
    Dim s As String
    s = s & "Schraube            : " & " M" & Me.Durchmesser & IIf(m_l > 0, " x " & Länge, "") & " " & Trim(Str(Me.Schraubengüte)) & vbCrLf
    s = s & "  Schraubennorm     : " & Me.SchraubenNorm & vbCrLf
    '.AddItem "Lochdurchmesser     : " & Lochdurchmesser & "[mm]"
    '.AddItem "Schraubengüte       : " & Str(Schraubengüte)
    s = s & "  fyb: " & fybk & "[N/mm2]; fub: " & fubk & "[N/mm2]" & vbCrLf
    s = s & "  Kategorie " & Kategorie & ", Ausf.form: " & Me.Ausführungsform & vbCrLf
    If m_isPass Then
        s = s & "  Passschraube Schaftdurchmesser: " & Me.Schaftdurchmesser & "[mm]" & vbCrLf
    Else
        s = s & "  Lage der Scherfuge: " & IIf(Me.IsScherfugeSchaft, "im Schaft", "im Gewinde") & vbCrLf
    End If
    If IsVorgespannt Then
        s = s & "  vorgespannte HV-Schraube " & vbCrLf '& Schlüsselweite & "[mm]"
    End If
    '.AddItem "Schlüsselweite      : " & Schlüsselweite & "[mm]"
    '.AddItem "Eckenmass           : " & Round(Eckenmass, 2) & "[mm]"
    '.AddItem "Eckenmass e_min     : " & EckenmassMin & "[mm]"
    'If m_isSenk Then
    '    .AddItem "Senkkopfhöhe        : " & Round(Kopfhöhe, 3) & "[mm]"
    'Else
    '    .AddItem "Kopfhöhe            : " & Round(Kopfhöhe, 3) & "[mm]"
    'End If
    '.AddItem "  Mutternhöhe       : " & Mutterhöhe & "[mm]"
    '.AddItem "Scheibendurchmesser : " & Scheibendurchmesser & "[mm]"
    '.AddItem "  Scheibendicke     : " & Scheibendicke & "[mm]"
    If Klemmlänge > 0 Then _
        s = s & "  Klemmlänge        : " & Me.Klemmlänge & "[mm]" & vbCrLf
    '.AddItem "Fiktive Blechstärke : " & FiktiveBlechstärke & "[mm]"
    '.AddItem "Min: Schraubenlänge : " & MinSchraubenlänge & "[mm]" 'könnte man auch als Klemmlänge ausgeben
    '.AddItem "Schaftquerschnitt    A  : " & Round(Schaftquerschnitt, 3) & "[cm2]"
    '.AddItem "Spannungsquerschnitt A_s: " & Round(Spannungsquerschnitt, 3) & "[cm2]"
    s = s & "  Grenzabscherkraft " & m_N.AbkNameGrenzabscherkraft & " = " & Round(Me.Grenzabscherkraft, 1) & "[kN]" & vbCrLf
    s = s & "  Grenzzugkraft     " & m_N.AbkNameGrenzzugkraft & " = " & Round(Me.Grenzzugkraft, 1) & "[kN]" & vbCrLf
    If m_isVorgesp Then
        s = s & "Grenzvorspannkraft " & m_N.AbkNameGrenzvorspannkraft & " = " & Round(Grenzvorspannkraft, 1) & "[kN]" & vbCrLf
    End If
    ToStr = s
End Function


