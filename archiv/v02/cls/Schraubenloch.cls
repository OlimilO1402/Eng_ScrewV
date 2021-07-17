VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Schraubenloch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_la As ELochart
Private m_s  As Schraube
Private m_IsVert As Boolean 'True=vertikal oder False=horizontal 'nur für Langloch
Private m_mue As Double

Public Sub New_(ByVal s As Schraube, ByVal la As ELochart, Optional ByVal isVert As Boolean = False)
    Set m_s = s
    If la < 0 Then la = 0
    m_la = la
    m_IsVert = isVert
    m_mue = 0.5
End Sub
Public Property Get Lochart() As ELochart
    Lochart = m_la
End Property
Public Property Let Lochart(ByVal Value As ELochart)
    m_la = Value
End Property

Public Property Get Schraube() As Schraube
    Set Schraube = m_s
End Property

Public Property Get IsVertikal() As Boolean
    IsVertikal = m_IsVert
End Property
Public Property Let IsVertikal(ByVal Value As Boolean)
    m_IsVert = Value
End Property


Public Property Get NennlochspielSenkr() As Double
    NennlochspielSenkr = nls(True)
End Property
Public Property Get Nennlochspiel() As Double
    Nennlochspiel = nls
End Property
Private Function nls(Optional ByVal getSenkr As Boolean = False) As Double
    'liefert das Nennlochspiel für den Schraubendurchmesser
    'wenn getSenk=True dann liefert Nls wie für Normalloch
    'd.h. bei Langloch ist der Lochdurchmeser Senkr zur Lochrichtung wie bei Normalloch
    Dim la As ELochart: la = m_la
    If Not (la = Übergroß) And getSenkr Then la = Normal
    If m_s.IsPassschraube Then
        nls = 0
        Exit Function
    End If
    Dim d As Double: d = m_s.Durchmesser
    Select Case la
    Case Normal
        Select Case True
        Case d <= 12:                nls = 1
        Case (12 < d) And (d <= 24): nls = 2
        Case (24 < d):               nls = 3
        End Select
    Case Übergroß
        Select Case True
        Case (d <= 12):              nls = 3
        Case (12 < d) And (d <= 22): nls = 4
        Case (22 < d) And (d <= 24): nls = 6
        Case (24 < d):               nls = 8
        End Select
    Case LanglochKurz
        Select Case True
        Case (d <= 12):              nls = 4
        Case (12 < d) And (d <= 22): nls = 6
        Case (22 < d) And (d <= 24): nls = 8
        Case (24 < d):               nls = 10
        End Select
    Case LanglochLang
        Select Case True
        Case d <= 12:                nls = 18
        Case (12 < d) And (d <= 16): nls = 24
        Case (16 < d) And (d <= 20): nls = 30
        Case (20 < d) And (d <= 22): nls = 33
        Case (22 < d) And (d <= 24): nls = 36
        Case (24 < d) And (d <= 27): nls = 40.5
        Case (27 < d) And (d <= 30): nls = 45
        Case (30 < d):               nls = 54
        End Select
    End Select
End Function

Public Property Get Durchmesser() As Double
    Durchmesser = m_s.Durchmesser + Nennlochspiel
End Property
Public Property Get DurchmesserSenkr() As Double
    DurchmesserSenkr = m_s.Durchmesser + NennlochspielSenkr
End Property

Public Property Get ks() As Double
    Select Case m_la
    Case Normal:       ks = 1#
    Case Übergroß:     ks = 0.85
    Case LanglochKurz: ks = IIf(IsVertikal, 0.85, 0.76) 'Vertikal heißt hier Längsachse ist quer zur Kraftrichtung
    Case LanglochLang: ks = IIf(IsVertikal, 0.7, 0.63) 'Vertikal heißt hier Längsachse ist quer zur Kraftrichtung
    End Select
End Property
Public Property Get Gleitflächenklasse() As String
    Dim GFK As String
    If m_mue >= 0.2 Then GFK = "D"
    If m_mue >= 0.3 Then GFK = "C"
    If m_mue >= 0.4 Then GFK = "B"
    If m_mue >= 0.5 Then GFK = "A"
    Gleitflächenklasse = GFK
End Property
Public Property Get Reibzahl_mue() As Double
    Reibzahl_mue = m_mue
End Property
Public Property Let Reibzahl_mue(ByVal Value As Double)
    m_mue = Value
End Property
Public Function Grenzgleitkraft() As Double
    'getestet alle Werte richtig berechnet
    'liefert die Grenzgleitkraft F_s,Rd(,ser) in kN je Gleitfuge für mue=0.5
    Dim mue As Double: mue = 0.5 'Reibzahl mue
    'Abhängig von Oberfläche des Trägers/Blech
    'Gleitflächenklasse A,B,C,D nach Bezugsnormengruppe 7 [-1-8/Tab. 3.7]
    'GFK_A : mue=0.5
    'GFK_B : mue=0.4
    'GFK_C : mue=0.3
    'GFK_D : mue=0.2
    Dim n As Double: n = 1 'Anzahl der Gleitfugen
    Dim gam As Double: gam = IIf(m_s.Kategorie = "C", m_s.Norm.gam_M3, m_s.Norm.gam_M3_ser)
    Grenzgleitkraft = ks * n * mue * m_s.Grenzvorspannkraft / gam
End Function

Public Sub ToListBox(aLB As ListBox)
    Dim isEC As Boolean: isEC = m_s.Norm.Norm = Norm_EuroCode3
    With aLB
        '.Clear
        .AddItem "Schraubenloch"
        .AddItem "  Lochdurchmesser    : " & Durchmesser & "[mm]"
        .AddItem "  Nennlochspiel      : " & Nennlochspiel & "[mm]"
        If Lochart >= LanglochKurz Then _
            .AddItem "  Nennlochspiel senkr: " & NennlochspielSenkr & "[mm]"
        .AddItem "  Lochart            : " & ELochart_ToStr(m_la)
        If Lochart >= LanglochKurz Then _
            .AddItem "  Richtung           : " & IIf(IsVertikal, "Vertikal", "Horizontal") & "[mm]"
        If m_s.IsGleitfest Then
            .AddItem "  Grenzgleitkraft " & _
                IIf(isEC, "F_s,Rd", "V_g,Rd") & " = " & Round(Grenzgleitkraft, 1) & "[kN]"
            .AddItem "  Gleitflächenklasse GFK: " & Gleitflächenklasse & " mue=" & m_mue
            
        End If
    End With
End Sub
