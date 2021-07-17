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
Public Enum Lochart
    Normal
    Übergroß
    LanglochKurz
    LanglochLang
End Enum
Private m_la As Lochart
Private m_S  As Schraube
Private m_IsVert As Boolean 'True=vertikal oder False=horizontal 'nur für Langloch


Public Sub New_(ByVal s As Schraube, ByVal la As Lochart, ByVal isVert As Boolean)
    Set m_S = s
    If la < 0 Then la = 0
    m_la = la
    m_IsVert = isVert
End Sub
Public Property Get Lochart() As Lochart
    Lochart = m_la
End Property
Public Property Let Lochart(ByVal Value As Lochart)
    m_la = Value
End Property

Public Property Get Schraube() As Schraube
    Set Schraube = m_S
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
Private Function Lochart_ToStr() As String
    Dim s As String
    Select Case m_la
    Case Normal:       s = "Normal"
    Case Übergroß:     s = "Übergroß"
    Case LanglochKurz: s = "LanglochKurz"
    Case LanglochLang: s = "LanglochLang"
    End Select
    Lochart_ToStr = s
End Function
Private Function nls(Optional ByVal getSenkr As Boolean = False) As Double
    'liefert das Nennlochspiel für den Schraubendurchmesser
    'wenn getSenk=True dann liefert Nls wie für Normalloch
    'd.h. bei Langloch ist der Lochdurchmeser Senkr zur Lochrichtung wie bei Normalloch
    Dim la As Lochart: la = m_la
    If Not (la = Übergroß) And getSenkr Then la = Normal
    Dim d As Double: d = m_S.Durchmesser
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
    Durchmesser = m_S.Durchmesser + Nennlochspiel
End Property

Public Sub ToListBox(aLB As ListBox)
    With aLB
        '.Clear
        .AddItem "Lochdurchmesser     : " & Durchmesser & "[mm]"
        .AddItem "Nennlochspiel       : " & Nennlochspiel & "[mm]"
        .AddItem "Nennlochspiel senkr : " & NennlochspielSenkr & "[mm]"
        .AddItem "Lochart             : " & Lochart_ToStr
        If Lochart >= LanglochKurz Then _
            .AddItem "Richtung            : " & IIf(IsVertikal, "Vertikal", "Horizontal") & "[mm]"
    End With
End Sub
