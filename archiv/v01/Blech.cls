VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Blech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_N   As Norm
Private m_sso As Stahlsorte  ' aka Stahlgüte z.B. S235, ST37
Private m_t   As Double   ' Dicke t in mm 4,6,8,10,12,14,16,18,20,22,24,26,28,30
                        ' eine gerade Zahl
Private m_lx As Double  ' Abmessungen Länge in x-Richtung
Private m_hz As Double  ' Abmessungen Höhe  in z-Richtung
' hmm naja was jetzt x, y und z ist, ist fraglich
' wie macht man das mit den angreifenden Kräften
' sollte man sich irgendwie von x,y,z loslösen?
'
' jedes Blech braucht Zugriff auf die Schraubengruppe
' und jedes Blech hat die eigenen Löcher in der Anzahl der Schrauben in der Schraubengruppe
Private m_SG  As Schraubengruppe

Public Sub New_(ByVal aNorm As Norm, ByVal SG As Schraubengruppe, ByVal sso As Stahlsorte, ByVal t As Double, ByVal l_x As Double, ByVal h_z As Double)
    Set m_N = aNorm
    Set m_SG = SG
    If sso < 0 Then sso = S235
    m_sso = sso
    m_t = t
    m_lx = l_x
    m_hz = h_z
End Sub

Public Property Get Blechdicke() As Double
    Blechdicke = m_t
End Property
Public Property Let Blechdicke(ByVal Value As Double)
    m_t = Value
End Property
Public Property Get Länge() As Double
    Länge = m_lx
End Property
Public Property Let Länge(ByVal Value As Double)
    m_lx = Value
End Property
Public Property Get Höhe() As Double
    Höhe = m_hz
End Property
Public Property Let Höhe(ByVal Value As Double)
    m_hz = Value
End Property

Public Property Get Schraubengruppe() As Schraubengruppe
    Set Schraubengruppe = m_SG
End Property

Public Property Get Stahlsorte() As Stahlsorte
    'liefert die Stahlsorte als Enumkonstante
    Stahlsorte = m_sso
End Property
Public Property Let Stahlsorte(ByVal Value As Stahlsorte)
    'liefert die Stahlsorte als Enumkonstante
    m_sso = Value
End Property
Public Function StahlsorteToStr() As String
    Dim s As String
    If m_N.Norm = Norm_DIN18800 Then
        Select Case m_sso
        Case S235: s = "ST37"
        Case S275: s = "ST42"
        Case S355: s = "ST52"
        End Select
    Else
        Select Case m_sso
        Case S235: s = "S235"
        Case S275: s = "S275"
        Case S355: s = "S355"
        End Select
    End If
    StahlsorteToStr = s
End Function
Public Property Get fy() As Double
'ist das nicht auch Norm-abhängig
    If m_t <= 40 Then
        Select Case m_sso
        Case S235: fy = 235
        Case S275: fy = 275
        Case S355: fy = 355
        Case Else: fy = 235
        End Select
    ElseIf m_t <= 80 Then
        Select Case m_sso
        Case S235: fy = 215
        Case S275: fy = 255
        Case S355: fy = 335
        Case Else: fy = 215
        End Select
    Else 'm_t > 80
        fy = 215
    End If
End Property
Private Property Get fyk() As Double
    'If m_N.Norm = Norm_DIN18800 Then
        Select Case m_sso
        Case S235: fyk = 240
        Case S275: fyk = 280
        Case S355: fyk = 360
        Case Else: fyk = 240
        End Select
End Property
Public Property Get fu() As Double
'ist das nicht auch Norm-abhängig
    If m_t <= 40 Then
        Select Case m_sso
        Case S235: fu = 360
        Case S275: fu = 430
        Case S355: fu = 490
        Case Else: fu = 360
        End Select
    ElseIf m_t <= 80 Then
        Select Case m_sso
        Case S235: fu = 360
        Case S275: fu = 410
        Case S355: fu = 490
        Case Else: fu = 360
        End Select
    Else
        fu = 235
    End If
End Property
Public Property Get Stahlgüte() As Double
    'liefert die Stahlgüte als Zahl in N/mm2
    Stahlgüte = fy
End Property

Public Function GrenzlochleibungskraftRA() As Double 'RA=RandAbstand
    'liefert die Grenzlochleibungskraft in kN
    'V_l_Rd = t * d_Sch * alpha_l * f_yk / gam_M
    Dim d_S As Double:     d_S = m_SG.Schraube.Durchmesser
    Dim d_L As Double:     d_L = m_SG.Schraubenloch(0).Durchmesser
    Dim e_0 As Double:     e_0 = m_SG.AbstandLochX
    Dim e_1 As Double:     e_1 = m_SG.AbstandRandX
    Dim e_2 As Double:     e_2 = m_SG.AbstandRandZ
    Dim e_3 As Double:     e_3 = m_SG.AbstandLochZ
    If m_N.Norm = Norm_DIN18800 Then
        Dim f_yk As Double:   f_yk = fyk
        Dim alp_l As Double
        Dim gam_M As Double: gam_M = m_N.gam_M
        
        'Randabstand in Kraftrichtung
        If (e_2 >= 1.5 * d_L) And (e_3 >= 3# * d_L) Then
            alp_l = 1.1 * Min(e_1, 3 * d_L) / d_L - 0.3
        Else
            alp_l = 0.73 * Min(e_1, 3 * d_L) / d_L - 0.2
        End If
        GrenzlochleibungskraftRA = m_t * d_S * alp_l * f_yk / gam_M / 1000
    Else
        
    End If
End Function
Public Function GrenzlochleibungskraftLA() As Double 'LA=LochAbstand
    'liefert die Grenzlochleibungskraft in kN
    'V_l_Rd = t * d_Sch * alpha_l * f_yk / gam_M
    Dim d_S As Double:     d_S = m_SG.Schraube.Durchmesser
    Dim d_L As Double:     d_L = m_SG.Schraubenloch(0).Durchmesser
    Dim e_0 As Double:     e_0 = m_SG.AbstandLochX
    Dim e_1 As Double:     e_1 = m_SG.AbstandRandX
    Dim e_2 As Double:     e_2 = m_SG.AbstandRandZ
    Dim e_3 As Double:     e_3 = m_SG.AbstandLochZ
    If m_N.Norm = Norm_DIN18800 Then
        Dim f_yk As Double:   f_yk = fyk
        Dim alp_l As Double
        Dim gam_M As Double: gam_M = m_N.gam_M
                
        'Lochabstand in Kraftrichtung
        If (e_2 >= 1.5 * d_L) And (e_3 >= 3# * d_L) Then
            alp_l = 1.08 * Min(e_0, 3.5 * d_L) / d_L - 0.77
        Else
            alp_l = 0.72 * Min(e_0, 3.5 * d_L) / d_L - 0.51
        End If
        GrenzlochleibungskraftLA = m_t * d_S * alp_l * f_yk / gam_M / 1000
    Else
        
    End If
End Function

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
    'liefert die Grenzdurchstanzkraft B_pRd in kN je Schraube bezogen auf 10mm Blechdicke
    'nur für vorgespannte Schrauben
    'berechnet die
    Dim d_m As Double:       d_m = m_SG.Schraube.Kopfdurchmesser
    Dim gam_M2 As Double: gam_M2 = 1.25
    Dim B_pRd As Double:   B_pRd = 0.6 * 4 * Atn(1) * d_m * m_t * fu / gam_M2 / 1000
    Grenzdurchstanzkraft = B_pRd
End Function


Public Sub ToListBox(aLB As ListBox)
    With aLB
        '.Clear
        .AddItem "Stahlsorte          : " & StahlsorteToStr & "[N/mm2]"
        .AddItem "Blechdicke          : " & Blechdicke & "[mm]"
        '.AddItem "Länge               : " & Länge & "[mm]"
        .AddItem "Höhe                : " & Höhe & "[mm]"
        .AddItem "GrenzlochleibungskraftLA: " & Round(GrenzlochleibungskraftLA, 2) & "[kN]"
        .AddItem "GrenzlochleibungskraftRA: " & Round(GrenzlochleibungskraftRA, 2) & "[kN]"
        .AddItem "Grenzdurchstanzkraft: " & Round(Grenzdurchstanzkraft, 1) & "[kN]"
    End With
End Sub

