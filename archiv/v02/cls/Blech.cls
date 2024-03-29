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
Private m_isLinks As Boolean 'True: Links; False: Rechts
Private m_sso As EStahlsorte  ' aka Stahlg�te z.B. S235, ST37
Private m_t   As Double   ' Dicke t in mm 4,6,8,10,12,14,16,18,20,22,24,26,28,30
                        ' eine gerade Zahl
Private m_lx As Double  ' Abmessungen L�nge in x-Richtung
Private m_hz As Double  ' Abmessungen H�he  in z-Richtung
' hmm naja was jetzt x, y und z ist, ist fraglich
' wie macht man das mit den angreifenden Kr�ften
' sollte man sich irgendwie von x,y,z losl�sen?

Private m_isZange As Boolean
Private m_isMehrschnittig As Boolean
' jedes Blech braucht Zugriff auf die Schraubengruppe
' und jedes Blech hat die eigenen L�cher in der Anzahl der Schrauben in der Schraubengruppe
'oder sollte man die Schraubengruppe mit Zugriff aufs Blech machen????
'Private m_sg  As Schraubengruppe
Public Event Changed(sender As Blech)

Public Sub New_(ByVal aNorm As Norm, ByVal sso As EStahlsorte, ByVal t As Double, ByVal l_x As Double, ByVal h_z As Double, _
                ByVal bIsLinks As Boolean, ByVal bIsZange As Boolean, ByVal bIsMehrschnittig As Boolean)
    Set m_N = aNorm
    'Set m_sg = sgByVal sg As Schraubengruppe,
    If sso < 0 Then sso = S235
    m_sso = sso
    m_t = t
    m_lx = l_x
    m_hz = h_z
    m_isLinks = bIsLinks
    m_isZange = bIsZange
    m_isMehrschnittig = bIsMehrschnittig
End Sub

Public Property Get Blechdicke() As Double
    Blechdicke = m_t
End Property
Public Property Let Blechdicke(ByVal Value As Double)
    m_t = Value
    RaiseEvent Changed(Me)
End Property
Public Property Get GesamtT() As Double
    GesamtT = m_t * IIf(IsZange, 2, 1)
End Property
Public Property Get L�nge() As Double
    L�nge = m_lx
End Property
Public Property Let L�nge(ByVal Value As Double)
    m_lx = Value
End Property
Public Property Get H�he() As Double
    H�he = m_hz
End Property
Public Property Let H�he(ByVal Value As Double)
    m_hz = Value
End Property
Public Property Get IsZange() As Boolean
    IsZange = m_isZange
End Property
Public Property Let IsZange(ByVal Value As Boolean)
    m_isZange = Value
    RaiseEvent Changed(Me)
End Property
Public Property Get IsLinks() As Boolean
    IsLinks = m_isLinks
End Property
Public Property Let IsLinks(ByVal Value As Boolean)
    m_isLinks = Value
End Property
Public Property Get IsMehrschnittig() As Boolean
    IsMehrschnittig = m_isMehrschnittig
End Property
Public Property Let IsMehrschnittig(ByVal Value As Boolean)
    m_isMehrschnittig = Value
    RaiseEvent Changed(Me)
End Property
Public Function SeiteToStr() As String
    SeiteToStr = IIf(m_isLinks, "Links", "Rechts")
End Function
'Public Property Get Schraubengruppe() As Schraubengruppe
'    Set Schraubengruppe = m_sg
'End Property

Public Property Get Stahlsorte() As EStahlsorte
    'liefert die Stahlsorte als Enumkonstante
    Stahlsorte = m_sso
End Property
Public Property Let Stahlsorte(ByVal Value As EStahlsorte)
    'legt die Stahlsorte als Enumkonstante fest
    m_sso = Value
End Property
Public Property Get fy() As Double
'ist das nicht auch Norm-abh�ngig
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
Public Property Get fyk() As Double
    'If m_N.Norm = Norm_DIN18800 Then
        Select Case m_sso
        Case S235: fyk = 240
        Case S275: fyk = 280
        Case S355: fyk = 360
        Case Else: fyk = 240
        End Select
End Property
Public Property Get fu() As Double
'ist das nicht auch Norm-abh�ngig
    If m_t <= 40 Then
        Select Case m_sso
        Case S235: fu = 360
        Case S275: fu = 430
        Case S355: fu = 490
        Case S420: fu = 520
        Case S460: fu = 540
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
Public Property Get Stahlg�te() As Double
    'liefert die Stahlg�te als Zahl in N/mm2
    Stahlg�te = fy
End Property

'Maximale Abst�nde volle Tragkraft
Public Function MaxE1() As Double 'AbstandRandX
    MaxE1 = RoundDown(4 * m_t + 40)
End Function
Public Function MaxE2() As Double 'AbstandRandZ
    MaxE2 = RoundDown(4 * m_t + 40)
End Function
Public Function MaxP1() As Double 'AbstandLochX aka e_0
    MaxP1 = RoundDown(Min(14 * m_t, 200))
End Function
Public Function MaxP2() As Double 'AbstandLochZ aka e_3
    MaxP2 = RoundDown(Min(14 * m_t, 200))
End Function


Public Sub ToListBox(aLB As ListBox)
'Try: On Error GoTo Catch
    Dim isEC As Boolean: isEC = m_N.Norm = Norm_EuroCode3
    With aLB
        '.Clear
        .AddItem "Tr�ger " & SeiteToStr
        .AddItem "  Stahlsorte          : " & EStahlsorte_ToStr(m_sso, m_N.Norm) & "[N/mm2]"
        .AddItem "  Blechdicke          : " & Blechdicke & "[mm]"
        If IsZange Then .AddItem "  Gesamtdicke         : " & GesamtT & "[mm]"
        '.AddItem "L�nge               : " & L�nge & "[mm]"
        .AddItem "  H�he                : " & H�he & "[mm]"
        
'mmh Blech hat schon die Schraubengruppe, dann kann Sg ja nicht auch gleichzeitig Blech haben
'hmm Pech gehabt
'f�r was braucht Blech die Schraubengruppe?
'sollte man es umdrehen?
    End With
    Exit Sub
Catch:
    If Err.Number = 6 Then Resume Next '�berlauf wg unendlich
End Sub

