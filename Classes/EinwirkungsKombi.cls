VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EinwirkungsKombi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Verflixt ich weiﬂ nicht mehr wie Schraubengruppe, die Vektor-Klasse VectorXZ der SchraubenNachweis und die Klasse einwirkungskombi zusammen h‰ngen
'wer macht was?
'Private m_sg   As Schraubengruppe
Private m_M_Ed As Double 'Moment      design
Private m_V_Ed As Double 'Querkraft   design
Private m_N_Ed As Double 'Normalkraft design

'Offset der Einwirkungsgruppe zum Schwerpunkt der Schraubengruppe
Private m_offX As Double
Private m_offZ As Double
'
'sollte hier ein Changed-Event vorhanden sein?
'OM_2016_12_05: OK jetzt neu, hier Changed Event
Public Event Changed(sender As EinwirkungsKombi)

Friend Sub New_(ByVal m_Ed As Double, ByVal V_Ed As Double, ByVal N_Ed As Double, ByVal OffX As Double, ByVal OffZ As Double)
             m_M_Ed = m_Ed:        m_V_Ed = V_Ed:        m_N_Ed = N_Ed:        m_offX = OffX:        m_offZ = OffZ
     'Set m_sg = sg 'sg As Schraubengruppe,
     'If Not m_sg Is Nothing Then m_sg.SetVxVz Me
End Sub

' ######################################## ' v   ISerializable   v ' ######################################## '
Public Property Get ID() As Long
    ID = ObjPtr(Me)
End Property
Public Property Let ID(ByVal Value As Long)
    'nix
End Property

Public Property Get MomentEd() As Double
    MomentEd = m_M_Ed
End Property
Public Property Let MomentEd(ByVal Value As Double)
    m_M_Ed = Value
    'm_sg.SetVxVz Me
    RaiseEvent Changed(Me)
End Property

Public Property Get QuerkraftEd() As Double
    QuerkraftEd = m_V_Ed
End Property
Public Property Let QuerkraftEd(ByVal Value As Double)
    m_V_Ed = Value
    RaiseEvent Changed(Me)
    'm_sg.SetVxVz Me
End Property

Public Property Get NormalkraftEd() As Double
    NormalkraftEd = m_N_Ed
End Property
Public Property Let NormalkraftEd(ByVal Value As Double)
    m_N_Ed = Value
    RaiseEvent Changed(Me)
    'm_sg.SetVxVz Me
End Property

Public Property Get OffX() As Double
    OffX = m_offX
End Property
Public Property Let OffX(ByVal Value As Double)
    m_offX = Value
    RaiseEvent Changed(Me)
    'm_sg.SetVxVz Me
End Property

Public Property Get OffZ() As Double
    OffZ = m_offZ
End Property
Public Property Let OffZ(ByVal Value As Double)
    m_offZ = Value
    RaiseEvent Changed(Me)
    'm_sg.SetVxVz Me
End Property
' ######################################## ' ^   ISerializable   ^ ' ######################################## '

Public Property Get MomentS() As Double
    'liefert das Moment bezogen auf den Schwerpunkt der Schrauben
    MomentS = MomentEd + QuerkraftEd * OffX / 1000 + NormalkraftEd * OffZ / 1000
End Property
Public Property Get QuerkraftS() As Double
    QuerkraftS = QuerkraftEd
End Property
Public Property Get NormalkraftS() As Double
    NormalkraftS = NormalkraftEd
End Property

'Public Property Get MaxVxi() As Double
''    Dim M_S As Double: M_S = Me.MomentS
''    Dim N_S As Double: N_S = Me.NormalkraftS
''    Dim n   As Double:   n = m_sg.Anzahl
''    Dim b   As Double:   b = m_sg.BRectSelInnW / 1000 'in m
''    Dim h   As Double:   h = m_sg.BRectSelInnH / 1000 'in m
''    Dim alp As Double: alp = m_sg.alpha
''    Dim bet As Double: bet = m_sg.beta
''    If n <> 0 And h <> 0 Then _
''        MaxVxi = N_S / n + M_S / h * alp / (1 + bet * (b / h) ^ 2)
'    MaxVxi = m_sg.Vmax.X
'End Property
'Public Property Get MaxVzi() As Double
''    Dim M_S As Double: M_S = Me.MomentS
''    Dim V_S As Double: V_S = Me.QuerkraftS
''    Dim n   As Double:   n = m_sg.Anzahl
''    Dim b   As Double:   b = m_sg.BRectSelInnW / 1000 'in m
''    Dim h   As Double:   h = m_sg.BRectSelInnH / 1000 'in m
''    Dim alp As Double: alp = m_sg.alpha
''    Dim bet As Double: bet = m_sg.beta
''    If n <> 0 And h <> 0 Then _
''        MaxVzi = V_S / n + M_S / h * alp / (1 + bet * (b / h) ^ 2) * (b / h)
'    MaxVzi = m_sg.Vmax.Z
'End Property
'Public Property Get MaxRi() As Double
''    MaxRi = VBA.Math.Sqr(MaxVxi ^ 2 + MaxVzi ^ 2)
'    MaxRi = m_sg.Rmax
'End Property
'
'OM 2016_10_24: durch ToStr ersetzt
'Public Sub ToListBox(aLB As ListBox)
'    'If Not m_sg Is Nothing Then m_sg.SetVxVz Me
'    With aLB
'        If Me.MomentS <> 0 Or Me.QuerkraftS <> 0 Or Me.NormalkraftS <> 0 Then
'            .AddItem "Einwirkungen"
'            If Me.MomentEd <> 0 Then _
'                .AddItem "  Moment Ed     : " & MomentEd & " kNm"
'            If Me.QuerkraftEd <> 0 Then _
'                .AddItem "  Querkraft Ed  : " & QuerkraftEd & " kN"
'            If Me.NormalkraftEd <> 0 Then _
'                .AddItem "  Normalkraft Ed: " & NormalkraftEd & " kN"
'            'If Not m_sg Is Nothing Then
'            '    Dim v As VectorXZ: Set v = m_sg.Vmax
'            If OffX <> 0 Or OffZ <> 0 Then
'                .AddItem "  Offset X: " & OffX & "mm" & "; Offset Z: " & OffZ & "mm"
'                If MomentS <> 0 Then _
'                    .AddItem "  Moment Sp     : " & MomentS & " kNm"
'                If QuerkraftS <> 0 Then _
'                    .AddItem "  Querkraft Sp  : " & QuerkraftS & " kN"
'                If NormalkraftS <> 0 Then _
'                    .AddItem "  Normalkraft Sp: " & NormalkraftS & " kN"
'            End If
'                '.AddItem "max V_xi : " & Round(v.X, 2) & " kN"
'                '.AddItem "max V_zi : " & Round(v.Z, 2) & " kN"
'                '.AddItem "max Ri   : " & Round(v.Length, 2) & " kN"
'            'End If
'        End If
'    End With
'End Sub

Friend Function ToStr() As String
    Dim s As String
    If Me.MomentS <> 0 Or Me.QuerkraftS <> 0 Or Me.NormalkraftS <> 0 Then
        s = s & "Einwirkungen" & vbCrLf
        If Me.MomentEd <> 0 Then _
            s = s & "  Moment Ed     : " & MomentEd & " kNm" & vbCrLf
        If Me.QuerkraftEd <> 0 Then _
            s = s & "  Querkraft Ed  : " & QuerkraftEd & " kN" & vbCrLf
        If Me.NormalkraftEd <> 0 Then _
            s = s & "  Normalkraft Ed: " & NormalkraftEd & " kN" & vbCrLf
        'If Not m_sg Is Nothing Then
        '    Dim v As VectorXZ: Set v = m_sg.Vmax
        If OffX <> 0 Or OffZ <> 0 Then
            s = s & "  Offset X: " & OffX & "mm" & "; Offset Z: " & OffZ & "mm" & vbCrLf
            If MomentS <> 0 Then _
                s = s & "  Moment Sp     : " & MomentS & " kNm" & vbCrLf
            If QuerkraftS <> 0 Then _
                s = s & "  Querkraft Sp  : " & QuerkraftS & " kN" & vbCrLf
            If NormalkraftS <> 0 Then _
                s = s & "  Normalkraft Sp: " & NormalkraftS & " kN" & vbCrLf
        End If
            '.AddItem "max V_xi : " & Round(v.X, 2) & " kN"
            '.AddItem "max V_zi : " & Round(v.Z, 2) & " kN"
            '.AddItem "max Ri   : " & Round(v.Length, 2) & " kN"
        'End If
    End If
    ToStr = s
End Function


