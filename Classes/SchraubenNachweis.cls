VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SchraubenNachweis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'SL Schraubennachweis auf Scher-Lochleibung
'Abscheren,   Grenzabscherkraft
'Lochleibung, Grenzlochleibungskraft
'
Private m_N As Norm
Private m_s As Schraube
Private m_sg As Schraubengruppe

Friend Sub New_(ByVal aCalcNorm As Norm, s As Schraube, sg As Schraubengruppe)
    Set m_N = aCalcNorm: Set m_s = s: Set m_sg = sg
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

Public Property Get Schraube() As Schraube
    Set Schraube = m_s
End Property
Public Property Set Schraube(ByVal Value As Schraube)
    Set m_s = Value
End Property

Public Property Get Schraubengruppe()
    Set Schraubengruppe = m_sg
End Property
Public Property Set Schraubengruppe(ByVal Value As Schraubengruppe)
    Set m_sg = Value
End Property
' ######################################## ' ^   ISerializable   ^ ' ######################################## '

Public Property Get EtaAbscheren() As Double
    Dim mgask As Double: mgask = m_sg.MinGrenzabscherkraft
    'Halt STopp für eine Schraube, aber weil zweischnittig -> * 2
    If mgask <> 0 Then
        If Not m_sg.TrägerLinks Is Nothing Then
            EtaAbscheren = m_sg.Rmax / (mgask * IIf(m_sg.TrägerLinks.IsMehrschnittig, 2, 1))
        End If
    End If
End Property

Public Property Get EtaLochleibung() As Double
    Dim mgllk As Double: mgllk = m_sg.MinGrenzlochleibungskraft
    If mgllk <> 0 Then _
        EtaLochleibung = m_sg.Rmax / mgllk
End Property

'OM 2016_10_24: durch ToStr ersetzt
'Public Sub ToListBox(aLB As ListBox)
'    'der SL-Nachweis der Verbindung
'    'Abscheren
'    Dim sF As String
'    With aLB
'        If Not m_sg Is Nothing Then
'            If m_sg.Rmax <> 0 Then
'                If EtaAbscheren <> 0 Then
'                    sF = m_N.AbkName_Grenzabscherkraft
'                    .AddItem "Nachweis auf Abscheren"
'                    .AddItem "  eta_A = max_R_i / " & sF & " = " & Round(EtaAbscheren, 4)
'                End If
'                If EtaLochleibung <> 0 Then
'                    sF = m_N.AbkName_Grenzlochleibungskraft
'                    .AddItem "Nachweis auf Lochleibung"
'                    .AddItem "  eta_L = max_R_i / " & sF & " = " & Round(EtaLochleibung, 4)
'                    'hier soll auch der alpha_b-Wert und k_1 ausgegeben werden.
'                    If m_N.Norm = Norm_EuroCode3 Then
'                        .AddItem "  mit alp = " & Round(m_sg.Alpha, 4) & " und k1 = " & Round(m_sg.k1, 4)
'                    End If
'                End If
'            End If
'        End If
'    End With
'End Sub

Friend Function ToStr() As String
    'der SL-Nachweis der Verbindung
    'Abscheren
    Dim s As String
    Dim sF As String
    If m_sg Is Nothing Then Exit Function
    'If Not m_sg Is Nothing Then
    If m_sg.Rmax = 0 Then Exit Function
    'If m_sg.Rmax <> 0 Then
    If EtaAbscheren <> 0 Then
        sF = m_N.AbkNameGrenzabscherkraft
        s = s & "Nachweis auf Abscheren" & vbCrLf
        s = s & "  eta_A = max_R_i / " & sF & " = " & Round(EtaAbscheren, 4) & vbCrLf
    End If
    If EtaLochleibung <> 0 Then
        sF = m_N.AbkNameGrenzlochleibungskraft
        s = s & "Nachweis auf Lochleibung" & vbCrLf
        s = s & "  eta_L = max_R_i / " & sF & " = " & Round(EtaLochleibung, 4) & vbCrLf
        'hier soll auch der alpha_b-Wert und k_1 ausgegeben werden.
        If m_N.ENorm = Norm_EuroCode3 Then
            s = s & "  mit alp = " & Round(m_sg.Alpha, 4) & " und k1 = " & Round(m_sg.k1, 4) & vbCrLf
        End If
    End If
    'End If
    'End If
    ToStr = s
End Function

