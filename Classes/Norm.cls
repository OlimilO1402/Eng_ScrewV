VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Norm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'hier schauts nach abstrakte Basisklasse aus, aber in wirklichkeit ist es die Manager-Klasse für die beiden Klassen NormDIN1800 und NormEuroCode3
'OK das Enum ENorm ist bald passee 'halt noch nicht so schnell!!
'Private m_N As ENorm
Private m_N As Norm 'entweder NormDIN18800 oder NormEuroCode3
'Private m_is_k As Boolean

'Friend Sub New_(ByVal aNormBerechnung As ENorm)
'    m_N = aNormBerechnung
'End Sub
Friend Sub New_(ByVal aNormBerechnung As Norm)
    Set m_N = aNormBerechnung
End Sub

' ######################################## ' v   ISerializable   v ' ######################################## '
Public Property Get ID() As Long
    ID = ObjPtr(Me)
End Property
Public Property Let ID(ByVal Value As Long)
    'nix
End Property

'Public Property Get Norm() As ENorm
'    Norm = m_N
'End Property
'Public Property Let Norm(ByVal Value As ENorm)
'    m_N = Value
'End Property
'wird für die Deserialisierung gebraucht:
Public Property Get Norm() As Norm
    Set Norm = m_N
End Property
Public Property Set Norm(ByVal Value As Norm)
    Set m_N = Value
End Property

'Public Property Get NormStr() As String
'    NormStr = MENorm.ENorm_ToStr(m_N)
'End Property
'Public Property Let NormStr(ByVal Value As String)
'    m_N = MENorm.ENorm_Parse(Value)
'End Property

Public Property Get IsCharacteristic() As Boolean
    'IsCharacteristic = m_is_k
    IsCharacteristic = m_N.IsCharacteristic
End Property
Public Property Let IsCharacteristic(ByVal Value As Boolean)
    'm_is_k = Value
    m_N.IsCharacteristic = Value
End Property
' ######################################## ' ^   ISerializable   ^ ' ######################################## '

'jetzt readonly:
Public Property Get ENorm() As ENorm
    ENorm = m_N.ENorm
End Property
Public Property Get gamM() As Double
    gamM = m_N.gamM ' IIf(m_is_k, 1#, 1.1)
'    If m_is_k Then
'        gamM = 1#
'    Else
'        gamM = 1.1
'    End If
End Property
Public Property Get gamM0() As Double
    gamM0 = m_N.gamM0 ' IIf(m_is_k, 1#, 1#)
'    If m_is_k Then
'        gamM0 = 1#
'    Else
'        gamM0 = 1#
'    End If
End Property
Public Property Get gamM2() As Double
    gamM2 = m_N.gamM2 ' IIf(m_is_k, 1#, 1.25)
'    If m_is_k Then
'        gamM2 = 1#
'    Else
'        gamM2 = 1.25
'    End If
End Property
Public Property Get gamM3() As Double
    gamM3 = m_N.gamM3 ' IIf(m_is_k, 1#, 1.25)
'    If m_is_k Then
'        gamM3 = 1#
'    Else
'        gamM3 = 1.25
'    End If
End Property
Public Property Get gamM3ser() As Double
    gamM3ser = m_N.gamM3ser ' IIf(m_is_k, 1#, 1.1)
'    If m_is_k Then
'        gamM3ser = 1#
'    Else
'        gamM3ser = 1.1
'    End If
End Property

Public Property Get AbkNameGrenzabscherkraft() As String
    AbkNameGrenzabscherkraft = _
        m_N.AbkNameGrenzabscherkraft 'IIf(m_N = ENorm.Norm_DIN18800, "V_a,Rd", "F_v,Rd")
End Property
Public Property Get AbkNameGrenzlochleibungskraft() As String
    AbkNameGrenzlochleibungskraft = _
        m_N.AbkNameGrenzlochleibungskraft 'IIf(m_N = ENorm.Norm_DIN18800, "V_l,Rd", "F_b,Rd")
End Property
Public Property Get AbkNameGrenzzugkraft() As String
    AbkNameGrenzzugkraft = _
        m_N.AbkNameGrenzzugkraft 'IIf(m_N = ENorm.Norm_DIN18800, "V_t,Rd", "F_t,Rd")
End Property
Public Property Get AbkNameGrenzdurchstanzkraft() As String
    AbkNameGrenzdurchstanzkraft = _
        m_N.AbkNameGrenzdurchstanzkraft 'IIf(m_N = ENorm.Norm_DIN18800, "V_b,Rd", "B_p,Rd")
End Property
Public Property Get AbkNameGrenzgleitkraft() As String
    AbkNameGrenzgleitkraft = _
        m_N.AbkNameGrenzgleitkraft 'IIf(m_N = ENorm.Norm_DIN18800, "V_s,Rd", "F_s,Rd")
End Property
Public Property Get AbkNameGrenzvorspannkraft() As String
    AbkNameGrenzvorspannkraft = _
        m_N.AbkNameGrenzvorspannkraft 'IIf(m_N = ENorm.Norm_DIN18800, "V_p,C", "F_p,C")
End Property
Public Property Get AbkNameBlockversagenskraftNzentr() As String
    AbkNameBlockversagenskraftNzentr = _
        m_N.AbkNameBlockversagenskraftNzentr 'IIf(m_N = ENorm.Norm_DIN18800, "V_eff,1,Rd", "V_eff,1,Rd")
End Property
Public Property Get AbkNameBlockversagenskraftVexztr() As String
    AbkNameBlockversagenskraftVexztr = _
        m_N.AbkNameBlockversagenskraftVexztr 'IIf(m_N = ENorm.Norm_DIN18800, "V_eff,2,Rd", "V_eff,2,Rd")
End Property

'OM 2016_10_24: durch ToStr ersetzt
'Public Sub ToListBox(aLB As ListBox)
'    With aLB
'        .AddItem "Norm                : " & ENorm_ToStr(m_N)
'        .AddItem "  gam_M: " & gam_M & "; gam_M2: " & gam_M2
'    End With
'End Sub
Friend Function ToStr() As String
    Dim s As String: s = ""
    s = s & "Norm                : " & ENorm_ToStr(m_N.ENorm) & vbCrLf
    s = s & "  gam_M: " & Me.gamM & "; gamM2: " & Me.gamM2 & vbCrLf
    ToStr = s
End Function

