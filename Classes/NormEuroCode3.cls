VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NormEuroCode3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Norm
Private m_is_k As Boolean

Private Property Get Norm_ENorm() As ENorm
    Norm_ENorm = Norm_EuroCode3
End Property

Private Property Get Norm_ID() As Long
    Norm_ID = ObjPtr(Me)
End Property
Private Property Let Norm_ID(ByVal Value As Long)
    'nix
End Property

Private Property Get Norm_Norm() As Norm
    Set Norm_Norm = Me
End Property
Private Property Set Norm_Norm(ByVal Value As Norm)
    'hier nix
End Property

Private Property Get Norm_IsCharacteristic() As Boolean
    'IsCharacteristic = m_is_k
    Norm_IsCharacteristic = m_is_k
End Property
Private Property Let Norm_IsCharacteristic(ByVal Value As Boolean)
    'm_is_k = Value
    m_is_k = Value
End Property
' ######################################## ' ^   ISerializable   ^ ' ######################################## '

Private Property Get Norm_gamM() As Double
    Norm_gamM = IIf(m_is_k, 1#, 1.1)
End Property
Private Property Get Norm_gamM0() As Double
    Norm_gamM0 = IIf(m_is_k, 1#, 1#)
End Property
Private Property Get Norm_gamM2() As Double
    Norm_gamM2 = IIf(m_is_k, 1#, 1.25)
End Property
Private Property Get Norm_gamM3() As Double
    Norm_gamM3 = IIf(m_is_k, 1#, 1.25)
End Property
Private Property Get Norm_gamM3ser() As Double
    Norm_gamM3ser = IIf(m_is_k, 1#, 1.1)
End Property

Private Property Get Norm_AbkNameGrenzabscherkraft() As String
    Norm_AbkNameGrenzabscherkraft = _
                                            "F_v,Rd"
End Property
Private Property Get Norm_AbkNameGrenzlochleibungskraft() As String
    Norm_AbkNameGrenzlochleibungskraft = _
                                            "F_b,Rd"
End Property
Private Property Get Norm_AbkNameGrenzzugkraft() As String
    Norm_AbkNameGrenzzugkraft = _
                                            "F_t,Rd"
End Property
Private Property Get Norm_AbkNameGrenzdurchstanzkraft() As String
    Norm_AbkNameGrenzdurchstanzkraft = _
                                            "B_p,Rd"
End Property
Private Property Get Norm_AbkNameGrenzgleitkraft() As String
    Norm_AbkNameGrenzgleitkraft = _
                                            "F_s,Rd"
End Property
Private Property Get Norm_AbkNameGrenzvorspannkraft() As String
    Norm_AbkNameGrenzvorspannkraft = _
                                            "F_p,C"
End Property
Private Property Get Norm_AbkNameBlockversagenskraftNzentr() As String
    Norm_AbkNameBlockversagenskraftNzentr = _
                                            "V_eff,1,Rd"
End Property
Private Property Get Norm_AbkNameBlockversagenskraftVexztr() As String
    Norm_AbkNameBlockversagenskraftVexztr = _
                                            "V_eff,2,Rd"
End Property

