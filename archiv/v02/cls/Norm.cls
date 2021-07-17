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
Private m_N As ENorm
Private m_is_k As Boolean

Friend Sub New_(ByVal aNormBerechnung As ENorm)
    m_N = aNormBerechnung
End Sub

Public Property Get Norm() As ENorm
    Norm = m_N
End Property
Public Property Let Norm(ByVal Value As ENorm)
    m_N = Value
End Property

Public Property Get IsCharacteristic() As Boolean
    IsCharacteristic = m_is_k
End Property
Public Property Let IsCharacteristic(ByVal Value As Boolean)
    m_is_k = Value
End Property

Public Property Get gam_M() As Double
    If m_is_k Then
        gam_M = 1#
    Else
        gam_M = 1.1
    End If
End Property
Public Property Get gam_M0() As Double
    If m_is_k Then
        gam_M0 = 1#
    Else
        gam_M0 = 1#
    End If
End Property
Public Property Get gam_M2() As Double
    If m_is_k Then
        gam_M2 = 1#
    Else
        gam_M2 = 1.25
    End If
End Property
Public Property Get gam_M3() As Double
    If m_is_k Then
        gam_M3 = 1#
    Else
        gam_M3 = 1.25
    End If
End Property
Public Property Get gam_M3_ser() As Double
    If m_is_k Then
        gam_M3_ser = 1#
    Else
        gam_M3_ser = 1.1
    End If
End Property

Public Property Get AbkName_Grenzabscherkraft() As String
    AbkName_Grenzabscherkraft = _
        IIf(m_N = ENorm.Norm_DIN18800, "V_a,Rd", "F_v,Rd")
End Property
Public Property Get AbkName_Grenzlochleibungskraft() As String
    AbkName_Grenzlochleibungskraft = _
        IIf(m_N = ENorm.Norm_DIN18800, "V_l,Rd", "F_b,Rd")
End Property
Public Property Get AbkName_Grenzzugkraft() As String
    AbkName_Grenzzugkraft = _
        IIf(m_N = ENorm.Norm_DIN18800, "V_t,Rd", "F_t,Rd")
End Property
Public Property Get AbkName_Grenzdurchstanzkraft() As String
    AbkName_Grenzdurchstanzkraft = _
        IIf(m_N = ENorm.Norm_DIN18800, "V_b,Rd", "B_p,Rd")
End Property
Public Property Get AbkName_Grenzgleitkraft() As String
    AbkName_Grenzgleitkraft = _
        IIf(m_N = ENorm.Norm_DIN18800, "V_s,Rd", "F_s,Rd")
End Property
Public Property Get AbkName_Grenzvorspannkraft() As String
    AbkName_Grenzvorspannkraft = _
        IIf(m_N = ENorm.Norm_DIN18800, "V_p,C", "F_p,C")
End Property
Public Property Get AbkName_BlockversagenskraftNzentr() As String
    AbkName_BlockversagenskraftNzentr = _
        IIf(m_N = ENorm.Norm_DIN18800, "V_eff,1,Rd", "V_eff,1,Rd")
End Property
Public Property Get AbkName_BlockversagenskraftVexztr() As String
    AbkName_BlockversagenskraftVexztr = _
        IIf(m_N = ENorm.Norm_DIN18800, "V_eff,2,Rd", "V_eff,2,Rd")
End Property



Public Sub ToListBox(aLB As ListBox)
    With aLB
        .AddItem "Norm                : " & ENorm_ToStr(m_N)
        .AddItem "  gam_M: " & gam_M & "; gam_M2: " & gam_M2
    End With
End Sub
