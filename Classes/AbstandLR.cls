VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AbstandLR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Klasse Abstand Loch Rand gewählt oder rechnerisch optimiert
'Achtung LR heißt nicht Links/Rechts sonder Loch/Rand!!!
Private WithEvents mLoch As VectorXZ
Attribute mLoch.VB_VarHelpID = -1
Private WithEvents mRand As VectorXZ
Attribute mRand.VB_VarHelpID = -1
Public Event Changed(sender As AbstandLR)

Friend Sub New_(ByVal LochXZ As VectorXZ, ByVal RandXZ As VectorXZ)
    Set mLoch = LochXZ
    Set mRand = RandXZ
End Sub
Friend Sub NewC(ByVal other As AbstandLR)
    With other
        Set mLoch = MNew.VectorXZCopy(.Loch)
        Set mRand = MNew.VectorXZCopy(.Rand)
    End With
End Sub

' ######################################## ' v   ISerializable   v ' ######################################## '
Public Property Get ID() As Long
    ID = ObjPtr(Me)
End Property
Public Property Let ID(ByVal Value As Long)
    'nix
End Property

Public Property Get Loch() As VectorXZ
    Set Loch = mLoch
End Property
Public Property Set Loch(ByVal Value As VectorXZ)
    Set mLoch = Value
    RaiseEvent Changed(Me)
End Property

Public Property Get Rand() As VectorXZ
    Set Rand = mRand
End Property
Public Property Set Rand(ByVal Value As VectorXZ)
    Set mRand = Value
    RaiseEvent Changed(Me)
End Property
' ######################################## ' ^   ISerializable   ^ ' ######################################## '

Private Sub mLoch_Changed(sender As VectorXZ)
    RaiseEvent Changed(Me)
End Sub
Private Sub mRand_Changed(sender As VectorXZ)
    RaiseEvent Changed(Me)
End Sub

Public Property Get RectInnWidth(ByVal nx As Long) As Double
    RectInnWidth = (nx - 1) * mLoch.X
End Property
Public Property Get RectInnHeight(ByVal nz As Long) As Double
    RectInnHeight = (nz - 1) * mLoch.Z
End Property

Public Property Get RectOutWidth(ByVal nx As Long) As Double
    RectOutWidth = RectInnWidth(nx) + 2 * mRand.X
End Property
Public Property Get RectOutHeight(ByVal nz As Long) As Double
    RectOutHeight = RectInnHeight(nz) + 2 * mRand.Z
End Property

'Public Property Get EAbstLoch() As EAbstand
'    EAbstLoch = m_eaL
'End Property
'Public Property Let EAbstLoch(ByVal Value As EAbstand)
'    m_eaL = Value
'End Property
'
'Public Property Get EAbstRand() As EAbstand
'    EAbstRand = m_eaR
'End Property
'Public Property Let EAbstRand(ByVal Value As EAbstand)
'    m_eaR = Value
'End Property
'
'Public Property Get AbstandRand() As VectorXZ
'    Set AbstandRand = m_AbstandRand 'D: e_1 ' E: e_1 'in Kraftrichtung
'End Property
'Public Property Get AbstandLoch() As VectorXZ
'    'soll man bei nx=1 hier 0 zurückliefern? oder vlt unendl?
'    'AbstandLochX = m_LochX 'D: e_0 ' E: p_1 'in Kraftrichtung
'    Set AbstandLoch = m_AbstandLoch 'IIf(m_NX = 1, MMath.posINF, )
'End Property
'Public Property Get AbstandLochZ() As Double
'    'soll man bei nz=1 hier 0 zurückliefern? oder vlt unendl?
'    'AbstandLochZ = m_LochZ 'D: e_3 ' E: p_2 'senkr zu Kraftrichtung
'    AbstandLochZ = m_LochZ 'IIf(m_NZ = 1, MMath.posINF, )
'End Property

'OM 2016_10_24: durch ToStr ersetzt
'Public Sub ToListBox(aLB As ListBox)
'    With aLB
'        'das ist so auch nicht sooo gut!!!
'        'die Klasse ist zu allgemein besser die richtigen Bezeichnungen je nach Norm also e_1, e_2, p_1, p_2 . . .
'        .AddItem "  Abstand Rand: " & Me.Rand.ToStr
'        .AddItem "  Abstand Loch: " & Me.Loch.ToStr
'    End With
'End Sub
Friend Function ToStr() As String
    Dim s As String
    s = s & "  Abstand Rand: " & Me.Rand.ToStr & vbCrLf
    s = s & "  Abstand Loch: " & Me.Loch.ToStr & vbCrLf
    ToStr = s
End Function

