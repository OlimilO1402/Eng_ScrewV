VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SGScrew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Klasse stellt eine Schraube in einer Schraubengruppe dar, mit Position Kräfte und Loch
'ehemals TScrew siehe Klasse Schraubengruppe
'sollten her die VectorXZ-Objekte das Changed-Event empfangen
Private m_Pos As VectorXZ 'die Position
Private m_v   As VectorXZ 'die Kräfte in x- und z-Ri
Private m_sl  As Schraubenloch 'das Schrabuenloch mit Schraube

Friend Sub New_(aPos As VectorXZ, VForces As VectorXZ, aSL As Schraubenloch)
    Set m_Pos = aPos
    Set m_v = VForces
    Set m_sl = aSL
End Sub
Friend Sub NewC(other As SGScrew)
    'copy-constructor
    With other
        Set m_Pos = .Position
        Set m_v = .VForces
        Set m_sl = .Schraubenloch
    End With
End Sub
' ######################################## ' v   ISerializable   v ' ######################################## '
Public Property Get ID() As Long
    ID = ObjPtr(Me)
End Property
Public Property Let ID(ByVal Value As Long)
    'nix
End Property

Public Property Get Position() As VectorXZ
    Set Position = m_Pos
End Property
Public Property Set Position(ByVal Value As VectorXZ)
    Set m_Pos = Value
End Property

Public Property Get VForces() As VectorXZ
    Set VForces = m_v
End Property
Public Property Set VForces(ByVal Value As VectorXZ)
    Set m_v = Value
End Property

Public Property Get Schraubenloch() As Schraubenloch
    Set Schraubenloch = m_sl
End Property
Public Property Set Schraubenloch(ByVal Value As Schraubenloch)
    Set m_sl = Value
End Property
' ######################################## ' ^   ISerializable   ^ ' ######################################## '

Friend Function ToStr() As String
    Dim s As String: s = ""
    ToStr = s
End Function
