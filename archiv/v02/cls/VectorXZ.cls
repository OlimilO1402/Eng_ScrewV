VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VectorXZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_X As Double
Private m_Z As Double
Public Event Changed(sender As VectorXZ)

Public Sub New_(ByVal aX As Double, aZ As Double)
    m_X = aX: m_Z = aZ
End Sub
Public Sub NewC(ByVal other As VectorXZ)
    With other:        m_X = .X:        m_Z = .Z:    End With
End Sub

Public Property Get X() As Double
    X = m_X
End Property
Public Property Let X(ByVal Value As Double)
    m_X = Value
    RaiseEvent Changed(Me)
End Property
Public Property Get Z() As Double
    Z = m_Z
End Property
Public Property Let Z(ByVal Value As Double)
    m_Z = Value
    RaiseEvent Changed(Me)
End Property

Public Property Get Length() As Double
    Length = VBA.Math.Sqr(X ^ 2 + Z ^ 2)
End Property
Public Function IsEqualXZ(ByVal aX As Double, ByVal aZ As Double) As Boolean
    IsEqualXZ = ((X = aX) And (Z = aZ))
End Function
Public Function ToStr() As String
    ToStr = "{X: " & X & "; Z: " & Z & "}"
End Function

Public Property Get JSONObject() As cCollection
    Set JSONObject = New_c.JSONObject
    With JSONObject
        .prop("X") = m_X
        .prop("Z") = m_Z
    End With
End Property
Public Property Set JSONObject(jo As cCollection)
    'woher weiﬂ ich dass ich den richtigen stream hab?
    m_X = jo.prop("X")
    m_Z = jo.prop("Z")
End Property

