VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RectWH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Width As Double
Private m_Height As Double

Friend Sub New_(ByVal W As Double, ByVal H As Double)
    m_Width = W
    m_Height = H
End Sub

' ######################################## ' v   ISerializable   v ' ######################################## '
Public Property Get ID() As Long
    ID = ObjPtr(Me)
End Property
Public Property Let ID(ByVal Value As Long)
    'nix
End Property

Public Property Get Width() As Double
    Width = m_Width
End Property
Public Property Let Width(ByVal Value As Double)
    m_Width = Value
End Property

Public Property Get Height() As Double
    Height = m_Height
End Property
Public Property Let Height(ByVal Value As Double)
    m_Height = Value
End Property
' ######################################## ' ^   ISerializable   ^ ' ######################################## '

