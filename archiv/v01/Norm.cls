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

Public Enum ENorm
    Norm_DIN18800
    Norm_EuroCode3
End Enum

Private m_N As ENorm

Friend Sub New_(ByVal aNormBerechnung As ENorm)
    m_N = aNormBerechnung
End Sub

Public Property Get Norm() As ENorm
    Norm = m_N
End Property
Public Property Let Norm(ByVal Value As ENorm)
    m_N = Value
End Property
Private Function Norm_ToStr() As String
    Dim s As String
    Select Case m_N
    Case Norm_DIN18800:  s = "DIN 18800"
    Case Norm_EuroCode3: s = "EuroCode 3"
    End Select
    Norm_ToStr = s
End Function
Public Property Get gam_M() As Double
    gam_M = 1.1
End Property
Public Property Get gam_M2() As Double
    gam_M2 = 1.25
End Property

Public Sub ToListBox(aLB As ListBox)
    With aLB
        .AddItem "Norm                : " & Norm_ToStr
        .AddItem "gam_M : " & gam_M & ";  gam_M2 : " & gam_M2
    End With
End Sub
