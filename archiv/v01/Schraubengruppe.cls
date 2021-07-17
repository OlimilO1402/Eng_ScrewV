VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Schraubengruppe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Die Klasse Schraubengruppe definiert eine 2D-Gruppe in X-Anzahl in x-Richtung mal Z-Anzahl in z-Richtung
'von gleichen Schrauben mit einem bestimmten Abstand zum Rand des Blechs und der Schrauben untereinander
Private m_S As Schraube
Private m_NX As Byte 'Anzahl der Schrauben in X-Richtung
Private m_NZ As Byte 'Anzahl der Schrauben in Z-Richtung
Private m_RandX As Double
Private m_RandZ As Double
Private m_LochX As Double
Private m_LochZ As Double
Public Enum EAbstand
    AbstandMinimal
    AbstandMittel
    AbstandMaximal
End Enum
Private m_SLs() As Schraubenloch

Public Sub New_(ByVal SL As Schraubenloch, ByVal nSchraubenX As Byte, ByVal nSchraubenZ As Byte, _
                                           ByVal RandX As Double, ByVal RandZ As Double, _
                                           ByVal LochX As Double, ByVal LochZ As Double)
       m_NX = nSchraubenX:    m_NZ = nSchraubenZ
    m_RandX = RandX:       m_RandZ = RandZ
    m_LochX = LochX:       m_LochZ = LochZ
    ReDim m_SLs(0 To Anzahl - 1)
    Set m_SLs(0) = SL
End Sub

Public Property Get Schraube() As Schraube
    Set Schraube = m_SLs(0).Schraube
End Property

Public Property Get Schraubenloch(ByVal Index As Long) As Schraubenloch
    If Index >= Anzahl Then Index = Anzahl - 1
    If Index < 0 Then Index = 0
    Set Schraubenloch = m_SLs(Index)
End Property

Public Property Get Anzahl() As Integer
    Anzahl = m_NX * m_NZ
End Property

Public Property Get AnzahlX() As Byte
    AnzahlX = m_NX
End Property
Public Property Let AnzahlX(ByVal Value As Byte)
    m_NX = Value
End Property
Public Property Get AnzahlZ() As Byte
    AnzahlZ = m_NZ
End Property
Public Property Let AnzahlZ(ByVal Value As Byte)
    m_NZ = Value
End Property

Public Property Get AbstandRandX() As Double
    AbstandRandX = m_RandX 'D: e_1 ' E: e_1 'in Kraftrichtung
End Property
Public Property Let AbstandRandX(ByVal Value As Double)
    m_RandX = Value        'D: e_1 ' E: e_1 'in Kraftrichtung
End Property
Public Property Get AbstandRandZ() As Double
    AbstandRandZ = m_RandZ 'D: e_2 ' E: e_2 'senkr zu Kraftrichtung
End Property
Public Property Let AbstandRandZ(ByVal Value As Double)
    m_RandZ = Value        'D: e_2 ' E: e_2 'senkr zu Kraftrichtung
End Property

Public Property Get AbstandLochX() As Double
    AbstandLochX = m_LochX 'D: e_0 ' E: p_1 'in Kraftrichtung
End Property
Public Property Let AbstandLochX(ByVal Value As Double)
    m_LochX = Value        'D: e_0 ' E: p_1 'in Kraftrichtung
End Property
Public Property Get AbstandLochZ() As Double
    AbstandLochZ = m_LochZ 'D: e_3 ' E: p_2 'senkr zu Kraftrichtung
End Property
Public Property Let AbstandLochZ(ByVal Value As Double)
    m_LochZ = Value        'D: e_3 ' E: p_2 'senkr zu Kraftrichtung
End Property

Public Function MinE1() As Double 'AbstandRandX
    Dim d_L As Double: d_L = m_SLs(0).Durchmesser
    MinE1 = RoundUp(1.2 * d_L)
End Function
Public Function MinE2() As Double 'AbstandRandZ
    Dim d_L As Double: d_L = m_SLs(0).Durchmesser
    MinE2 = RoundUp(1.2 * d_L)
End Function
Public Function MinP1() As Double 'AbstandLochX aka e_0
    Dim d_L As Double: d_L = m_SLs(0).Durchmesser
    MinP1 = RoundUp(2.2 * d_L)
End Function
Public Function MinP2() As Double 'AbstandLochZ aka e_3
    Dim d_L As Double: d_L = m_SLs(0).Durchmesser
    MinP2 = RoundUp(2.4 * d_L)
End Function

Public Function MaxE1() As Double 'AbstandRandX
    Dim d_L As Double: d_L = m_SLs(0).Durchmesser
    MaxE1 = RoundUp(3# * d_L)
End Function
Public Function MaxE2() As Double 'AbstandRandZ
    Dim d_L As Double: d_L = m_SLs(0).Durchmesser
    MaxE2 = RoundUp(1.5 * d_L)
End Function
Public Function MaxP1() As Double 'AbstandLochX aka e_0
    Dim d_L As Double: d_L = m_SLs(0).Durchmesser
    MaxP1 = RoundUp(3.5 * d_L)
End Function
Public Function MaxP2() As Double 'AbstandLochZ aka e_3
    Dim d_L As Double: d_L = m_SLs(0).Durchmesser
    MaxP2 = RoundUp(3# * d_L)
End Function

Public Sub SetAbstandRand(A As EAbstand)
    'setzt a-Schraubenabstand in Abhängigkeit vom Schraubendurchmesser
    Dim d_L As Double: d_L = m_SLs(0).Durchmesser
    Select Case A
    Case AbstandMinimal:
        m_RandX = RoundUp(1.2 * d_L) ' e_1
        m_RandZ = RoundUp(1.2 * d_L) ' e_2
    Case AbstandMittel:
        m_RandX = RoundUp(1.2 + 3) / 2 * d_L ' e_1
        m_RandZ = RoundUp(1.2 + 3) / 2 * d_L ' e_2
    Case AbstandMaximal:
        m_RandX = RoundUp(3# * d_L)  ' e_1
        m_RandZ = RoundUp(1.5 * d_L) ' e_2
    End Select
End Sub
Public Sub SetAbstandLoch(A As EAbstand)
    'setzt a-Schraubenabstand in Abhängigkeit vom Schraubendurchmesser
    Dim d_L As Double: d_L = m_SLs(0).Durchmesser
    Select Case A
    Case AbstandMinimal:
        m_LochX = RoundUp(2.2 * d_L) ' e_0
        m_LochZ = RoundUp(3# * d_L)  ' e_3
    Case AbstandMittel:
        m_LochX = RoundUp((2.2 + 3.5) / 2 * d_L) ' e_0
        m_LochZ = RoundUp(3# * d_L) ' e_3
    Case AbstandMaximal:
        m_LochX = RoundUp(3.5 * d_L) ' e_0
        m_LochZ = RoundUp(3# * d_L)  ' e_3
    End Select
End Sub

Public Sub ToListBox(aLB As ListBox)
    With aLB
        .AddItem "Anzahl-X       : " & AnzahlX & ";  Anzahl-Z       : " & AnzahlZ
        .AddItem "Abstand Rand-X : " & AbstandRandX & ";  Abstand Rand-Z : " & AbstandRandZ
        .AddItem "Abstand Loch-X : " & AbstandLochX & ";  Abstand Loch-Z : " & AbstandLochZ
    End With
End Sub
