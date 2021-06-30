VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "List2DOfSGScrew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Private m_u1 As Long 'Zeilen
'Private m_u2 As Long 'Spalten
Private m_nZeilen_i  As Long 'N1
Private m_nSpalten_j As Long 'N2

'hmm könnten wir das nicht besser in eine Collection umwandeln?
'Private m_List() As SGScrew
Private m_List As Collection 'Of SGScrew
Public Event Changed(sender As List2DOfSGScrew)

Private Sub Class_Initialize()
    Clear
    'Delete
End Sub

Friend Sub New_(ByVal aN1 As Long, ByVal aN2 As Long)
    Resize aN1, aN2
End Sub

' ######################################## ' v   ISerializable   v ' ######################################## '
Public Property Get N1() As Long 'Zeilen
    N1 = m_nZeilen_i 'm_u1 + 1
End Property
Public Property Let N1(ByVal Value As Long)
    Resize Value, Me.N2
    RaiseEvent Changed(Me)
End Property

Public Property Get N2() As Long 'Spalten
    N2 = m_nSpalten_j ''m_u2 + 1
End Property
Public Property Let N2(ByVal Value As Long)
    Resize Me.N1, Value
    RaiseEvent Changed(Me)
End Property

Public Property Get List() As Collection
    Set List = m_List
End Property
Public Property Set List(acol As Collection)
    Set m_List = acol
End Property
' ######################################## ' ^   ISerializable   ^ ' ######################################## '

'hmm die Frage ist wie man das serialisieren können soll?
'ich glaub das geht so garnicht!
Friend Property Get Item(ByVal i1 As Long, ByVal i2 As Long) As SGScrew
    'OK die zwei Indizes in einen Index umrechnen
    i1 = i1 + 1
    i2 = i2 + 1
    Set Item = m_List(CalcIndex(i1, i2))
End Property
Friend Property Set Item(ByVal i1 As Long, ByVal i2 As Long, ByVal Value As SGScrew)
    'i1 und i2 sind 0 basiert ->
    i1 = i1 + 1
    i2 = i2 + 1
    Dim sgs As SGScrew: Set sgs = m_List.Item(CalcIndex(i1, i2))
    sgs.NewC Value
    
    'Set m_List.Item(CalcIndex(i1, i2)) = Value
    
End Property

Private Function CalcIndex(ByVal Zeile_i As Long, ByVal Spalte_j As Long) As Long
    CalcIndex = m_nSpalten_j * (Zeile_i - 1) + Spalte_j '- 1
End Function
'Array 4 Zeilen * 6 Spalten => n = 4 * 6 = 24 Elemente
'das Element mit den Indices: Zeile_i=3, Spalte_j=4
'hat den 1-basierten Index: i = 16 = n_Spalten * (Zeile_i - 1) + Spalte_j
'
'*  *  *  *  *  *
'*  *  *  *  *  *
'*  *  * [*] *  *
'*  *  *  *  *  *

'Public Property Get Count() As Long
'    Count = m_nZeilen_i * m_nSpalten_j
'End Property
Public Property Get Count() As Long
    Count = Me.N1 * Me.N2
End Property

Public Sub Resize(ByVal aN1 As Long, ByVal aN2 As Long)
'    m_u1 = aN1 - 1
'    m_u2 = aN2 - 1
    If aN1 < 1 Then Exit Sub
    If aN2 < 1 Then Exit Sub
    Dim n As Long: n = Me.Count
    m_nZeilen_i = aN1
    m_nSpalten_j = aN2
    If aN1 * aN2 > n Then
        'If m_u1 < 0 Then Exit Sub
        'If m_u2 < 0 Then Exit Sub
        'ReDim m_List(0 To m_u1, 0 To m_u2)
        Dim i As Long
        For i = 1 To Me.Count
            m_List.Add New SGScrew
        Next
    End If
End Sub

'Friend Sub Delete()
'    'ReDim m_List(0 To 0, 0 To 0)
'End Sub
Friend Sub Clear()
    Set m_List = New Collection
End Sub
'
'Friend Function GetEnumerator() As IUnknown
'    'GetEnumerator = m_List
'    Set GetEnumerator = m_List.[_NewEnum]
'End Function
'
