VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SchrVerbFL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Diese Klasse hat alle Objekte der Verbindung
Private m_no As Norm
Private m_sc As Schraube
Private m_sl As Schraubenloch
Private m_sg As Schraubengruppe
Private m_bl As Blech 'Blech von links  kann auch doppelt sein
Private m_br As Blech 'Blech von rechts kann nur  einfach sein
Private m_nw As SchraubenNachweis
Private m_Ed As EinwirkungsKombi

Friend Sub New_(no As Norm, sc As Schraube, sl As Schraubenloch, sg As Schraubengruppe, bl As Blech, br As Blech, nw As SchraubenNachweis, ed As EinwirkungsKombi)
    Set m_no = no:    Set m_sc = sc:    Set m_sl = sl:    Set m_sg = sg:    Set m_bl = bl:    Set m_br = br:    Set m_nw = nw: Set m_Ed = ed
End Sub

' ######################################## ' v   ISerializable   v ' ######################################## '
Public Property Get ID() As Long
    ID = ObjPtr(Me)
End Property
Public Property Let ID(ByVal Value As Long)
    'nix
End Property

Public Property Get Norm() As Norm
    Set Norm = m_no
End Property
Public Property Set Norm(ByVal Value As Norm)
    Set m_no = Value
End Property

Public Property Get Schraube() As Schraube
    Set Schraube = m_sc
End Property
Public Property Set Schraube(ByVal Value As Schraube)
    Set m_sc = Value
End Property

Public Property Get Schraubenloch() As Schraubenloch
    Set Schraubenloch = m_sl
End Property
Public Property Set Schraubenloch(ByVal Value As Schraubenloch)
    Set m_sl = Value
End Property

Public Property Get Schraubengruppe() As Schraubengruppe
    Set Schraubengruppe = m_sg
End Property
Public Property Set Schraubengruppe(ByVal Value As Schraubengruppe)
    Set m_sg = Value
End Property

Public Property Get BlechLinks() As Blech
    Set BlechLinks = m_bl
End Property
Public Property Set BlechLinks(ByVal Value As Blech)
    Set m_bl = Value
End Property

Public Property Get BlechRechts() As Blech
    Set BlechRechts = m_br
End Property
Public Property Set BlechRechts(ByVal Value As Blech)
    Set m_br = Value
End Property

Public Property Get SchraubenNachweis() As SchraubenNachweis
    Set SchraubenNachweis = m_nw
End Property
Public Property Set SchraubenNachweis(ByVal Value As SchraubenNachweis)
    Set m_nw = Value
End Property

Public Property Get EinwirkungsKombi() As EinwirkungsKombi
    Set EinwirkungsKombi = m_Ed
End Property
Public Property Set EinwirkungsKombi(ByVal Value As EinwirkungsKombi)
    Set m_Ed = Value
End Property
' ######################################## ' ^   ISerializable   ^ ' ######################################## '

Public Function ToStr() As String
    Dim s As String: s = ""
    s = s & m_no.ToStr & vbCrLf
    s = s & m_sc.ToStr & vbCrLf
    s = s & m_sl.ToStr & vbCrLf
    s = s & m_sg.ToStr & vbCrLf
    s = s & m_bl.ToStr & vbCrLf
    s = s & m_br.ToStr & vbCrLf
    s = s & m_Ed.ToStr & vbCrLf
    ToStr = s
End Function
'Public Sub ToListBox(aLB As ListBox)
'    'Alle Schreiben in die Gleiche ListBox
'    'eigentlich Käsig, in eine ListBox zu schreiben, wie wäre es mit einer TextBox?
'    'OK man könnte das abändern indem man einen Text rausschreibt und eine Routine die Text automatisch in eine ListBox einträgt
'    'Aber warum eigentlich überhaupt die ListBox?
'    'wieso nicht in eine RichTextBox?
'    Dim sa() As String: sa = Split(Me.ToStr, vbCrLf)
'    Dim v
'    For Each v In sa
'        aLB.AddItem s
'    Next
''    Call m_no.ToListBox(aLB)
''    Call m_sc.ToListBox(aLB)
''    Call m_sl.ToListBox(aLB)
''    Call m_sg.ToListBox(aLB)
''    Call m_bl.ToListBox(aLB)
''    Call m_br.ToListBox(aLB)
''    Call M_Ed.ToListBox(aLB)
'End Sub
