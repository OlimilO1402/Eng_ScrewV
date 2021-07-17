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
Private m_Ed As EinwirkungsKombi

Public Sub New_(no As Norm, sc As Schraube, sl As Schraubenloch, sg As Schraubengruppe, bl As Blech, br As Blech, ed As EinwirkungsKombi)
    Set m_no = no:    Set m_sc = sc:    Set m_sl = sl:    Set m_sg = sg:    Set m_bl = bl:    Set m_br = br:    Set m_Ed = ed
End Sub

Public Property Get Norm() As Norm
    Set Norm = m_no
End Property
Public Property Get Schraube() As Schraube
    Set Schraube = m_sc
End Property
Public Property Get Schraubenloch() As Schraubenloch
    Set Schraubenloch = m_sl
End Property
Public Property Get Schraubengruppe() As Schraubengruppe
    Set Schraubengruppe = m_sg
End Property
Public Property Get BlechLinks() As Blech
    Set BlechLinks = m_bl
End Property
Public Property Get BlechRechts() As Blech
    Set BlechRechts = m_br
End Property
Public Property Get EinwirkungsKombi() As EinwirkungsKombi
    Set EinwirkungsKombi = m_Ed
End Property

Public Sub ToListBox(aLB As ListBox)
    Call m_no.ToListBox(aLB)
    Call m_sc.ToListBox(aLB)
    Call m_sl.ToListBox(aLB)
    Call m_sg.ToListBox(aLB)
    Call m_bl.ToListBox(aLB)
    Call m_br.ToListBox(aLB)
    Call m_Ed.ToListBox(aLB)
End Sub
