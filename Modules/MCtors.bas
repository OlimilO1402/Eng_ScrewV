Attribute VB_Name = "MNew"
Option Explicit


Private Declare Function GetWindow Lib "user32" (ByVal hwnd As _
        Long, ByVal wCmd As Long) As Long
        
Private Declare Function GetClassName Lib "user32" Alias _
        "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName _
        As String, ByVal nMaxCount As Long) As Long

Const GW_OWNER = 4


'eigentlich müßte für diese Funktion auch erst die Datei gesucht werden
'und mit LoadLibrary geladen werden
Declare Function GetInstanceEx Lib "DirectCOM.dll" _
    (StrPtr_FName As Long, _
     StrPtr_ClassName As Long, _
     ByVal UseAlteredSearchPath As Boolean) As Object
     
Private m_NormInst As Norm
Private sVBRC5pfn  As String
Public IsInIDE As Boolean

'Public Function Norm(Optional aNorm As ENorm) As Norm
''Norm soll ein Singleton/Monoton sein
'    If m_Norm Is Nothing Then
'        Set m_Norm = New Norm: m_Norm.New_ aNorm
'        Set Norm = m_Norm
'    Else
'        If m_Norm.Norm = aNorm Then
'            Set Norm = m_Norm
'        Else
'            'darf nicht vorkommen
'
'        End If
'    End If
'End Function
'Scheiße wie kommt die Norm hierein wenn man deserialisiert?
'OK man muss einfach nach dem Deserialiseren hiereinsetzen
Public Function Norm(ByVal aNorm As Norm) As Norm
'Norm soll ein Singleton/Monoton sein
    If m_NormInst Is Nothing Then
        Set m_NormInst = New Norm
    'Else
        'If m_Norm.Norm = aNorm Then
            'Set Norm = m_Norm
        'Else
            'darf nicht vorkommen
        'End If
    End If
    m_NormInst.New_ aNorm
    Set Norm = m_NormInst
End Function

Public Sub SetNormAfterDeserializing(aNorm As Norm)
    Set m_NormInst = aNorm
End Sub

Public Function NormFromENorm(e As ENorm) As Norm
    Select Case e
    Case Norm_DIN18800:  Set NormFromENorm = New NormDIN18800
    Case Norm_EuroCode3: Set NormFromENorm = New NormEuroCode3
    End Select
End Function
    
Public Function Schraube(ByVal aCalcNorm As Norm, _
                         ByVal Schraubendurchmesser As Double, _
                         Optional ByVal Schraubengüte As Double = 4.6, _
                         Optional ByVal isSenk As Boolean = False, _
                         Optional ByVal isPass As Boolean = False, _
                         Optional ByVal isSFS As Boolean = False, _
                         Optional ByVal isGlf As Boolean = False, _
                         Optional ByVal isZug As Boolean = False, _
                         Optional ByVal isVor As Boolean = False) As Schraube
    Set Schraube = New Schraube: Schraube.New_ aCalcNorm, Schraubendurchmesser, Schraubengüte, isSenk, isPass, isSFS, isGlf, isZug, isVor
End Function

Public Function Schraubenloch(ByVal s As Schraube, ByVal la As ELochart, _
                              Optional ByVal isVert As Boolean = False) As Schraubenloch
    Set Schraubenloch = New Schraubenloch: Schraubenloch.New_ s, la, isVert
End Function

Public Function Schraubengruppe(ByVal sl As Schraubenloch, _
                                ByVal nSchraubenX As Long, _
                                ByVal nSchraubenZ As Long, _
                                ByVal AbstandSelected As AbstandLR, _
                                ByVal eaL As EAbstand, ByVal eaR As EAbstand, _
                                Ewk As EinwirkungsKombi, _
                                bl As Blech, br As Blech) As Schraubengruppe
    Set Schraubengruppe = New Schraubengruppe: Schraubengruppe.New_ sl, nSchraubenX, nSchraubenZ, AbstandSelected, eaL, eaR, Ewk, bl, br
End Function

Public Function VectorXZ(ByVal aX As Double, ByVal aZ As Double) As VectorXZ
    Set VectorXZ = New VectorXZ: VectorXZ.New_ aX, aZ
End Function
Public Function VectorXZCopy(ByVal other As VectorXZ) As VectorXZ
    Set VectorXZCopy = New VectorXZ: VectorXZCopy.NewC other
End Function

Public Function AbstandLR(ByVal LochXZ As VectorXZ, ByVal RandXZ As VectorXZ) As AbstandLR
    Set AbstandLR = New AbstandLR: Call AbstandLR.New_(LochXZ, RandXZ)
End Function
Public Function AbstandLRCopy(ByVal other As AbstandLR) As AbstandLR
    Set AbstandLRCopy = New AbstandLR: AbstandLRCopy.NewC other
End Function

Public Function Blech(ByVal aNorm As Norm, _
                      ByVal sso As EStahlsorte, _
                      ByVal t As Double, _
                      ByVal l_x As Double, _
                      ByVal h_z As Double, _
                      ByVal bIsLinks As Boolean, _
                      ByVal bIsZange As Boolean, _
                      ByVal bIsMehrschnittig As Boolean) As Blech
    Set Blech = New Blech: Blech.New_ aNorm, sso, t, l_x, h_z, bIsLinks, bIsZange, bIsMehrschnittig
'sg,                       ByVal sg As Schraubengruppe,
End Function

Public Function EinwirkungsKombi(ByVal MEd As Double, ByVal VEd As Double, ByVal NEd As Double, _
                                 ByVal OffX As Double, ByVal OffZ As Double) As EinwirkungsKombi
    Set EinwirkungsKombi = New EinwirkungsKombi: EinwirkungsKombi.New_ MEd, VEd, NEd, OffX, OffZ
End Function

Public Function SchraubenNachweis(ByVal aCalcNorm As Norm, s As Schraube, sg As Schraubengruppe) As SchraubenNachweis
    Set SchraubenNachweis = New SchraubenNachweis: SchraubenNachweis.New_ aCalcNorm, s, sg
End Function

Public Function SchrVerbFL(no As Norm, sc As Schraube, sl As Schraubenloch, sg As Schraubengruppe, bl As Blech, br As Blech, nw As SchraubenNachweis, ed As EinwirkungsKombi) As SchrVerbFL
    Set SchrVerbFL = New SchrVerbFL: SchrVerbFL.New_ no, sc, sl, sg, bl, br, nw, ed
End Function

Public Function List2DOfSGScrew(ByVal aN1 As Long, ByVal aN2 As Long) As List2DOfSGScrew
    Set List2DOfSGScrew = New List2DOfSGScrew: List2DOfSGScrew.New_ aN1, aN2
End Function

Public Function SGScrew(aPos As VectorXZ, VForces As VectorXZ, aSL As Schraubenloch) As SGScrew
    Set SGScrew = New SGScrew: SGScrew.New_ aPos, VForces, aSL
End Function
Public Function SGScrewC(other As SGScrew) As SGScrew
    'copy-constructor
    Set SGScrewC = New SGScrew: SGScrewC.NewC other
End Function

Public Function Constructor() As cConstructor
    'OK was machen?
    'sodale jetzt erstmal den Pfad suchen auf die Datei "DirectCOM.dll"
    'sodale jetzt erstmal den Pfad suchen auf die Datei "vbRichClient5.dll"
    'die Datei könnte sein im Pfad:
    '* "C:\RC5\"
    '* "C:\RC5\bin\"
    '* "C:\
    '* App.Path
    '* App.Path & "\RC5\"
    '* App.Path & "\bin\"
    '* App.Path & "\..\" & "\RC5\"
    '* App.Path & "\..\" & "\bin\"
    'oder in einem Unterordner von "C:\ProgramData\"
    'oder in einem Shared-Verzeichnis
    '"DirectCOM.dll"
    If IsInIDE Then
        Set Constructor = New vbrichclient5.cConstructor
    Else
        sVBRC5pfn = App.Path
        If Right(sVBRC5pfn, 1) <> "\" Then sVBRC5pfn = sVBRC5pfn & "\"
        'Debug.Print sVBRC5pfn
        If Not FileExists(sVBRC5pfn & "DirectCOM.dll") Then
            MsgBox "konnte die Datei nicht finden:" & vbCrLf & sVBRC5pfn & "DirectCOM.dll"
        End If
        If Not FileExists(sVBRC5pfn & "vbRichClient5.dll") Then
            MsgBox "konnte die Datei nicht finden:" & vbCrLf & sVBRC5pfn & "vbRichClient5.dll"
        End If
        If Not FileExists(sVBRC5pfn & "vb_cairo_sqlite.dll") Then
            MsgBox "konnte die Datei nicht finden:" & vbCrLf & sVBRC5pfn & "vb_cairo_sqlite.dll"
        End If
        sVBRC5pfn = sVBRC5pfn & "vbRichClient5.dll"
        Set Constructor = GetInstanceEx(StrPtr(sVBRC5pfn), StrPtr("cConstructor"), True)
    End If
End Function

Public Function FileExists(ByVal FileName As String) As Boolean
    On Error Resume Next
    FileExists = Not CBool(GetAttr(FileName) And (vbDirectory Or vbVolume))
    On Error GoTo 0
End Function

Public Function DirExists(ByVal DirectoryName As String) As Boolean
    On Error Resume Next
    DirExists = CBool(GetAttr(DirectoryName) And vbDirectory)
    On Error GoTo 0
End Function

Public Function CairoPicBox(aPb As PictureBox, c As cCairo) As CairoPicBox
    Set CairoPicBox = New CairoPicBox: CairoPicBox.New_ aPb, c
End Function
Public Function CairoPdfDoc(c As cCairo, ByVal epo As EPageOrientation, ByVal epf As EPageFormat, ByVal zoom As Double) As CairoPdfDoc
    Set CairoPdfDoc = New CairoPdfDoc: CairoPdfDoc.New_ c, epo, epf, zoom
End Function

'Public Function IsInIDE() As Boolean
'Try: On Error GoTo Catch
'    Debug.Print 1 / 0
'    Exit Function
'Catch: IsInIDE = True
'End Function
Public Function IsRunningInIDE(ByVal ahWnd As Long) As Boolean
'Private Function IsInIDE() As Boolean
    'evtl hier die Länge des Srings vergrößern:
    Dim Buffer     As String:     Buffer = Space(128)
    Dim ParenthWnd As Long:   ParenthWnd = GetWindow(ahWnd, GW_OWNER)
    Call GetClassName(ParenthWnd, Buffer, Len(Buffer))

    IsRunningInIDE = (LCase(Left(Buffer, 11)) = "thundermain")
    'If LCase(Left(Buffer, 11)) = "thundermain" Then
    '    IsRunningInIDE = True
    'Else
    '    IsRunningInIDE = False
    'End If
End Function

'Private Function IsRunningInIDE() As Boolean
'Private Function IsInIDE() As Boolean
'    'evtl hier die Länge des Srings vergrößern:
'    Dim Buffer     As String:     Buffer = Space(128)
'    Dim ParenthWnd As Long:   ParenthWnd = GetWindow(Me.hWnd, GW_OWNER)
'    Call GetClassName(ParenthWnd, Buffer, Len(Buffer))
'
'    IsRunningInIDE = (LCase(Left(Buffer, 11)) = "thundermain")
'    'If LCase(Left(Buffer, 11)) = "thundermain" Then
'    '    IsRunningInIDE = True
'    'Else
'    '    IsRunningInIDE = False
'    'End If
'End Function
