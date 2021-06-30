VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ScrewV - It's groovy"
   ClientHeight    =   1455
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1004.267
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   120
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   960
      Width           =   1260
   End
   Begin VB.Label lblDescription 
      Caption         =   "Beschreibung"
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Name der Anwendung"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   1725
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   3840
      TabIndex        =   4
      Top             =   240
      Width           =   1245
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Registrierungsschl�ssel-Sicherheitsoptionen...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Registrierungsschl�ssel-Stammtypen...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Null-terminierte Unicode-Zeichenfolge
Const REG_DWORD = 4                      ' 32-Bit-Zahl

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


'Private Sub cmdSysInfo_Click()
'  Call StartSysInfo
'End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Info zu " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    picIcon.Picture = frmSchrauben.Icon
    Dim s As String
    s = App.CompanyName & vbCrLf & _
        App.FileDescription & " mit " & App.Title & " " & App.Comments & vbCrLf '& _

    lblDescription.Caption = s
End Sub

'Public Sub StartSysInfo()
'    On Error GoTo SysInfoErr
'
'    Dim rc As Long
'    Dim SysInfoPath As String
'
'    ' Versuchen, den Systeminfo-Programmpfad/-namen aus der Registrierung abzurufen...
'    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
'    ' Versuchen, nur den Systeminfo-Programmpfad aus der Registrierung abzurufen...
'    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
'        ' �berpr�fen, ob bekannte 32-Dateiversion vorhanden ist
'        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
'            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
'
'        ' Fehler - Datei wurde nicht gefunden...
'        Else
'            GoTo SysInfoErr
'        End If
'    ' Fehler - Registrierungseintrag wurde nicht gefunden...
'    Else
'        GoTo SysInfoErr
'    End If
'
'    Call Shell(SysInfoPath, vbNormalFocus)
'
'    Exit Sub
'SysInfoErr:
'    MsgBox "Systeminformationen sind momentan nicht verf�gbar", vbOKOnly
'End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Schleifenz�hler
    Dim rc As Long                                          ' R�ckgabe-Code
    Dim hKey As Long                                        ' Zugriffsnummer f�r einen offenen Registrierungsschl�ssel
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Datentyp eines Registrierungsschl�ssels
    Dim tmpVal As String                                    ' Tempor�rer Speicher eines Registrierungsschl�sselwertes
    Dim KeyValSize As Long                                  ' Gr��e der Registrierungsschl�sselvariablen
    '------------------------------------------------------------
    ' Registrierungsschl�ssel unter KeyRoot {HKEY_LOCAL_MACHINE...} �ffnen
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Registrierungsschl�ssel �ffnen
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Fehler behandeln...
    
    tmpVal = String$(1024, 0)                             ' Platz f�r Variable reservieren
    KeyValSize = 1024                                       ' Gr��e der Variable markieren
    
    '------------------------------------------------------------
    ' Registrierungsschl�sselwert abrufen...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Schl�sselwert abrufen/erstellen
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Fehler behandeln
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 f�gt null-terminierte Zeichenfolge hinzu...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null gefunden, aus Zeichenfolge extrahieren
    Else                                                    ' Keine null-terminierte Zeichenfolge f�r WinNT...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null nicht gefunden, nur Zeichenfolge extrahieren
    End If
    '------------------------------------------------------------
    ' Schl�sselwerttyp f�r Konvertierung bestimmen...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Datentypen durchsuchen...
    Case REG_SZ                                             ' Zeichenfolge f�r Registrierungsschl�sseldatentyp
        KeyVal = tmpVal                                     ' Zeichenfolgenwert kopieren
    Case REG_DWORD                                          ' Registrierungsschl�sseldatentyp DWORD
        For i = Len(tmpVal) To 1 Step -1                    ' Jedes Bit konvertieren
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Wert Zeichen f�r Zeichen erstellen
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' DWORD in Zeichenfolge konvertieren
    End Select
    
    GetKeyValue = True                                      ' Erfolgreiche Ausf�hrung zur�ckgeben
    rc = RegCloseKey(hKey)                                  ' Registrierungsschl�ssel schlie�en
    Exit Function                                           ' Beenden
    
GetKeyError:      ' Bereinigen, nachdem ein Fehler aufgetreten ist...
    KeyVal = ""                                             ' R�ckgabewert auf leere Zeichenfolge setzen
    GetKeyValue = False                                     ' Fehlgeschlagene Ausf�hrung zur�ckgeben
    rc = RegCloseKey(hKey)                                  ' Registrierungsschl�ssel schlie�en
End Function

