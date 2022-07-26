VERSION 5.00
Begin VB.Form frmSchrauben 
   Caption         =   "Form1"
   ClientHeight    =   13935
   ClientLeft      =   450
   ClientTop       =   1095
   ClientWidth     =   14355
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   13935
   ScaleWidth      =   14355
   Begin VB.ListBox LBSchraube 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   5160
      TabIndex        =   51
      Top             =   4320
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   5880
      MultiLine       =   -1  'True
      TabIndex        =   49
      Text            =   "Form1.frx":8882
      Top             =   600
      Width           =   4215
   End
   Begin VB.TextBox TBSchraube 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9855
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   78
      Text            =   "Form1.frx":8888
      Top             =   480
      Width           =   4695
   End
   Begin VB.TextBox TxOffX 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   960
      TabIndex        =   41
      Top             =   10800
      Width           =   855
   End
   Begin VB.TextBox TxOffZ 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   1920
      TabIndex        =   42
      Top             =   10800
      Width           =   855
   End
   Begin VB.CheckBox CkPdfQuer 
      Caption         =   "Querformat"
      Height          =   255
      Left            =   7800
      TabIndex        =   76
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox TxNormalkraftEd 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   1920
      TabIndex        =   40
      Top             =   10440
      Width           =   855
   End
   Begin VB.TextBox TxQuerkraftEd 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   1920
      TabIndex        =   39
      Top             =   10080
      Width           =   855
   End
   Begin VB.TextBox TxMomentEd 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   1920
      TabIndex        =   38
      Top             =   9720
      Width           =   855
   End
   Begin VB.ComboBox CbZoom 
      Height          =   315
      Left            =   9360
      TabIndex        =   47
      Text            =   "Combo1"
      Top             =   120
      Width           =   975
   End
   Begin VB.OptionButton OpErgGra 
      Caption         =   "Erg | Gra"
      Height          =   375
      Left            =   5040
      Style           =   1  'Grafisch
      TabIndex        =   45
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox CbGFKmue 
      Height          =   315
      Left            =   1920
      TabIndex        =   10
      Top             =   1920
      Width           =   855
   End
   Begin VB.CheckBox CkDrawMutter 
      Caption         =   "Mutter"
      Height          =   255
      Left            =   1440
      TabIndex        =   17
      Top             =   3360
      Width           =   855
   End
   Begin VB.CheckBox CkDrawUScheibe 
      Caption         =   "U-Scheibe"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CheckBox CkSFKAlle 
      Caption         =   "alle"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton BtnOptions 
      Caption         =   "Optionen"
      Height          =   375
      Left            =   6360
      TabIndex        =   46
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton OpGrafik 
      Caption         =   "Grafik"
      Height          =   375
      Left            =   3960
      Style           =   1  'Grafisch
      TabIndex        =   44
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton OpCalc 
      Caption         =   "Ergebnisse"
      Height          =   375
      Left            =   2880
      Style           =   1  'Grafisch
      TabIndex        =   43
      Top             =   120
      Width           =   1095
   End
   Begin VB.CheckBox CkZug 
      Caption         =   "Zugverb."
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ComboBox CbGleitGZ 
      Height          =   315
      Left            =   1080
      TabIndex        =   9
      Top             =   1920
      Width           =   855
   End
   Begin VB.CheckBox CkGleitf 
      Caption         =   "Gleitfest"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.CheckBox CkVorgesp 
      Caption         =   "Vorgespannt"
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox CkSenk 
      Caption         =   "Senkkopf"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CheckBox CkSFSchaft 
      Caption         =   "Scherfuge im Schaft"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.ComboBox CbNorm 
      Height          =   315
      ItemData        =   "Form1.frx":888E
      Left            =   1080
      List            =   "Form1.frx":8890
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.PictureBox PnlSG 
      BorderStyle     =   0  'Kein
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5535
      ScaleWidth      =   2775
      TabIndex        =   53
      Top             =   4080
      Width           =   2775
      Begin VB.PictureBox PnlBL 
         BorderStyle     =   0  'Kein
         Height          =   975
         Left            =   120
         ScaleHeight     =   975
         ScaleWidth      =   2775
         TabIndex        =   72
         Top             =   2160
         Width           =   2775
         Begin VB.ComboBox CBBLStahl 
            Height          =   315
            ItemData        =   "Form1.frx":8892
            Left            =   1680
            List            =   "Form1.frx":8894
            TabIndex        =   29
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox TxBLh 
            Alignment       =   1  'Rechts
            Height          =   285
            Left            =   600
            TabIndex        =   27
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox TxBLt 
            Alignment       =   1  'Rechts
            Height          =   285
            Left            =   2040
            TabIndex        =   28
            Top             =   0
            Width           =   615
         End
         Begin VB.CheckBox CkZange 
            Caption         =   "Träger Links ist Zange"
            Height          =   255
            Left            =   0
            TabIndex        =   30
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label LbBLS 
            AutoSize        =   -1  'True
            Caption         =   "Stahlsorte"
            Height          =   195
            Left            =   0
            TabIndex        =   75
            Top             =   360
            Width           =   705
         End
         Begin VB.Label LbBLh 
            AutoSize        =   -1  'True
            Caption         =   "Höhe h"
            Height          =   195
            Left            =   0
            TabIndex        =   74
            Top             =   0
            Width           =   525
         End
         Begin VB.Label LbBLt 
            AutoSize        =   -1  'True
            Caption         =   "Dicke t"
            Height          =   195
            Left            =   1440
            TabIndex        =   73
            Top             =   0
            Width           =   510
         End
      End
      Begin VB.CheckBox CkRoundUp5 
         Caption         =   "auf 5mm aufrunden"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CheckBox CkUpdateAbstand 
         Caption         =   "Abstände aktualisieren"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   5280
         Width           =   2655
      End
      Begin VB.ComboBox CmBAbstLoch 
         Height          =   315
         Left            =   1800
         TabIndex        =   36
         Top             =   4920
         Width           =   975
      End
      Begin VB.ComboBox CmBAbstRand 
         Height          =   315
         Left            =   1800
         TabIndex        =   35
         Top             =   4560
         Width           =   975
      End
      Begin VB.CheckBox CkBBeamRight 
         Caption         =   "Träger von Rechts (FL)"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   3360
         Width           =   2535
      End
      Begin VB.TextBox TxBRt 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   2160
         TabIndex        =   33
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox TxBRh 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   720
         TabIndex        =   32
         Top             =   3720
         Width           =   615
      End
      Begin VB.ComboBox CBBRStahl 
         Height          =   315
         ItemData        =   "Form1.frx":8896
         Left            =   1800
         List            =   "Form1.frx":8898
         TabIndex        =   34
         Top             =   4080
         Width           =   975
      End
      Begin VB.CheckBox CkBBeamLeft 
         Caption         =   "Träger von Links (FL)"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox TxNX 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   720
         TabIndex        =   19
         Top             =   0
         Width           =   615
      End
      Begin VB.TextBox TxNZ 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   2160
         TabIndex        =   20
         Top             =   0
         Width           =   615
      End
      Begin VB.TextBox TxRX 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   2160
         TabIndex        =   23
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox TxLX 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   720
         TabIndex        =   22
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox TxRZ 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   1440
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox TxLZ 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label LbAbstLoch 
         AutoSize        =   -1  'True
         Caption         =   "Setze Abstand Loch"
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   4920
         Width           =   1440
      End
      Begin VB.Label LbAbstRand 
         AutoSize        =   -1  'True
         Caption         =   "Setze Abstand Rand"
         Height          =   195
         Left            =   120
         TabIndex        =   63
         Top             =   4560
         Width           =   1470
      End
      Begin VB.Label LbBRt 
         AutoSize        =   -1  'True
         Caption         =   "Dicke t"
         Height          =   195
         Left            =   1560
         TabIndex        =   62
         Top             =   3720
         Width           =   510
      End
      Begin VB.Label LbBRh 
         AutoSize        =   -1  'True
         Caption         =   "Höhe h"
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   3720
         Width           =   525
      End
      Begin VB.Label LbBRS 
         AutoSize        =   -1  'True
         Caption         =   "Stahlsorte"
         Height          =   195
         Left            =   120
         TabIndex        =   60
         Top             =   4080
         Width           =   705
      End
      Begin VB.Label LbNX 
         AutoSize        =   -1  'True
         Caption         =   "nx"
         Height          =   195
         Left            =   360
         TabIndex        =   59
         Top             =   0
         Width           =   165
      End
      Begin VB.Label LbNZ 
         AutoSize        =   -1  'True
         Caption         =   "nz"
         Height          =   195
         Left            =   1800
         TabIndex        =   58
         Top             =   0
         Width           =   165
      End
      Begin VB.Label LbRX 
         AutoSize        =   -1  'True
         Caption         =   "Rand-x"
         Height          =   195
         Left            =   1560
         TabIndex        =   57
         Top             =   720
         Width           =   510
      End
      Begin VB.Label LbLX 
         AutoSize        =   -1  'True
         Caption         =   "Loch-x"
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   720
         Width           =   480
      End
      Begin VB.Label LbRZ 
         AutoSize        =   -1  'True
         Caption         =   "Rand-z"
         Height          =   195
         Left            =   840
         TabIndex        =   55
         Top             =   360
         Width           =   510
      End
      Begin VB.Label LbLZ 
         AutoSize        =   -1  'True
         Caption         =   "Loch-z"
         Height          =   195
         Left            =   840
         TabIndex        =   54
         Top             =   1080
         Width           =   480
      End
   End
   Begin VB.CheckBox CkBBoltGroup 
      Caption         =   "Schraubengruppe"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CheckBox CkBIsVert 
      Caption         =   "Vert."
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   2640
      Width           =   735
   End
   Begin VB.ComboBox CBLochart 
      Height          =   315
      Left            =   840
      TabIndex        =   13
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CheckBox CkDrawHole 
      Caption         =   "Schraubenloch"
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CheckBox CkBPass 
      Caption         =   "Passschr."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox CBGüte 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6240
      TabIndex        =   48
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox CBDurchmesser 
      Height          =   315
      ItemData        =   "Form1.frx":889A
      Left            =   2040
      List            =   "Form1.frx":889C
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox PbSchraube 
      BackColor       =   &H00FFFFFF&
      Height          =   9855
      Left            =   7560
      ScaleHeight     =   653
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   437
      TabIndex        =   66
      Top             =   480
      Width           =   6615
   End
   Begin VB.Label Label10 
      Caption         =   "Offset x,z"
      Height          =   255
      Left            =   120
      TabIndex        =   77
      Top             =   10800
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Normalkraft [kN]"
      Height          =   255
      Left            =   120
      TabIndex        =   71
      Top             =   10440
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Querkraft [kN]"
      Height          =   255
      Left            =   120
      TabIndex        =   70
      Top             =   10080
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Moment [kNm]"
      Height          =   255
      Left            =   120
      TabIndex        =   69
      Top             =   9720
      Width           =   1575
   End
   Begin VB.Label LbZoom 
      Caption         =   "Massstab M 1:"
      Height          =   255
      Left            =   10560
      TabIndex        =   68
      Top             =   165
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Zeichne:"
      Height          =   255
      Left            =   120
      TabIndex        =   67
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Norm"
      Height          =   195
      Left            =   120
      TabIndex        =   65
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Lochart"
      Height          =   195
      Left            =   120
      TabIndex        =   52
      Top             =   2640
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Schraubengüte"
      Height          =   195
      Left            =   120
      TabIndex        =   50
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Schraubendurchmesser"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1680
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Datei"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Öffnen"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Speichern"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Speichern &unter..."
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "E&xport"
         Begin VB.Menu mnuFileExpResTxt 
            Caption         =   "Ergebnisse Text"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileExpResGrf 
            Caption         =   "Ergebnisse Grafik"
         End
         Begin VB.Menu mnuFileExpResTxtGrf 
            Caption         =   "Ergebnisse Text&&Grafik"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBsp1 
         Caption         =   "Bsp1"
      End
      Begin VB.Menu mnuFileBsp2 
         Caption         =   "Bsp2"
      End
      Begin VB.Menu mnuFileBsp3 
         Caption         =   "Bsp3"
      End
      Begin VB.Menu mnuFileBsp4 
         Caption         =   "Bsp4"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Be&enden"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   " ? "
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "Info"
      End
   End
End
Attribute VB_Name = "frmSchrauben"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'OM: 2016_10_24: werden jetzt in der doc-klasse SchrVerbFL verwendet
'Private m_N As Norm
'Private m_s As Schraube
'Private m_l As Schraubenloch 'Standardloch für alle Schrauben der Gruppe
'Private m_g As Schraubengruppe
'Private m_bl As Blech 'Blech links
'Private m_br As Blech 'Blech rechts
'Private m_nw As SchraubenNachweis
'Private m_Einwk_d  As EinwirkungsKombi

'dient als doc:
Private m_SVFL As SchrVerbFL


'om 2022-04-03 wenn die Option Schraubengruppe ausgeschaltet wird
'dann zuerst Schraubengruppe und Bleche hier sichern
Private m_SGr_old As Schraubengruppe
Private m_BlL_old As Blech
Private m_BlR_old As Blech

Private bInitText As Boolean
Private PbCanvas  As CairoPicBox
Private isUpdating As Boolean
Private m_FSO As cFSO
'OK das Programm erweitern?
'als Biegesteife Träger-Verbindung
'was soll alles möglich sein?
'Auswahl des Trägers
'Nein wir lassen das Programm erstmal so als Verbindung zwei Flachprofile
'oder als Verbindung Zange mit Flachprofil

'Warum brauchen wir überhaupt ein Changed event?
'verflixt umd warum wird der Nachweis  nicht aktualisiert wenn die Schnittsgößen verändert werden?

'Private Sub Form_Load()
'Private Sub UserForm_Initialize()
Private Sub Form_Initialize()
    MNew.IsInIDE = MNew.IsRunningInIDE(Me.hwnd)
    MMain.NewInit
    'OM: 2016_10_24 neu wird jetzt als doc verwendet:
    Set m_SVFL = New SchrVerbFL
    TBSchraube.ZOrder 0
    PbSchraube.ZOrder 0
    'LBSchraube.ZOrder 0
    'Text1.ZOrder 0
    'PbSchraube.ScaleMode = vbPixels
    Set PbCanvas = MNew.CairoPicBox(PbSchraube, Cairo)
    Me.Caption = "ScrewV is groovy - Trägerstoß biegesteif mit Laschen"
    
    ENorm_FillComboBox CbNorm
    'With CbNorm:      .AddItem "DIN 18800": .AddItem "EuroCode 3": End With
    With CBDurchmesser
        .AddItem "8":  .AddItem "10": .AddItem "12": .AddItem "16": .AddItem "20":
        .AddItem "22": .AddItem "24": .AddItem "27": .AddItem "30": .AddItem "36"
    End With
    'CBDurchmesser.ListIndex = 7
    InitCBGüte
    ELochart_FillComboBox CBLochart
    'With CBLochart:   .AddItem "Normal": .AddItem "Übergroß": .AddItem "Langloch Kurz": .AddItem "Langloch Lang":    End With
    EAbstand_FillComboBox CmBAbstLoch
    EAbstand_FillComboBox CmBAbstRand
    'With CbGleitGZ:   .AddItem "GZ-Gebrauchst.": .AddItem "GZ-Tragfähigkeit": End With
    With CbGleitGZ: .AddItem "GZ-G": .AddItem "GZ-T": End With
    With CbGFKmue: .AddItem "u=0,5": .AddItem "u=0,4": .AddItem "u=0,3": .AddItem "u=0,2": End With
    CbGFKmue.ToolTipText = "Reibzahl mue"
    CbGleitGZ.Enabled = False 'CkGleitf.Value = vbChecked
    CbGFKmue.Enabled = False 'CkGleitf.Value = vbChecked
    ComboFillArr CbZoom, Array(100, 75, 66.667, 50, 40, 33.333, 30, 25, 20, 16.667, 15, 13.333, 12.5, 11.111, _
                               10#, 7.5, 6.667, 5#, 4#, 3.333, 3#, 2.5, 2#, 1.667, 1.5, 1.333, 1.25, 1.111, _
                               1#, 0.75, 0.667, 0.5, 0.4, 0.333, 0.3, 0.25)
    CbZoom.ListIndex = 28
    'CbZoom.Text = "1.0"
    CkZange.Enabled = True
    'CbNorm.ListIndex = 0
    'CBDurchmesser.ListIndex = 2
    'CBGüte.ListIndex = 0
    'CkBPass.Value = vbUnchecked
    'CBLochart.ListIndex = 0
    CreateNorm
    InitCBStahl
    CreateSchraube
    CreateSchraubenLoch
    CreateEd
    CreateNw
    'Schraubengruppe und Bleche existieren hier noch nicht
    PnlSG_Enabled = False
    InitEPTipps
    'OpCalc.Value = True ' vbChecked 'pseudo TabControl
    OpErgGra.Value = True
    Call UpdateView
End Sub
Private Sub Form_Resize()
    Dim L, t, W, H, brdr
    brdr = 8 * Screen.TwipsPerPixelX
    L = TBSchraube.Left
    t = TBSchraube.Top
    If OpErgGra.Value Then
        W = 4695 'Me.ScaleWidth - l - brdr
        H = Me.ScaleHeight - t - brdr
        If W > 0 And H > 0 Then TBSchraube.Move L, t, W, H
        L = L + W
        W = Me.ScaleWidth - L - brdr
        If W > 0 And H > 0 Then PbSchraube.Move L, t, W, H
    Else
        W = Me.ScaleWidth - L - brdr
        H = Me.ScaleHeight - t - brdr
        If W > 0 And H > 0 Then TBSchraube.Move L, t, W, H
        If W > 0 And H > 0 Then PbSchraube.Move L, t, W, H
    End If
    PbSchraube.Refresh
End Sub

Private Sub BtnOptions_Click()
    Dim s As String: s = PbCanvas.ScreenDiagonaleInch ' MCDraw.ScreenDiag
    s = InputBox("Geben Sie ihre Bildschirmdiagonale in Zoll ein: ", "Bildschirmdiagonale", s)
    If StrPtr(s) = 0 Then Exit Sub
    Dim d As Double
    If Double_TryParse(s, d) Then
        PbCanvas.ScreenDiagonaleInch = d
    End If
    'MCDraw.ScreenDiag = Val(s)
    'InitScale
    'UpdateView
End Sub
Private Sub BtnExportGrafik_Click()
    Dim pdf As CairoPdfDoc: Set pdf = MNew.CairoPdfDoc(MMain.Cairo, IIf(CkPdfQuer.Value = vbChecked, poLandscape, poPortrait), pfDIN_A4, PbCanvas.ZoomFactor)
    'Dim t As Double: t = MCDraw.PixProMM
    'MCDraw.PixProMM = pdf.PunkteProMM
    Call DrawSystem(pdf.Canvas, m_SVFL, pdf.PageWith(euMM) / 2, pdf.PageHeight(euMM) / 2) 'm_s, m_l, m_g, m_bl, m_br)
    pdf.WriteToFile "C:\users\Oliver Meyer\Documents\test.pdf"
    'MCDraw.PixProMM = t
End Sub
Private Sub CbZoom_Click()
    If isUpdating Then Exit Sub
    Dim s As String: s = CbZoom.Text
    Dim d As Double
    If Double_TryParse(s, d) Then 'MCDraw.ZoomFact = 1 / d
        PbCanvas.ZoomFactor = 1 / d
    End If
    LbZoom.Caption = "Masstab: 1:" & s
    UpdateView
End Sub

Private Sub Draw()
    If CkDrawHole.Value = vbChecked Then
        'Call MCDraw.DrawSystem(PbCanvas.Canvas, PbCanvas.PicBox.ScaleWidth, PbCanvas.PicBox.ScaleHeight, m_s, m_L, m_g, m_bl, m_br)
        'Call MCDraw.DrawSystem(PbCanvas.Canvas, PbCanvas.Surface.Width, PbCanvas.Surface.Height, m_s, m_L, m_g, m_bl, m_br)
        Call MCDraw.DrawSystem(PbCanvas.Canvas, m_SVFL, PbCanvas.CenterX, PbCanvas.CenterY) ' m_s, m_l, m_g, m_bl, m_br)
    Else
        'Call MCDraw.DrawSystem(PbCanvas.Canvas, PbCanvas.PicBox.ScaleWidth, PbCanvas.PicBox.ScaleHeight, m_s, Nothing, m_g, m_bl, m_br)
        'Call MCDraw.DrawSystem(PbCanvas.Canvas, PbCanvas.Surface.Width, PbCanvas.Surface.Height, m_s, Nothing, m_g, m_bl, m_br)
        Call MCDraw.DrawSystem(PbCanvas.Canvas, m_SVFL, PbCanvas.CenterX, PbCanvas.CenterY) 'm_s, Nothing, m_g, m_bl, m_br)
    End If
End Sub


Private Sub CkRoundUp5_Click()
    If isUpdating Then Exit Sub
    CmBAbstLoch_Click
    CmBAbstRand_Click
End Sub

Private Sub CkUpdateAbstand_Click()
    'OK wenn man ein Beispiel lädt kommts hier zum problem weil diese Option nirgends gespeichert ist
    If isUpdating Then Exit Sub
    CmBAbstLoch_Click
    CmBAbstRand_Click
End Sub

Private Sub CkZange_Click()
    If isUpdating Then Exit Sub
    m_SVFL.BlechLinks.IsZange = CkZange.Value = vbChecked
    m_SVFL.BlechLinks.IsMehrschnittig = CkZange.Value = vbChecked
    'm_SVFL.BlechLinks.IsMehrschnittig = CkZange.Value = vbChecked
    
    'If Not m_br Is Nothing Then m_br.IsMehrschnittig = True
    'CalcSchraubenlänge
    UpdateView
End Sub


Private Sub Command2_Click()
    Dim d As Double
    Dim v As Double
    
    d = 2.5
    v = Min(d, MMath.negINF)
    'MsgBox v
    
    v = Min(d, MMath.posINF)
    'MsgBox v
    
    d = -2.5
    v = Min(d, MMath.negINF)
    'MsgBox v
    
    v = Min(d, MMath.posINF)
    'MsgBox v
    
    'MsgBox IsPositive(MMath.posINF)
    
    'MsgBox IsPositive(MMath.negINF)
    v = MMath.negINF
    'If MMath.posINF > 0 Then MsgBox "größer 0"
    If v > 0 Then MsgBox "größer 0"
    'If MMath.negINF < 0 Then MsgBox "kleiner 0"
End Sub

Private Sub Form_Terminate()
    If Forms.Count = 0 Then New_c.CleanupRichClientDll
End Sub

Private Sub mnuFileBsp1_Click()
    Set m_SVFL = MBsps.Bsp1(isUpdating) ',  ' m_N, m_s, m_l, m_Einwk_d, m_bl, m_br, m_g, m_nw
    'entwede den Befehl durchführen oder hier die Option zuerst ausschalten
    CkUpdateAbstand.Value = vbUnchecked
    'CkUpdateAbstand_Click
    UpdateView
End Sub
Private Sub mnuFileBsp2_Click()
    Set m_SVFL = MBsps.Bsp2(isUpdating) ', m_SVFL ' m_N, m_s, m_l, m_Einwk_d, m_bl, m_br, m_g, m_nw
    'CkUpdateAbstand_Click
    CkUpdateAbstand.Value = vbUnchecked
    UpdateView
End Sub
Private Sub mnuFileBsp3_Click()
    Set m_SVFL = MBsps.Bsp3(isUpdating) ', m_SVFL ' m_N, m_s, m_l, m_Einwk_d, m_bl, m_br, m_g, m_nw
    CkUpdateAbstand.Value = vbUnchecked
    'CkUpdateAbstand_Click
    UpdateView
End Sub
Private Sub mnuFileBsp4_Click()
    Set m_SVFL = MBsps.Bsp4(isUpdating) ', m_SVFL ' m_N, m_s, m_l, m_Einwk_d, m_bl, m_br, m_g, m_nw
    CkUpdateAbstand.Value = vbUnchecked
    'CkUpdateAbstand_Click
    UpdateView
End Sub

Private Sub mnuFileExpResGrf_Click()
    Dim pdf As CairoPdfDoc: Set pdf = MNew.CairoPdfDoc(MMain.Cairo, IIf(CkPdfQuer.Value = vbChecked, poLandscape, poPortrait), pfDIN_A4, PbCanvas.ZoomFactor)
    'Dim t As Double: t = MCDraw.PixProMM
    'MCDraw.PixProMM = pdf.PunkteProMM
    Call DrawSystem(pdf.Canvas, m_SVFL, pdf.PageWith(euMM) / 2, pdf.PageHeight(euMM) / 2) ' m_s, m_l, m_g, m_bl, m_br)
    
    If m_FSO Is Nothing Then Set m_FSO = New_c.FSO
    Dim DefDir As String
    DefDir = m_FSO.CurrentDirectory
    DefDir = m_FSO.GetSpecialFolder(CSIDL_MYDOCUMENTS)
    Dim Flt As String: Flt = "pdf Ergebnisse Grafik|*.pdf|Alle Dateien|*.*"
    Dim FNm As String: FNm = m_FSO.ShowSaveDialog(, DefDir, , "XScrew.xsx", Flt, "XSx", Me.hwnd)
    If StrPtr(FNm) Then
        pdf.WriteToFile FNm
    End If
    'MCDraw.PixProMM = t

End Sub

Private Sub mnuFileOpen_Click()
    'den Datei-Öffnen-Dialog anzeigen
    If m_FSO Is Nothing Then Set m_FSO = New_c.FSO
    Dim DefDir As String
    DefDir = m_FSO.CurrentDirectory
    DefDir = m_FSO.GetSpecialFolder(CSIDL_MYDOCUMENTS)
    'Dim Ttl as String: ttl = "
    Dim Flt As String: Flt = "Schrauben-Dateien|*.XSx|Alle Dateien|*.*"
    Dim FNm As String: FNm = m_FSO.ShowOpenDialog(OFN_ALLOWMULTISELECT, DefDir, , "XScrew.xsx", Flt, "XSx", Me.hwnd)
    'If Len(FNm) > 0 Then
        'MsgBox FNm
    'End If
End Sub

Private Sub mnuFileSave_Click()
    '
End Sub

Private Sub mnuFileSaveAs_Click()
    'den Datei-Speichern-Unter-Dialog anzeigen
    'auch die export-Formate svg und pdf dxf?
    'den Datei-Speichern-unter-Dialog anzeigen
    Dim FNm As String
    Dim DefDir As String
    If m_FSO Is Nothing Then Set m_FSO = New_c.FSO
    DefDir = m_FSO.CurrentDirectory
    DefDir = m_FSO.GetSpecialFolder(CSIDL_MYDOCUMENTS)
    'Dim Ttl as String: ttl = "
    'nö der Filter hier ist irgendwie unsauberer Pfusch
    Dim Flt As String: Flt = "Schrauben-Dateien|*.XSx|pdf Ergebnisse|*.Erg_pdf|pdf nur Grafik|*.Grf_pdf|pdf Ergeb.+Grafik|*.ErGr_pdf|Alle Dateien|*.*"
    FNm = m_FSO.ShowSaveDialog(, DefDir, , "XScrew.xsx", Flt, "XSx", Me.hwnd)
    If Len(FNm) > 0 Then
    '    MsgBox FNm
        'MSer.JSONSerialize FNm
    End If
End Sub
Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpInfo_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub PbSchraube_Click()
    CbZoom.SetFocus
End Sub

Private Sub PbSchraube_Paint()
    Draw
    PbCanvas.DrawAll
End Sub

Private Sub ComboFillArr(aCB As ComboBox, arr)
    Dim v: For Each v In arr: aCB.AddItem Str(v): Next
End Sub
Private Sub UpdateView()
    'CkBScrwHole.Value = Abs(Not CkBPass.Value)
    'CkBScrwHole.Enabled = Not CkBPass.Value
    'Static i
    'i = i + 1
    'Debug.Print i
    Dim sb As cStringBuilder: Set sb = New_c.StringBuilder
    isUpdating = True
    'LBSchraube.Clear
    Dim n As Norm: Set n = m_SVFL.Norm
    If Not n Is Nothing Then
        With n
            'erst den linken View aktualisieren
            CbNorm.Text = ENorm_ToStr(.ENorm)
            'dann die ListBox aktualisieren
            sb.AppendNL .ToStr 'ListBox LBSchraube
        End With
    End If
    Dim s As Schraube: Set s = m_SVFL.Schraube
    If Not s Is Nothing Then
        With s
            CkSFSchaft.Enabled = Not .IsPassschraube
            CbGleitGZ.Enabled = .IsGleitfest ' CkGleitf.Value = vbChecked
            CkVorgesp.Enabled = Not .IsGleitfest '.IsVorgespannt ' Not (CkGleitf.Value = vbChecked)
            CbGFKmue.Enabled = .IsGleitfest ' CkGleitf.Value = vbChecked
            CkZug.Enabled = Not .IsGleitfest  ' Not (CkGleitf.Value = vbChecked)
            
            CBDurchmesser.Text = .Durchmesser
            CBGüte.Text = Trim$(Str$(.Schraubengüte))
            CkBPass.Value = Abs(.IsPassschraube)
            CkSenk.Value = Abs(.IsSenkschraube)
            CkSFSchaft.Value = Abs(.IsScherfugeSchaft)
            CkGleitf.Value = Abs(.IsGleitfest)
            CbGleitGZ.Text = IIf(.IsGleitfestImGZT, "GZ-T", "GZ-G")
            CkZug.Value = Abs(.IsZugverbindung)
            CkVorgesp.Value = Abs(.IsVorgespannt)
            sb.AppendNL .ToStr
        End With
        Dim L As Schraubenloch: Set L = m_SVFL.Schraubenloch
        If Not L Is Nothing Then
            With L
                'Nur dann Vertikal einblenden wenn Langloch
                CkBIsVert.Enabled = (.Lochart = LanglochKurz) Or (.Lochart = LanglochLang)
                CkBIsVert.Value = Abs(.IsVertikal)
                CBLochart.Text = ELochart_ToStr(.Lochart)
                CbGFKmue.Text = "u=" & CStr(.Reibzahl_mue)
                sb.AppendNL .ToStr '.ToListBox LBSchraube
            End With
        End If
        Dim ewkd As EinwirkungsKombi: Set ewkd = m_SVFL.EinwirkungsKombi
        If Not ewkd Is Nothing Then
            With ewkd
                TxMomentEd.Text = .MomentEd
                TxQuerkraftEd.Text = .QuerkraftEd
                TxNormalkraftEd.Text = .NormalkraftEd
                TxOffX.Text = .OffX
                TxOffZ.Text = .OffZ
                sb.AppendNL .ToStr '.ToListBox LBSchraube
            End With
        End If
        Dim g As Schraubengruppe: Set g = m_SVFL.Schraubengruppe
        PnlSG_Enabled = (Not g Is Nothing)

        If Not g Is Nothing Then
            With g
                CmBAbstLoch.ListIndex = .EAbstLoch ' 0
                CmBAbstRand.ListIndex = .EAbstRand ' 0
                TxNX.Text = .Schraubenliste.N1
                TxNZ.Text = .Schraubenliste.N2
                TxRX.Text = .AbstandSel.Rand.X
                TxRZ.Text = .AbstandSel.Rand.Z
                TxLX.Text = .AbstandSel.Loch.X
                TxLZ.Text = .AbstandSel.Loch.Z
                sb.AppendNL .ToStr '.ToListBox LBSchraube
                TxLX.Enabled = Not (.Schraubenliste.N1 = 1#)
                TxLZ.Enabled = Not (.Schraubenliste.N2 = 1#)
            End With
            Dim bl As Blech: Set bl = m_SVFL.BlechLinks
            If Not bl Is Nothing Then
                With bl
                    TxBLt.Text = .Blechdicke
                    TxBLh.Text = .Höhe
                    CBBLStahl.Text = EStahlsorte_ToStr(.Stahlsorte, n.ENorm)
                    CkZange.Value = IIf(.IsZange, vbChecked, vbUnchecked)
                    sb.AppendNL .ToStr '.ToListBox LBSchraube
                End With
            End If
            Dim br As Blech: Set br = m_SVFL.BlechRechts
            If Not br Is Nothing Then
                With br
                    TxBRt.Text = .Blechdicke
                    TxBRh.Text = .Höhe
                    CBBRStahl.Text = EStahlsorte_ToStr(.Stahlsorte, n.ENorm)
                    sb.AppendNL .ToStr '.ToListBox LBSchraube
                End With
            End If
            PnlAbstLR_Enabled = (Not bl Is Nothing) Or (Not br Is Nothing)
        End If
        PnlBR_Enabled = (Not br Is Nothing)
        PnlBL_Enabled = (Not bl Is Nothing)
        
        Dim nw As SchraubenNachweis: Set nw = m_SVFL.SchraubenNachweis
        With nw
            sb.AppendNL .ToStr '.ToListBox LBSchraube
        End With
    End If
    TBSchraube.Text = sb.ToString
    PbSchraube.Refresh
    
    isUpdating = False
End Sub

Sub InitCBGüte()
    Call MESFK.ESFK_FillComboBox(CBGüte, Not (CkSFKAlle.Value = vbChecked), CkVorgesp.Value = vbChecked)
End Sub
Sub InitCBStahl()
    With m_SVFL.Norm
        EStahlsorte_FillComboBox CBBLStahl, .ENorm
        EStahlsorte_FillComboBox CBBRStahl, .ENorm
    End With
End Sub
Sub InitEPTipps()
    If m_SVFL.Norm.ENorm = ENorm.Norm_DIN18800 Then
        Call TTxt(LbRX, TxRX, "e_1"):    Call TTxt(LbRZ, TxRZ, "e_2")
        Call TTxt(LbLX, TxLX, "e_0"):    Call TTxt(LbLZ, TxLZ, "e_3")
    Else
        Call TTxt(LbRX, TxRX, "e_1"):    Call TTxt(LbRZ, TxRZ, "e_2")
        Call TTxt(LbLX, TxLX, "p_1"):    Call TTxt(LbLZ, TxLZ, "p_2")
    End If
End Sub
Sub TTxt(lb As Label, tx As TextBox, s As String)
    lb.ToolTipText = s: tx.ToolTipText = s
End Sub


Sub CreateNorm()
    With m_SVFL
        If Not .Norm Is Nothing Then Exit Sub
        Set .Norm = MNew.Norm(MNew.NormFromENorm(ENorm.Norm_EuroCode3))
    End With
End Sub
Sub CreateSchraube()
    With m_SVFL
        If Not .Schraube Is Nothing Then Exit Sub
        Set .Schraube = MNew.Schraube(.Norm, 12, 4.6)
    End With
End Sub
Sub CreateSchraubenLoch()
    With m_SVFL
        If .Schraube Is Nothing Then Exit Sub
        Set .Schraubenloch = MNew.Schraubenloch(.Schraube, ELochart.Normal)
    End With
End Sub
Sub CreateSchraubenGruppe()
    With m_SVFL
        If .Schraubenloch Is Nothing Then Exit Sub
        Set .Schraubengruppe = MNew.Schraubengruppe(.Schraubenloch, 2, 1, _
                                                    MNew.AbstandLR( _
                                                        MNew.VectorXZ(30, 30), _
                                                        MNew.VectorXZ(30, 30)), _
                                                    AbstandMinVol, _
                                                    AbstandMinVol, .EinwirkungsKombi, Nothing, Nothing)
        Set .SchraubenNachweis.Schraubengruppe = .Schraubengruppe
    End With
End Sub
Sub CreateBlechL()
    'If m_SVFL Is Nothing Then Exit Sub
    With m_SVFL
        Set .BlechLinks = IIf(CkBBeamLeft.Value = vbChecked, MNew.Blech(.Norm, S235, 10, 0, 100, True, False, False), Nothing)
        If .Schraubengruppe Is Nothing Then Exit Sub
        Set .Schraubengruppe.TrägerLinks = .BlechLinks
    End With
End Sub
Sub CreateBlechR()
    'If m_SVFL Is Nothing Then Exit Sub
    With m_SVFL
        Set .BlechRechts = IIf(CkBBeamRight.Value = vbChecked, MNew.Blech(.Norm, S235, 10, 0, 100, False, False, False), Nothing)
        If .Schraubengruppe Is Nothing Then Exit Sub
        Set .Schraubengruppe.TrägerRechts = m_SVFL.BlechRechts
    End With
End Sub
Sub CreateEd()
    Set m_SVFL.EinwirkungsKombi = MNew.EinwirkungsKombi(0, 0, 0, 0, 0)
End Sub
Sub CreateNw()
    With m_SVFL
        Set .SchraubenNachweis = MNew.SchraubenNachweis(.Norm, .Schraube, .Schraubengruppe)
    End With
End Sub

Private Sub Command1_Click()
    Dim t As String: t = Text1.Text
    If Len(t) = 0 Then t = Clipboard.GetText
    Dim ta() As String: ta = Split(t, vbCrLf)
    Dim i As Long
    For i = LBound(ta) + 1 To UBound(ta)
        If Len(ta(i)) <> 0 Then
            If Len(ta(i / 2)) = 0 Then
                ta(i / 2) = ta(i)
                ta(i) = ""
            End If
        End If
    Next
    ReDim Preserve ta(UBound(ta) / 2)
    Text1.Text = Join(ta(), vbCrLf)
    Clipboard.SetText Text1.Text
End Sub

Private Sub CbNorm_Click()
    With m_SVFL
        If .Norm Is Nothing Or isUpdating Then Exit Sub
        Set .Norm = MNew.Norm(MNew.NormFromENorm(CbNorm.ListIndex))
        InitCBStahl
        InitEPTipps
        Call UpdateView
        If Not .BlechLinks Is Nothing Then CBBLStahl.ListIndex = .BlechLinks.Stahlsorte
        If Not .BlechRechts Is Nothing Then CBBRStahl.ListIndex = .BlechRechts.Stahlsorte
    End With
End Sub

Private Sub CBDurchmesser_Change()
    With m_SVFL
        If .Schraube Is Nothing Or isUpdating Then Exit Sub
        Dim d As Double
        If Double_TryParse(CBDurchmesser.Text, d) Then .Schraube.Durchmesser = d
        If CkUpdateAbstand.Value = vbChecked Then
            CmBAbstLoch_Click
            CmBAbstRand_Click
        End If
    End With
    Call UpdateView
End Sub
Private Sub CBDurchmesser_Click()
    With m_SVFL
        If .Schraube Is Nothing Or isUpdating Then Exit Sub
        Dim d As Double
        If Double_TryParse(CBDurchmesser.Text, d) Then .Schraube.Durchmesser = d
        If CkUpdateAbstand.Value = vbChecked Then
            CmBAbstLoch_Click
            CmBAbstRand_Click
        End If
    End With
    Call UpdateView
End Sub
Private Sub CBGüte_Click()
    'Dim s As Schraube: Set s = m_SVFL.Schraube
    With m_SVFL
        If .Schraube Is Nothing Or isUpdating Then Exit Sub
        Dim d As Double
        If Double_TryParse(CBGüte.Text, d) Then .Schraube.Schraubengüte = d
    End With
    Call UpdateView
End Sub
Private Sub CkSFKAlle_Click()
    If isUpdating Then Exit Sub
    InitCBGüte
    UpdateView
End Sub

Private Sub CkBPass_Click()
    'Dim s As Schraube: Set s = m_SVFL.Schraube
    With m_SVFL
        If .Schraube Is Nothing Or isUpdating Then Exit Sub
        .Schraube.IsPassschraube = CkBPass.Value = vbChecked
    End With
    Call UpdateView
End Sub
Private Sub CkSenk_Click()
    With m_SVFL
        If .Schraube Is Nothing Or isUpdating Then Exit Sub
        .Schraube.IsSenkschraube = CkSenk.Value = vbChecked
    End With
    Call UpdateView
End Sub
Private Sub CkSFSchaft_Click()
    With m_SVFL
        If .Schraube Is Nothing Or isUpdating Then Exit Sub
        .Schraube.IsScherfugeSchaft = CkSFSchaft.Value = vbChecked
    End With
    UpdateView
End Sub
Private Sub CkGleitf_Click()
    With m_SVFL
        If .Schraube Is Nothing Or isUpdating Then Exit Sub
        .Schraube.IsGleitfest = CkGleitf.Value = vbChecked
    End With
    'm_S.IsVorgespannt = CkGleitf.Value = vbChecked
    'm_S.IsZugverbindung = False 'CkGleitf.Value = vbChecked
    If CbGleitGZ.ListIndex = -1 Then CbGleitGZ.ListIndex = 0
    'CkVorgesp.Value = CkGleitf.Value
    'CkZug.Value = vbUnchecked 'CkGleitf.Value = vbChecked
    InitCBGüte
    Call CheckSchrGüte(CkGleitf.Value = vbChecked)
    UpdateView
End Sub
Private Sub CbGleitGZ_Click()
    'Dim s As Schraube: Set s = m_SVFL.Schraube
    With m_SVFL
        If .Schraube Is Nothing Or isUpdating Then Exit Sub
        .Schraube.IsGleitfestImGZT = CbGleitGZ.ListIndex = 1
    End With
    If CbGleitGZ.ListIndex = -1 Then CbGleitGZ.ListIndex = 0
    UpdateView
End Sub
Private Sub CbGFKmue_Click()
    With m_SVFL
        If .Schraubenloch Is Nothing Or isUpdating Then Exit Sub
    'hier den Text parsen
        Dim s As String: s = CbGFKmue.Text
        Dim Pos As Long: Pos = InStr(1, s, "=")
        Dim mue As Double
        If Double_TryParse(Mid(s, Pos + 1, Len(s) - Pos), mue) Then .Schraubenloch.Reibzahl_mue = mue
    End With
    UpdateView
End Sub

Private Sub CkZug_Click()
    With m_SVFL
        If .Schraube Is Nothing Or isUpdating Then Exit Sub
        .Schraube.IsZugverbindung = CkZug.Value = vbChecked
    End With
    UpdateView
End Sub
Private Sub CkVorgesp_Click()
    With m_SVFL
        If .Schraube Is Nothing Or isUpdating Then Exit Sub
        .Schraube.IsVorgespannt = CkVorgesp.Value = vbChecked
    End With
    InitCBGüte
    Call CheckSchrGüte(CkVorgesp.Value = vbChecked)
    UpdateView
End Sub
Private Sub CheckSchrGüte(ByVal vorgesp As Boolean)
    With m_SVFL
        If isUpdating Then Exit Sub
        If vorgesp And (m_SVFL.Schraube.Schraubengüte < 8.8) Then
            .Schraube.Schraubengüte = 8.8
            MsgBox "Die Schraubengüte wurde geändert auf: " & Str(.Schraube.Schraubengüte)
        End If
    End With
End Sub
Private Sub CBLochart_Click()
    With m_SVFL
        If .Schraubenloch Is Nothing Or isUpdating Then Exit Sub
        .Schraubenloch.Lochart = CBLochart.ListIndex
    End With
    Call UpdateView
End Sub
Private Sub CkBIsVert_Click()
    With m_SVFL
        If .Schraubenloch Is Nothing Or isUpdating Then Exit Sub
        .Schraubenloch.IsVertikal = CkBIsVert.Value = vbChecked
    End With
    Call UpdateView
End Sub
Private Sub CkDrawHole_Click()
    If isUpdating Then Exit Sub
    Call UpdateView
End Sub
Private Sub CkDrawUScheibe_Click()
    If isUpdating Then Exit Sub
    Call UpdateView
End Sub
Private Sub CkDrawMutter_Click()
    If isUpdating Then Exit Sub
    UpdateView
End Sub

Private Sub CkBBoltGroup_Click()
    With m_SVFL
        If isUpdating Then Exit Sub
        If Not CkBBoltGroup.Value = vbChecked Then
            'Schraubengruppe vorher sichern!
            'und bei Checked wieder die gleichen Einstellungen herstellen
            Set m_SGr_old = .Schraubengruppe
            Set m_BlL_old = .BlechLinks
            Set m_BlR_old = .BlechRechts
            
            Set .Schraubengruppe = Nothing
            Set .BlechLinks = Nothing:  'CkBBeamLeft.Value = vbUnchecked
            Set .BlechRechts = Nothing: 'CkBBeamRight.Value = vbUnchecked
        Else
            If m_SGr_old Is Nothing Then
                Call CreateEd
                Call CreateSchraubenGruppe
            Else
                Set .Schraubengruppe = m_SGr_old
                Set .BlechLinks = m_BlL_old
                Set .BlechRechts = m_BlR_old
                
                'CkBBeamLeft.Value = vbChecked
                'CkBBeamRight.Value = vbChecked
            End If
        End If
    End With
    UpdateView
End Sub
Private Property Let PnlSG_Enabled(ByVal ben As Boolean)
    'ben = be enabled
    PnlSG.Enabled = ben
    LbNX.Enabled = ben: TxNX.Enabled = ben
    LbNZ.Enabled = ben: TxNZ.Enabled = ben
    LbRZ.Enabled = ben: TxRZ.Enabled = ben
    LbRX.Enabled = ben: TxRX.Enabled = ben
    LbLZ.Enabled = ben: TxLZ.Enabled = ben
    LbLX.Enabled = ben: TxLX.Enabled = ben
    CkRoundUp5.Enabled = ben
    CkBBeamLeft.Enabled = ben
    PnlBL_Enabled = CkBBeamLeft.Value = vbChecked
    CkBBeamRight.Enabled = ben
    PnlBR_Enabled = CkBBeamRight.Value = vbChecked
    PnlAbstLR_Enabled = CkBBeamLeft.Value = vbChecked Or CkBBeamRight.Value = vbChecked
    'wenn es eine Schraubengruppe gibt dann kann Abstand Loch bereits enabled sein
    'Abstand Rand kann erst enanbled sein wenn Träger vorhanden
    LbAbstLoch.Enabled = ben
    CmBAbstLoch.Enabled = ben
    CkUpdateAbstand.Enabled = ben
End Property
Private Property Let PnlBL_Enabled(ByVal ben As Boolean)
    LbBLh.Enabled = ben: TxBLh.Enabled = ben
    LbBLt.Enabled = ben: TxBLt.Enabled = ben
    LbBLS.Enabled = ben: CBBLStahl.Enabled = ben
    CkZange.Enabled = ben 'CkBBeamLeft.Value = vbChecked And CkBBeamRight.Value = vbChecked
End Property
Private Property Let PnlBR_Enabled(ByVal ben As Boolean)
    LbBRh.Enabled = ben: TxBRh.Enabled = ben
    LbBRt.Enabled = ben: TxBRt.Enabled = ben
    LbBRS.Enabled = ben: CBBRStahl.Enabled = ben
    CkZange.Enabled = CkBBeamLeft.Value = vbChecked And CkBBeamRight.Value = vbChecked
End Property
Private Property Let PnlAbstLR_Enabled(ByVal ben As Boolean)
    LbAbstLoch.Enabled = ben: CmBAbstLoch.Enabled = ben
    LbAbstRand.Enabled = ben: CmBAbstRand.Enabled = ben
    CkUpdateAbstand.Enabled = ben
End Property

Private Sub OpCalc_Click()
    'LBSchraube.ZOrder 0
    TBSchraube.ZOrder 0
    PbSchraube.Visible = False
    Form_Resize
End Sub
Private Sub OpGrafik_Click()
    PbSchraube.Visible = True
    PbSchraube.ZOrder 0
    Form_Resize
End Sub
Private Sub OpErgGra_Click()
    PbSchraube.Visible = True
    PbSchraube.ZOrder 0
    Form_Resize
End Sub

Private Sub TxMomentEd_KeyDown(KeyCode As Integer, Shift As Integer)
    UpdateDataNumeric TxMomentEd, KeyCode, m_SVFL.EinwirkungsKombi, "MomentEd"
End Sub
Private Sub TxMomentEd_LostFocus()
    UpdateDataNumeric TxMomentEd, vbKeyTab, m_SVFL.EinwirkungsKombi, "MomentEd"
End Sub

'Private Sub TxMomentEd_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Not ((KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab)) Then Exit Sub
'    If isUpdating Then Exit Sub
'    Dim d As Double
'    If Double_TryParse(TxMomentEd, d) Then
'        m_Einwk_d.MomentEd = d
'    End If
'    UpdateView
'End Sub

Private Sub TxOffX_KeyDown(KeyCode As Integer, Shift As Integer)
    UpdateDataNumeric TxOffX, KeyCode, m_SVFL.EinwirkungsKombi, "OffX"
End Sub
Private Sub TxOffX_LostFocus()
    UpdateDataNumeric TxOffX, vbKeyTab, m_SVFL.EinwirkungsKombi, "OffX"
End Sub
'Private Sub TxOffX_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Not ((KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab)) Then Exit Sub
'    If isUpdating Then Exit Sub
'    Dim d As Double
'    If Double_TryParse(TxOffX, d) Then
'        m_Einwk_d.OffX = d
'    End If
'    UpdateView
'End Sub
Private Sub TxOffZ_KeyDown(KeyCode As Integer, Shift As Integer)
    UpdateDataNumeric TxOffZ, KeyCode, m_SVFL.EinwirkungsKombi, "OffZ"
End Sub
Private Sub TxOffZ_LostFocus()
    UpdateDataNumeric TxOffZ, vbKeyTab, m_SVFL.EinwirkungsKombi, "OffZ"
End Sub
'Private Sub TxOffZ_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Not ((KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab)) Then Exit Sub
'    If isUpdating Then Exit Sub
'    Dim d As Double
'    If Double_TryParse(TxOffZ, d) Then
'        m_Einwk_d.OffZ = d
'    End If
'    UpdateView
'End Sub
'
Private Sub TxQuerkraftEd_KeyDown(KeyCode As Integer, Shift As Integer)
    UpdateDataNumeric TxQuerkraftEd, KeyCode, m_SVFL.EinwirkungsKombi, "QuerkraftEd"
End Sub
Private Sub TxQuerkraftEd_LostFocus()
    UpdateDataNumeric TxQuerkraftEd, vbKeyTab, m_SVFL.EinwirkungsKombi, "QuerkraftEd"
End Sub

'Private Sub TxQuerkraftEd_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Not ((KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab)) Then Exit Sub
'    If isUpdating Then Exit Sub
'    Dim d As Double
'    If Double_TryParse(TxQuerkraftEd, d) Then
'        m_Einwk_d.QuerkraftEd = d
'    End If
'    UpdateView
'End Sub
Private Sub TxNormalkraftEd_KeyDown(KeyCode As Integer, Shift As Integer)
    UpdateDataNumeric TxNormalkraftEd, KeyCode, m_SVFL.EinwirkungsKombi, "NormalkraftEd"
End Sub
Private Sub TxNormalkraftEd_LostFocus()
    UpdateDataNumeric TxNormalkraftEd, vbKeyTab, m_SVFL.EinwirkungsKombi, "NormalkraftEd"
End Sub
'Private Sub TxNormalkraftEd_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Not ((KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab)) Then Exit Sub
'    If isUpdating Then Exit Sub
'    Dim d As Double
'    If Double_TryParse(TxNormalkraftEd, d) Then
'        m_Einwk_d.NormalkraftEd = d
'    End If
'    UpdateView
'End Sub

Private Sub TxNX_KeyDown(KeyCode As Integer, Shift As Integer)
    UpdateDataNumeric TxNX, KeyCode, m_SVFL.Schraubengruppe.Schraubenliste, "N1", True
End Sub
Private Sub TxNX_LostFocus()
    UpdateDataNumeric TxNX, vbKeyTab, m_SVFL.Schraubengruppe.Schraubenliste, "N1", True
End Sub
Private Sub TxNZ_KeyDown(KeyCode As Integer, Shift As Integer)
    UpdateDataNumeric TxNZ, KeyCode, m_SVFL.Schraubengruppe.Schraubenliste, "N2", True
End Sub
Private Sub TxNZ_LostFocus()
    UpdateDataNumeric TxNZ, vbKeyTab, m_SVFL.Schraubengruppe.Schraubenliste, "N2", True
End Sub

'Private Sub TxNX_Change()
'    If m_g Is Nothing Or Not IsNumeric(TxNX) Or isUpdating Then Exit Sub
'    Dim d As Double
'    If Double_TryParse(TxNX, d) Then
'        'Hier könnte man eine Beschränkung für Demo-Version einbauen
'        'nicht mehr als 2 Schrauben
'        d = Round(d, 0)
'        If d > 100 Then d = 100
'        m_g.AnzahlX = d
'    End If
'    Call UpdateView
'End Sub
'Private Sub TxNZ_Change()
'    If m_g Is Nothing Or Not IsNumeric(TxNZ) Or isUpdating Then Exit Sub
'    Dim d As Double
'    If Double_TryParse(TxNZ, d) Then
'        'Hier könnte man eine Beschränkung für Demo-Version einbauen
'        'nicht mehr als 2 Schrauben
'        'soll man das hier in der Oberfläche tun oder in der Schrauben-Klasse???
'        d = Round(d, 0)
'        If d > 100 Then d = 100
'        m_g.AnzahlZ = d
'    End If
'    Call UpdateView
'End Sub

Private Sub TxRX_KeyDown(KeyCode As Integer, Shift As Integer)
    UpdateDataNumeric TxRX, KeyCode, m_SVFL.Schraubengruppe.AbstandSel.Rand, "X"
End Sub
Private Sub TxRX_LostFocus()
    UpdateDataNumeric TxRX, vbKeyTab, m_SVFL.Schraubengruppe.AbstandSel.Rand, "X"
End Sub

Private Sub TxRZ_KeyDown(KeyCode As Integer, Shift As Integer)
    UpdateDataNumeric TxRZ, KeyCode, m_SVFL.Schraubengruppe.AbstandSel.Rand, "Z"
End Sub
Private Sub TxRZ_LostFocus()
    UpdateDataNumeric TxRZ, vbKeyTab, m_SVFL.Schraubengruppe.AbstandSel.Rand, "Z"
End Sub

Private Sub TxLX_KeyDown(KeyCode As Integer, Shift As Integer)
    UpdateDataNumeric TxLX, KeyCode, m_SVFL.Schraubengruppe.AbstandSel.Loch, "X"
End Sub
Private Sub TxLX_LostFocus()
    UpdateDataNumeric TxLX, vbKeyTab, m_SVFL.Schraubengruppe.AbstandSel.Loch, "X"
End Sub

Private Sub TxLZ_KeyDown(KeyCode As Integer, Shift As Integer)
    UpdateDataNumeric TxLZ, KeyCode, m_SVFL.Schraubengruppe.AbstandSel.Loch, "Z"
End Sub
Private Sub TxLZ_LostFocus()
    UpdateDataNumeric TxLZ, vbKeyTab, m_SVFL.Schraubengruppe.AbstandSel.Loch, "Z"
End Sub

Private Sub UpdateDataNumeric(tb As TextBox, KeyCode As Integer, Obj As Object, Prop As String, Optional ByVal isInt As Boolean = False)
    If Not ((KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab)) Then Exit Sub
    Dim s As String: s = tb.Text
    If Obj Is Nothing Or Not IsNumeric(s) Or isUpdating Then Exit Sub
    Dim d As Double
    If Double_TryParse(tb, d) Then
        If isInt Then d = Round(d)
        CallByName Obj, Prop, VbLet, d
    End If
    UpdateView
End Sub

Private Sub CkBBeamLeft_Click()
    If isUpdating Then Exit Sub
    Call CreateBlechL
    'CalcSchraubenlänge
    Call UpdateView
End Sub

Private Sub TxBLh_KeyDown(KeyCode As Integer, Shift As Integer)
    UpdateDataNumeric TxBLh, KeyCode, m_SVFL.BlechLinks, "Höhe"
'    If Not ((KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab)) Then Exit Sub
'    If m_bl Is Nothing Or isUpdating Then Exit Sub
'    Dim d As Double: If Double_TryParse(TxBLh, d) Then m_bl.Höhe = d
'    Call UpdateView
End Sub
Private Sub TxBLh_LostFocus()
    UpdateDataNumeric TxBLh, vbKeyTab, m_SVFL.BlechLinks, "Höhe"
End Sub

Private Sub TxBLt_KeyDown(KeyCode As Integer, Shift As Integer)
    UpdateDataNumeric TxBLt, KeyCode, m_SVFL.BlechLinks, "Blechdicke"
End Sub
Private Sub TxBLt_LostFocus()
    UpdateDataNumeric TxBLt, vbKeyTab, m_SVFL.BlechLinks, "Blechdicke"
End Sub

'Private Sub TxBLt_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Not ((KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab)) Then Exit Sub
'    If m_bl Is Nothing Or isUpdating Then Exit Sub
'    Dim d As Double: If Double_TryParse(TxBLt, d) Then m_bl.Blechdicke = d
'    'CalcSchraubenlänge
'    Call UpdateView
'End Sub
Private Sub CBBLStahl_Click()
    Dim bl As Blech: Set bl = m_SVFL.BlechLinks
    If bl Is Nothing Or isUpdating Then Exit Sub
    bl.Stahlsorte = CBBLStahl.ListIndex
    Call UpdateView
End Sub
Private Sub CkBBeamRight_Click()
    If isUpdating Then Exit Sub
    Call CreateBlechR
    'CalcSchraubenlänge
    Call UpdateView
End Sub

Private Sub TxBRh_KeyDown(KeyCode As Integer, Shift As Integer)
    UpdateDataNumeric TxBRh, KeyCode, m_SVFL.BlechRechts, "Höhe"
End Sub
Private Sub TxBRh_LostFocus()
    UpdateDataNumeric TxBRh, vbKeyTab, m_SVFL.BlechRechts, "Höhe"
End Sub
'Private Sub TxBRh_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Not ((KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab)) Then Exit Sub
'    If m_br Is Nothing Or isUpdating Then Exit Sub
'    Dim d As Double: If Double_TryParse(TxBRh, d) Then m_br.Höhe = d
'    Call UpdateView
'End Sub
Private Sub TxBRt_KeyDown(KeyCode As Integer, Shift As Integer)
    UpdateDataNumeric TxBRt, KeyCode, m_SVFL.BlechRechts, "Blechdicke"
End Sub
Private Sub TxBRt_LostFocus()
    UpdateDataNumeric TxBRt, vbKeyTab, m_SVFL.BlechRechts, "Blechdicke"
End Sub
'Private Sub TxBRt_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Not ((KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab)) Then Exit Sub
'    If m_br Is Nothing Or isUpdating Then Exit Sub
'    Dim d As Double: If Double_TryParse(TxBRt, d) Then m_br.Blechdicke = d
'    CalcSchraubenlänge
'    Call UpdateView
'End Sub
Private Sub CBBRStahl_Click()
    Dim br As Blech: Set br = m_SVFL.BlechRechts
    If br Is Nothing Or isUpdating Then Exit Sub
    br.Stahlsorte = CBBRStahl.ListIndex
    Call UpdateView
End Sub

Private Sub CmBAbstRand_Click()
    Dim g As Schraubengruppe: Set g = m_SVFL.Schraubengruppe
    If g Is Nothing Or isUpdating Then Exit Sub
    Dim b  As Blech
    Dim bl As Blech: Set bl = m_SVFL.BlechLinks
    Dim br As Blech: Set br = m_SVFL.BlechRechts
    If Not bl Is Nothing Then
        If Not br Is Nothing Then
            Set b = IIf(bl.GesamtT < br.GesamtT, bl, br)
        Else
            Set b = bl
        End If
    Else
        If Not br Is Nothing Then
            Set b = br
        End If
    End If
    Call g.SetAbstandRandOpt(CmBAbstRand.ListIndex, b) ', CkRoundUp5.Value = vbChecked)
    If Me.CkUpdateAbstand.Value = vbChecked Then
        g.SyncAbstandRandSel Me.CkRoundUp5.Value = vbChecked
    End If
    Call UpdateView
End Sub
Private Sub CmBAbstLoch_Click()
    Dim g As Schraubengruppe: Set g = m_SVFL.Schraubengruppe
    If g Is Nothing Or isUpdating Then Exit Sub
    Dim b  As Blech
    Dim bl As Blech: Set bl = m_SVFL.BlechLinks
    Dim br As Blech: Set br = m_SVFL.BlechRechts
    If Not bl Is Nothing Then
        If Not br Is Nothing Then
            Set b = IIf(bl.GesamtT < br.GesamtT, bl, br)
        Else
            Set b = bl
        End If
    Else
        If Not br Is Nothing Then
            Set b = br
        End If
    End If
    Call g.SetAbstandLochOpt(CmBAbstLoch.ListIndex, b) ', CkRoundUp5.Value = vbChecked)
    If Me.CkUpdateAbstand.Value = vbChecked Then
        g.SyncAbstandLochSel Me.CkRoundUp5.Value = vbChecked
    End If
    Call UpdateView
End Sub
'
'Private Sub BtnOK_Click()
'    'so jetzt Schrauben und Bleche zeichnen
'    ''''Call ZeichneSchraube
'    Me.Hide
'    'Unload Me
'End Sub
'Private Sub BtnCancel_Click()
'    Me.Hide
'    'Unload Me
'End Sub

