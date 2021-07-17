VERSION 5.00
Begin VB.Form frmSchrauben 
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox CkZug 
      Caption         =   "Zugverb."
      Height          =   255
      Left            =   120
      TabIndex        =   51
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ComboBox CbGleitGZ 
      Height          =   315
      Left            =   1080
      TabIndex        =   50
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CheckBox CkGleitf 
      Caption         =   "Gleitfest"
      Height          =   255
      Left            =   120
      TabIndex        =   49
      Top             =   1920
      Width           =   975
   End
   Begin VB.CheckBox CkVorgesp 
      Caption         =   "Vorgespannt"
      Height          =   255
      Left            =   1440
      TabIndex        =   48
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox CkSenk 
      Caption         =   "Senkkopf"
      Height          =   255
      Left            =   1440
      TabIndex        =   47
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CheckBox CkSFSchaft 
      Caption         =   "Scherfuge im Schaft"
      Height          =   255
      Left            =   120
      TabIndex        =   46
      Top             =   1560
      Width           =   1815
   End
   Begin VB.ComboBox CbNorm 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1080
      List            =   "Form1.frx":0002
      TabIndex        =   44
      Top             =   120
      Width           =   1695
   End
   Begin VB.PictureBox PnlSG 
      BorderStyle     =   0  'Kein
      Height          =   4695
      Left            =   0
      ScaleHeight     =   4695
      ScaleWidth      =   2775
      TabIndex        =   13
      Top             =   4080
      Width           =   2775
      Begin VB.ComboBox CmBAbstLoch 
         Height          =   315
         Left            =   1800
         TabIndex        =   41
         Text            =   "min"
         Top             =   3960
         Width           =   975
      End
      Begin VB.ComboBox CmBAbstRand 
         Height          =   315
         Left            =   1800
         TabIndex        =   40
         Text            =   "min"
         Top             =   4320
         Width           =   975
      End
      Begin VB.CheckBox CkBBeamRight 
         Caption         =   "Träger von Rechts (FL)"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox TxBRt 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   2160
         TabIndex        =   35
         Text            =   "10"
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox TxBRh 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   720
         TabIndex        =   34
         Text            =   "100"
         Top             =   3120
         Width           =   615
      End
      Begin VB.ComboBox CBBRStahl 
         Height          =   315
         ItemData        =   "Form1.frx":0004
         Left            =   1800
         List            =   "Form1.frx":0006
         TabIndex        =   33
         Top             =   3480
         Width           =   975
      End
      Begin VB.CheckBox CkBBeamLeft 
         Caption         =   "Träger von Links (FL)"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox TxBLt 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   2160
         TabIndex        =   28
         Text            =   "10"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox TxBLh 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   720
         TabIndex        =   27
         Text            =   "100"
         Top             =   1920
         Width           =   615
      End
      Begin VB.ComboBox CBBLStahl 
         Height          =   315
         ItemData        =   "Form1.frx":0008
         Left            =   1800
         List            =   "Form1.frx":000A
         TabIndex        =   26
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox TxNX 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   720
         TabIndex        =   19
         Text            =   "2"
         Top             =   0
         Width           =   615
      End
      Begin VB.TextBox TxNZ 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   2160
         TabIndex        =   18
         Text            =   "1"
         Top             =   0
         Width           =   615
      End
      Begin VB.TextBox TxRX 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   2160
         TabIndex        =   17
         Text            =   "30"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox TxLX 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   720
         TabIndex        =   16
         Text            =   "30"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox TxRZ 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Text            =   "30"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox TxLZ 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Text            =   "30"
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label LbAbstLoch 
         AutoSize        =   -1  'True
         Caption         =   "Setze Abstand Loch"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   3960
         Width           =   1440
      End
      Begin VB.Label LbAbstRand 
         AutoSize        =   -1  'True
         Caption         =   "Setze Abstand Rand"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   4320
         Width           =   1470
      End
      Begin VB.Label LbBRt 
         AutoSize        =   -1  'True
         Caption         =   "Dicke t"
         Height          =   195
         Left            =   1560
         TabIndex        =   39
         Top             =   3120
         Width           =   510
      End
      Begin VB.Label LbBRh 
         AutoSize        =   -1  'True
         Caption         =   "Höhe h"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   3120
         Width           =   525
      End
      Begin VB.Label LbBRS 
         AutoSize        =   -1  'True
         Caption         =   "Stahlsorte"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   3480
         Width           =   705
      End
      Begin VB.Label LbBLt 
         AutoSize        =   -1  'True
         Caption         =   "Dicke t"
         Height          =   195
         Left            =   1560
         TabIndex        =   32
         Top             =   1920
         Width           =   510
      End
      Begin VB.Label LbBLh 
         AutoSize        =   -1  'True
         Caption         =   "Höhe h"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   1920
         Width           =   525
      End
      Begin VB.Label LbBLS 
         AutoSize        =   -1  'True
         Caption         =   "Stahlsorte"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   2280
         Width           =   705
      End
      Begin VB.Label LbNX 
         AutoSize        =   -1  'True
         Caption         =   "nx"
         Height          =   195
         Left            =   360
         TabIndex        =   25
         Top             =   0
         Width           =   165
      End
      Begin VB.Label LbNZ 
         AutoSize        =   -1  'True
         Caption         =   "nz"
         Height          =   195
         Left            =   1800
         TabIndex        =   24
         Top             =   0
         Width           =   165
      End
      Begin VB.Label LbRX 
         AutoSize        =   -1  'True
         Caption         =   "Rand-x"
         Height          =   195
         Left            =   1560
         TabIndex        =   23
         Top             =   720
         Width           =   510
      End
      Begin VB.Label LbLX 
         AutoSize        =   -1  'True
         Caption         =   "Loch-x"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   480
      End
      Begin VB.Label LbRZ 
         AutoSize        =   -1  'True
         Caption         =   "Rand-z"
         Height          =   195
         Left            =   840
         TabIndex        =   21
         Top             =   360
         Width           =   510
      End
      Begin VB.Label LbLZ 
         AutoSize        =   -1  'True
         Caption         =   "Loch-z"
         Height          =   195
         Left            =   840
         TabIndex        =   20
         Top             =   1080
         Width           =   480
      End
   End
   Begin VB.CheckBox CkBBoltGroup 
      Caption         =   "Schraubengruppe"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CheckBox CkBIsVert 
      Caption         =   "Vertikal"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   1455
   End
   Begin VB.ComboBox CBLochart 
      Height          =   315
      Left            =   1080
      TabIndex        =   9
      Top             =   2640
      Width           =   1695
   End
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
      Height          =   8160
      Left            =   2880
      TabIndex        =   8
      Top             =   600
      Width           =   7335
   End
   Begin VB.CheckBox CkBDrawHole 
      Caption         =   "Schraubenloch zeichnen"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CheckBox CkBPass 
      Caption         =   "Passschr."
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox CBGüte 
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   5880
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":000C
      Top             =   600
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox CBDurchmesser 
      Height          =   315
      ItemData        =   "Form1.frx":0012
      Left            =   2040
      List            =   "Form1.frx":0014
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Norm"
      Height          =   195
      Left            =   120
      TabIndex        =   45
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Lochart"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Schraubengüte"
      Height          =   195
      Left            =   120
      TabIndex        =   5
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
End
Attribute VB_Name = "frmSchrauben"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_N As Norm
Private m_S As Schraube
Private m_L As Schraubenloch 'Standardloch für alle Schrauben der Gruppe
Private m_g As Schraubengruppe
Private m_BL As Blech 'Blech links
Private m_BR As Blech 'Blech rechts
Private bInitText As Boolean

'Private Sub Form_Load()
'Private Sub UserForm_Initialize()
Private Sub Form_Initialize()
    Me.Caption = "Schrauben-Berechnungen nach DIN18800 und EuroCode 3"
    With CbNorm:      .AddItem "DIN 18800": .AddItem "EuroCode 3": End With
    With CBDurchmesser
        .AddItem "8":  .AddItem "10": .AddItem "12": .AddItem "16": .AddItem "20":
        .AddItem "22": .AddItem "24": .AddItem "27": .AddItem "30": .AddItem "36"
    End With
    With CBGüte:      .AddItem "4.6": .AddItem "5.6": .AddItem "8.8": .AddItem "10.9":     End With
    With CBLochart:   .AddItem "Normal": .AddItem "Übergroß": .AddItem "Langloch Kurz": .AddItem "Langloch Lang":    End With
    With CmBAbstLoch: .AddItem "minimal": .AddItem "mittel": .AddItem "maximal":    End With
    With CmBAbstRand: .AddItem "minimal": .AddItem "mittel": .AddItem "maximal":    End With
    With CbGleitGZ:   .AddItem "GZ-Gebrauchst.": .AddItem "GZ-Tragfähigkeit": End With
    CbNorm.ListIndex = 0
    CBDurchmesser.ListIndex = 2
    CBGüte.ListIndex = 0
    CkBPass.Value = vbUnchecked
    CBLochart.ListIndex = 0
    CreateNorm
    InitCBStahl
    CreateSchraube
    CreateSchraubenLoch
    PnlSG_Enabled = False
    Call UpdateView
    InitEPTipps
    
End Sub
Sub InitCBStahl()
    CBBLStahl.Clear
    CBBRStahl.Clear
    If m_N.Norm = Norm_DIN18800 Then
        With CBBLStahl:   .AddItem "ST37": .AddItem "ST42": .AddItem "ST52":    End With
        With CBBRStahl:   .AddItem "ST37": .AddItem "ST42": .AddItem "ST52":    End With
    Else
        With CBBLStahl:   .AddItem "S235": .AddItem "S275": .AddItem "S355":    End With
        With CBBRStahl:   .AddItem "S235": .AddItem "S275": .AddItem "S355":    End With
    End If
End Sub
Sub InitEPTipps()
    If m_N.Norm = Norm_DIN18800 Then
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
    If Not m_N Is Nothing Then Exit Sub
    Set m_N = New_Norm(Norm_DIN18800)
    'Set m_N = New_Norm(Norm_EuroCode3)
End Sub
Sub CreateSchraube()
    If Not IsNumeric(CBDurchmesser.Text) Then Exit Sub
    Set m_S = New_Schraube(m_N, CBDurchmesser.Text, Val(CBGüte.Text))
End Sub
Sub CreateSchraubenLoch()
    If m_S Is Nothing Then Exit Sub
    Set m_L = New_Schraubenloch(m_S, CBLochart.ListIndex, CkBIsVert.Value)
End Sub
Sub CreateSchraubenGruppe()
    If Not (IsNumeric(TxNX) And IsNumeric(TxNZ) And _
            IsNumeric(TxRX) And IsNumeric(TxRZ) And _
            IsNumeric(TxLX) And IsNumeric(TxLZ)) Then Exit Sub
    If m_L Is Nothing Then Exit Sub
    Set m_g = New_Schraubengruppe(m_L, TxNX, TxNZ, TxRX, TxRZ, TxLX, TxLZ)
End Sub
Sub CreateBlechL()
    If Not CkBBeamLeft.Value = vbChecked Then
        Set m_BL = Nothing
        Exit Sub
    End If
    If Not (IsNumeric(TxBLh) And IsNumeric(TxBLt)) Then Exit Sub
    Set m_BL = New_Blech(m_N, m_g, TxBLt, CBBLStahl.ListIndex, 0, TxBLh)
End Sub
Sub CreateBlechR()
    If Not CkBBeamRight.Value = vbChecked Then
        Set m_BR = Nothing
        Exit Sub
    End If
    If Not (IsNumeric(TxBRh) And IsNumeric(TxBRt)) Then Exit Sub
    Set m_BR = New_Blech(m_N, m_g, TxBRt, CBBRStahl.ListIndex, 0, TxBRh)
End Sub
Private Sub Command1_Click()
    Dim T As String: T = Text1.Text
    If Len(T) = 0 Then T = Clipboard.GetText
    Dim ta() As String: ta = Split(T, vbCrLf)
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
 
Private Sub UpdateView()
    'CkBScrwHole.Value = Abs(Not CkBPass.Value)
    'CkBScrwHole.Enabled = Not CkBPass.Value
    LBSchraube.Clear
    If Not m_N Is Nothing Then Call m_N.ToListBox(LBSchraube)
    If Not m_S Is Nothing Then
        Call m_S.ToListBox(LBSchraube)
        If Not m_L Is Nothing Then
            Call m_L.ToListBox(LBSchraube)
        End If
        If Not m_g Is Nothing Then
            Call m_g.ToListBox(LBSchraube)
            If Not m_BL Is Nothing Then
                Call m_BL.ToListBox(LBSchraube)
            End If
            If Not m_BR Is Nothing Then
                Call m_BR.ToListBox(LBSchraube)
            End If
        End If
    End If
End Sub

Private Sub CbNorm_Click()
    If m_N Is Nothing Then Exit Sub
    m_N.Norm = CbNorm.ListIndex
    InitCBStahl
    InitEPTipps
    Call UpdateView
    If Not m_BL Is Nothing Then CBBLStahl.ListIndex = m_BL.Stahlsorte
    If Not m_BR Is Nothing Then CBBRStahl.ListIndex = m_BR.Stahlsorte
End Sub

Private Sub CBDurchmesser_Change()
    If m_S Is Nothing Then Exit Sub
    m_S.Durchmesser = Val(CBDurchmesser.Text)
    Call UpdateView
End Sub
Private Sub CBDurchmesser_Click()
    If m_S Is Nothing Then Exit Sub
    m_S.Durchmesser = Val(CBDurchmesser.Text)
    Call UpdateView
End Sub
Private Sub CBGüte_Click()
    If m_S Is Nothing Then Exit Sub
    m_S.Schraubengüte = Val(CBGüte.Text)
    Call UpdateView
End Sub

Private Sub CkBPass_Click()
    If m_S Is Nothing Then Exit Sub
    m_S.IsPassschraube = CkBPass.Value = vbChecked
    CkSFSchaft.Enabled = Not CkBPass.Value = vbChecked
    Call UpdateView
End Sub
Private Sub CkSenk_Click()
    If m_S Is Nothing Then Exit Sub
    m_S.IsSenkschraube = CkSenk.Value = vbChecked
    Call UpdateView
End Sub
Private Sub CkSFSchaft_Click()
    If m_S Is Nothing Then Exit Sub
    m_S.IsScherfugeSchaft = CkSFSchaft.Value = vbChecked
    UpdateView
End Sub
Private Sub CkGleitf_Click()
    If m_S Is Nothing Then Exit Sub
    m_S.IsGleitfest = CkGleitf.Value = vbChecked
    m_S.IsVorgespannt = CkGleitf.Value = vbChecked
    m_S.IsZugverbindung = False 'CkGleitf.Value = vbChecked
    If CbGleitGZ.ListIndex = -1 Then CbGleitGZ.ListIndex = 0
    CkVorgesp.Enabled = Not CkGleitf.Value = vbChecked
    CkZug.Enabled = Not CkGleitf.Value = vbChecked
    CkZug.Value = vbUnchecked 'CkGleitf.Value = vbChecked
    UpdateView
End Sub
Private Sub CbGleitGZ_Click()
    If m_S Is Nothing Then Exit Sub
    m_S.IsGleitfestImGZT = CbGleitGZ.ListIndex = 1
    If CbGleitGZ.ListIndex = -1 Then CbGleitGZ.ListIndex = 0
    UpdateView
End Sub

Private Sub CkZug_Click()
    If m_S Is Nothing Then Exit Sub
    m_S.IsZugverbindung = CkZug.Value = vbChecked
    UpdateView
End Sub
Private Sub CkVorgesp_Click()
    If m_S Is Nothing Then Exit Sub
    m_S.IsVorgespannt = CkVorgesp.Value = vbChecked
    UpdateView
End Sub

Private Sub CBLochart_Click()
    If m_L Is Nothing Then Exit Sub
    m_L.Lochart = CBLochart.ListIndex
    Call UpdateView
End Sub
Private Sub CkBIsVert_Click()
    If m_L Is Nothing Then Exit Sub
    m_L.IsVertikal = CkBIsVert.Value = vbChecked
    Call UpdateView
End Sub
Private Sub CkBDrawHole_Click()
    Call UpdateView
End Sub

Private Sub CkBBoltGroup_Click()
    If Not CkBBoltGroup.Value = vbChecked Then
        Set m_g = Nothing
        Set m_BL = Nothing: CkBBeamLeft.Value = vbUnchecked
        Set m_BR = Nothing: CkBBeamRight.Value = vbUnchecked
    Else
        Call CreateSchraubenGruppe
    End If
    PnlSG_Enabled = CkBBoltGroup.Value = vbChecked
    UpdateView
End Sub
Private Property Let PnlSG_Enabled(ByVal ben As Boolean)
    PnlSG.Enabled = ben
    LbNX.Enabled = ben: TxNX.Enabled = ben
    LbNZ.Enabled = ben: TxNZ.Enabled = ben
    LbRZ.Enabled = ben: TxRZ.Enabled = ben
    LbRX.Enabled = ben: TxRX.Enabled = ben
    LbLZ.Enabled = ben: TxLZ.Enabled = ben
    LbLX.Enabled = ben: TxLX.Enabled = ben
    CkBBeamLeft.Enabled = ben
    PnlBL_Enabled = CkBBeamLeft.Value = vbChecked
    CkBBeamRight.Enabled = ben
    PnlBR_Enabled = CkBBeamRight.Value = vbChecked
    PnlAbstLR_Enabled = CkBBeamLeft.Value = vbChecked Or CkBBeamRight.Value = vbChecked
End Property
Private Property Let PnlBL_Enabled(ByVal ben As Boolean)
    LbBLh.Enabled = ben: TxBLh.Enabled = ben
    LbBLt.Enabled = ben: TxBLt.Enabled = ben
    LbBLS.Enabled = ben: CBBLStahl.Enabled = ben
End Property
Private Property Let PnlBR_Enabled(ByVal ben As Boolean)
    LbBRh.Enabled = ben: TxBRh.Enabled = ben
    LbBRt.Enabled = ben: TxBRt.Enabled = ben
    LbBRS.Enabled = ben: CBBRStahl.Enabled = ben
End Property
Private Property Let PnlAbstLR_Enabled(ByVal ben As Boolean)
    LbAbstLoch.Enabled = ben: CmBAbstLoch.Enabled = ben
    LbAbstRand.Enabled = ben: CmBAbstRand.Enabled = ben
End Property

Private Sub Form_Resize()
    Dim L, T, W, H, brdr
    brdr = 8 * Screen.TwipsPerPixelX
    L = LBSchraube.Left
    T = LBSchraube.Top
    W = Me.ScaleWidth - L - brdr
    H = Me.ScaleHeight - T - brdr
    LBSchraube.Move L, T, W, H
End Sub

Private Sub TxNX_Change()
    If m_g Is Nothing Or Not IsNumeric(TxNX) Then Exit Sub
    m_g.AnzahlX = TxNX
    Call UpdateView
End Sub
Private Sub TxNZ_Change()
    If m_g Is Nothing Or Not IsNumeric(TxNZ) Then Exit Sub
    m_g.AnzahlZ = TxNZ
    Call UpdateView
End Sub
Private Sub TxRX_Change()
    If m_g Is Nothing And Not IsNumeric(TxRX) Then Exit Sub
    m_g.AbstandRandX = TxRX
    Call UpdateView
End Sub
Private Sub TxRZ_Change()
    If m_g Is Nothing And Not IsNumeric(TxRZ) Then Exit Sub
    m_g.AbstandRandZ = TxRZ
    Call UpdateView
End Sub
Private Sub TxLX_Change()
    If m_g Is Nothing And Not IsNumeric(TxLX) Then Exit Sub
    m_g.AbstandLochX = TxLX
    Call UpdateView
End Sub
Private Sub TxLZ_Change()
    If m_g Is Nothing And Not IsNumeric(TxLZ) Then Exit Sub
    m_g.AbstandLochZ = TxLZ
    Call UpdateView
End Sub

Private Sub CkBBeamLeft_Click()
    Call CreateBlechL
    PnlBL_Enabled = CkBBeamLeft = vbChecked
    PnlAbstLR_Enabled = CkBBeamLeft.Value = vbChecked Or CkBBeamRight.Value = vbChecked
    Call UpdateView
    If Not m_BL Is Nothing Then CBBLStahl.ListIndex = m_BL.Stahlsorte
End Sub
Private Sub TxBLh_Change()
    If m_BL Is Nothing Then Exit Sub
    m_BL.Höhe = TxBLh
    Call UpdateView
End Sub
Private Sub TxBLt_Change()
    If m_BL Is Nothing Then Exit Sub
    m_BL.Blechdicke = TxBLt
    Call UpdateView
End Sub
Private Sub CBBLStahl_Click()
    If m_BL Is Nothing Then Exit Sub
    m_BL.Stahlsorte = CBBLStahl.ListIndex
    Call UpdateView
End Sub

Private Sub CkBBeamRight_Click()
    Call CreateBlechR
    PnlBR_Enabled = CkBBeamRight = vbChecked
    PnlAbstLR_Enabled = CkBBeamLeft.Value = vbChecked Or CkBBeamRight.Value = vbChecked
    Call UpdateView
    If Not m_BR Is Nothing Then CBBRStahl.ListIndex = m_BR.Stahlsorte
End Sub
Private Sub TxBRh_Change()
    If m_BR Is Nothing Then Exit Sub
    m_BR.Höhe = TxBRh
    Call UpdateView
End Sub
Private Sub TxBRt_Change()
    If m_BR Is Nothing Then Exit Sub
    m_BR.Blechdicke = TxBRt
    Call UpdateView
End Sub
Private Sub CBBRStahl_Click()
    If m_BR Is Nothing Then Exit Sub
    m_BR.Stahlsorte = CBBRStahl.ListIndex
    Call UpdateView
End Sub

Private Sub CmBAbstLoch_Click()
    If Not m_g Is Nothing Then
        Call m_g.SetAbstandLoch(CmBAbstLoch.ListIndex)
        TxLX.Text = m_g.AbstandLochX
        TxLZ.Text = m_g.AbstandLochZ
        Call UpdateView
    End If
End Sub
Private Sub CmBAbstRand_Click()
    If Not m_g Is Nothing Then
        Call m_g.SetAbstandRand(CmBAbstRand.ListIndex)
        TxRX.Text = m_g.AbstandRandX
        TxRZ.Text = m_g.AbstandRandZ
        Call UpdateView
    End If
End Sub

Private Sub BtnOK_Click()
    'so jetzt Schrauben und Bleche zeichnen
    ''''Call ZeichneSchraube
    Me.Hide
    'Unload Me
End Sub
Private Sub BtnCancel_Click()
    Me.Hide
    'Unload Me
End Sub
