VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CairoPicBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mPB As PictureBox
Attribute mPB.VB_VarHelpID = -1
Private m_sdig As Double
Private m_zoom As Double
Private m_PixProMM As Double
Private m_cair As cCairo
Private m_csrf As cCairoSurface 'a cCairoSurface is something like an allocated InMemory-Bitmap (a hDIB)
Private m_cctx As cCairoContext 'a cCairoContext is something like an hDC in GDI, but ...

Friend Sub New_(aPb As PictureBox, c As cCairo)
    Set mPB = aPb
    mPB.ScaleMode = vbPixels
    'mPB.AutoRedraw = True
    'mpb.Cls
    Set m_cair = c
    m_sdig = 23.4 'Zoll
    m_zoom = 1
    'mPB_Resize
End Sub

' ######################################## ' v   ISerializable   v ' ######################################## '
Public Property Get ID() As Long
    ID = ObjPtr(Me)
End Property
Public Property Let ID(ByVal Value As Long)
    'nix
End Property

Public Property Get ScreenDiagonaleInch() As Double
    ScreenDiagonaleInch = m_sdig
End Property
Public Property Let ScreenDiagonaleInch(ByVal Value As Double)
    m_sdig = Value
    mPB_Resize
End Property

Public Property Get ZoomFactor() As Double
    ZoomFactor = m_zoom
End Property
Public Property Let ZoomFactor(ByVal Value As Double)
    m_zoom = Value
    mPB_Resize
End Property
' ######################################## ' ^   ISerializable   ^ ' ######################################## '

Public Property Get CenterX() As Double
    CenterX = m_csrf.Width / 2 / (m_PixProMM * m_zoom)
End Property
Public Property Get CenterY() As Double
    CenterY = m_csrf.Height / 2 / (m_PixProMM * m_zoom)
End Property
Private Sub UpdateScale()
    If m_cctx Is Nothing Then Exit Sub
    Dim AuflHori As Double: AuflHori = Screen.Width / Screen.TwipsPerPixelX ' z.B. 1920
    Dim AuflVert As Double: AuflVert = Screen.Height / Screen.TwipsPerPixelY ' z.B. 1080
    'If ScreenDiag = 0 Then ScreenDiag = 23.4     ' Zoll 'oder Abfrage!!!
    m_PixProMM = Sqr(AuflHori ^ 2 + AuflVert ^ 2) / m_sdig / 25.4
    m_cctx.ScaleDrawings m_PixProMM * m_zoom, m_PixProMM * m_zoom
    'mPB.Refresh
End Sub

Private Sub mPB_Resize()
    Dim W As Long: W = mPB.ScaleWidth  '/ Screen.TwipsPerPixelX
    Dim H As Long: H = mPB.ScaleHeight '/ Screen.TwipsPerPixelY
    Set m_csrf = m_cair.CreateSurface(W, H, ImageSurface)
    Set m_cctx = m_csrf.CreateContext
    UpdateScale
End Sub
Public Sub DrawAll()
    'm_cctx.Paint
    m_csrf.DrawToDC mPB.hDC, 0, 0
End Sub
Public Property Get Canvas() As cCairoContext
    Set Canvas = m_cctx
End Property
Public Property Get Surface() As cCairoSurface
    Set Surface = m_csrf
End Property
Public Property Get PicBox() As PictureBox
    Set PicBox = mPB
End Property
