VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CairoPdfDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum EPageOrientation
    poPortrait  'Hochformat
    poLandscape 'Querformat
End Enum
Public Enum EPageFormat
    pfDIN_A0
    pfDIN_A1
    pfDIN_A2
    pfDIN_A3
    pfDIN_A4
    pfDIN_A5
    pfDIN_A6
End Enum
Public Enum EUnit
    euMM
    euCM
    euInch
    euPoints 'Einheit Punkt =1/72 eines Inch
End Enum
Private m_cairo As cCairo
Private m_pdfsrf As cCairoSurface
Private m_pdfctx As cCairoContext
Private m_pageO As EPageOrientation 'Quer- oder Hochformat
Private m_pageF As EPageFormat      'DIN A4, A5. . .
Private m_pageW As Double 'Breite in mm
Private m_pageH As Double 'Höhe   in mm
Private m_zoom  As Double
Private m_FNam  As String
Public Sub New_(c As cCairo, ByVal epo As EPageOrientation, ByVal epf As EPageFormat, ByVal zoom As Double)
    Set m_cairo = c
    m_zoom = 1
    m_pageO = epo:    m_pageF = epf
    UpdatePageSize
    Dim w As Double: w = Me.PageWith
    Dim h As Double: h = Me.PageHeight
    Set m_pdfsrf = c.CreateSurface(w, h, PDFSurface)
    Set m_pdfctx = m_pdfsrf.CreateContext
    m_zoom = zoom 'muss hier nach CreateSurface stehen!!!!
    m_pdfctx.ScaleDrawings PunkteProMM * m_zoom, PunkteProMM * m_zoom
    'm_pdfctx.SetSourceColor vbWhite 'nicht erforderlich in pdf
    'm_pdfctx.Paint
    w = m_pdfsrf.Width
    h = m_pdfsrf.Height
    m_pdfctx.DeviceToUserDistance w, h
End Sub
Public Property Get PathFilename() As String

End Property
Public Property Get PageOrientation() As EPageOrientation
    PageOrientation = m_pageO
End Property
Public Property Let PageOrientation(ByVal epo As EPageOrientation)
    m_pageF = epo
    UpdatePageOrientation
End Property

Public Property Get PageFormat() As EPageFormat
    PageOrientation = m_pageO
End Property
Public Property Let PageFormat(ByVal epf As EPageFormat)
    m_pageF = epf
    UpdatePageSize
End Property
Public Property Get PunkteProMM() 'As Double
    PunkteProMM = 72 / 25.4 '2.83464566929134
End Property
Public Property Get PageWith(Optional ByVal e As EUnit = euPoints) As Double
    'liefert die SeitenBreite in der angegebenen Einheit
    PageWith = RecalcUnit(m_pageW, e) / m_zoom
End Property
Public Property Get PageHeight(Optional ByVal e As EUnit = euPoints) As Double
    'liefert die SeitenHöhe in der angegebenen Einheit
    PageHeight = RecalcUnit(m_pageH, e) / m_zoom
End Property
Function RecalcUnit(ByVal ValueMM As Double, ByVal e As EUnit) As Double
    RecalcUnit = ValueMM:           If e = euMM Then Exit Function
    RecalcUnit = RecalcUnit / 10:   If e = euCM Then Exit Function
    RecalcUnit = RecalcUnit / 2.54: If e = euInch Then Exit Function
    RecalcUnit = RecalcUnit * 72:   'If e = euPoints Then Exit Function
End Function
Private Sub UpdatePageSize()
    Dim w2 As Single: w2 = Sqr(2)
    m_pageW = Round((840 / (w2 ^ m_pageF)), 1)
    m_pageH = Round(m_pageW * w2, 1)
    UpdatePageOrientation
End Sub
Private Sub UpdatePageOrientation()
    Dim w As Double: w = IIf(m_pageO = poPortrait, Min(m_pageW, m_pageH), Max(m_pageW, m_pageH))
    Dim h As Double: h = IIf(m_pageO = poPortrait, Max(m_pageW, m_pageH), Min(m_pageW, m_pageH))
    m_pageW = w
    m_pageH = h
End Sub
'Private Sub UpdatePage()
    'UpdatePageSize

'End Sub
Public Property Get Canvas() As cCairoContext
    Set Canvas = m_pdfctx
End Property
Public Property Get Surface() As cCairoSurface
    Set Surface = m_pdfsrf
End Property

Public Sub WriteToFile(aPFN As String)
    'm_pdfctx.Save
    'm_pdfctx.Restore
    Set m_pdfctx = Nothing 'wir sind fertig, wichtig hier Objekt gleich zerstören!!!
    m_pdfsrf.EmitPDFPage 'keine weiteren Seiten mehr
    'Dim b() As Byte
    Dim s As cairo_status_enm
    's = m_pdfsrf.WriteContentToPDFByteArray(b())
    s = m_pdfsrf.WriteContentToPDFFile(aPFN)
    If s = CAIRO_STATUS_SUCCESS Then
        'MsgBox UBound(b)
        'Dim FNr As Integer: FNr = FreeFile
        'Open aPFN For Binary As FNr
        'Put FNr, , b
        'Close FNr
    Else
        Err.Raise "Error, writing PDF-ByteStream: " & s
    End If

    'm_pdfsrf.WriteContentToPDFFile aPFN
End Sub
