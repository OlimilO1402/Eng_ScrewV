Attribute VB_Name = "MMain"
Option Explicit
Public New_c As cConstructor
Public Cairo As cCairo

Public Sub Main()
    frmSchrauben.Show
End Sub
Public Sub NewInit()
    'Dim AppPath As String: AppPath = App.Path & "\"
    Set New_c = MNew.Constructor()
    Set Cairo = New_c.Cairo
    'Call MCDraw.InitScale
    Call MMath.Init
    'Load frmSchrauben
End Sub
Public Function CairoStatus_ToStr(e As cairo_status_enm) As String
    Dim s As String
    Select Case e
    Case cairo_status_enm.CAIRO_STATUS_CLIP_NOT_REPRESENTABLE:    s = "CAIRO_STATUS_CLIP_NOT_REPRESENTABLE"
    Case cairo_status_enm.CAIRO_STATUS_DEVICE_ERROR:              s = "CAIRO_STATUS_DEVICE_ERROR"
    Case cairo_status_enm.CAIRO_STATUS_DEVICE_TYPE_MISMATCH:      s = "CAIRO_STATUS_DEVICE_TYPE_MISMATCH"
    Case cairo_status_enm.CAIRO_STATUS_FILE_NOT_FOUND:            s = "CAIRO_STATUS_FILE_NOT_FOUND"
    Case cairo_status_enm.CAIRO_STATUS_FONT_TYPE_MISMATCH:        s = "CAIRO_STATUS_FONT_TYPE_MISMATCH"
    Case cairo_status_enm.CAIRO_STATUS_INVALID_CLUSTERS:          s = "CAIRO_STATUS_INVALID_CLUSTERS"
    Case cairo_status_enm.CAIRO_STATUS_INVALID_CONTENT:           s = "CAIRO_STATUS_INVALID_CONTENT"
    Case cairo_status_enm.CAIRO_STATUS_INVALID_DASH:              s = "CAIRO_STATUS_INVALID_DASH"
    Case cairo_status_enm.CAIRO_STATUS_INVALID_DSC_COMMENT:       s = "CAIRO_STATUS_INVALID_DSC_COMMENT"
    Case cairo_status_enm.CAIRO_STATUS_INVALID_FORMAT:            s = "CAIRO_STATUS_INVALID_FORMAT"
    Case cairo_status_enm.CAIRO_STATUS_INVALID_INDEX:             s = "CAIRO_STATUS_INVALID_INDEX"
    Case cairo_status_enm.CAIRO_STATUS_INVALID_MATRIX:            s = "CAIRO_STATUS_INVALID_MATRIX"
    Case cairo_status_enm.CAIRO_STATUS_INVALID_PATH_DATA:         s = "CAIRO_STATUS_INVALID_PATH_DATA"
    Case cairo_status_enm.CAIRO_STATUS_INVALID_POP_GROUP:         s = "CAIRO_STATUS_INVALID_POP_GROUP"
    Case cairo_status_enm.CAIRO_STATUS_INVALID_RESTORE:           s = "CAIRO_STATUS_INVALID_RESTORE"
    Case cairo_status_enm.CAIRO_STATUS_INVALID_SIZE:              s = "CAIRO_STATUS_INVALID_SIZE"
    Case cairo_status_enm.CAIRO_STATUS_INVALID_SLANT:             s = "CAIRO_STATUS_INVALID_SLANT"
    Case cairo_status_enm.CAIRO_STATUS_INVALID_STATUS:            s = "CAIRO_STATUS_INVALID_STATUS"
    Case cairo_status_enm.CAIRO_STATUS_INVALID_STRIDE:            s = "CAIRO_STATUS_INVALID_STRIDE"
    Case cairo_status_enm.CAIRO_STATUS_INVALID_STRING:            s = "CAIRO_STATUS_INVALID_STRING"
    Case cairo_status_enm.CAIRO_STATUS_INVALID_VISUAL:            s = "CAIRO_STATUS_INVALID_VISUAL"
    Case cairo_status_enm.CAIRO_STATUS_INVALID_WEIGHT:            s = "CAIRO_STATUS_INVALID_WEIGHT"
    Case cairo_status_enm.CAIRO_STATUS_LAST_STATUS:               s = "CAIRO_STATUS_LAST_STATUS"
    Case cairo_status_enm.CAIRO_STATUS_NEGATIVE_COUNT:            s = "CAIRO_STATUS_NEGATIVE_COUNT"
    Case cairo_status_enm.CAIRO_STATUS_NO_CURRENT_POINT:          s = "CAIRO_STATUS_NO_CURRENT_POINT"
    Case cairo_status_enm.CAIRO_STATUS_NO_MEMORY:                 s = "CAIRO_STATUS_NO_MEMORY"
    Case cairo_status_enm.CAIRO_STATUS_NULL_POINTER:              s = "CAIRO_STATUS_NULL_POINTER"
    Case cairo_status_enm.CAIRO_STATUS_PATTERN_TYPE_MISMATCH:     s = "CAIRO_STATUS_PATTERN_TYPE_MISMATCH"
    Case cairo_status_enm.CAIRO_STATUS_READ_ERROR:                s = "CAIRO_STATUS_READ_ERROR"
    Case cairo_status_enm.CAIRO_STATUS_SUCCESS:                   s = "CAIRO_STATUS_SUCCESS"
    Case cairo_status_enm.CAIRO_STATUS_SURFACE_FINISHED:          s = "CAIRO_STATUS_SURFACE_FINISHED"
    Case cairo_status_enm.CAIRO_STATUS_TEMP_FILE_ERROR:           s = "CAIRO_STATUS_TEMP_FILE_ERROR"
    Case cairo_status_enm.CAIRO_STATUS_USER_FONT_ERROR:           s = "CAIRO_STATUS_USER_FONT_ERROR"
    Case cairo_status_enm.CAIRO_STATUS_USER_FONT_IMMUTABLE:       s = "CAIRO_STATUS_USER_FONT_IMMUTABLE"
    Case cairo_status_enm.CAIRO_STATUS_USER_FONT_NOT_IMPLEMENTED: s = "CAIRO_STATUS_USER_FONT_NOT_IMPLEMENTED"
    Case cairo_status_enm.CAIRO_STATUS_WRITE_ERROR:               s = "CAIRO_STATUS_WRITE_ERROR"
    End Select
    CairoStatus_ToStr = s
End Function

