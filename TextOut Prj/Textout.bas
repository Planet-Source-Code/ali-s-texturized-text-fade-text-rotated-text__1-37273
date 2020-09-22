Attribute VB_Name = "Module1"
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(1 To 32) As Byte
End Type
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Const FW_NORMAL = 400
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Const TRANSPARENT = 1
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Const FW_DONTCARE = 0
Public Const FW_THIN = 100
Public Const FW_EXTRALIGHT = 200
Public Const FW_ULTRALIGHT = 200
Public Const FW_LIGHT = 300
Public Const FW_REGULAR = 400
Public Const FW_MEDIUM = 500
Public Const FW_SEMIBOLD = 600
Public Const FW_DEMIBOLD = 600
Public Const FW_BOLD = 700
Public Const FW_EXTRABOLD = 800
Public Const FW_ULTRABOLD = 800
Public Const FW_HEAVY = 900
Public Const FW_BLACK = 900
Public Const ANSI_CHARSET = 0
Public Const ARABIC_CHARSET = 178
Public Const BALTIC_CHARSET = 186
Public Const CHINESEBIG5_CHARSET = 136
Public Const DEFAULT_CHARSET = 1
Public Const EASTEUROPE_CHARSET = 238
Public Const GB2312_CHARSET = 134
Public Const GREEK_CHARSET = 161
Public Const HANGEUL_CHARSET = 129
Public Const HEBREW_CHARSET = 177
Public Const MAC_CHARSET = 77
Public Const OEM_CHARSET = 255
Public Const SHIFTJIS_CHARSET = 128
Public Const SYMBOL_CHARSET = 2
Public Const TURKISH_CHARSET = 162
Public Const OUT_DEVICE_PRECIS = 5
Public Const OUT_RASTER_PRECIS = 6
Public Const OUT_STRING_PRECIS = 1
Public Const OUT_STROKE_PRECIS = 3
Public Const OUT_TT_ONLY_PRECIS = 7
Public Const OUT_TT_PRECIS = 4
Public Const CLIP_DEFAULT_PRECIS = 0
Public Const CLIP_EMBEDDED = 128
Public Const CLIP_LH_ANGLES = 16
Public Const CLIP_STROKE_PRECIS = 2
Public Const ANTIALIASED_QUALITY = 4
Public Const DEFAULT_QUALITY = 0
Public Const DRAFT_QUALITY = 1
Public Const NONANTIALIASED_QUALITY = 3
Public Const PROOF_QUALITY = 2
Public Const DEFAULT_PITCH = 0
Public Const FIXED_PITCH = 1
Public Const VARIABLE_PITCH = 2
Public Const FF_DECORATIVE = 80
Public Const FF_DONTCARE = 0
Public Const FF_MODERN = 48
Public Const FF_ROMAN = 16
Public Const FF_SCRIPT = 64
Public Const FF_SWISS = 32
