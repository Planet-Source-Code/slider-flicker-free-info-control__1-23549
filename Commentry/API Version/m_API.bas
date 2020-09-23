Attribute VB_Name = "m_API"
Option Explicit

Public Enum ECGTextAlignFlags
   DT_TOP = &H0&
   DT_LEFT = &H0&
   DT_CENTER = &H1&
   DT_RIGHT = &H2&
   DT_VCENTER = &H4&
   DT_BOTTOM = &H8&
   DT_WORDBREAK = &H10&
   DT_SINGLELINE = &H20&
   DT_EXPANDTABS = &H40&
   DT_TABSTOP = &H80&
   DT_NOCLIP = &H100&
   DT_EXTERNALLEADING = &H200&
   DT_CALCRECT = &H400&
   DT_NOPREFIX = &H800&
   DT_INTERNAL = &H1000&
'#if(WINVER >= =&H0400)
   DT_EDITCONTROL = &H2000&
   DT_PATH_ELLIPSIS = &H4000&
   DT_END_ELLIPSIS = &H8000&
   DT_MODIFYSTRING = &H10000
   DT_RTLREADING = &H20000
   DT_WORD_ELLIPSIS = &H40000
End Enum

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

' Rectangle functions:
Public Declare Function EqualRect Lib "user32" (lpRect1 As RECT, lpRect2 As RECT) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptX As Long, ByVal ptY As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Const CLR_INVALID = -1

Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
    Public Const OPAQUE = 2
    Public Const TRANSPARENT = 1

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Declare Function GetStockObject& Lib "gdi32" (ByVal nIndex As Long)
    Public Const SYSTEM_FONT = 13
    Public Const LF_FACESIZE = 32
    Type LOGFONT
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
        lfFaceName(LF_FACESIZE) As Byte
    End Type

Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const LOGPIXELSX = 88    '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90    '  Logical pixels/inch in Y

Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
           lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Private Const FW_NORMAL = 400
    Private Const FW_BOLD = 700
    Private Const FF_DONTCARE = 0
    Private Const DEFAULT_QUALITY = 0
    Private Const DEFAULT_PITCH = 0
    Private Const DEFAULT_CHARSET = 1

' Corrected Draw State function declarations:
Public Declare Function DrawState Lib "user32" Alias "DrawStateA" _
   (ByVal hdc As Long, _
   ByVal hBrush As Long, _
   ByVal lpDrawStateProc As Long, _
   ByVal lParam As Long, _
   ByVal wParam As Long, _
   ByVal X As Long, _
   ByVal Y As Long, _
   ByVal cx As Long, _
   ByVal cy As Long, _
   ByVal fuFlags As Long) As Long
Public Declare Function DrawStateString Lib "user32" Alias "DrawStateA" _
   (ByVal hdc As Long, _
   ByVal hBrush As Long, _
   ByVal lpDrawStateProc As Long, _
   ByVal lpString As String, _
   ByVal cbStringLen As Long, _
   ByVal X As Long, _
   ByVal Y As Long, _
   ByVal cx As Long, _
   ByVal cy As Long, _
   ByVal fuFlags As Long) As Long

' Missing Draw State constants declarations:
'/* Image type */
Public Const DST_COMPLEX = &H0
Public Const DST_TEXT = &H1
Public Const DST_PREFIXTEXT = &H2
Public Const DST_ICON = &H3
Public Const DST_BITMAP = &H4

' /* State type */
Public Const DSS_NORMAL = &H0
Public Const DSS_UNION = &H10
Public Const DSS_DISABLED = &H20
Public Const DSS_MONO = &H80
Public Const DSS_RIGHT = &H8000

' Create a new icon based on an image list icon:
Public Declare Function ImageList_GetIcon Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal diIgnore As Long _
    ) As Long
' Draw an item in an ImageList:
Public Declare Function ImageList_Draw Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal hdcDst As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal fStyle As Long _
    ) As Long
' Draw an item in an ImageList with more control over positioning
' and colour:
Public Declare Function ImageList_DrawEx Lib "COMCTL32.DLL" ( _
      ByVal hIml As Long, _
      ByVal i As Long, _
      ByVal hdcDst As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal dx As Long, _
      ByVal dy As Long, _
      ByVal rgbBk As Long, _
      ByVal rgbFg As Long, _
      ByVal fStyle As Long _
   ) As Long
' Built in ImageList drawing methods:
Public Const ILD_NORMAL = 0
Public Const ILD_TRANSPARENT = 1
Public Const ILD_BLEND25 = 2
Public Const ILD_SELECTED = 4
Public Const ILD_FOCUS = 4
Public Const ILD_OVERLAYMASK = 3840
' Use default rgb colour:
Public Const CLR_NONE = -1
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function ImageList_GetIconSize Lib "comctl32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
Public Declare Function ImageList_GetImageCount Lib "comctl32" (ByVal hImageList As Long) As Long

' Standard GDI draw icon function:
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Const DI_MASK = &H1
Public Const DI_IMAGE = &H2
Public Const DI_NORMAL = &H3
Public Const DI_COMPAT = &H4
Public Const DI_DEFAULTSIZE = &H8

Public Declare Function LoadImageByNum Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADTRANSPARENT = &H20
Public Const IMAGE_BITMAP = 0

Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
    
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENINNER = &H8

Public Const BDR_OUTER = &H3
Public Const BDR_INNER = &HC
Public Const BDR_RAISED = &H5
Public Const BDR_SUNKEN = &HA

Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8

Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Const BF_DIAGONAL = &H10

' For diagonal lines, the BF_RECT flags specify the end point of
' the vector bounded by the rectangle parameter.
Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP _
             Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP _
             Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM _
             Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM _
             Or BF_RIGHT)

Public Const BF_MIDDLE = &H800    ' Fill in the middle.
Public Const BF_SOFT = &H1000     ' Use for softer buttons.
Public Const BF_ADJUST = &H2000   ' Calculate the space left over.
Public Const BF_FLAT = &H4000     ' For flat rather than 3-D borders.
Public Const BF_MONO = &H8000     ' For monochrome borders.

Public Sub pOLEFontToLogFont(fntThis As StdFont, hdc As Long, tLF As LOGFONT)
Dim sFont As String
Dim iChar As Integer

   ' Convert an OLE StdFont to a LOGFONT structure:
   With tLF
       sFont = fntThis.Name
       ' There is a quicker way involving StrConv and CopyMemory, but
       ' this is simpler!:
       For iChar = 1 To Len(sFont)
           .lfFaceName(iChar - 1) = CByte(Asc(Mid$(sFont, iChar, 1)))
       Next iChar
       ' Based on the Win32SDK documentation:
       .lfHeight = -MulDiv((fntThis.Size), (GetDeviceCaps(hdc, LOGPIXELSY)), 72)
       .lfItalic = fntThis.Italic
       If (fntThis.Bold) Then
           .lfWeight = FW_BOLD
       Else
           .lfWeight = FW_NORMAL
       End If
       .lfUnderline = fntThis.Underline
       .lfStrikeOut = fntThis.Strikethrough
       .lfCharSet = fntThis.Charset
   End With

End Sub

Public Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Public Sub DrawGraduatedBackdrop(ByVal lHDC As Long, _
                                 ByVal lLeft As Long, _
                                 ByVal lTop As Long, _
                                 ByVal lRight As Long, _
                                 ByVal lBottom As Long, _
                                 Optional ByVal eStartColour As OLE_COLOR = &H0&, _
                                 Optional ByVal eEndColour As OLE_COLOR = vbButtonShadow, _
                                 Optional ByVal bVertical As Boolean = False)

    Dim lSRed As Long, lSGreen As Long, lSBlue As Long
    Dim lERed As Long, lEGreen As Long, lEBlue As Long
    Dim lRed As Long, lGreen As Long, lBlue As Long
    Dim lLastRed As Long, lLastGreen As Long, lLastBlue As Long
    Dim lRGB As Long
    Dim hBr As Long
    Dim tR As RECT
    Dim iPos As Long, lSize As Long, lMinStep As Long

    With tR
        .Left = lLeft
        .Top = lTop
        .Right = lRight
        .Bottom = lBottom
    End With

    If (eStartColour = eEndColour) Then
        ' Simple! (but dull...)
        hBr = CreateSolidBrush(eStartColour)
        FillRect lHDC, tR, hBr
        DeleteObject hBr
    Else
        ' Create a gradation:
        lSRed = eStartColour And &HFF&
        lSGreen = (eStartColour And &HFF00&) \ &H100&
        lSBlue = (eStartColour And &HFF0000) \ &H10000
        lERed = eEndColour And &HFF&
        lEGreen = (eEndColour And &HFF00&) \ &H100&
        lEBlue = (eEndColour And &HFF0000) \ &H10000

        If (bVertical) Then
            ' Vertical graduation:
            lSize = lBottom - lTop
            tR.Bottom = tR.Top + 1
            For iPos = 1 To lSize + 1
                lRed = Abs(lSRed + ((lERed - lSRed) * iPos) \ lSize)
                lGreen = Abs(lSGreen + ((lEGreen - lSGreen) * iPos) \ lSize)
                lBlue = Abs(lSBlue + ((lEBlue - lSBlue) * iPos) \ lSize)
                lRGB = RGB(lRed, lGreen, lBlue)
                hBr = CreateSolidBrush(lRGB)
                FillRect lHDC, tR, hBr
                DeleteObject hBr
                tR.Top = tR.Top + 1
                tR.Bottom = tR.Top + 1
            Next
        Else
            ' Horizontal graduation:
            lSize = lRight - lLeft
            lMinStep = lSize \ 64
            If (lMinStep = 0) Then lMinStep = 1
            lLastRed = lSRed: lLastGreen = lSGreen: lLastBlue = lSBlue
            tR.Right = tR.Left + lMinStep
            For iPos = 1 To lSize + 1 Step lMinStep
                lRed = lSRed + ((lERed - lSRed) * iPos) \ lSize
                lGreen = lSGreen + ((lEGreen - lSGreen) * iPos) \ lSize
                lBlue = lSBlue + ((lEBlue - lSBlue) * iPos) \ lSize
                If (lGreen = lLastGreen) And (lRed = lLastRed) And (lBlue = lLastBlue) Then
                    tR.Right = tR.Right + lMinStep
                Else
                    hBr = CreateSolidBrush(RGB(lLastRed, lLastGreen, lLastBlue))
                    FillRect lHDC, tR, hBr
                    DeleteObject hBr
                    tR.Left = tR.Right
                    tR.Right = tR.Left + lMinStep
                    lLastRed = lRed
                    lLastGreen = lGreen
                    lLastBlue = lBlue
                End If
            Next
        End If
    End If

End Sub

Public Sub TileArea(ByVal hdc As Long, _
                    ByVal X As Long, _
                    ByVal Y As Long, _
                    ByVal Width As Long, _
                    ByVal Height As Long, _
                    ByVal lSrcDC As Long, _
                    ByVal lBitmapW As Long, _
                    ByVal lBitmapH As Long, _
                    ByVal lSrcOffsetX As Long, _
                    ByVal lSrcOffsetY As Long)

    Dim lSrcX           As Long
    Dim lSrcY           As Long
    Dim lSrcStartX      As Long
    Dim lSrcStartY      As Long
    Dim lSrcStartWidth  As Long
    Dim lSrcStartHeight As Long
    Dim lDstX           As Long
    Dim lDstY           As Long
    Dim lDstWidth       As Long
    Dim lDstHeight      As Long

    lSrcStartX = ((X + lSrcOffsetX) Mod lBitmapW)
    lSrcStartY = ((Y + lSrcOffsetY) Mod lBitmapH)
    lSrcStartWidth = (lBitmapW - lSrcStartX)
    lSrcStartHeight = (lBitmapH - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY

    lDstY = Y
    lDstHeight = lSrcStartHeight

    Do While lDstY < (Y + Height)
        If (lDstY + lDstHeight) > (Y + Height) Then
            lDstHeight = Y + Height - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = X
        lSrcX = lSrcStartX
        Do While lDstX < (X + Width)
            If (lDstX + lDstWidth) > (X + Width) Then
                lDstWidth = X + Width - lDstX
                If (lDstWidth = 0) Then
                    lDstWidth = 4
                End If
            End If
            'If (lDstWidth > Width) Then lDstWidth = Width
            'If (lDstHeight > Height) Then lDstHeight = Height
            BitBlt hdc, lDstX, lDstY, lDstWidth, lDstHeight, lSrcDC, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = lBitmapW
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = lBitmapH
    Loop

End Sub

Public Sub DrawImage(ByVal hIml As Long, _
                     ByVal iIndex As Long, _
                     ByVal hdc As Long, _
                     ByVal xPixels As Integer, _
                     ByVal yPixels As Integer, _
                     ByVal lIconSizeX As Long, ByVal lIconSizeY As Long, _
                     Optional ByVal bSelected = False, _
                     Optional ByVal bCut = False, _
                     Optional ByVal bDisabled = False, _
                     Optional ByVal oCutDitherColour As OLE_COLOR = vbWindowBackground, _
                     Optional ByVal hExternalIml As Long = 0)

    Dim hIcon     As Long
    Dim lFlags    As Long
    Dim lhIml     As Long
    Dim lColor    As Long
    Dim iImgIndex As Long

    ' Draw the image at 1 based index or key supplied in vKey.
    ' on the hDC at xPixels,yPixels with the supplied options.
    ' You can even draw an ImageList from another ImageList control
    ' if you supply the handle to hExternalIml with this function.

    iImgIndex = iIndex
    If (iImgIndex > -1) Then
        If (hExternalIml <> 0) Then
            lhIml = hExternalIml
        Else
            lhIml = hIml
        End If

        lFlags = ILD_TRANSPARENT
        If (bSelected) Or (bCut) Then
            lFlags = lFlags Or ILD_SELECTED
        End If

        If (bCut) Then
            ' Draw dithered:
            lColor = TranslateColor(oCutDitherColour)
            If (lColor = -1) Then lColor = TranslateColor(vbWindowBackground)
            ImageList_DrawEx lhIml, _
                             iImgIndex, _
                             hdc, _
                             xPixels, yPixels, 0, 0, _
                             CLR_NONE, lColor, _
                             lFlags
        ElseIf (bDisabled) Then
            ' extract a copy of the icon:
            hIcon = ImageList_GetIcon(hIml, iImgIndex, 0)
            ' Draw it disabled at x,y:
            DrawState hdc, 0, 0, hIcon, 0, xPixels, yPixels, lIconSizeX, lIconSizeY, DST_ICON Or DSS_DISABLED
            ' Clear up the icon:
            DestroyIcon hIcon

        Else
            ' Standard draw:
            ImageList_Draw lhIml, _
                           iImgIndex, _
                           hdc, _
                           xPixels, _
                           yPixels, _
                           lFlags
        End If
    End If

End Sub
