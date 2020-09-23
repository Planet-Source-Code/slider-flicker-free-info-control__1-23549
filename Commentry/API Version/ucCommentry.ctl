VERSION 5.00
Begin VB.UserControl ucCommentry 
   BackColor       =   &H8000000C&
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1245
   ScaleHeight     =   705
   ScaleWidth      =   1245
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   720
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   210
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   210
      Width           =   375
   End
End
Attribute VB_Name = "ucCommentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const clSTEP     As Long = 20
Private Const csDCAPTION As String = "Commentry Control"
Private Const csDDESC    As String = "A Control that has a caption and detailed comment text with:-" + _
                                     vbCrLf + "  * Different font settings and colours for the Caption and Comments" + _
                                     vbCrLf + "  * Ellipses for the Caption if too long" + _
                                     vbCrLf + "  * Auto-Wordwrap with comments" + _
                                     vbCrLf + "  * Background bitmap or selectable colour"

Private mcMemDC          As cMemDC
Private mhDCSrc          As Long

Private mbBitmap         As Boolean
Private mlBitmapW        As Long
Private mlBitmapH        As Long

Private Type udtTextInfo
    Text      As String
    Colour    As OLE_COLOR
    Font      As New StdFont
    hFntDC    As Long
    hFntOldDC As Long
    Flags     As Long
End Type

Private mtCaption            As udtTextInfo
Private mtDesc               As udtTextInfo
Private mlCaptionShadowColor As OLE_COLOR

Private mlGap            As Long

Public Property Get BackColor() As OLE_COLOR
    BackColor = picContainer.BackColor
End Property

Public Property Let BackColor(ByVal NewColor As OLE_COLOR)
    picContainer.BackColor() = NewColor
    PropertyChanged "BackColor"
    pDraw
End Property

Public Property Get CommentryColour() As OLE_COLOR
    CommentryColour = mtDesc.Colour
End Property

Public Property Let CommentryColour(ByVal NewColor As OLE_COLOR)
    mtDesc.Colour = NewColor
    PropertyChanged "DescriptionColour"
    pDraw
End Property

Public Property Get CaptionColor() As OLE_COLOR
    CaptionColor = mtCaption.Colour
End Property

Public Property Let CaptionColor(ByVal New_CaptionColor As OLE_COLOR)
    mtCaption.Colour = New_CaptionColor
    PropertyChanged "CaptionColor"
    pDraw
End Property

Public Property Get CaptionShadowColor() As OLE_COLOR
    CaptionShadowColor = mlCaptionShadowColor
End Property

Public Property Let CaptionShadowColor(ByVal NewColor As OLE_COLOR)
    mlCaptionShadowColor = NewColor
    PropertyChanged "CaptionShadowColor"
    pDraw
End Property

Public Property Get BackgroundPicture() As StdPicture
    Set BackgroundPicture = picImage.Picture
End Property

Public Property Set BackgroundPicture(sPic As StdPicture)

    On Error Resume Next

    Set picImage.Picture = sPic
    picImage.Refresh
    If (Err.Number <> 0) Or (picImage.ScaleWidth = 0) Or (sPic Is Nothing) Then
        mhDCSrc = 0
        mbBitmap = False
    Else
        mbBitmap = True
        mhDCSrc = picImage.hdc
        mlBitmapW = picImage.ScaleWidth \ Screen.TwipsPerPixelX
        mlBitmapH = picImage.ScaleHeight \ Screen.TwipsPerPixelY
    End If
    pDraw
    PropertyChanged "BackgroundPicture"

End Property

Public Property Get Caption() As String
    Caption = mtCaption.Text
End Property

Public Property Let Caption(Text As String)
    mtCaption.Text = Text
    PropertyChanged "Caption"
    pDraw
End Property

Public Property Get Description() As String
    Description = mtDesc.Text
End Property

Public Property Let Description(Text As String)
    mtDesc.Text = Text
    PropertyChanged "Description"
    pDraw
End Property

Public Property Get CaptionFont() As StdFont
Attribute CaptionFont.VB_ProcData.VB_Invoke_Property = "StandardFont;Font"
Attribute CaptionFont.VB_UserMemId = -512
   Set CaptionFont = mtCaption.Font
End Property

Public Property Set CaptionFont(ByVal sFont As StdFont)
    Set mtCaption.Font = sFont
    PropertyChanged "CaptionFont"
    pDraw
End Property

Public Property Get DescFont() As StdFont
    Set DescFont = mtDesc.Font
End Property

Public Property Set DescFont(ByVal sFont As StdFont)
   Set mtDesc.Font = sFont
   PropertyChanged "DescriptionFont"
   pDraw
End Property

Public Property Get hwnd() As Long
   hwnd = UserControl.hwnd
End Property

Private Sub UserControl_Initialize()
    mlGap = 2 * Screen.TwipsPerPixelX
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set picContainer.Font = Ambient.Font
    mtCaption.Text = csDCAPTION
    mtDesc.Text = csDDESC
    mtCaption.Colour = vbInfoBackground
    mtDesc.Colour = vbHighlightText
    mlCaptionShadowColor = &H0&
    picContainer.BackColor = vbApplicationWorkspace

    Dim sFnt1 As New StdFont
    With sFnt1
        .Name = "Arial"
        .Size = 24
        .Bold = True
    End With
    Set mtCaption.Font = sFnt1

    Dim sFnt2 As New StdFont
    With sFnt2
        .Name = "Arial"
        .Size = 12
        .Bold = False
    End With
    Set mtDesc.Font = sFnt2

    pInitialise
End Sub

Public Sub Refresh()
    pDraw
End Sub

Private Sub UserControl_Paint()
    pDraw
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    pInitialise

    Dim sFnt1 As New StdFont
    With sFnt1
        .Name = "Arial"
        .Size = 24
        .Bold = True
    End With
    Set mtCaption.Font = PropBag.ReadProperty("CaptionFont", sFnt1)
    Dim sFnt2 As New StdFont
    With sFnt2
        .Name = "Arial"
        .Size = 12
        .Bold = False
    End With
    Set mtDesc.Font = PropBag.ReadProperty("DescriptionFont", sFnt2)

    picContainer.BackColor = PropBag.ReadProperty("BackColor", vbApplicationWorkspace)
    mtCaption.Colour = PropBag.ReadProperty("CaptionColor", vbInfoBackground)
    mlCaptionShadowColor = PropBag.ReadProperty("CaptionShadowColor", &H0&)
    mtDesc.Colour = PropBag.ReadProperty("DescriptionColour", vbHighlightText)
    Set BackgroundPicture = PropBag.ReadProperty("BackgroundPicture", Nothing)
    mtCaption.Text = PropBag.ReadProperty("Caption", csDCAPTION)
    mtDesc.Text = PropBag.ReadProperty("Description", csDDESC)

End Sub

Private Sub UserControl_Resize()
    LockControl picContainer, True
    With UserControl
        picContainer.Move .ScaleLeft, _
                          .ScaleTop, _
                          .ScaleWidth, _
                          .ScaleHeight
    End With
    pDraw
    LockControl picContainer, False
End Sub

Private Sub UserControl_Terminate()
    Set mcMemDC = Nothing
'    Debug.Print "UserControl_Terminate"
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Dim sFnt1 As New StdFont
    With sFnt1
        .Name = "Arial"
        .Size = 24
        .Bold = True
    End With
    PropBag.WriteProperty "CaptionFont", mtCaption.Font, sFnt1
    PropBag.WriteProperty "Caption", mtCaption.Text, csDCAPTION
    PropBag.WriteProperty "CaptionColor", mtCaption.Colour, vbInfoBackground
    PropBag.WriteProperty "CaptionShadowColor", mlCaptionShadowColor, &H0&
    Dim sFnt2 As New StdFont
    With sFnt2
        .Name = "Arial"
        .Size = 12
        .Bold = False
    End With
    PropBag.WriteProperty "DescriptionFont", mtDesc.Font, sFnt2
    PropBag.WriteProperty "Description", mtDesc.Text, csDDESC
    PropBag.WriteProperty "DescriptionColour", mtDesc.Colour, vbHighlightText
    PropBag.WriteProperty "BackColor", picContainer.BackColor, vbApplicationWorkspace
    PropBag.WriteProperty "BackgroundPicture", BackgroundPicture, Nothing

End Sub

Private Sub pDraw()

    Dim PicRect   As RECT
    Dim CaptRect  As RECT
    Dim DescRect  As RECT
    Dim lWidth    As Long
    Dim lHeight   As Long
    Dim sText     As String
    Dim eFlags    As ECGTextAlignFlags
    Dim lHDC      As Long
    Dim lhDCU     As Long
    Dim hFontOld  As Long
    Dim bMemDC    As Boolean
    Dim tLF       As LOGFONT

    '## 1. Preparation
    LockControl picContainer, True
    pPrepareMemDC lHDC, lhDCU, bMemDC
    GetClientRect picContainer.hwnd, PicRect

    '## 2. Set Background
    pFillBackground lHDC, PicRect, 0, 0

    '## 3. Print Caption
    Set Font = mtCaption.Font
    lHeight = TextHeight(mtCaption.Text) \ Screen.TwipsPerPixelY  '+ 4
    lWidth = TextWidth(mtCaption.Text) \ Screen.TwipsPerPixelX  '+ 4

    '## 3-1. Prepare Caption area
    CaptRect = PicRect
    CaptRect.Top = CaptRect.Top + 1
    CaptRect.Bottom = CaptRect.Top + lHeight
    CaptRect.Left = CaptRect.Left + 5
    CaptRect.Right = CaptRect.Right - CaptRect.Left

    '## 3-2. Prepare Caption Text
    sText = mtCaption.Text & vbNullChar
    eFlags = mtCaption.Flags

    '## 3-3. Set Caption Text Font & Colour
    If (mtCaption.hFntDC <> 0) Then
        If (mtCaption.hFntOldDC <> 0) Then
            If (lHDC <> 0) Then
                SelectObject lHDC, mtCaption.hFntOldDC
            End If
        End If
        DeleteObject mtCaption.hFntDC
    End If
    Set picContainer.Font = mtCaption.Font
    pOLEFontToLogFont mtCaption.Font, picContainer.hdc, tLF
    mtCaption.hFntDC = CreateFontIndirect(tLF)
    If (lHDC <> 0) Then
        mtCaption.hFntOldDC = SelectObject(lHDC, mtCaption.hFntDC)
    End If

    '## 3-4. Write Caption text to Memory DC
    '-- (a) Shadow
    SetTextColor lHDC, TranslateColor(mlCaptionShadowColor)
    DrawText lHDC, sText, -1, CaptRect, eFlags
    '-- (b) Text
    CaptRect.Left = CaptRect.Left - 1
    CaptRect.Top = CaptRect.Top - 1
    SetTextColor lHDC, TranslateColor(mtCaption.Colour)
    DrawText lHDC, sText, -1, CaptRect, eFlags

    '## 4. Print Description
    Set Font = mtCaption.Font
    lWidth = TextWidth(mtCaption.Text) \ Screen.TwipsPerPixelX  '+ 4

    '## 4-1. Prepare Caption area
    DescRect = PicRect
    DescRect.Top = CaptRect.Bottom - 3
    DescRect.Left = CaptRect.Left + 10
    DescRect.Right = DescRect.Right - DescRect.Left

    '## 4-2. Prepare Caption Text
    sText = mtDesc.Text & vbNullChar
    eFlags = mtDesc.Flags

    '## 4-3. Set Caption Text Font & Colour
    If (mtDesc.hFntDC <> 0) Then
        If (mtDesc.hFntOldDC <> 0) Then
            If (lHDC <> 0) Then
                SelectObject lHDC, mtDesc.hFntOldDC
            End If
        End If
        DeleteObject mtDesc.hFntDC
    End If
    Set UserControl.Font = mtDesc.Font
    pOLEFontToLogFont mtDesc.Font, UserControl.hdc, tLF
    mtDesc.hFntDC = CreateFontIndirect(tLF)
    If (lHDC <> 0) Then
        mtDesc.hFntOldDC = SelectObject(lHDC, mtDesc.hFntDC)
    End If

    '## 4-4. Write Caption text to Memory DC
    SetTextColor lHDC, TranslateColor(mtDesc.Colour)
    DrawText lHDC, sText, -1, DescRect, eFlags

    '## 5. Draw Frame
    ThinBorder picContainer.hwnd, True

    '## 6. Copy result to the Containter/Usercontrol
    pMemDCToDC lhDCU, lHDC, bMemDC, PicRect
    LockControl picContainer, False

End Sub

Private Sub pFillBackground(ByVal lHDC As Long, _
                            ByRef tR As RECT, _
                            ByVal lOffsetX As Long, _
                            ByVal lOffsetY As Long)

    Dim hBr As Long

    If (mbBitmap) Then
        TileArea lHDC, _
                 tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, _
                 mhDCSrc, _
                 mlBitmapW, mlBitmapH, _
                 lOffsetX, lOffsetY
    Else
        hBr = CreateSolidBrush(TranslateColor(picContainer.BackColor))
        FillRect lHDC, tR, hBr
        DeleteObject hBr
    End If

End Sub

Private Sub pInitialise()
    Set mcMemDC = New cMemDC
    'Set mcMemDC2 = New cMemDC
    mtCaption.Flags = DT_WORD_ELLIPSIS Or _
                      DT_PATH_ELLIPSIS Or _
                      DT_MODIFYSTRING Or _
                      DT_END_ELLIPSIS Or _
                      DT_SINGLELINE
    mtDesc.Flags = DT_TOP Or _
                   DT_LEFT Or _
                   DT_WORDBREAK Or _
                   DT_EDITCONTROL

End Sub

Private Sub pMemDCToDC(ByVal lhDCU As Long, ByVal lHDC As Long, ByVal bMemDC As Boolean, ByRef tR As RECT)
   If bMemDC Then
      With tR
          BitBlt lhDCU, .Left, .Top, .Right - .Left, .Bottom - .Top, lHDC, 0, 0, vbSrcCopy
      End With
   End If
End Sub

Private Sub pPrepareMemDC(ByRef lHDC As Long, ByRef lhDCU As Long, ByRef bMemDC As Boolean)
   
   lhDCU = picContainer.hdc
   If Not mcMemDC Is Nothing Then
      mcMemDC.Width = picContainer.ScaleWidth \ Screen.TwipsPerPixelY
      mcMemDC.Height = picContainer.ScaleHeight \ Screen.TwipsPerPixelX
      lHDC = mcMemDC.hdc
   End If
   If lHDC = 0 Then
      lHDC = lhDCU
   Else
      bMemDC = True
   End If
   SetBkColor lHDC, TranslateColor(picContainer.BackColor)
   SetBkMode lHDC, TRANSPARENT
   SetTextColor lHDC, TranslateColor(picContainer.ForeColor)

End Sub

