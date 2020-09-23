VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMemDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================
' Name:     cMemDC.cls
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     20 October 1999
'
' Requires: SSUBTMR.DLL
'
' Copyright © 1999 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
'    http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' Memory DC for flicker free drawing.
'
' FREE SOURCE CODE - ENJOY!
' Do not sell this code.  Credit vbAccelerator.
' ======================================================================================


Private m_hDC As Long
Private m_hBmpOld As Long
Private m_hBmp As Long
Private m_lWidth As Long
Private m_lheight As Long

Public Property Get hdc() As Long
   hdc = m_hDC
End Property
Public Property Let Width(ByVal lW As Long)
   If lW > m_lWidth Then
      pCreate lW, m_lheight
   End If
End Property
Public Property Let Height(ByVal lH As Long)
   If lH > m_lheight Then
      pCreate m_lWidth, lH
   End If
End Property
Private Sub pCreate(ByVal lW As Long, ByVal lH As Long)
Dim lHDC As Long
   pDestroy
   lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   m_hDC = CreateCompatibleDC(lHDC)
   m_hBmp = CreateCompatibleBitmap(lHDC, lW, lH)
   m_hBmpOld = SelectObject(m_hDC, m_hBmp)
   If m_hBmpOld = 0 Then
      pDestroy
   Else
      m_lWidth = lW
      m_lheight = lH
   End If
   DeleteDC lHDC
End Sub
Private Sub pDestroy()
   If Not m_hBmpOld = 0 Then
      SelectObject m_hDC, m_hBmpOld
      m_hBmpOld = 0
   End If
   If Not m_hBmp = 0 Then
      DeleteObject m_hBmp
      m_hBmp = 0
   End If
   m_lWidth = 0
   m_lheight = 0
   If Not m_hDC = 0 Then
      DeleteDC m_hDC
      m_hDC = 0
   End If
End Sub

Private Sub Class_Terminate()
   pDestroy
End Sub

