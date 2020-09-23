Attribute VB_Name = "Module1"
'===========================================================================
'
' Module Name:  x
' Author:       Graeme Grant
' Date:         xx/xx/2000
' Version:      00.01.00 Beta
' Description:  xx
' Edit History:
'
'===========================================================================

Option Private Module
Option Explicit

'##------------------------------------------------------------------
' For setting up a thin border on a picture box control:
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_SHOWWINDOW = &H40

Public Function ThinBorder(ByVal lHwnd As Long, ByVal bState As Boolean)

    Dim lS As Long

    lS = GetWindowLong(lHwnd, GWL_EXSTYLE)
    If Not (bState) Then
        lS = lS Or WS_EX_CLIENTEDGE And Not WS_EX_STATICEDGE
    Else
        lS = lS Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
    End If
    SetWindowLong lHwnd, GWL_EXSTYLE, lS
    SetWindowPos lHwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED

End Function

