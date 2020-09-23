Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
        ByVal lParam As Long) As Long
                   
Private Const WM_SETREDRAW = &HB
Public Function LockControl(objX As Object, cLock As Boolean)

   Dim i As Long
   
   If cLock Then
      ' Disable the Redraw flag for the specified window
      Call SendMessage(objX.hwnd, WM_SETREDRAW, False, 0)
   Else
      ' Enable the Redraw flag for the specified window, and repaint
      Call SendMessage(objX.hwnd, WM_SETREDRAW, True, 0)
      objX.Refresh
   End If

End Function
