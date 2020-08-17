Attribute VB_Name = "MyHOTKEY"
Option Explicit
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long) As Long
Public Const WM_HOTKEY = &H312
Public Const GWL_WNDPROC = (-4)
Public Const KEY_Shift = &H4
Public preWinProc As Long, MyhWnd As Long

Public Function WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If Msg = WM_HOTKEY Then
    MainForm.JieTu_Click
  End If
  
  WndProc = CallWindowProc(preWinProc, hwnd, Msg, wParam, lParam)
End Function
