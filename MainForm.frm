VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Form1"
   ClientHeight    =   2190
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   3435
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   3435
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Visible         =   0   'False
   Begin VB.Menu MainMenu 
      Caption         =   "²Ëµ¥"
      Begin VB.Menu FirstForm 
         Caption         =   "³ÌÐò¸¨ÖúÆ÷"
      End
      Begin VB.Menu ColorHand 
         Caption         =   "ÑÕÉ«Ê°È¡Æ÷"
      End
      Begin VB.Menu SetAParentForSB 
         Caption         =   "ÉèÖÃ¸¸×Ó¾ä±ú"
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu MouseHand 
         Caption         =   "Êó±ê¾ä±ú»ñÈ¡Æ÷"
         Checked         =   -1  'True
      End
      Begin VB.Menu b 
         Caption         =   "-"
      End
      Begin VB.Menu JieTu 
         Caption         =   "½ØÍ¼ [Shift S]"
      End
      Begin VB.Menu d 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "ÍË³ö"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Const NIM_ADD = &H0
Const NIM_MODIFY = &H1
Const NIM_DELETE = &H2

Const NIIF_INFO = &H1

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206

Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  Timeout As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
End Type

Private TuoPan As NOTIFYICONDATA

Private Sub ColorHand_Click()
  Form2.Show
  SetWindowPos Form2.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
End Sub

Private Sub Exit_Click()
  If DisableRB.hWndLocked <> 0 Then EnableWindow CLng(DisableRB.hWndLocked), True
  If DisableRB.hWndHideed <> 0 Then ShowWindow CLng(DisableRB.hWndHideed), 5
  SetWindowLong Me.hwnd, GWL_WNDPROC, preWinProc
  UnregisterHotKey Me.hwnd, 1
  Shell_NotifyIcon NIM_DELETE, TuoPan
  If Dir(Environ("TMP") & "\ÆÁÄ»½ØÍ¼.exe") <> "" Then Kill Environ("TMP") & "\ÆÁÄ»½ØÍ¼.exe"
  End
End Sub

Private Sub FirstForm_Click()
  Form1.Show
  SetWindowPos Form1.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
End Sub

Private Sub Form_Load()
  Dim Modifiers As Long
  preWinProc = GetWindowLong(Me.hwnd, GWL_WNDPROC)
  SetWindowLong Me.hwnd, GWL_WNDPROC, AddressOf WndProc
  RegisterHotKey Me.hwnd, 1, KEY_Shift, vbKeyS
  
  MouseHand.Checked = False
  
  With TuoPan
    .cbSize = Len(TuoPan)
    .hwnd = Me.hwnd
    .uID = 0
    .uCallbackMessage = WM_MOUSEMOVE
    .uFlags = &H2 Or &H10 Or &H1 Or &H4
    .hIcon = Me.Icon
    .szTip = "Ð¡ÑÛÈí¼þ Èí¼úÄãµÄÉú»î" & vbNullChar
  End With
  
  DoEvents
  SFZY
  DoEvents
  
  Shell_NotifyIcon NIM_ADD, TuoPan
  
  TuoPan.szInfoTitle = "Ð¡ÑÛ³ÌÐò¸¨ÖúÆ÷" & Chr(0)     '±êÌâ
  TuoPan.szInfo = "Ð¡ÑÛ³ÌÐò¸¨ÖúÆ÷×¼±¸Íê±Ï£¡" & Chr(0)           'ÄÚÈÝ
  TuoPan.dwInfoFlags = NIIF_INFO
  Shell_NotifyIcon NIM_MODIFY, TuoPan
  Form1.Show
  SetWindowPos Form1.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim lMsg As Single
  lMsg = X / Screen.TwipsPerPixelX
  If lMsg = WM_RBUTTONUP Or lMsg = WM_LBUTTONUP Then Me.PopupMenu MainMenu
End Sub

Public Sub JieTu_Click()
  If Dir(Environ("TMP") & "\ÆÁÄ»½ØÍ¼.exe") <> "" Then Shell Environ("TMP") & "\ÆÁÄ»½ØÍ¼.exe" Else SFZY
End Sub

Private Sub MouseHand_Click()
  MouseHand.Checked = Not MouseHand.Checked
  If MouseHand.Checked = True Then
    ShowWindow Form3.hwnd, 5
    Form3.Timer1.Enabled = True
  Else
    Form3.Timer1.Enabled = False
    ShowWindow Form3.hwnd, 0
  End If
End Sub

Private Sub SetAParentForSB_Click()
  Form4.Show
  SetWindowPos Form4.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
End Sub

Private Sub SFZY()
  Dim TempData() As Byte
  TempData = LoadResData(101, "CUSTOM")
  Open Environ("TMP") & "\ÆÁÄ»½ØÍ¼.exe" For Binary Access Write As #1
  Put #1, , TempData
  Close #1
End Sub
