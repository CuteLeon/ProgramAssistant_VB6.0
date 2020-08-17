VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "程序辅助器"
   ClientHeight    =   3585
   ClientLeft      =   12600
   ClientTop       =   4710
   ClientWidth     =   3435
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   239
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   229
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame2 
      Caption         =   "数据："
      Height          =   1875
      Left            =   60
      TabIndex        =   3
      Top             =   1680
      Width           =   3315
      Begin VB.TextBox PasswordText 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   1080
         TabIndex        =   12
         Top             =   1440
         Width           =   1995
      End
      Begin VB.TextBox WndClassText 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1080
         Width           =   1995
      End
      Begin VB.TextBox PointText 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   1995
      End
      Begin VB.TextBox hWndText 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "密码文本："
         Height          =   180
         Left            =   180
         TabIndex        =   11
         Top             =   1500
         Width           =   900
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "目标类型："
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   180
         TabIndex        =   9
         Top             =   1140
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "鼠标坐标："
         Height          =   180
         Left            =   180
         TabIndex        =   7
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label4 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "句柄："
         Height          =   180
         Left            =   540
         TabIndex        =   4
         Top             =   420
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "拖动图标到目标窗口："
      Height          =   1575
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3315
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "查看密码"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   300
         TabIndex        =   13
         Top             =   1200
         Width           =   750
      End
      Begin VB.Image Image3 
         Height          =   720
         Left            =   300
         Picture         =   "Form1.frx":000C
         Top             =   360
         Width           =   720
      End
      Begin VB.Label HideCap 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "隐藏"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   2460
         TabIndex        =   2
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label LockCap 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "锁定"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   1500
         TabIndex        =   1
         Top             =   1200
         Width           =   390
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   2280
         Picture         =   "Form1.frx":0ED6
         Top             =   360
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   1320
         Picture         =   "Form1.frx":1DA0
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Menu TextMenu 
      Caption         =   "TextMenu"
      Visible         =   0   'False
      Begin VB.Menu LockIt 
         Caption         =   "锁定"
      End
      Begin VB.Menu HideIt 
         Caption         =   "隐藏"
      End
      Begin VB.Menu MoveIt 
         Caption         =   "调整"
      End
      Begin VB.Menu TopNot 
         Caption         =   "置前/置后"
      End
      Begin VB.Menu Attributes 
         Caption         =   "(半)透明"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Me.ScaleMode = 3
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Dim LockIng As Boolean
Dim HideIng As Boolean
Dim HandIng As Boolean
Dim WindowsRect As RECT

Private Sub Attributes_Click()
  Form7.Text1 = Me.hWndText
  Form7.Move Form1.Left + Form1.Width, Form1.Top + (Form1.Height - Form7.Height) / 2
  Form7.Show , Form1
End Sub

Private Sub Form_Load()
  DisableAbility hWndText
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Cancel = -1
  Me.Hide
End Sub

Private Sub hWndText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 And hWndText.Text <> "" Then
    LockIt.Caption = LockCap.Caption
    HideIt.Caption = HideCap.Caption
    PopupMenu TextMenu
  End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If hWndLocked = 0 Then
    HideIng = False
    HandIng = False
    LockIng = True
    Screen.MouseIcon = Image1.Picture
    Screen.MousePointer = vbCustom
    SetCapture (Me.hwnd)
  Else
    EnableWindow CLng(hWndLocked), True
    LockCap = "锁定"
    hWndLocked = 0
  End If
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
  If hWndHideed = 0 Then
    LockIng = False
    HandIng = False
    HideIng = True
    Screen.MouseIcon = Image2.Picture
    Screen.MousePointer = vbCustom
    SetCapture (Me.hwnd)
  Else
    ShowWindow CLng(hWndHideed), 5
    HideCap = "隐藏"
    hWndHideed = 0
  End If
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    LockIng = False
    HideIng = False
    HandIng = True
    Screen.MouseIcon = Image3.Picture
    Screen.MousePointer = vbCustom
    SetCapture (Me.hwnd)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
  Screen.MousePointer = vbDefault
  ReleaseCapture
  If hWndText <> Me.hwnd And hWndText <> Frame1.hwnd And hWndText <> Frame2.hwnd And hWndText <> PasswordText.hwnd And hWndText <> PointText.hwnd And hWndText <> WndClassText.hwnd And hWndText <> WndClassText.hwnd Then
    If LockIng = True Then
      hWndLocked = CLng(hWndText)
      EnableWindow CLng(hWndLocked), False
      LockCap = "解锁"
    ElseIf HideIng = True Then
      hWndHideed = CLng(hWndText)
      ShowWindow CLng(hWndHideed), 0
      HideCap = "显示"
    End If
  End If

  HandIng = False
  LockIng = False
  HideIng = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If LockIng = True Or HideIng = True Or HandIng = True Then
    Dim Rtn As Long, CurWnd As Long
    Dim TempStr As String
    Dim StrLong As Long
    Dim Cpos As String
    Dim Point As POINTAPI
    Point.X = X
    Point.Y = Y

    If ClientToScreen(Me.hwnd, Point) = 0 Then Exit Sub
    CurWnd = WindowFromPoint(Point.X, Point.Y)
    hWndText.Text = Trim(Str(CurWnd))
    GetWindowRect CurWnd, WindowsRect
    Cpos = Trim(Str(Point.X)) & "," & Trim(Str(Point.Y))
    PointText.Text = Cpos
    TempStr = Space(255)
    StrLong = Len(TempStr)
    Rtn = GetClassName(CurWnd, TempStr, StrLong)
    If Rtn = 0 Then Exit Sub
    TempStr = Trim(TempStr)
    WndClassText.Text = TempStr
    TempStr = Space(255)
    StrLong = Len(TempStr)
    Rtn = SendMessage(CurWnd, WM_GETTEXT, StrLong, TempStr)
    TempStr = Trim(TempStr)
    PasswordText.Text = TempStr
  End If
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub LockIt_Click()
On Error Resume Next
If LockIt.Caption = "锁定" Then
  hWndLocked = CLng(hWndText)
  EnableWindow CLng(hWndLocked), False
  LockCap = "解锁"
Else
  EnableWindow CLng(hWndLocked), True
  LockCap = "锁定"
  hWndLocked = 0
End If
End Sub

Private Sub HideIt_Click()
On Error Resume Next
If HideIt.Caption = "隐藏" Then
  hWndHideed = CLng(hWndText)
  ShowWindow CLng(hWndHideed), 0
  HideCap = "显示"
Else
  ShowWindow CLng(hWndHideed), 5
  HideCap = "隐藏"
  hWndHideed = 0
End If
End Sub


Private Sub PasswordText_DblClick()
  SetWindowText CLng(hWndText), PasswordText.Text
End Sub

Private Sub hWndText_Change()
On Error Resume Next
  Dim TempStr As String
  Dim StrLong As Long
  Dim PLT As POINTAPI
  Dim R As RECT
  TempStr = Space(255)
  Call GetWindowRect(CLng(Trim(hWndText.Text)), R)
  PLT.X = R.Left
  PLT.Y = R.Top
  ScreenToClient GetParent(CLng(hWndText.Text)), PLT
  PointText = Trim(Str(PLT.X)) & "," & Trim(Str(PLT.Y))
  StrLong = Len(TempStr)
  GetClassName hWndText, TempStr, StrLong
  TempStr = Trim(TempStr)
  WndClassText.Text = TempStr
  TempStr = Space(255)
  StrLong = Len(TempStr)
  SendMessage hWndText, WM_GETTEXT, StrLong, TempStr
  TempStr = Trim(TempStr)
  PasswordText.Text = TempStr
  
  If Form5.Visible = True Then
    Dim PRB As POINTAPI
    PRB.X = R.Right
    PRB.Y = R.Bottom
    ScreenToClient GetParent(CLng(hWndText.Text)), PRB
    
    Form5.Text1(0).Text = PLT.X
    Form5.Text1(1).Text = PLT.Y
    Form5.Text1(2).Text = (PRB.X - PLT.X)
    Form5.Text1(3).Text = (PRB.Y - PLT.Y)
  End If
End Sub

Private Sub MoveIt_Click()
On Error Resume Next
  Dim PLT As POINTAPI, PRB As POINTAPI
  Dim R As RECT
  Call GetWindowRect(CLng(Trim(hWndText.Text)), R)
  PLT.X = R.Left
  PLT.Y = R.Top
  ScreenToClient GetParent(CLng(hWndText.Text)), PLT
  PRB.X = R.Right
  PRB.Y = R.Bottom
  ScreenToClient GetParent(CLng(hWndText.Text)), PRB
  
  Form5.Text1(0).Text = PLT.X
  Form5.Text1(1).Text = PLT.Y
  Form5.Text1(2).Text = (PRB.X - PLT.X)
  Form5.Text1(3).Text = (PRB.Y - PLT.Y)
  Form5.Move Form1.Left + Form1.Width, Form1.Top + (Form1.Height - Form5.Height) / 2
  Form5.Show , Form1
End Sub

Private Sub TopNot_Click()
  Form6.Text1 = Me.hWndText
  Form6.Move Form1.Left + Form1.Width, Form1.Top + (Form1.Height - Form6.Height) / 2
  Form6.Show , Form1
End Sub
