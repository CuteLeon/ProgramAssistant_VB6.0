VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置父子句柄"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3615
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   189
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   241
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame2 
      Caption         =   "数据："
      Height          =   1035
      Left            =   60
      TabIndex        =   3
      Top             =   1560
      Width           =   3495
      Begin 程序辅助器.XYQQButton Command2 
         Height          =   315
         Left            =   2700
         TabIndex        =   11
         Top             =   600
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         Caption         =   "桌面"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin 程序辅助器.XYQQButton Command1 
         Height          =   315
         Left            =   2700
         TabIndex        =   10
         Top             =   240
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         Caption         =   "关联"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   1
         Left            =   1020
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   0
         Left            =   1020
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "父句柄："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   660
         Width           =   840
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "子句柄："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   300
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "拖动图标到目标窗口或控件："
      Height          =   1455
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3495
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "取消关系"
         ForeColor       =   &H000000FF&
         Height          =   210
         Index           =   4
         Left            =   2400
         TabIndex        =   8
         Top             =   1080
         Width           =   750
      End
      Begin VB.Image Image3 
         Height          =   720
         Left            =   2400
         Picture         =   "Form4.frx":000C
         Top             =   240
         Width           =   720
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   1260
         Picture         =   "Form4.frx":0ED6
         Top             =   240
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   300
         Picture         =   "Form4.frx":1DA0
         Top             =   420
         Width           =   480
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "父句柄"
         ForeColor       =   &H000000FF&
         Height          =   210
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "子句柄"
         ForeColor       =   &H000000FF&
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   570
      End
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 打开主窗口 "
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   1260
      TabIndex        =   9
      Top             =   2625
      Width           =   1110
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Me.ScaleMode = 3

Option Explicit
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Dim Cing As Boolean
Dim Ping As Boolean
Dim Qing As Boolean

Private Sub Command1_Click()
On Error Resume Next
  SetParent CLng(Trim(Text1(0))), CLng(Trim(Text1(1)))
End Sub

Private Sub Command2_Click()
On Error Resume Next
  SetParent CLng(Trim(Text1(0))), GetDesktopWindow
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Screen.MousePointer = vbDefault
  ReleaseCapture
  If Cing = True Then
    Cing = False
  ElseIf Ping = True Then
    Ping = False
  ElseIf Qing = True Then
    Qing = False
    SetParent CLng(Trim(Text1(0))), GetDesktopWindow
  End If
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Cancel = -1
  Me.Hide
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Ping = False
  Qing = False
  Cing = True
  Screen.MouseIcon = Image1.Picture
  Screen.MousePointer = vbCustom
  SetCapture (Me.hwnd)
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Cing = False
  Qing = False
  Ping = True
  Screen.MouseIcon = Image2.Picture
  Screen.MousePointer = vbCustom
  SetCapture (Me.hwnd)
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Ping = False
  Cing = False
  Qing = True
  Screen.MouseIcon = Image3.Picture
  Screen.MousePointer = vbCustom
  SetCapture (Me.hwnd)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim CurWnd As Long
  Dim Point As POINTAPI
  Point.X = X
  Point.Y = Y

  If ClientToScreen(Me.hwnd, Point) = 0 Then Exit Sub
  CurWnd = WindowFromPoint(Point.X, Point.Y)
  If Cing = True Then
    Text1(0).Text = Trim(Str(CurWnd))
  ElseIf Ping = True Then
    Text1(1).Text = Trim(Str(CurWnd))
  ElseIf Qing = True Then
    Text1(1).Text = Trim(Str(CurWnd))
  End If
End Sub

Private Sub Label2_Click()
  Form1.Show
End Sub
