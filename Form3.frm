VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3510
   DrawMode        =   1  'Blackness
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1440
      Top             =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1010100"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      FillColor       =   &H00FF0000&
      Height          =   1455
      Left            =   60
      Top             =   60
      Visible         =   0   'False
      Width           =   3375
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'获得指定句柄的位置
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long   'left + width
        Bottom As Long  'top + height
End Type
'鼠标下的句柄
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
'鼠标穿透
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Const WS_EX_TRANSPARENT As Long = &H20&
'窗体置前
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

'窗体透明
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long

'鼠标位置
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Dim MousePos As POINTAPI

Private Sub Form_Initialize()
  Me.Move 0, 0, Screen.Width, Screen.Height                            '窗体大小与屏幕同步
End Sub

Private Sub Form_Load()
  SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED Or WS_EX_TRANSPARENT  '鼠标穿透
  Call SetWindowLong(Me.hwnd, -20, GetWindowLong(Me.hwnd, -20) Or &H80000) '窗体透明
  Call SetLayeredWindowAttributes(Me.hwnd, &HE0E0E0, 0, 1)
End Sub

Private Sub Timer1_Timer()
  SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1     '窗体置前
  GetCursorPos MousePos                                                '获取鼠标位置
  Dim R As RECT
  Call GetWindowRect(WindowFromPoint(MousePos.X, MousePos.Y), R)       '得到鼠标位置句柄
  Shape1.Visible = False
  Label1.Visible = False
  Shape1.Left = R.Left * 15
  Shape1.Width = (R.Right - R.Left) * 15
  Shape1.Top = R.Top * 15
  Shape1.Height = (R.Bottom - R.Top) * 15
  Label1.Move Shape1.Left, Shape1.Top
  Label1.Caption = Str(WindowFromPoint(MousePos.X, MousePos.Y))
  Form1.Caption = Label1.Caption
  Shape1.Visible = True
  Label1.Visible = True
End Sub

