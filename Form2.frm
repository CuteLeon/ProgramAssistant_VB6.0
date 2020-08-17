VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "颜色拾取器(拖动风车取色)"
   ClientHeight    =   3495
   ClientLeft      =   9180
   ClientTop       =   5310
   ClientWidth     =   5250
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5250
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "数据："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   3300
      TabIndex        =   1
      Top             =   60
      Width           =   1875
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   600
         TabIndex        =   10
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   600
         TabIndex        =   9
         Top             =   1740
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   600
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         Height          =   495
         Left            =   840
         ScaleHeight     =   435
         ScaleWidth      =   795
         TabIndex        =   7
         Top             =   2460
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   600
         TabIndex        =   6
         Top             =   1140
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   5
         Top             =   660
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   2100
         Width           =   210
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   1740
         Width           =   210
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "颜色:"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   180
         Picture         =   "Form2.frx":000C
         Top             =   2460
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Y："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   3
         Top             =   660
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   210
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   3195
      Left            =   60
      ScaleHeight     =   3135
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 打开主窗口 "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3540
      TabIndex        =   16
      Top             =   3180
      Width           =   1290
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "点击下图中某点锁定颜色"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   45
      TabIndex        =   15
      Top             =   15
      Width           =   3150
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Me.ScaleMode = 3

Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Const Srccopy = &HCC0020
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Dim Pos As POINTAPI
Dim ColorIng As Boolean
Dim Chosed As Boolean
Dim R, G, b As Integer

Private Sub Label6_Click()
  Form1.Show
End Sub

Private Sub Picture1_Click()
  GetColor
  Chosed = Not Chosed
  If Chosed = True Then
    Label5.Caption = "点击图像重新锁定颜色"
  Else
    Label5.Caption = "点击下图中某点锁定颜色"
  End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Chosed = False Then GetColor
End Sub

Private Sub FangDa()
  GetCursorPos Pos
  StretchBlt Picture1.hdc, 0, 0, Picture1.ScaleWidth / 5, Picture1.ScaleHeight / 5, GetDC(0), Pos.X - (Picture1.ScaleWidth / 90), Pos.Y - (Picture1.ScaleHeight / 90), Picture1.ScaleWidth / 15, Picture1.ScaleHeight / 15, Srccopy
  Text1 = Pos.X
  Text2 = Pos.Y
End Sub

Private Sub GetColor()
  Dim P1 As POINTAPI, h As Long, h1 As Long, r1 As Long
  GetCursorPos P1
  h1 = GetDC(h)
  r1 = GetPixel(h1, P1.X, P1.Y)
  If r1 = -1 Then
     BitBlt Picture2.hdc, 0, 0, 1, 1, h1, P1.X, P1.Y, vbSrcCopy
     r1 = Picture2.Point(0, 0)
   Else
     Picture2.PSet (0, 0), r1
  End If
  
  Picture2.BackColor = r1      '将颜色应用到Picture1
  R = Picture2.BackColor Mod 256          'Red
  G = Picture2.BackColor \ 256 Mod 256    'Green
  b = Picture2.BackColor \ 65536          'Blue
  Text3(0) = Picture2.BackColor
  Text3(1) = R
  Text3(2) = G
  Text3(3) = b
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Image1.Visible = False
  Label5.Caption = "点击下图中某点锁定颜色"
  ColorIng = True
  Chosed = False
  Screen.MouseIcon = Image1.Picture
  Screen.MousePointer = vbCustom
  SetCapture (Me.hwnd)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If ColorIng = True Then
    Screen.MousePointer = vbDefault
    Image1.Visible = True
    ColorIng = False
    ReleaseCapture
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If ColorIng = True Then
    FangDa
    GetColor
  End If
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Cancel = -1
  Me.Hide
End Sub
