VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置透明"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3570
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1140
      TabIndex        =   8
      Text            =   "171781"
      Top             =   120
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "半透明"
      Height          =   255
      Left            =   1980
      TabIndex        =   6
      Top             =   2580
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "透明颜色"
      Height          =   255
      Left            =   540
      TabIndex        =   5
      Top             =   2580
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      Caption         =   "设置透明颜色："
      Height          =   915
      Left            =   60
      TabIndex        =   1
      Top             =   1560
      Width           =   3435
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   180
         ScaleHeight     =   435
         ScaleWidth      =   2955
         TabIndex        =   3
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置透明度："
      Height          =   915
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   3435
      Begin 程序辅助器.SliderBar SliderBar1 
         Height          =   300
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   529
         Value           =   255
         MyMax           =   255
         LargeChang      =   5
         MyStyle         =   8
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "225/225"
         Height          =   180
         Left            =   2580
         TabIndex        =   4
         Top             =   600
         Width           =   630
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "句柄:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   7
      Top             =   180
      Width           =   600
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type ChooseColor
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Declare Function ChooseColorAPI Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long
Dim RES As Long

Private Sub Check1_Click()
  a
End Sub

Private Sub Check2_Click()
  a
End Sub

Private Sub Picture1_Click()
On Error Resume Next
  Dim MyColor As ChooseColor
  MyColor.lStructSize = Len(MyColor)
  MyColor.hInstance = App.hInstance
  MyColor.hwndOwner = Me.hwnd
  MyColor.flags = 0
  MyColor.lpCustColors = String$(16 * 4, 125)
  ChooseColorAPI MyColor

  Picture1.BackColor = MyColor.rgbResult
  a
End Sub

Private Sub SliderBar1_Change()
On Error Resume Next
  Label1 = SliderBar1 & "/225"
  a
End Sub

Sub a()
  If Check1.Value = 1 And Check2.Value = 1 Then
    RES = 1 Or 2
  ElseIf Check1.Value = 1 And Check2.Value = 0 Then
    RES = 1
  ElseIf Check1.Value = 0 And Check2.Value = 1 Then
    RES = 2
  ElseIf Check1.Value = 0 And Check2.Value = 0 Then
    RES = 0
  End If
  Call SetWindowLong(CLng(Trim(Text1)), -20, GetWindowLong(CLng(Trim(Text1)), -20) Or &H80000)
  Call SetLayeredWindowAttributes(CLng(Trim(Text1)), Picture1.BackColor, SliderBar1.Value, RES)
End Sub
