VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MoveIt"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2775
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   315
      Left            =   60
      TabIndex        =   9
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "位置与大小："
      Height          =   1695
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2655
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   3
         Left            =   840
         TabIndex        =   8
         Top             =   1260
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   2
         Left            =   840
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   1
         Left            =   840
         TabIndex        =   4
         Top             =   660
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   0
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height："
         Height          =   180
         Index           =   3
         Left            =   60
         TabIndex        =   7
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width："
         Height          =   180
         Index           =   2
         Left            =   60
         TabIndex        =   5
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Top："
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   3
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Left："
         Height          =   180
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   420
         Width           =   720
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Sub Command1_Click()
On Error Resume Next
  If Text1(Index) <> "" Then
    MoveWindow CLng(Trim(Form1.hWndText.Text)), CLng(Text1(0)), CLng(Text1(1)), CLng(Text1(2)), CLng(Text1(3)), True
  End If
End Sub
